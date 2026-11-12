# -*- coding: utf-8 -*-
"""
KIỂM TRA ỔN ĐỊNH TRƯỢT MÁI ĐẬP – PHƯƠNG PHÁP BISHOP ĐƠN GIẢN (mục 5.3 của đồ án)
  K = Σ{[c'·b + (W − u·b)·tgφ'] / mα} / Σ(W·sinα + Mngoài/R) ,  mα = cosα·(1 + tgα·tgφ'/K)
- Mặt trượt trụ tròn; tìm tâm trượt nguy hiểm bằng lưới tâm × bán kính (như SLOPE/W "Grid & Radius").
- Áp lực nước lỗ rỗng u lấy từ lời giải thấm phần tử hữu hạn (tham_fem.py): u = γw·(H − z) ≥ 0.
- Nước hồ trên mái thượng lưu: tải trọng áp lực nước vuông góc mặt mái (thành phần đứng cộng vào W,
  thành phần ngang tạo mô men) – cách xử lý giống "ponded water" trong SLOPE/W.
- Rút nước nhanh: áp lực lỗ rỗng trong thân đập giữ như trạng thái thấm ổn định ở mực nước cao
  (giả thiết đường bão hòa chưa kịp hạ – thiên an toàn), tải nước hồ lấy theo mực nước thấp.
- Động đất (giả tĩnh): lực quán tính ngang kh·W đặt tại trọng tâm dải, chiều gây trượt.
"""
import numpy as np
import matplotlib.tri as mtri
import so_lieu as S

GW = S.GAMMA_W


class LuoiRaster:
    """Rời rạc hóa miền thành lưới ô vuông để tra nhanh vật liệu và áp lực nước."""

    def __init__(self, xy, tri, mat, H, mc, dx=0.25, x0=None, x1=None, z0=None, z1=None, H_ap=None):
        self.mc = mc
        x0 = mc.x_trai if x0 is None else x0
        x1 = mc.x_phai if x1 is None else x1
        z0 = mc.Zbot if z0 is None else z0
        z1 = mc.Zc + 1.0 if z1 is None else z1
        self.xs = np.arange(x0, x1 + dx, dx); self.zs = np.arange(z0, z1 + dx, dx); self.dx = dx
        X, Z = np.meshgrid(self.xs, self.zs, indexing='ij')
        T = mtri.Triangulation(xy[:, 0], xy[:, 1], tri)
        f = T.get_trifinder()
        it = f(X, Z)
        self.mat = np.where(it >= 0, mat[np.maximum(it, 0)], 0)
        Hap = H if H_ap is None else H_ap
        interp = mtri.LinearTriInterpolator(T, Hap)
        Hg = interp(X, Z)
        Hg = np.ma.filled(Hg, np.nan)
        self.psi = Hg - Z            # cột nước áp lực (m)
        self.X, self.Z = X, Z
        self.ztop = mc.mat_tren(self.xs)

    def tra(self, x, z):
        i = np.clip(np.round((x - self.xs[0]) / self.dx).astype(int), 0, len(self.xs) - 1)
        j = np.clip(np.round((z - self.zs[0]) / self.dx).astype(int), 0, len(self.zs) - 1)
        return i, j


class OnDinh:
    def __init__(self, ras, vat_lieu, Zw=None, kh=0.0, chieu=+1, dung_cI=True):
        """vat_lieu: dict id -> (gamma_tn, gamma_bh, c, phi) [kN/m3, kPa, độ]
        chieu = +1: mái hạ lưu (trượt về phía +x); −1: mái thượng lưu (trượt về −x).
        Zw: mực nước hồ tác dụng lên mái thượng lưu (None nếu không có)."""
        self.r, self.vl, self.Zw, self.kh, self.d = ras, vat_lieu, Zw, kh, chieu
        n = max(vat_lieu) + 1
        self.g = np.zeros(n); self.gs = np.zeros(n); self.c = np.zeros(n); self.tphi = np.zeros(n)
        for k, (g, gs, c, phi) in vat_lieu.items():
            self.g[k], self.gs[k], self.c[k], self.tphi[k] = g, gs, c, np.tan(np.radians(phi))

    def K_vong(self, xc, zc, R, ns=40, tra_chi_tiet=False):
        r = self.r; mc = r.mc
        # giao điểm vòng tròn với mặt đất: lấy các điểm của lưới x nằm trong vòng tròn và dưới mặt đất
        xs = r.xs
        inside = np.abs(xs - xc) < R
        zb_all = np.full_like(xs, np.inf)
        zb_all[inside] = zc - np.sqrt(R ** 2 - (xs[inside] - xc) ** 2)
        below = zb_all < r.ztop - 1e-6
        if below.sum() < 4:
            return np.inf, None
        idx = np.where(below)[0]
        # chỉ lấy đoạn liên tục đầu tiên chứa phần sâu nhất
        splits = np.where(np.diff(idx) > 1)[0]
        if len(splits):
            segs = np.split(idx, splits + 1)
            depth = [np.max(r.ztop[s] - zb_all[s]) for s in segs]
            idx = segs[int(np.argmax(depth))]
        xa, xb = xs[idx[0]], xs[idx[-1]]
        if idx[0] == 0 or idx[-1] == len(xs) - 1:      # cung trượt ra ngoài miền mô hình -> loại
            return np.inf, None
        if xb - xa < 3.0 or np.max(r.ztop[idx] - zb_all[idx]) < 1.0:
            return np.inf, None
        edges = np.linspace(xa, xb, ns + 1)
        xm = 0.5 * (edges[1:] + edges[:-1]); b = np.diff(edges)
        zb = zc - np.sqrt(np.maximum(R ** 2 - (xm - xc) ** 2, 0))
        zt = mc.mat_tren(xm)
        h = zt - zb
        if np.any(h < -1e-6):
            h = np.maximum(h, 0)
        # trọng lượng dải: tích phân dọc cột với 30 điểm
        nz = 30
        tt = (np.arange(nz) + 0.5) / nz
        zz = zb[:, None] + tt[None, :] * h[:, None]
        XX = np.repeat(xm[:, None], nz, 1)
        ii, jj = r.tra(XX, zz)
        mm = r.mat[ii, jj]
        ps = r.psi[ii, jj]
        ps = np.nan_to_num(ps, nan=-1.0)
        gam = np.where(ps > 0, self.gs[mm], self.g[mm])
        gam = np.where(mm == 0, 0.0, gam)
        W = (gam * (h[:, None] / nz)).sum(1) * b
        # trọng tâm theo phương đứng (cho lực quán tính)
        zg = np.where(W > 0, (gam * zz * (h[:, None] / nz)).sum(1) * b / np.maximum(W, 1e-9), zb)
        # đáy dải: vật liệu và áp lực nước
        ib, jb = r.tra(xm, zb + 0.05)
        mb = r.mat[ib, jb]
        u = GW * np.maximum(np.nan_to_num(r.psi[ib, jb], nan=0.0), 0.0)
        c = self.c[mb]; tphi = self.tphi[mb]
        # nước hồ trên mặt mái (ngập)
        Wn = np.zeros_like(W); Mn = 0.0
        if self.Zw is not None:
            dep = np.maximum(self.Zw - zt, 0.0)
            Wn = GW * dep * b
            # áp lực ngang trên mặt nghiêng của dải: Fx = ∫p dz (đẩy vào thân đập), mô men quanh tâm
            z1 = mc.mat_tren(edges[:-1]); z2 = mc.mat_tren(edges[1:])
            p = GW * dep
            Fx = p * (z2 - z1) * (-1.0)   # mặt đất tăng theo x (mái TL): nước đẩy theo +x khi z2>z1 -> lực trên đất hướng +x
            # với mái thượng lưu z2 > z1: lực nước tác dụng lên đất hướng +x
            Fx = p * (z2 - z1)
            Mn = np.sum(-(zt - zc) * Fx)          # M = (x−xc)·Fz − (z−zc)·Fx ; ở đây chỉ phần Fx
        # mô men (ngược chiều kim đồng hồ dương); chiều gây trượt: d=+1 -> CCW, d=-1 -> CW
        MW = np.sum(-(xm - xc) * (W + Wn))         # trọng lượng (Fz = −W)
        Meq = np.sum(-(zg - zc) * (self.d * self.kh * W))   # lực quán tính ngang theo chiều trượt
        Mdrive = self.d * (MW + Mn) + self.d * Meq * 1.0
        if Mdrive <= 0:
            return np.inf, None
        # góc đáy dải theo quy ước Bishop (dương khi đáy dốc xuống theo chiều trượt)
        sina = self.d * (xc - xm) / R
        sina = np.clip(sina, -0.999, 0.999)
        cosa = np.sqrt(1 - sina ** 2)
        Wt = W + Wn
        N_eff = np.maximum(Wt - u * b, 0.0)    # thành phần pháp tuyến hiệu quả (không âm)
        K = 1.5
        for _ in range(50):
            ma = cosa * (1 + (sina / cosa) * tphi / K)
            ma = np.maximum(ma, 0.2)
            S_ = (c * b + N_eff * tphi) / ma
            Kn = R * np.sum(S_) / Mdrive
            if abs(Kn - K) < 1e-5:
                K = Kn; break
            K = Kn
        if tra_chi_tiet:
            return K, dict(xm=xm, b=b, zb=zb, zt=zt, W=W, Wn=Wn, u=u, c=c, phi=np.degrees(np.arctan(tphi)),
                           sina=sina, Mdrive=Mdrive, xa=xa, xb=xb)
        return K, (xa, xb)

    def tim(self, xc_rng, zc_rng, nx=25, nz=20, nr=25, R_min=None, x_gioi_han=None, tinh_chinh=True, z_day=None):
        best = (np.inf, None)
        xcs = np.linspace(*xc_rng, nx); zcs = np.linspace(*zc_rng, nz)
        mc = self.r.mc
        for xc in xcs:
            for zc in zcs:
                # bán kính: từ tiếp xúc độ sâu nông đến tiếp xúc đáy mô hình phần nền
                Rmax = zc - (mc.Zday - 8.0) if z_day is None else zc - z_day
                Rmin = max(zc - mc.mat_tren(xc) + 1.0, 3.0) if R_min is None else R_min
                if Rmax <= Rmin:
                    continue
                for R in np.linspace(Rmin, Rmax, nr):
                    K, ex = self.K_vong(xc, zc, R)
                    if ex is None:
                        continue
                    if x_gioi_han is not None and not (x_gioi_han[0] <= ex[0] and ex[1] <= x_gioi_han[1]):
                        continue
                    if K < best[0]:
                        best = (K, (xc, zc, R))
        if tinh_chinh and best[1] is not None:
            # tinh chỉnh quanh điểm tốt nhất
            xc0, zc0, R0 = best[1]
            dxc = (xc_rng[1] - xc_rng[0]) / nx; dzc = (zc_rng[1] - zc_rng[0]) / nz
            for xc in np.linspace(xc0 - dxc, xc0 + dxc, 9):
                for zc in np.linspace(zc0 - dzc, zc0 + dzc, 9):
                    for R in np.linspace(R0 - 2.0, R0 + 2.0, 17):
                        if z_day is not None and zc - R < z_day:
                            continue
                        K, ex = self.K_vong(xc, zc, R)
                        if ex is None:
                            continue
                        if x_gioi_han is not None and not (x_gioi_han[0] <= ex[0] and ex[1] <= x_gioi_han[1]):
                            continue
                        if K < best[0]:
                            best = (K, (xc, zc, R))
        return best


def vat_lieu_on_dinh(dung='I'):
    """Chỉ tiêu tính toán ổn định theo TTGH I (c_I, φ_I) nếu có, ngược lại dùng giá trị tiêu chuẩn."""
    D = S.DAT
    ma = {1: 'DD-I', 2: 'DC', 3: 'DD-II', 4: 'DC', 5: 'DA', 6: 'DD-I', 7: 'L4', 8: 'LI', 9: 'LII',
          11: 'DAP_MOI', 12: 'DA'}
    out = {}
    for k, key in ma.items():
        d = D[key]
        c, phi = (d['cI'], d['phiI']) if dung == 'I' else (d['c'], d['phi'])
        out[k] = (d['g'], d['gs'], c, phi)
    return out
