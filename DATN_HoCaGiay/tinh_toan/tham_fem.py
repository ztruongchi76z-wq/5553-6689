# -*- coding: utf-8 -*-
"""
TÍNH THẤM QUA ĐẬP ĐẤT BẰNG PHƯƠNG PHÁP PHẦN TỬ HỮU HẠN (mục 5.2 của đồ án)
- Phương trình thấm ổn định bão hòa – không bão hòa (Richards, dạng cột nước tổng H):
      ∂/∂x(k(ψ)·∂H/∂x) + ∂/∂z(k(ψ)·∂H/∂z) = 0 ,   ψ = H − z
- Hàm thấm tương đối k_r(ψ) = 1 khi ψ ≥ 0; k_r = max(e^{αψ}, k_min) khi ψ < 0
  (cách xử lý giống SEEP/W: không cần biết trước đường bão hòa, đường bão hòa là đường ψ = 0)
- Phần tử tam giác bậc 1, lưới tạo bằng thư viện Triangle (J. R. Shewchuk)
- Biên mặt thấm ra (seepage face) được xác định lặp: nút có ψ > 0 trên biên tiềm năng → H = z
"""
import numpy as np
import scipy.sparse as sp
import scipy.sparse.linalg as spla
import triangle as tr
import so_lieu as S

# ---------------------------------------------------------------------------
# HÌNH HỌC MẶT CẮT LÒNG SÔNG (cọc C16 – Km0+400, chiều cao đập lớn nhất)
# Trục x = 0 tại tim đập, dương về hạ lưu; cao độ z (m).
# Nguồn: bản vẽ trắc ngang đập (HCG-TNNC, 2023), mặt cắt đại diện quan trắc (TKCS Hình 2-10),
#        thông số mái [KTTV Bảng 1-2]; các kích thước không có trong hồ sơ ghi "giả thiết".
# ---------------------------------------------------------------------------
class MatCat:
    def __init__(self, Zc=78.35, Zday=54.0, B=5.0, m1u=3.0, Zcu=70.0, bcu=3.0, m2u=3.5,
                 m1d=3.0, Zcd=68.0, bcd=14.0, m2d=3.5, Z_ltr=59.0, b_ltr=3.0, m_ltr_ngoai=1.5,
                 m_ltr_trong=1.0, t_ong_khoi=1.0, t_dem=0.8, Z4=48.0, ZI=43.0, Zbot=35.0,
                 x_trai=-150.0, x_phai=160.0, chan_khay_den=48.0, b_ck=6.0, x_ck=-2.0,
                 man_kp=None):
        self.__dict__.update(locals()); del self.__dict__['self']
        # các điểm đặc trưng mặt đập
        self.xcd = B / 2                       # mép đỉnh hạ lưu
        self.xcu = -B / 2
        self.x_b1 = self.xcd + m1d * (Zc - Zcd)
        self.x_b2 = self.x_b1 + bcd
        self.x_l1 = self.x_b2 + m2d * (Zcd - Z_ltr)      # đỉnh lăng trụ phía trong
        self.x_l2 = self.x_l1 + b_ltr
        self.x_l3 = self.x_l2 + m_ltr_ngoai * (Z_ltr - Zday)   # chân lăng trụ
        self.x_l0 = self.x_l1 - m_ltr_trong * (Z_ltr - Zday)   # chân mái trong lăng trụ
        self.x_u1 = self.xcu - m1u * (Zc - Zcu)
        self.x_u2 = self.x_u1 - bcu
        self.x_u3 = self.x_u2 - m2u * (Zcu - Zday)             # chân mái thượng lưu
        # ống khói: mặt thượng lưu đi qua (1,0 ; Zc-1) dốc 1:1 xuống nền
        self.Zok = Zc - 1.0
        self.dx_ok = t_ong_khoi * np.sqrt(2)
        self.x_ok_u_bot = 1.0 + (self.Zok - Zday)
        self.x_ok_d_dem = 1.0 + self.dx_ok + (self.Zok - (Zday + t_dem))
        self.x_dem_end = self.x_l1 - m_ltr_trong * (Z_ltr - (Zday + t_dem))

    # mặt đất / mặt đập (đường bao trên) theo x
    def mat_tren(self, x):
        xs = [self.x_trai, self.x_u3, self.x_u2, self.x_u1, self.xcu, self.xcd, self.x_b1, self.x_b2,
              self.x_l1, self.x_l2, self.x_l3, self.x_phai]
        zs = [self.Zday, self.Zday, self.Zcu, self.Zcu, self.Zc, self.Zc, self.Zcd, self.Zcd,
              self.Z_ltr, self.Z_ltr, self.Zday, self.Zday]
        return np.interp(x, xs, zs)

    def duong_mat(self):
        xs = [self.x_trai, self.x_u3, self.x_u2, self.x_u1, self.xcu, self.xcd, self.x_b1, self.x_b2,
              self.x_l1, self.x_l2, self.x_l3, self.x_phai]
        zs = [self.Zday, self.Zday, self.Zcu, self.Zcu, self.Zc, self.Zc, self.Zcd, self.Zcd,
              self.Z_ltr, self.Z_ltr, self.Zday, self.Zday]
        return np.array(xs), np.array(zs)

    # ------------------------------------------------------------------
    def pslg(self):
        """Tạo đồ thị phẳng (đỉnh + đoạn) cho Triangle, kèm các điểm hạt giống vùng vật liệu."""
        V = []; idx = {}

        def v(x, z):
            key = (round(x, 4), round(z, 4))
            if key not in idx:
                idx[key] = len(V); V.append(key)
            return idx[key]

        segs = []

        def line(pts):
            ids = [v(*p) for p in pts]
            for a, b in zip(ids[:-1], ids[1:]):
                if a != b:
                    segs.append((a, b))
            return ids

        Zd, Z4, ZI, Zb = self.Zday, self.Z4, self.ZI, self.Zbot
        xL, xR = self.x_trai, self.x_phai
        # cấu trúc chân khay
        ck_bot = self.chan_khay_den
        hck = Zd - ck_bot
        xck1 = self.x_ck - self.b_ck / 2 - hck; xck2 = self.x_ck - self.b_ck / 2
        xck3 = self.x_ck + self.b_ck / 2; xck4 = self.x_ck + self.b_ck / 2 + hck
        # các điểm trên đường nền z = Zd (theo thứ tự x)
        nen_pts = sorted({xL, self.x_u3, xck1, xck4, self.x_ok_u_bot, self.x_l0, self.x_l3, xR})
        # biên ngoài
        line([(xL, Zb), (xR, Zb), (xR, ZI), (xR, Z4), (xR, Zd)])
        # mặt đất hạ lưu + lăng trụ + mái + đỉnh + mái TL + mặt đất TL
        line([(xR, Zd), (self.x_l3, Zd), (self.x_l2, self.Z_ltr), (self.x_l1, self.Z_ltr),
              (self.x_b2, self.Zcd), (self.x_b1, self.Zcd), (self.xcd, self.Zc), (0.0, self.Zc), (self.xcu, self.Zc),
              (self.x_u1, self.Zcu), (self.x_u2, self.Zcu), (self.x_u3, Zd), (xL, Zd),
              (xL, Z4), (xL, ZI), (xL, Zb)])
        # đường nền z = Zd dưới thân đập
        line([(x, Zd) for x in nen_pts if self.x_u3 <= x <= self.x_l3])
        # ranh giới lớp 4 / lớp I và lớp I / lớp II
        if ck_bot > Z4 + 1e-6:
            line([(xL, Z4), (xR, Z4)])
        else:
            line([(xL, Z4), (xck1 + (Zd - Z4), Z4)]); line([(xck4 - (Zd - Z4), Z4), (xR, Z4)])
        if ck_bot > ZI + 1e-6:
            line([(xL, ZI), (xR, ZI)])
        else:
            line([(xL, ZI), (xck2, ZI)]); line([(xck2, ZI), (xck3, ZI)]); line([(xck3, ZI), (xR, ZI)])
        # chân khay
        line([(xck1, Zd), (xck2, ck_bot), (xck3, ck_bot), (xck4, Zd)])
        if ck_bot < Z4 - 1e-6:   # chân khay cắt qua ranh giới z = Z4: thêm đoạn nối
            pass
        # ống khói + đệm tiêu nước
        Zok, dx = self.Zok, self.dx_ok
        line([(1.0, Zok), (1.0 + dx, Zok)])
        line([(1.0, Zok), (0.0, self.Zc)])        # ranh giới khối lõi / khối hạ lưu phía trên ống khói
        line([(1.0, Zok), (self.x_ok_u_bot, Zd)])
        line([(1.0 + dx, Zok), (self.x_ok_d_dem, Zd + self.t_dem)])
        line([(self.x_ok_d_dem, Zd + self.t_dem), (self.x_dem_end, Zd + self.t_dem)])
        # mái trong lăng trụ
        line([(self.x_l1, self.Z_ltr), (self.x_dem_end, Zd + self.t_dem), (self.x_l0, Zd)])
        # màng khoan phụt (chỉ dùng cho bài toán chuyên đề, None ở đây)
        regions = [
            (-20.0, Zd + 6.0, 1, 0.30),                      # ĐĐ-I thân đập TL + lõi
            (1.0 + dx / 2 + (Zok - 66.0), 66.0, 2, 0.15),    # ống khói
            (self.x_b1 + 5.0, Zd + 4.0, 3, 0.30),            # ĐĐ-II khối hạ lưu
            (0.5 * (self.x_ok_d_dem + self.x_dem_end), Zd + self.t_dem / 2, 4, 0.15),   # đệm
            (0.5 * (self.x_l1 + self.x_l2), 0.5 * (Zd + self.Z_ltr), 5, 0.40),         # lăng trụ
            (self.x_ck, 0.5 * (Zd + ck_bot), 6, 0.40),       # chân khay
            (xL + 5.0, 0.5 * (Zd + Z4), 7, 2.0), (xR - 5.0, 0.5 * (Zd + Z4), 7, 2.0),  # lớp 4
            (xL + 5.0, 0.5 * (Z4 + ZI), 8, 3.0), (xR - 5.0, 0.5 * (Z4 + ZI), 8, 3.0),  # lớp I
            (0.0, 0.5 * (ZI + Zb), 9, 6.0),                  # lớp II
        ]
        if ck_bot < Z4 - 1e-6:   # chân khay qua lớp 4 vào lớp I: hai phần lớp I ở hai bên
            pass
        return np.array(V, float), np.array(segs, int), regions


VAT_LIEU = {   # id: (tên, khóa trong S.DAT, k (m/s))
    1: ('ĐĐ-I', 'DD-I'), 2: ('Ống khói (cát lọc)', 'DC'), 3: ('ĐĐ-II', 'DD-II'),
    4: ('Đệm tiêu nước', 'DC'), 5: ('Lăng trụ đá', 'DA'), 6: ('Chân khay (ĐĐ-I)', 'DD-I'),
    7: ('Lớp 4', 'L4'), 8: ('Lớp I', 'LI'), 9: ('Lớp II', 'LII'), 10: ('Màng khoan phụt', None),
}


def k_cua(mid, kdat=None):
    if kdat and mid in kdat:
        return kdat[mid]
    ten, key = VAT_LIEU[mid]
    return S.DAT[key]['k']


def tao_luoi(mc, min_angle=28):
    V, segs, regions = mc.pslg()
    A = dict(vertices=V, segments=segs, regions=np.array([[r[0], r[1], r[2], r[3]] for r in regions]))
    T = tr.triangulate(A, 'pq%dAa' % min_angle)
    xy = T['vertices']; tri = T['triangles']; att = T['triangle_attributes'][:, 0].astype(int)
    return xy, tri, att


class BaiToanTham:
    def __init__(self, xy, tri, mat, kdat=None, alpha=1.0, kmin=1e-4, anis=1.0):
        self.xy, self.tri, self.mat = xy, tri, mat
        self.n = len(xy)
        self.ksat = np.array([k_cua(m, kdat) for m in mat])
        self.alpha, self.kmin, self.anis = alpha, kmin, anis
        x = xy[tri, 0]; z = xy[tri, 1]
        self.xc = x.mean(1); self.zc = z.mean(1)
        # gradient hàm dạng
        b = np.stack([z[:, 1] - z[:, 2], z[:, 2] - z[:, 0], z[:, 0] - z[:, 1]], 1)
        c = np.stack([x[:, 2] - x[:, 1], x[:, 0] - x[:, 2], x[:, 1] - x[:, 0]], 1)
        self.A2 = x[:, 1] * z[:, 2] - x[:, 2] * z[:, 1] - x[:, 0] * z[:, 2] + x[:, 2] * z[:, 0] + x[:, 0] * z[:, 1] - x[:, 1] * z[:, 0]
        self.area = 0.5 * np.abs(self.A2)
        self.bx = b / self.A2[:, None]; self.bz = c / self.A2[:, None]

    def kr(self, psi):
        return np.where(psi >= 0, 1.0, np.maximum(np.exp(self.alpha * np.minimum(psi, 0)), self.kmin))

    def lap_ma_tran(self, H):
        psi_e = H[self.tri].mean(1) - self.zc
        ke = self.ksat * self.kr(psi_e)
        Ke = (ke * self.area)[:, None, None] * (self.anis * self.bx[:, :, None] * self.bx[:, None, :]
                                                + self.bz[:, :, None] * self.bz[:, None, :])
        rows = np.repeat(self.tri, 3, axis=1).ravel()
        cols = np.tile(self.tri, (1, 3)).ravel()
        K = sp.coo_matrix((Ke.ravel(), (rows, cols)), shape=(self.n, self.n)).tocsr()
        return K, ke

    def giai(self, dir_nodes, dir_vals, seep_nodes, H0=None, tol=1e-4, itmax=400, relax=0.6, verbose=False):
        n = self.n
        z = self.xy[:, 1]
        H = np.full(n, np.mean(dir_vals)) if H0 is None else H0.copy()
        H[dir_nodes] = dir_vals
        seep_nodes = np.asarray(seep_nodes, int)
        seep_on = np.zeros(len(seep_nodes), bool)
        for it in range(itmax):
            K, ke = self.lap_ma_tran(H)
            fix = np.zeros(n, bool); val = np.zeros(n)
            fix[dir_nodes] = True; val[dir_nodes] = dir_vals
            sn = seep_nodes[seep_on]
            fix[sn] = True; val[sn] = z[sn]
            free = ~fix
            rhs = -K[free][:, fix] @ val[fix]
            Hn = H.copy(); Hn[fix] = val[fix]
            Hn[free] = spla.spsolve(K[free][:, free].tocsc(), rhs)
            # cập nhật biên mặt thấm ra
            changed = False
            psi_s = Hn[seep_nodes] - z[seep_nodes]
            # nút đang tự do mà áp lực dương -> thành nút mặt thấm
            new_on = seep_on | (psi_s > 1e-6)
            # nút đang cố định mà lưu lượng đi vào miền -> trả lại tự do
            if seep_on.any():
                R = K @ Hn            # phản lực nút R = K·H = lưu lượng ĐI VÀO miền tại nút cố định
                release = seep_on & (R[seep_nodes] > 1e-14)
                new_on = new_on & ~release
            if it < 40 and not np.array_equal(new_on, seep_on):
                changed = True
                seep_on = new_on
            dH = np.max(np.abs(Hn - H))
            w = relax if it < 30 else 0.3
            H = H + w * (Hn - H) if it > 2 else Hn
            if verbose:
                print('it', it, 'dH', dH, 'seep', seep_on.sum())
            if dH < tol and not changed:
                break
        self.H = H; self.seep_on = seep_nodes[seep_on]; self.niter = it
        K, ke = self.lap_ma_tran(H)
        self.ke = ke
        R = K @ H
        self.R = R
        return H

    def gradient(self, H=None):
        H = self.H if H is None else H
        gx = (self.bx * H[self.tri]).sum(1); gz = (self.bz * H[self.tri]).sum(1)
        return gx, gz, np.hypot(gx, gz)

    def luu_luong(self, nodes):
        """Tổng lưu lượng qua tập nút có điều kiện biên cố định (m3/s/m): tổng phản lực nút."""
        return float(np.sum(self.R[nodes]))


def bien(mc, xy, Zw, Zw_hl=None):
    """Xác định nút biên: nút thượng lưu ngập nước (H = Zw), nút mặt thấm tiềm năng phía hạ lưu."""
    x, z = xy[:, 0], xy[:, 1]
    zt = mc.mat_tren(x)
    on_top = np.abs(z - zt) < 1e-6
    # thượng lưu: các nút trên mặt đất/mái thượng lưu (x <= mép đỉnh TL) thấp hơn mực nước
    up = on_top & (x <= mc.xcu + 1e-6) & (z <= Zw + 1e-9)
    # mặt thấm tiềm năng: mặt đất/mái hạ lưu (x >= mép đỉnh hạ lưu)
    seep = on_top & (x >= mc.xcd - 1e-6)
    if Zw_hl is not None:
        dn = seep & (z <= Zw_hl)
        return np.where(up)[0], np.where(dn)[0], np.where(seep & ~dn)[0]
    return np.where(up)[0], np.array([], int), np.where(seep)[0]


def chay(mc, Zw, kdat=None, alpha=1.0, verbose=False, Zw_hl=None, luoi=None):
    xy, tri, mat = luoi if luoi is not None else tao_luoi(mc)
    bt = BaiToanTham(xy, tri, mat, kdat=kdat, alpha=alpha)
    up, dn, seep = bien(mc, xy, Zw, Zw_hl)
    dir_nodes = np.concatenate([up, dn]); dir_vals = np.concatenate([np.full(len(up), Zw), np.full(len(dn), Zw_hl if Zw_hl else 0)])
    H0 = np.interp(xy[:, 0], [mc.x_u3, mc.x_l3], [Zw, mc.Zday])
    bt.giai(dir_nodes, dir_vals, seep, H0=H0, verbose=verbose)
    q_in = bt.luu_luong(up)           # lưu lượng vào từ thượng lưu (m3/s/m)
    q_out = -bt.luu_luong(bt.seep_on)  # lưu lượng ra qua mặt thấm
    return bt, dict(q_in=q_in, q_out=q_out, up=up, seep=bt.seep_on)


if __name__ == '__main__':
    import time
    mc = MatCat()
    t = time.time()
    xy, tri, mat = tao_luoi(mc)
    print('nodes', len(xy), 'elems', len(tri), 'time', round(time.time() - t, 2))
    for m in sorted(set(mat)):
        print(m, VAT_LIEU[m][0], (mat == m).sum())
    t = time.time()
    bt, r = chay(mc, S.MNDBT, luoi=(xy, tri, mat), verbose=False)
    print('solve time', round(time.time() - t, 1), 'iters', bt.niter, 'q_in %.3e q_out %.3e' % (r['q_in'], r['q_out']))
