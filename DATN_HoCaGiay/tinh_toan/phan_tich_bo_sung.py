# -*- coding: utf-8 -*-
"""
PHÂN TÍCH BỔ SUNG (hoàn thiện thuyết minh theo góp ý, 10/10/2026)
  (1) Kiểm chứng chéo hệ số ổn định Bishop bằng phương pháp Spencer (GLE, λ = tgθ không đổi) – thỏa mãn
      đồng thời cân bằng lực và mô men (TCVN 8216:2018 điều 11.3.1, mục 4).
  (2) Độ hội tụ của lưới tâm trượt (lưới tâm và số bán kính tăng gấp đôi).
  (3) Kiểm tra ổn định thời kỳ thi công – sửa chữa (TCVN 8216:2018 Bảng 1: [K] = 1,24 đập cấp II; điều 11.3.2):
      hồ ở MNDBT, tải trọng thi công phân bố q trên toàn bề rộng đỉnh đập.
Mặt cắt sau sửa chữa với đỉnh +78,85 m (so_lieu.Z_DINH_DAP_SC).
"""
import json, time
import numpy as np
import so_lieu as S
import tham_fem as T
import on_dinh as O
import chay_tham_on_dinh as C

GW = S.GAMMA_W


def chi_tiet(od, xc, zc, R, ns=40, q_s=0.0, x_qs=None):
    """Chia dải giống OnDinh.K_vong; trả về đủ đại lượng cho Bishop/Spencer. q_s (kPa) đặt trên [x_qs]."""
    r = od.r; mc = r.mc
    xs = r.xs
    inside = np.abs(xs - xc) < R
    zb_all = np.full_like(xs, np.inf)
    zb_all[inside] = zc - np.sqrt(R ** 2 - (xs[inside] - xc) ** 2)
    below = zb_all < r.ztop - 1e-6
    if below.sum() < 4:
        return None
    idx = np.where(below)[0]
    splits = np.where(np.diff(idx) > 1)[0]
    if len(splits):
        segs = np.split(idx, splits + 1)
        depth = [np.max(r.ztop[s] - zb_all[s]) for s in segs]
        idx = segs[int(np.argmax(depth))]
    xa, xb = xs[idx[0]], xs[idx[-1]]
    if idx[0] == 0 or idx[-1] == len(xs) - 1:
        return None
    if xb - xa < 3.0 or np.max(r.ztop[idx] - zb_all[idx]) < 1.0:
        return None
    edges = np.linspace(xa, xb, ns + 1)
    xm = 0.5 * (edges[1:] + edges[:-1]); b = np.diff(edges)
    zb = zc - np.sqrt(np.maximum(R ** 2 - (xm - xc) ** 2, 0))
    zt = mc.mat_tren(xm)
    h = np.maximum(zt - zb, 0)
    nz = 30
    tt = (np.arange(nz) + 0.5) / nz
    zz = zb[:, None] + tt[None, :] * h[:, None]
    XX = np.repeat(xm[:, None], nz, 1)
    ii, jj = r.tra(XX, zz)
    mm = r.mat[ii, jj]
    ps = np.nan_to_num(r.psi[ii, jj], nan=-1.0)
    gam = np.where(ps > 0, od.gs[mm], od.g[mm])
    gam = np.where(mm == 0, 0.0, gam)
    W = (gam * (h[:, None] / nz)).sum(1) * b
    zg = np.where(W > 0, (gam * zz * (h[:, None] / nz)).sum(1) * b / np.maximum(W, 1e-9), zb)
    ib, jb = r.tra(xm, zb + 0.05)
    mb = r.mat[ib, jb]
    u = GW * np.maximum(np.nan_to_num(r.psi[ib, jb], nan=0.0), 0.0)
    c = od.c[mb]; tphi = od.tphi[mb]
    Wn = np.zeros_like(W); Fx = np.zeros_like(W)
    if od.Zw is not None:
        dep = np.maximum(od.Zw - zt, 0.0)
        Wn = GW * dep * b
        z1 = mc.mat_tren(edges[:-1]); z2 = mc.mat_tren(edges[1:])
        Fx = GW * dep * (z2 - z1)
    if q_s and x_qs is not None:          # phụ tải thi công trên đỉnh đập: cộng vào tải đứng ngoài
        lo = np.maximum(edges[:-1], x_qs[0]); hi = np.minimum(edges[1:], x_qs[1])
        Wn = Wn + q_s * np.maximum(hi - lo, 0.0)
    return dict(xm=xm, b=b, zb=zb, zt=zt, zg=zg, W=W, Wn=Wn, Fx=Fx, u=u, c=c, tphi=tphi, xc=xc, zc=zc, R=R,
                d=od.d, kh=od.kh, xa=xa, xb=xb)


def _mdrive(s):
    d = s['d']; xc, zc = s['xc'], s['zc']
    MW = np.sum(-(s['xm'] - xc) * (s['W'] + s['Wn']))
    Mn = np.sum(-(s['zt'] - zc) * s['Fx'])
    Meq = np.sum(-(s['zg'] - zc) * (d * s['kh'] * s['W']))
    return d * (MW + Mn) + d * Meq


def bishop(s):
    """Bishop đơn giản – đúng như OnDinh.K_vong (để đối chiếu)."""
    d, R = s['d'], s['R']
    sina = np.clip(d * (s['xc'] - s['xm']) / R, -0.999, 0.999); cosa = np.sqrt(1 - sina ** 2)
    Md = _mdrive(s)
    Neff = np.maximum(s['W'] + s['Wn'] - s['u'] * s['b'], 0.0)
    K = 1.5
    for _ in range(100):
        ma = np.maximum(cosa * (1 + (sina / cosa) * s['tphi'] / K), 0.2)
        Kn = R * np.sum((s['c'] * s['b'] + Neff * s['tphi']) / ma) / Md
        if abs(Kn - K) < 1e-7:
            break
        K = Kn
    return Kn


def spencer(s):
    """Spencer (GLE, lực tương tác giữa các dải nghiêng không đổi góc θ).
    Hệ quy chiếu lật sao cho khối trượt chuyển động theo +x. Với mỗi (F, θ): giải cân bằng lực 2 phương của từng dải
    để được N_i (pháp tuyến đáy, toàn phần) và Q_i (hợp lực tương tác). F_m: cân bằng mô men toàn khối quanh tâm cung;
    F_f: ΣQ_i = 0 (cân bằng lực ngang toàn khối). Nghiệm Spencer: θ sao cho F_m = F_f."""
    d, R = s['d'], s['R']
    xm, xc = d * s['xm'], d * s['xc']          # lật trục x nếu mái thượng lưu
    Fx = d * s['Fx']
    W, Wn, u, c, T_, b = s['W'], s['Wn'], s['u'], s['c'], s['tphi'], s['b']
    kh = s['kh']
    sina = np.clip((xc - xm) / R, -0.999, 0.999); cosa = np.sqrt(1 - sina ** 2)
    l = b / cosa
    Cc = c * l - u * l * T_
    Md = _mdrive(s)

    def NQ(F, th):
        a11 = sina - T_ * cosa / F; a12 = np.cos(th)
        a21 = cosa + T_ * sina / F; a22 = np.sin(th)
        b1 = Cc * cosa / F - kh * W - Fx
        b2 = W + Wn - Cc * sina / F
        det = a11 * a22 - a12 * a21
        N = (b1 * a22 - a12 * b2) / det
        Q = (a11 * b2 - a21 * b1) / det
        return N, Q

    def Fm(th):
        F = 1.5
        for _ in range(200):
            N, Q = NQ(F, th)
            Sres = (c * l + (N - u * l) * T_)        # = F·S_i
            Fn = R * np.sum(Sres) / Md
            if abs(Fn - F) < 1e-8:
                break
            F = 0.5 * (F + Fn)
        return Fn

    def Ff(th):
        lo, hi = 0.3, 6.0
        g = lambda F: np.sum(NQ(F, th)[1])
        glo, ghi = g(lo), g(hi)
        if glo * ghi > 0:
            return np.nan
        for _ in range(100):
            mid = 0.5 * (lo + hi); gm = g(mid)
            if glo * gm <= 0:
                hi, ghi = mid, gm
            else:
                lo, glo = mid, gm
        return 0.5 * (lo + hi)

    ths = np.radians(np.linspace(-30, 45, 151))
    fm = np.array([Fm(t) for t in ths]); ff = np.array([Ff(t) for t in ths])
    dlt = fm - ff
    k = np.where(np.isfinite(dlt[:-1]) & np.isfinite(dlt[1:]) & (dlt[:-1] * dlt[1:] <= 0))[0]
    if not len(k):
        return dict(F=np.nan, theta=np.nan, Fm0=float(Fm(0.0)))
    i = k[np.argmin(np.abs(ths[k]))]
    lo, hi = ths[i], ths[i + 1]
    for _ in range(60):
        mid = 0.5 * (lo + hi)
        if (Fm(lo) - Ff(lo)) * (Fm(mid) - Ff(mid)) <= 0:
            hi = mid
        else:
            lo = mid
    th = 0.5 * (lo + hi)
    return dict(F=float(Fm(th)), theta=float(np.degrees(th)), Fm0=float(Fm(0.0)), Ff0=float(Ff(0.0)))


class OnDinhPhuTai(O.OnDinh):
    """OnDinh có phụ tải phân bố q_s (kPa) trên đoạn x_qs của mặt đỉnh – dùng cho thời kỳ thi công."""
    def __init__(self, *a, q_s=0.0, x_qs=None, **k):
        super().__init__(*a, **k)
        self.q_s, self.x_qs = q_s, x_qs

    def K_vong(self, xc, zc, R, ns=40, tra_chi_tiet=False):
        s = chi_tiet(self, xc, zc, R, ns=ns, q_s=self.q_s, x_qs=self.x_qs)
        if s is None or _mdrive(s) <= 0:
            return np.inf, None
        K = bishop(s)
        if tra_chi_tiet:
            return K, s
        return K, (s['xa'], s['xb'])


def tim(od, mc, chieu, min_=1):
    """Lưới tìm tâm như chay_tham_on_dinh.chay_on_dinh; min_ = 2: lưới mịn gấp đôi (tâm và bán kính)."""
    nx, nz, nr = 28 * min_, 26 * min_, 26 * min_
    if chieu > 0:
        return od.tim((-10, 95), (70, 170), nx=nx, nz=nz, nr=nr, x_gioi_han=(mc.xcu - 1, mc.x_phai - 2))
    return od.tim((-95, 10), (70, 170), nx=nx, nz=nz, nr=nr, x_gioi_han=(mc.x_trai + 2, mc.xcd + 1))


if __name__ == '__main__':
    t0 = time.time()
    KQ = {}
    mc = T.MatCat(Zc=S.Z_DINH_DAP_SC)
    print('Đỉnh đập mô hình', mc.Zc)
    bt_dbt, _ = C.chay_tham(mc, S.MNDBT, 'MNDBT')
    bt_tac, _ = C.chay_tham(mc, S.MNDBT, 'MNDBT_tac', kdat={2: S.DAT['DD-II']['k'], 4: S.DAT['DD-II']['k']})
    vl = O.vat_lieu_on_dinh('I')
    TH = [  # mã, bt, Zw ngoài mái, chiều, kh, mưa
        ('TH1', bt_dbt, None, +1, 0.0, None),
        ('TH4', bt_tac, None, +1, 0.0, None),
        ('TH5', bt_dbt, None, +1, 0.05, None),
        ('TH7', bt_dbt, S.MNDBT, -1, 0.0, None),
        ('TH10', bt_dbt, S.MNDBT, -1, 0.05, None),
    ]
    for ma, bt, Zng, ch, kh, mua in TH:
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        od = O.OnDinh(ras, vl, Zw=Zng, kh=kh, chieu=ch)
        K1, c1 = tim(od, mc, ch, 1)
        K2, c2 = tim(od, mc, ch, 2)
        s = chi_tiet(od, *c1)
        sp = spencer(s)
        KQ[ma] = dict(K_bishop=K1, tam=c1, K_bishop_luoi_min=K2, tam_min=c2, K_bishop_kt=bishop(s), spencer=sp)
        print(ma, 'Bishop %.3f (lưới mịn %.3f) | kiểm tra %.3f | Spencer F=%.3f θ=%.1f° (Fm(θ=0)=%.3f)' % (
            K1, K2, bishop(s), sp['F'], sp['theta'], sp['Fm0']), ' %.0fs' % (time.time() - t0))
    # (3) thời kỳ thi công – sửa chữa: q = 20 kPa trên toàn bề rộng đỉnh đập (xe, máy lu 9–16 T, vật liệu tập kết)
    for q in (10.0, 20.0):
        for ch, ma in ((+1, 'TC_HL'), (-1, 'TC_TL')):
            ras = O.LuoiRaster(bt_dbt.xy, bt_dbt.tri, bt_dbt.mat, bt_dbt.H, mc, dx=0.25)
            od = OnDinhPhuTai(ras, vl, Zw=(S.MNDBT if ch < 0 else None), kh=0.0, chieu=ch, q_s=q, x_qs=(mc.xcu, mc.xcd))
            K1, c1 = tim(od, mc, ch, 1)
            s = chi_tiet(od, *c1, q_s=q, x_qs=(mc.xcu, mc.xcd))
            sp = spencer(s)
            KQ['%s_q%.0f' % (ma, q)] = dict(q=q, K_bishop=K1, tam=c1, spencer=sp)
            print(ma, 'q=%.0f kPa' % q, 'Bishop %.3f | Spencer %.3f  [K] = 1,24' % (K1, sp['F']), ' %.0fs' % (time.time() - t0))
    json.dump(KQ, open('../ket_qua/phan_tich_bo_sung_on_dinh.json', 'w'), ensure_ascii=False, indent=1, default=float)
    print('xong', round(time.time() - t0, 1))
