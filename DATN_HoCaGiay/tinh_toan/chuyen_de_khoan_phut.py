# -*- coding: utf-8 -*-
"""
CHUYÊN ĐỀ (Chương 8): ĐÁNH GIÁ HIỆU QUẢ MÀNG KHOAN PHỤT CHỐNG THẤM MANG TRÁI TRÀN SỰ CỐ
- Mặt cắt tính toán: "Cắt ngang đại diện khoan phụt" (TKCS Hình 2-4, tỷ lệ 1/250), cao độ mặt đất đọc từ
  hàng "Cao độ tự nhiên – Khoảng cách mia"; ranh giới lớp 5 / I / II số hóa từ bản vẽ (sai số ~0,2 m).
- Hệ số thấm: lớp 5 và lớp I theo thí nghiệm hiện trường tại CG7, CG8 (TKCS mục 2.1.2);
  lớp II không có thí nghiệm -> giả thiết 1e-5 cm/s, có phân tích độ nhạy.
- Màng khoan phụt: 3 hàng, a = 1,5 m, A = 3,0 m -> T = 4,8 m; đỉnh +77,00; đáy +67,0 ÷ +70,0
  (mặt cắt đại diện ≈ +68,3); K_màng thiết kế = 1e-5 cm/s.
- Bài toán thấm ổn định bão hòa – không bão hòa, dùng bộ giải tham_fem.BaiToanTham.
"""
import json
import os
import numpy as np
import triangle as tr
import tham_fem as F
import ve_hinh as VH
import matplotlib.pyplot as plt
import matplotlib.tri as mtri

KQ = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'ket_qua')

# --- mặt đất (x = 0 tại mép trái bản vẽ, dương về hạ lưu) -------------------------------------------
KC = [3.86, 4.00, 2.70, 15.37, 1.63, 1.85, 2.26, 2.88, 2.93, 3.64, 4.04, 3.83, 2.44, 2.28, 3.67, 2.40]
CD = [75.62, 76.53, 78.09, 78.75, 78.75, 77.65, 76.92, 76.29, 75.54, 74.75, 74.65, 74.55, 74.42, 74.32,
      74.23, 74.13, 74.07]
X_MAT = np.concatenate([[0.0], np.cumsum(KC)])
# mặt đập lấy theo cao trình sau sửa chữa +78,75 (hiện trạng +78,38 ÷ +78,42) để so sánh 'có/không có màng'
X_TUONG = X_MAT[3]          # tường chắn sóng (mép thượng lưu mặt đập) x = 10,56
X_MEP_HL = X_MAT[4]         # mép hạ lưu mặt đập x = 25,93
X_TIM = 17.81               # tim đập = hàng khoan B
MAT_X = np.concatenate([[-25.0, -16.9], X_MAT, [100.0]])
MAT_Z = np.concatenate([[70.0, 70.0], CD, [73.70]])
# ranh giới lớp (số hóa)
L5_I = (np.array([-25.0, 0.0, 14.7, 20.8, 39.2, 60.0, 100.0]), np.array([74.40, 74.53, 74.84, 74.70, 72.84, 72.00, 71.00]))
LI_II = (np.array([-25.0, 0.0, 15.0, 20.5, 39.2, 60.0, 100.0]), np.array([72.50, 72.50, 72.45, 72.35, 70.60, 69.55, 68.20]))
Z_DAY_MH = 58.0
X0, X1 = MAT_X[0], MAT_X[-1]

ID5, IDI, IDII, IDM = 21, 22, 23, 24
TEN = {ID5: 'Lớp 5 (K = 6,5e-5 cm/s)', IDI: 'Lớp I (K = 4,3e-4 cm/s)', IDII: 'Lớp II (đá PH vừa)',
       IDM: 'Màng khoan phụt'}
K_CM = {ID5: 6.54e-5, IDI: 4.35e-4, IDII: 1.0e-5}       # cm/s


def mat_tren(x):
    return np.interp(x, MAT_X, MAT_Z)


def z_lop(L, x):
    return np.interp(x, L[0], L[1])


def _luoi_x(man):
    xs_extra = [] if man is None else [man['xa'], man['xb']]
    X = np.unique(np.round(np.concatenate([MAT_X, L5_I[0], LI_II[0], xs_extra]), 6))
    # thêm giao điểm ranh giới lớp với mặt đất (lớp lộ ra trên mái thượng lưu)
    them = []
    for L in (L5_I, LI_II):
        f = mat_tren(X) - z_lop(L, X)
        for i in range(len(X) - 1):
            if f[i] * f[i + 1] < 0:
                them.append(X[i] - f[i] * (X[i + 1] - X[i]) / (f[i + 1] - f[i]))
    return np.unique(np.round(np.concatenate([X, them]), 6))


def luoi(man=None, a_ngoai=0.3, a_man=0.08):
    """man = dict(xa, xb, ztop, zbot) hoặc None (hiện trạng)."""
    X = _luoi_x(man)
    V = []; idx = {}

    def v(x, z):
        k = (round(x, 5), round(z, 5))
        if k not in idx:
            idx[k] = len(V); V.append([x, z])
        return idx[k]
    segs = []

    def line(pts):
        ids = [v(*p) for p in pts]
        for a, b in zip(ids[:-1], ids[1:]):
            if a != b:
                segs.append([a, b])
    zt = mat_tren(X)
    line(list(zip(X, zt)))                                    # mặt đất
    line([(X0, Z_DAY_MH), (X1, Z_DAY_MH)])                    # đáy
    for xb_ in (X0, X1):                                      # biên trái / phải
        zz = [Z_DAY_MH] + sorted(float(z_lop(L, xb_)) for L in (LI_II, L5_I) if z_lop(L, xb_) < mat_tren(xb_) - 1e-6)
        line([(xb_, z) for z in zz] + [(xb_, float(mat_tren(xb_)))])
    for L in (L5_I, LI_II):
        zl = z_lop(L, X)
        zz = np.where(np.abs(zl - zt) < 1e-4, zt, np.minimum(zl, zt))
        duoi = zl <= zt + 1e-4                                # chỉ phần nằm dưới mặt đất
        cat = False
        if man is not None:
            za, zb = z_lop(L, man['xa']), z_lop(L, man['xb'])
            cat = man['zbot'] < min(za, zb) and max(za, zb) < man['ztop']
        pts = list(zip(X, zz, duoi))
        nhom = [pts] if not cat else [[p for p in pts if p[0] <= man['xa'] + 1e-9],
                                      [p for p in pts if p[0] >= man['xb'] - 1e-9]]
        for g in nhom:
            doan = []
            for xi, zi, ok in g:
                if ok:
                    doan.append((xi, zi))
                else:
                    if len(doan) > 1:
                        line(doan)
                    doan = []
            if len(doan) > 1:
                line(doan)
    if man is not None:
        for xm in (man['xa'], man['xb']):
            zc = sorted({man['zbot'], man['ztop']} | {float(z_lop(L, xm)) for L in (L5_I, LI_II)
                                                      if man['zbot'] < z_lop(L, xm) < man['ztop']})
            line([(xm, z) for z in zc])
        line([(man['xa'], man['ztop']), (man['xb'], man['ztop'])])
        line([(man['xa'], man['zbot']), (man['xb'], man['zbot'])])
    reg = []
    xs_seed = [5.0, 30.0, 90.0] + ([] if man is None else [man['xa'] - 0.8, man['xb'] + 0.8])
    for xr in xs_seed:
        a, b, c = mat_tren(xr), z_lop(L5_I, xr), z_lop(LI_II, xr)
        if man is not None and man['xa'] < xr < man['xb']:
            continue
        reg.append([xr, 0.5 * (a + b), ID5, a_ngoai])
        reg.append([xr, 0.5 * (b + c), IDI, a_ngoai])
        reg.append([xr, 0.5 * (c + Z_DAY_MH), IDII, a_ngoai * 2])
    reg.append([-20.0, 64.0, IDII, a_ngoai * 2])
    if man is not None:
        reg.append([0.5 * (man['xa'] + man['xb']), 0.5 * (man['ztop'] + man['zbot']) + 0.01, IDM, a_man])
        reg.append([0.5 * (man['xa'] + man['xb']), 0.5 * (man['ztop'] + mat_tren(man['xa'])), ID5, a_ngoai])
    A = dict(vertices=np.array(V), segments=np.array(segs), regions=np.array(reg))
    T = tr.triangulate(A, 'pq30Aa')
    return T['vertices'], T['triangles'], T['triangle_attributes'][:, 0].astype(int)


def giai(Zw, man=None, K_man=1e-5, K_II=1e-5, msh=None):
    xy, tri, mat = luoi(man) if msh is None else msh
    kd = {ID5: K_CM[ID5] / 100, IDI: K_CM[IDI] / 100, IDII: K_II / 100, IDM: K_man / 100}
    bt = F.BaiToanTham(xy, tri, mat, kdat=kd, alpha=1.0)
    x, z = xy[:, 0], xy[:, 1]
    on_top = np.abs(z - mat_tren(x)) < 1e-6
    up = np.where(on_top & (x <= X_TUONG + 1e-6) & (z <= Zw + 1e-9))[0]
    seep = np.where(on_top & (x >= X_MEP_HL - 1e-6))[0]
    H0 = np.interp(x, [X0, X1], [Zw, mat_tren(X1)])
    bt.giai(up, np.full(len(up), Zw), seep, H0=H0)
    q_in = bt.luu_luong(up); q_out = -bt.luu_luong(bt.seep_on)
    gx, gz, J = bt.gradient()
    psi_e = bt.H[tri].mean(1) - bt.zc
    # gradient thoát ra: phần tử có nút thuộc mặt thấm ra đang hoạt động, vùng bão hòa
    sset = np.zeros(len(xy), bool); sset[bt.seep_on] = True
    e_ra = sset[tri].any(1) & (psi_e > -0.05)
    J_ra = float(J[e_ra].max()) if e_ra.any() else 0.0
    J_ra_tb = float(np.average(J[e_ra], weights=bt.area[e_ra])) if e_ra.any() else 0.0
    out = dict(Zw=Zw, q_in=q_in, q_out=q_out, q_ngay=q_in * 86400, J_ra=J_ra, J_ra_tb=J_ra_tb,
               n_seep=int(len(bt.seep_on)), niter=int(bt.niter))
    if len(bt.seep_on):
        out['z_seep_max'] = float(z[bt.seep_on].max()); out['x_seep_min'] = float(x[bt.seep_on].min())
    if man is not None:
        em = (mat == IDM) & (psi_e > 0)
        if em.any():
            out['J_man_tb'] = float(np.average(J[em], weights=bt.area[em]))
            out['J_man_95'] = float(np.percentile(J[em], 95))
        # chênh cột nước qua màng ở các cao độ
        Ti = mtri.LinearTriInterpolator(mtri.Triangulation(x, z, tri), bt.H)
        dH = []
        for zz in np.arange(man['zbot'] + 0.25, man['ztop'], 0.5):
            hu, hd = Ti(man['xa'] - 0.05, zz), Ti(man['xb'] + 0.05, zz)
            if np.ma.is_masked(hu) or np.ma.is_masked(hd):
                continue
            if hu - zz > 0:
                dH.append((float(zz), float(hu - hd)))
        out['dH_man'] = dH
        out['dH_man_max'] = max([d for _, d in dH]) if dH else 0.0
    # đường bão hòa ψ = 0 tại một số mặt cắt đứng
    Ti = mtri.LinearTriInterpolator(mtri.Triangulation(x, z, tri), bt.H - z)
    for xx in (X_TIM - 3.0, X_TIM + 3.0, X_MEP_HL, 30.0, X_MAT[9], 45.0):
        zs = np.linspace(Z_DAY_MH + 0.1, mat_tren(xx) - 1e-3, 400)
        ps = np.array([Ti(xx, zz) for zz in zs], float)
        k = np.where(ps > 0)[0]
        out['zbh_%.1f' % xx] = float(zs[k[-1]]) if len(k) else None
    return bt, (xy, tri, mat), out


def ve(bt, msh, man, Zw, ten, tieu_de):
    xy, tri, mat = msh
    fig, ax = plt.subplots(figsize=(11, 3.8))
    VH.ve_vat_lieu(ax, xy, tri, mat, TEN, xlim=(-10, 70), ylim=(62, 80.5))
    T = mtri.Triangulation(xy[:, 0], xy[:, 1], tri)
    cs = ax.tricontour(T, bt.H, levels=np.arange(70, 80, 0.5), colors='b', linewidths=0.5)
    ax.clabel(cs, fmt=lambda v: ('%.1f' % v).replace('.', ','), fontsize=6.5)
    ax.tricontour(T, bt.H - xy[:, 1], levels=[0.0], colors='r', linewidths=1.6)
    ax.plot(MAT_X, MAT_Z, 'k-', lw=1.0)
    xw = np.interp(Zw, MAT_Z[1:5], MAT_X[1:5])
    ax.fill_between([-10, xw], [Zw, Zw], [mat_tren(-10), Zw], color='#9ecae1', alpha=0.5, lw=0)
    ax.text(-9, Zw + 0.3, ('MN hồ %.2f' % Zw).replace('.', ','), fontsize=8, color='navy')
    if len(bt.seep_on):
        ax.plot(xy[bt.seep_on, 0], xy[bt.seep_on, 1], 'o', ms=2.5, color='magenta', label='mặt thấm ra')
    ax.set_title(tieu_de); ax.set_xlabel('x (m)'); ax.set_ylabel('Cao độ (m)')
    return VH.luu(fig, ten)


if __name__ == '__main__':
    MAN = dict(xa=X_TIM - 2.4, xb=X_TIM + 2.4, ztop=77.00, zbot=68.30)
    res = {'hien_trang': {}, 'thiet_ke': {}, 'nhay_K': [], 'nhay_day': [], 'nhay_KII': []}
    msh0 = luoi(None); msh1 = luoi(MAN)
    print('lưới HT', len(msh0[0]), len(msh0[1]), ' lưới SC', len(msh1[0]), len(msh1[1]))
    for Zw in (74.0, 74.70, 77.63, 78.54):
        bt0, _, r0 = giai(Zw, None, msh=msh0)
        bt1, _, r1 = giai(Zw, MAN, msh=msh1)
        res['hien_trang']['%.2f' % Zw] = r0; res['thiet_ke']['%.2f' % Zw] = r1
        print('Zw %.2f HT q=%.3e (%.2f m3/ng.đ/m) Jra=%.2f seep_zmax=%s | SC q=%.3e (%.2f) Jra=%.2f Jman=%.2f dHmax=%.2f giảm %.0f%%' % (
            Zw, r0['q_in'], r0['q_ngay'], r0['J_ra'], r0.get('z_seep_max'), r1['q_in'], r1['q_ngay'], r1['J_ra'],
            r1.get('J_man_tb', 0), r1.get('dH_man_max', 0), 100 * (1 - r1['q_in'] / r0['q_in'])))
        if Zw in (74.70, 78.54):
            ve(bt0, msh0, None, Zw, 'kp_HT_%.0f.png' % (Zw * 100), ('Chưa khoan phụt – MN hồ %.2f' % Zw).replace('.', ','))
            ve(bt1, msh1, MAN, Zw, 'kp_SC_%.0f.png' % (Zw * 100),
               'Sau khoan phụt (T = 4,8 m; K = 1×10$^{-5}$ cm/s; đáy +68,3) – MN hồ ' + ('%.2f' % Zw).replace('.', ','))
    Zw = 77.63
    q0 = res['hien_trang']['77.63']['q_in']
    for Km in (1e-6, 5e-6, 1e-5, 5e-5, 1e-4):
        _, _, r = giai(Zw, MAN, K_man=Km, msh=msh1)
        r['K_man'] = Km; r['giam'] = 1 - r['q_in'] / q0; res['nhay_K'].append(r)
        print('K_man %.0e q=%.3e giảm %.0f%% Jra %.2f Jman %.2f' % (Km, r['q_in'], 100 * r['giam'], r['J_ra'], r.get('J_man_tb', 0)))
    for zb in (74.0, 72.0, 70.0, 68.3, 66.0):
        man = dict(MAN, zbot=zb)
        _, _, r = giai(Zw, man)
        r['zbot'] = zb; r['giam'] = 1 - r['q_in'] / q0; res['nhay_day'].append(r)
        print('đáy %.1f q=%.3e giảm %.0f%% Jra %.2f Jman %.2f' % (zb, r['q_in'], 100 * r['giam'], r['J_ra'], r.get('J_man_tb', 0)))
    for KII in (1e-6, 1e-5, 1e-4):
        _, _, ra = giai(Zw, None, K_II=KII, msh=msh0)
        _, _, rb = giai(Zw, MAN, K_II=KII, msh=msh1)
        res['nhay_KII'].append(dict(K_II=KII, q_HT=ra['q_in'], q_SC=rb['q_in'], giam=1 - rb['q_in'] / ra['q_in'],
                                    J_ra_HT=ra['J_ra'], J_ra_SC=rb['J_ra']))
        print('K_II %.0e HT q=%.3e SC q=%.3e giảm %.0f%%' % (KII, ra['q_in'], rb['q_in'], 100 * (1 - rb['q_in'] / ra['q_in'])))
    json.dump(res, open(os.path.join(KQ, 'khoan_phut.json'), 'w'), ensure_ascii=False, indent=1, default=float)
    # biểu đồ độ nhạy
    fig, axs = plt.subplots(1, 2, figsize=(10, 3.6))
    a = res['nhay_K']
    axs[0].semilogx([r['K_man'] for r in a], [100 * r['giam'] for r in a], 'o-')
    axs[0].set_xlabel('Hệ số thấm màng K (cm/s)'); axs[0].set_ylabel('Mức giảm lưu lượng thấm (%)')
    axs[0].set_title('a) Ảnh hưởng của chất lượng màng (MNLTK)')
    b = res['nhay_day']
    axs[1].plot([r['zbot'] for r in b], [100 * r['giam'] for r in b], 's-')
    axs[1].invert_xaxis(); axs[1].set_xlabel('Cao độ đáy màng (m)'); axs[1].set_ylabel('Mức giảm lưu lượng thấm (%)')
    axs[1].set_title('b) Ảnh hưởng của chiều sâu màng (MNLTK)')
    VH.luu(fig, 'kp_do_nhay.png')


def nhay_T_dinh():
    """Độ nhạy chiều dày màng (1, 2, 3 hàng theo T = Bb + 0,6A) và cao trình đỉnh màng."""
    out = []
    for Zw in (77.63, 78.54):
        q0 = giai(Zw, None)[2]['q_in']
        for so_hang, T in ((1, 1.8), (2, 3.3), (3, 4.8)):
            for ztop in (77.0, 78.0):
                man = dict(xa=X_TIM - T / 2, xb=X_TIM + T / 2, ztop=ztop, zbot=68.30)
                r = giai(Zw, man)[2]
                out.append(dict(Zw=Zw, so_hang=so_hang, T=T, ztop=ztop, q=r['q_in'], giam=1 - r['q_in'] / q0,
                                J_ra=r['J_ra'], J_man=r.get('J_man_tb', 0), J_man95=r.get('J_man_95', 0),
                                dH=r.get('dH_man_max', 0)))
                print(Zw, so_hang, T, ztop, 'q %.3e giảm %.0f%% Jra %.2f Jman %.2f dH %.2f' % (
                    r['q_in'], 100 * out[-1]['giam'], r['J_ra'], out[-1]['J_man'], out[-1]['dH']))
    return out
