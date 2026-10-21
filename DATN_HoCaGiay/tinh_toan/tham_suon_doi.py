# -*- coding: utf-8 -*-
"""
TÍNH THẤM MẶT CẮT SƯỜN ĐỒI (VAI ĐẬP) – BỔ SUNG BẢNG LƯU LƯỢNG VÀ GRADIEN THẤM CÒN THIẾU
Hình học số hóa gần đúng từ mô hình SEEP/W mặt cắt sườn đồi của SV (Hình 2.2 bản thuyết minh SV):
  đỉnh +78,50 rộng 5 m; mái thượng lưu m = 3,0 đến cơ +68,0 rộng 3 m, dưới cơ m = 3,5 đến mặt đất +64,6;
  mái hạ lưu m = 3,0 đến mặt đất +68,3 (không có cơ); khối lõi ĐĐ-I có chân khay đáy +63,8 rộng 5 m;
  khối thượng lưu, hạ lưu bằng ĐĐ-II; đệm tiêu nước nằm ngang ở chân hạ lưu (x = 80 ÷ 96,3 m, đỉnh +68,7);
  nền một lớp từ +48,0 lên mặt đất ("Lop 6" trong mô hình).
Mô hình này là Mặt cắt V – T9 của tư vấn (BCNCKT, Phụ lục 3 mục 3.5, Hình PL 3-65 ÷ 3-80). Hệ số thấm hiệu chỉnh cho khớp
lưu lượng tư vấn đọc trên hình (1,613 / 3,271 / 3,616·10⁻⁶ ở +74,70 / +77,43 / +78,25; 0,988·10⁻⁶ khi đệm tắc):
  lõi k = 1·10⁻⁴ cm/s (thí nghiệm trong phòng – BCNCKT), nền lớp 4 k = 4,11·10⁻⁵ cm/s, khối TL, HL ĐĐ-II
  → lệch +4 / −7 / −1 / +13 %. Với nền lớp I (5·10⁻⁴ cm/s) lưu lượng lớn gấp gần 5 lần – không phù hợp.
Cùng bộ giải phần tử hữu hạn (tham_fem.BaiToanTham) và cùng cách lấy gradien với mặt cắt lòng sông
(chay_tham_on_dinh.phan_tich_tham): chỉ xét vùng bão hòa, lấy phân vị 95 %
  J vào – phần tử kề biên thượng lưu ngập nước; J ra – phần tử đất kề mặt thấm ra;
  J tiếp xúc – phần tử đất kề đệm tiêu nước; J chân khay – phần tử lõi nằm trong chân khay.
  python3 tham_suon_doi.py
"""
import json
import os

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.tri as mtri
import numpy as np
import triangle as tr

import so_lieu as S
import tham_fem as F

DIR = os.path.dirname(os.path.abspath(__file__))
plt.rcParams.update({'font.family': 'Liberation Serif'})
X0, X1, Z_DAY = 0.0, 110.0, 48.0
DINH, X_DU, X_DD = 78.50, 61.9, 66.9                 # đỉnh đập, mép đỉnh TL, mép đỉnh HL
MAT = [(0.0, 65.7), (15.5, 64.6), (27.4, 68.0), (30.4, 68.0), (X_DU, DINH), (X_DD, DINH), (96.3, 68.7),
       (97.5, 68.3), (X1, 68.3)]                      # mặt đất – mái – đỉnh (đường bao trên)
NEN = [(15.5, 64.6), (31.0, 65.5), (40.0, 66.5), (55.5, 67.2), (60.4, 63.8), (65.4, 63.8), (69.5, 67.4),
       (72.0, 67.5), (80.0, 67.8), (97.5, 68.3)]     # mặt nền dưới thân đập, kể cả chân khay
CK = (55.5, 67.2, 69.5, 67.4)                         # miệng chân khay (mặt đất tự nhiên)
VL = {1: ('Lõi ĐĐ-I', 'DD-I'), 12: ('Khối thượng lưu ĐĐ-II', 'DD-II'), 3: ('Khối hạ lưu ĐĐ-II', 'DD-II'),
      4: ('Đệm tiêu nước', 'DC'), 20: ('Nền', None)}


def z_nen(x):
    xs, zs = zip(*[p for p in NEN if p[0] not in (60.4, 65.4)])     # bỏ đáy chân khay khi nội suy mặt nền
    return float(np.interp(x, xs, zs))


def z_mat(x):
    return np.interp(x, *zip(*MAT))


def luoi():
    V, idx, segs = [], {}, []

    def v(p):
        key = (round(p[0], 4), round(p[1], 4))
        if key not in idx:
            idx[key] = len(V)
            V.append(key)
        return idx[key]

    def line(pts):
        ids = [v(p) for p in pts]
        segs.extend((a, b) for a, b in zip(ids[:-1], ids[1:]) if a != b)

    # điểm đầu các ranh giới bên trong phải là đỉnh của đường bao / mặt nền thì các vùng mới khép kín
    mat_day_du = sorted(MAT + [(62.4, DINH), (66.4, DINH)])
    nen_day_du = NEN[:1] + [(16.5, z_nen(16.5))] + NEN[1:8] + [(77.5, z_nen(77.5))] + NEN[8:]
    line([(X0, Z_DAY), (X1, Z_DAY), (X1, 68.3)] + mat_day_du[::-1][1:] + [(X0, Z_DAY)])
    line(nen_day_du)
    line([(16.5, z_nen(16.5)), (30.4, 65.9), (59.6, 75.6), (62.4, DINH)])      # ranh giới lõi – khối TL
    line([(66.4, DINH), (77.5, z_nen(77.5))])                                   # ranh giới lõi – khối HL
    line([(80.0, z_nen(80.0)), (80.0, 68.7), (96.3, 68.7)])                     # đệm tiêu nước
    regions = [(50.0, 55.0, 20, 1.0), (64.0, 72.0, 1, 0.2), (45.0, 71.8, 12, 0.2), (85.0, 70.5, 3, 0.2),
               (88.0, 68.4, 4, 0.03)]
    T = tr.triangulate(dict(vertices=np.array(V, float), segments=np.array(segs, int),
                            regions=np.array(regions, float)), 'pq30Aa')
    return T['vertices'], T['triangles'], T['triangle_attributes'][:, 0].astype(int)


def k_vat_lieu(nen='L5'):
    return {m: S.DAT[nen if key is None else key]['k'] for m, (ten, key) in VL.items()}


NEN_CHINH = 'L4'
K_HIEU_CHINH = {1: 1.0e-6}       # lõi: k = 1·10⁻⁴ cm/s (thí nghiệm trong phòng, BCNCKT) – hiệu chỉnh theo kết quả T9
TU_VAN_T9 = {74.70: 1.6132e-6, 77.43: 3.2713e-6, 78.25: 3.6157e-6, 'tac': 0.98831e-6}   # q đọc trên Hình PL 3-66 ÷ 3-78


def chay(Zw, nen=NEN_CHINH, mesh=None, k_them=None):
    """k_them: thay hệ số thấm vật liệu, vd. {4: k_ĐĐ-II} – đệm tiêu nước bị tắc."""
    xy, tri, mat = mesh if mesh is not None else luoi()
    kd = k_vat_lieu(nen)
    kd.update(K_HIEU_CHINH)
    kd.update(k_them or {})
    bt = F.BaiToanTham(xy, tri, mat, kdat=kd, alpha=1.0)
    x, z = xy[:, 0], xy[:, 1]
    tren = (np.abs(z - z_mat(x)) < 1e-6) & (z > Z_DAY + 1e-6)
    up = np.where(tren & (x <= X_DU + 1e-6) & (z <= Zw + 1e-9))[0]
    seep = np.where(tren & (x >= X_DD - 1e-6))[0]
    H0 = np.interp(x, [15.5, 97.5], [Zw, 68.3])
    bt.giai(up, np.full(len(up), Zw), seep, H0=H0)
    gx, gz, J = bt.gradient()
    psi_e = bt.H[tri].mean(1) - xy[tri, 1].mean(1)
    bh = psi_e >= -0.05

    def ke(nodes):
        return np.where(np.isin(tri, np.asarray(nodes, int)).any(1))[0]

    def tk(idx):
        idx = idx[bh[idx]]
        if len(idx) == 0:
            return dict(tb=np.nan, p95=np.nan, max=np.nan)
        return dict(tb=float(np.average(J[idx], weights=bt.area[idx])), p95=float(np.percentile(J[idx], 95)),
                    max=float(J[idx].max()))

    dat = np.isin(mat, [1, 12, 3, 20])
    e_vao = ke(up)
    e_vao = e_vao[dat[e_vao]]
    e_ra = ke(bt.seep_on) if len(bt.seep_on) else np.array([], int)
    e_ra = e_ra[dat[e_ra]]
    nut_dem = np.unique(tri[mat == 4])
    e_tiep = ke(nut_dem)
    e_tiep = e_tiep[dat[e_tiep]]
    xc, zc = xy[tri, 0].mean(1), xy[tri, 1].mean(1)
    trong_ck = (mat == 1) & (xc > CK[0]) & (xc < CK[2]) & (zc < np.interp(xc, [CK[0], CK[2]], [CK[1], CK[3]]))
    e_ck = np.where(trong_ck)[0]
    q_in = bt.luu_luong(up)
    q_out = -bt.luu_luong(bt.seep_on)
    return bt, dict(Zw=Zw, nen=nen, q=q_in, q_out=q_out, J_vao=tk(e_vao), J_ra=tk(e_ra), J_tiep=tk(e_tiep),
                    J_ck=tk(e_ck), so_nut=len(xy), so_pt=len(tri), niter=bt.niter,
                    seep_cao=float(xy[bt.seep_on, 1].max()) if len(bt.seep_on) else None)


def ve(ket, ten_file):
    mau = {20: '#C9C3E6', 1: '#F7E27C', 12: '#A8D98A', 3: '#A8D98A', 4: '#E879C8'}
    fig, axs = plt.subplots(len(ket), 1, figsize=(7.4, 3.0 * len(ket)))
    for ax, (bt, kq, nhan) in zip(np.atleast_1d(axs), ket):
        xy, tri, mat = bt.xy, bt.tri, bt.mat
        for m, c in mau.items():
            sel = mat == m
            if not sel.any():
                continue
            ax.tripcolor(xy[:, 0], xy[:, 1], tri[sel], facecolors=np.ones(sel.sum()), cmap=matplotlib.colors.ListedColormap([c]),
                         edgecolors='none')
        Tg = mtri.Triangulation(xy[:, 0], xy[:, 1], tri)
        lv = np.arange(64.0, kq['Zw'] + 0.01, 0.5)
        cs = ax.tricontour(Tg, bt.H, levels=lv, colors='#1F4E9C', linewidths=0.45)
        ax.clabel(cs, cs.levels[::2], fmt=lambda v: ('%.1f' % v).replace('.', ','), fontsize=5.5)
        ax.tricontour(Tg, bt.H - xy[:, 1], levels=[0.0], colors='#D62728', linewidths=1.4)
        ax.plot(*zip(*MAT), color='k', lw=0.8)
        xw = 27.4 - 3.5 * (68.0 - kq['Zw']) if kq['Zw'] <= 68 else 30.4 + 3.0 * (kq['Zw'] - 68.0)
        xn = np.linspace(X0, xw, 200)
        ax.fill_between(xn, kq['Zw'], z_mat(xn), color='#BFE3FF', alpha=0.6, lw=0)
        ax.plot([X0, xw], [kq['Zw']] * 2, color='#2A78D6', lw=1.0)
        if len(bt.seep_on):
            ax.plot(xy[bt.seep_on, 0], xy[bt.seep_on, 1], 'o', color='#2CA02C', ms=2.0)
        mq, ex = ('%.2e' % kq['q']).split('e')
        ax.set_title('%s – MN hồ +%s: q = %s·10$^{%d}$ m³/s/m' % (nhan, ('%.2f' % kq['Zw']).replace('.', ','),
                                                                   mq.replace('.', ','), int(ex)), fontsize=8.6, loc='left')
        ax.set_xlim(X0, X1)
        ax.set_ylim(Z_DAY, 82)
        ax.set_aspect('equal')
        ax.tick_params(labelsize=7)
        ax.set_ylabel('Cao độ (m)', fontsize=7.5)
    np.atleast_1d(axs)[-1].set_xlabel('Khoảng cách (m)', fontsize=7.5)
    axs_ = np.atleast_1d(axs)[0]
    axs_.plot([], [], color='#1F4E9C', lw=0.6, label='Đường đẳng cột nước (0,5 m)')
    axs_.plot([], [], color='#D62728', lw=1.4, label='Đường bão hòa')
    axs_.plot([], [], 'o', color='#2CA02C', ms=3, label='Mặt thấm ra')
    axs_.legend(loc='upper right', fontsize=6.5, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(ten_file, dpi=220, bbox_inches='tight', facecolor='white')
    plt.close(fig)


if __name__ == '__main__':
    mesh = luoi()
    print('nút %d, phần tử %d' % (len(mesh[0]), len(mesh[1])))
    KQ = {}
    ket_ve = []
    K_TAC = {4: S.DAT['DD-II']['k']}
    for ten, Zw, kt in (('TH1_MNDBT', S.MNDBT, None), ('TH2_MNLTK', S.MNLTK, None), ('TH3_MNLKT', S.MNLKT, None),
                        ('TH4_MNDBT_tac', S.MNDBT, K_TAC), ('TV_74.70', 74.70, None), ('TV_77.43', 77.43, None),
                        ('TV_78.25', 78.25, None), ('TV_tac', 74.70, K_TAC)):
        bt, kq = chay(Zw, mesh=mesh, k_them=kt)
        KQ[ten] = kq
        tv = TU_VAN_T9.get('tac' if ten == 'TV_tac' else Zw) if ten.startswith('TV') else None
        print('%-14s Zw %.2f q %.3e%s | Jvao %.3f Jra %.3f Jtiep %.3f Jck %.3f | lặp %d'
              % (ten, Zw, kq['q'], (' (TV %.3e, lệch %+.0f%%)' % (tv, 100 * (kq['q'] / tv - 1))) if tv else '',
                 kq['J_vao']['p95'], kq['J_ra']['p95'], kq['J_tiep']['p95'], kq['J_ck']['p95'], kq['niter']))
        if ten in ('TH1_MNDBT', 'TH2_MNLTK'):
            ket_ve.append((bt, kq, {'TH1_MNDBT': 'TH1 – MNDBT', 'TH2_MNLTK': 'TH2 – MNLTK'}[ten]))
    # độ nhạy: nền lớp 5, lớp I; lõi theo TKCS (ĐĐ-I)
    for ten, nen, kt in (('nen_L5', 'L5', None), ('nen_LI', 'LI', None), ('loi_DDI', NEN_CHINH, {1: S.DAT['DD-I']['k']})):
        for Zw in (S.MNDBT, S.MNLTK):
            bt, kq = chay(Zw, nen=nen, mesh=mesh, k_them=kt)
            KQ['%s_%.2f' % (ten, Zw)] = kq
            print('%-8s Zw %.2f q %.3e Jvao %.3f Jra %.3f Jtiep %.3f Jck %.3f' % (
                ten, Zw, kq['q'], kq['J_vao']['p95'], kq['J_ra']['p95'], kq['J_tiep']['p95'], kq['J_ck']['p95']))
    ve(ket_ve, os.path.join(DIR, '..', 'TAI_VE', 'Tham_MatCatSuonDoi.png'))
    json.dump(KQ, open(os.path.join(DIR, 'ket_qua', 'tham_suon_doi.json'), 'w'), ensure_ascii=False, indent=1)
    print('Đã ghi ket_qua/tham_suon_doi.json, ../TAI_VE/Tham_MatCatSuonDoi.png')
