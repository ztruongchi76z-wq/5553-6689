# -*- coding: utf-8 -*-
"""
ỔN ĐỊNH TRƯỢT MÁI ĐẬP – MẶT CẮT SƯỜN ĐỒI
(1) 7 trường hợp như Bảng 2-24 BCNCKT (Bảng 2.21 bản thuyết minh SV) – hiện trạng (chương 2).
(2) 12 trường hợp như Bảng 5.2 của mặt cắt lòng sông (chay_tham_on_dinh.py, phan_tich_bo_sung.py) – hiện trạng và
    sau sửa chữa (mặt đỉnh nâng +78,50 → +78,85 trong bề rộng đỉnh, lớp đắp bù K ≥ 0,97 – tham_suon_doi.luoi):
    động đất kh = 0,05; mưa kéo dài – tầng bão hòa song song mái hạ lưu dày 3 m (hiện trạng) / 1 m (sau sửa chữa);
    thi công – tải xe máy 20 kPa trên toàn bề rộng đỉnh, [K] = 1,24 (chỉ sau sửa chữa) – chương 5.
- Bishop đơn giản, mặt trượt trụ tròn, tìm tâm theo lưới tâm × bán kính (on_dinh.py – như SLOPE/W "Grid & Radius").
- Áp lực nước lỗ rỗng lấy từ lời giải thấm phần tử hữu hạn của chính mặt cắt sườn đồi (tham_suon_doi.py).
- Chỉ tiêu tính toán theo TTGH I (c_I, φ_I – TKCS Bảng 2-1, 2-3, 2-4): lõi ĐĐ-I; khối thượng lưu, hạ lưu ĐĐ-II;
  đệm cát; lớp đắp bù mặt đập – đất đắp mới K ≥ 0,97; nền "Lop 6" – thấm và cường độ như lớp 4 (hiệu chỉnh theo kết
  quả thấm Mặt cắt V – T9 của tư vấn, xem tham_suon_doi.py). Độ nhạy: nền lớp 5.
- Mặt đất hạ lưu +68,3 cao hơn mực nước hạ lưu ở lòng sông nên các trường hợp "HL có nước" thành HL không có nước.
- Rút nước nhanh: áp lực lỗ rỗng giữ như thấm ổn định ở mực nước cao (đường bão hòa chưa kịp hạ – thiên an toàn),
  tải nước hồ lấy theo mực nước thấp.
  python3 on_dinh_suon_doi.py
"""
import json
import os
import time

import matplotlib
matplotlib.use('Agg')
import matplotlib.lines
import matplotlib.patches
import matplotlib.pyplot as plt
import numpy as np

import on_dinh as O
import phan_tich_bo_sung as PT
import so_lieu as S
import tham_suon_doi as TS

DIR = os.path.dirname(os.path.abspath(__file__))
plt.rcParams.update({'font.family': 'Liberation Serif'})
MNDBT, MNLTK, MNLKT, MNC = S.MNDBT, S.MNLTK, S.MNLKT, S.MNC   # mực nước lũ theo kết quả điều tiết của SV
K_TAC = {4: S.DAT['DD-II']['k']}                              # đệm tiêu nước bị tắc: thấm như đất đắp
TH = [  # mã, mô tả, mái (+1 hạ lưu, −1 thượng lưu), MN tính thấm, thay k, MN hồ trên mái TL, [K]
    ('TH1', 'TL MNDBT; HL không có nước', +1, MNDBT, None, None, 1.30),
    ('TH2', 'TL MNLTK; HL không có nước', +1, MNLTK, None, None, 1.30),
    ('TH3', 'TL MNLKT; HL không có nước', +1, MNLKT, None, None, 1.13),
    ('TH4', 'TL MNDBT; đệm tiêu nước bị tắc', +1, MNDBT, K_TAC, None, 1.13),
    ('TH5', 'Rút từ MNLTK xuống ngưỡng tràn +74,70', -1, MNLTK, None, 74.70, 1.25),
    ('TH6', 'Rút từ MNLKT xuống ngưỡng tràn +74,70', -1, MNLKT, None, 74.70, 1.13),
    ('TH7', 'Rút từ MNDBT xuống MNC', -1, MNDBT, None, MNC, 1.13),
]
# bộ 12 trường hợp của chương 5 – cùng mã, tổ hợp và [K] với Bảng 5.2 mặt cắt lòng sông
KH = 0.05                          # hệ số động đất ngang giả tĩnh (như chay_tham_on_dinh.py)
D_MUA = {False: 3.0, True: 1.0}    # chiều sâu tầng bão hòa do mưa kéo dài: hiện trạng / sau sửa chữa (giả thiết)
Q_TC = 20.0                        # tải trọng xe máy thi công trên toàn bề rộng đỉnh (kPa)
K_CP = {'CB': 1.30, 'DB': 1.17, 'TC': 1.24}
TH_C5 = [  # mã, mô tả, mái, MN tính thấm, thay k, MN hồ trên mái TL, tổ hợp, kh, mưa, tải thi công
    ('TH1', 'MNDBT, thấm ổn định', +1, MNDBT, None, None, 'CB', 0.0, False, 0.0),
    ('TH2', 'MNLTK, thấm ổn định', +1, MNLTK, None, None, 'CB', 0.0, False, 0.0),
    ('TH3', 'MNLKT, thấm ổn định', +1, MNLKT, None, None, 'DB', 0.0, False, 0.0),
    ('TH4', 'MNDBT, thiết bị thoát nước bị tắc', +1, MNDBT, K_TAC, None, 'DB', 0.0, False, 0.0),
    ('TH5', 'MNDBT + động đất', +1, MNDBT, None, None, 'DB', KH, False, 0.0),
    ('TH6', 'MNDBT + mưa kéo dài', +1, MNDBT, None, None, 'DB', 0.0, True, 0.0),
    ('TH7', 'MNDBT, thấm ổn định (mái TL)', -1, MNDBT, None, MNDBT, 'CB', 0.0, False, 0.0),
    ('TH8', 'Hạ MN từ MNDBT xuống MNC (đường bão hòa chưa kịp hạ)', -1, MNDBT, None, MNC, 'CB', 0.0, False, 0.0),
    ('TH9', 'Rút nhanh từ MNLTK xuống MNDBT', -1, MNLTK, None, MNDBT, 'DB', 0.0, False, 0.0),
    ('TH10', 'MNDBT + động đất (mái TL)', -1, MNDBT, None, MNDBT, 'DB', KH, False, 0.0),
    ('TH11', 'Thi công: MNDBT + tải thi công 20 kPa trên đỉnh', +1, MNDBT, None, None, 'TC', 0.0, False, Q_TC),
    ('TH12', 'Thi công: MNDBT + tải thi công 20 kPa trên đỉnh', -1, MNDBT, None, MNDBT, 'TC', 0.0, False, Q_TC),
]


class MatCatSuon:
    """Bộ tiếp hợp hình học cho on_dinh.LuoiRaster / OnDinh; sau_sc: mặt đỉnh +78,85 (lớp đắp bù)."""
    x_trai, x_phai, Zbot, Zday = TS.X0, TS.X1, TS.Z_DAY, 64.6

    def __init__(self, sau_sc=False):
        self.sau_sc = sau_sc
        self.Zc = TS.DINH_SC if sau_sc else TS.DINH

    def mat_tren(self, x):
        return TS.z_mat(x, self.sau_sc)


def vat_lieu(nen):
    ma = {1: 'DD-I', 12: 'DD-II', 3: 'DD-II', 4: 'DC', 20: nen, 21: 'DAP_MOI'}
    return {k: (S.DAT[v]['g'], S.DAT[v]['gs'], S.DAT[v]['cI'], S.DAT[v]['phiI']) for k, v in ma.items()}


def tim_cung(od, chieu):
    if chieu > 0:
        return od.tim((60.0, 112.0), (79.0, 135.0), nx=26, nz=24, nr=24, x_gioi_han=(40.0, 108.5))
    return od.tim((-2.0, 70.0), (79.0, 135.0), nx=28, nz=24, nr=24, x_gioi_han=(1.5, 75.0))


def mua_keo_dai(ras, d_bh):
    """Tầng bão hòa do mưa kéo dài dày d_bh song song mái hạ lưu (m = 3, từ mép đỉnh hạ lưu đến chân mái), dòng chảy
    song song mái -> cột nước áp lực (d_bh − độ sâu)·cos²β – như chay_tham_on_dinh.mua_keo_dai."""
    X, Z = ras.X, ras.Z
    depth = ras.mc.mat_tren(X[:, 0])[:, None] - Z
    zone = (X > TS.X_DD) & (X < 96.3) & (depth >= 0) & (depth <= d_bh) & (ras.mat > 0)
    psi_mua = (d_bh - depth) * np.cos(np.arctan(1 / 3.0)) ** 2
    ras.psi = np.where(zone, np.maximum(np.nan_to_num(ras.psi, nan=-1), psi_mua), ras.psi)
    return ras


def tinh(nen=TS.NEN_CHINH, mesh=None, in_ra=True):
    mesh = mesh if mesh is not None else TS.luoi()
    mc = MatCatSuon()
    bts, kq = {}, []
    for ma, mota, chieu, Zt, kthem, Zng, Kcp in TH:
        t0 = time.time()
        key = (Zt, kthem is not None)
        if key not in bts:
            bts[key] = TS.chay(Zt, nen=nen, mesh=mesh, k_them=kthem)[0]
        bt = bts[key]
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        od = O.OnDinh(ras, vat_lieu(nen), Zw=Zng, chieu=chieu)
        K, (xc, zc, R) = tim_cung(od, chieu)
        _, ch = od.K_vong(xc, zc, R, tra_chi_tiet=True)
        kq.append(dict(ma=ma, mo_ta=mota, mai='Hạ lưu' if chieu > 0 else 'Thượng lưu', Zw_tham=Zt, Zw_ngoai=Zng,
                       K=float(K), K_cp=Kcp, dat=bool(K >= Kcp), xc=float(xc), zc=float(zc), R=float(R),
                       xa=float(ch['xa']), xb=float(ch['xb'])))
        if in_ra:
            print('%s %-40s K = %.3f [K] = %.2f  tâm (%.1f; %.1f) R = %.1f  vào–ra %.1f ÷ %.1f  %.0f s'
                  % (ma, mota, K, Kcp, xc, zc, R, ch['xa'], ch['xb'], time.time() - t0))
        kq[-1]['_bt'], kq[-1]['_ras'] = bt, ras
    return kq


def tinh_c5(sau_sc, nen=TS.NEN_CHINH, mesh=None, in_ra=True):
    """12 trường hợp của chương 5; TH11, TH12 (thi công) chỉ tính cho mặt cắt sau sửa chữa."""
    mesh = mesh if mesh is not None else TS.luoi(sau_sc=sau_sc)
    mc = MatCatSuon(sau_sc)
    vl = vat_lieu(nen)
    bts, kq = {}, []
    for ma, mota, chieu, Zt, kthem, Zng, toh, kh, mua, q_tc in TH_C5:
        if q_tc and not sau_sc:
            continue
        t0 = time.time()
        key = (Zt, kthem is not None)
        if key not in bts:
            bts[key] = TS.chay(Zt, nen=nen, mesh=mesh, k_them=kthem, sau_sc=sau_sc)[0]
        bt = bts[key]
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        if mua:
            ras = mua_keo_dai(ras, D_MUA[sau_sc])
        if q_tc:
            od = PT.OnDinhPhuTai(ras, vl, Zw=Zng, kh=kh, chieu=chieu, q_s=q_tc, x_qs=(TS.X_DU, TS.X_DD))
        else:
            od = O.OnDinh(ras, vl, Zw=Zng, kh=kh, chieu=chieu)
        K, (xc, zc, R) = tim_cung(od, chieu)
        _, ch = od.K_vong(xc, zc, R, tra_chi_tiet=True)
        kq.append(dict(ma=ma, mo_ta=mota, mai='HL' if chieu > 0 else 'TL', to_hop=toh, Zw_tham=Zt, Zw_ngoai=Zng,
                       kh=kh, d_mua=D_MUA[sau_sc] if mua else 0.0, q_tc=q_tc, K=float(K), K_cp=K_CP[toh],
                       dat=bool(K >= K_CP[toh]), xc=float(xc), zc=float(zc), R=float(R), xa=float(ch['xa']),
                       xb=float(ch['xb'])))
        if in_ra:
            print('%s %s %-52s K = %.3f [K] = %.2f  tâm (%.1f; %.1f) R = %.1f  vào–ra %.1f ÷ %.1f  %.0f s'
                  % ('SC' if sau_sc else 'HT', ma, mota, K, K_CP[toh], xc, zc, R, ch['xa'], ch['xb'], time.time() - t0))
        kq[-1]['_bt'], kq[-1]['_ras'] = bt, ras
    return kq


def ve(kq, ten_file, sau_sc=False, chu_thich=None, so_cot=2):
    """Lưới hình con: mỗi trường hợp một hình (vật liệu, đường bão hòa, nước hồ, cung trượt nguy hiểm);
    chu_thich: thêm một ô chữ ở cuối (ô trống còn lại)."""
    n = len(kq) + (1 if chu_thich else 0)
    so_hang = int(np.ceil(n / so_cot))
    fig, axs = plt.subplots(so_hang, so_cot, figsize=(7.4, 1.35 * so_hang + 0.0))
    mau = {20: '#C9C3E6', 1: '#F7E27C', 12: '#A8D98A', 3: '#A8D98A', 4: '#E879C8', 21: '#E3A35C'}
    mat_tren = [(x, float(TS.z_mat(x, sau_sc))) for x, _ in TS.MAT]
    if sau_sc:
        mat_tren = sorted(mat_tren + [(TS.X_DU - 1e-3, TS.DINH), (TS.X_DD + 1e-3, TS.DINH)])
    for ax, r in zip(axs.ravel(), kq):
        bt, ras = r['_bt'], r['_ras']
        xy, tri, mat = bt.xy, bt.tri, bt.mat
        for m, c in mau.items():
            sel = mat == m
            if sel.any():
                ax.tripcolor(xy[:, 0], xy[:, 1], tri[sel], facecolors=np.ones(sel.sum()),
                             cmap=matplotlib.colors.ListedColormap([c]), edgecolors='none')
        psi = np.ma.masked_where(ras.mat == 0, np.nan_to_num(ras.psi, nan=-1))   # bỏ đường giả dọc mặt đất
        ax.contour(ras.X, ras.Z, psi, levels=[0.0], colors='#D62728', linewidths=1.0)
        if r['Zw_ngoai'] is not None:
            xn = np.linspace(TS.X0, 62.0, 300)
            zt = TS.z_mat(xn, sau_sc)
            ax.fill_between(xn, r['Zw_ngoai'], zt, where=zt < r['Zw_ngoai'], color='#BFE3FF', alpha=0.7, lw=0)
        ax.plot(*zip(*mat_tren), color='k', lw=0.7)
        th = np.linspace(0, 2 * np.pi, 600)
        xx, zz = r['xc'] + r['R'] * np.cos(th), r['zc'] + r['R'] * np.sin(th)
        m = (zz <= TS.z_mat(xx, sau_sc)) & (xx >= r['xa'] - 0.3) & (xx <= r['xb'] + 0.3)
        ax.plot(np.where(m, xx, np.nan), np.where(m, zz, np.nan), color='k', lw=1.4)
        mai = {'HL': 'hạ lưu', 'TL': 'thượng lưu'}.get(r['mai'], r['mai'].lower())
        ax.set_title('%s – %s: K = %s' % (r['ma'], mai, ('%.3f' % r['K']).replace('.', ',')), fontsize=8.2, loc='left')
        ax.set_xlim(TS.X0, TS.X1)
        ax.set_ylim(55, 82)
        ax.set_aspect('equal')
        ax.tick_params(labelsize=6.5)
    for ax in axs.ravel()[len(kq):]:
        ax.axis('off')
    if chu_thich:
        axs.ravel()[-1].text(0.02, 0.95, chu_thich, transform=axs.ravel()[-1].transAxes, va='top', fontsize=8)
    else:   # chú giải chung ở dưới hình
        P = matplotlib.patches.Patch
        tay = [matplotlib.lines.Line2D([], [], color='k', lw=1.4, label='Cung trượt nguy hiểm'),
               matplotlib.lines.Line2D([], [], color='#D62728', lw=1.0, label='Đường bão hòa'),
               P(color='#BFE3FF', label='Nước hồ'), P(color='#F7E27C', label='Lõi ĐĐ-I'),
               P(color='#A8D98A', label='Khối TL, HL ĐĐ-II'), P(color='#E879C8', label='Đệm tiêu nước'),
               P(color='#C9C3E6', label='Nền')] + ([P(color='#E3A35C', label='Đắp nâng đỉnh')] if sau_sc else [])
        fig.legend(handles=tay, loc='lower center', ncol=4, fontsize=7.5, frameon=False, bbox_to_anchor=(0.5, -0.045))
    fig.tight_layout()
    fig.savefig(ten_file, dpi=220, bbox_inches='tight', facecolor='white')
    plt.close(fig)


def sach(kq):
    return [{k: v for k, v in r.items() if not k.startswith('_')} for r in kq]


if __name__ == '__main__':
    mesh = TS.luoi()
    KQ = {}
    kq = tinh(mesh=mesh)
    ve(kq, os.path.join(DIR, '..', 'TAI_VE', 'OnDinh_MatCatSuonDoi.png'),
       chu_thich='Mặt cắt sườn đồi – ổn định trượt (Bishop)\n— cung trượt nguy hiểm nhất\n— đường bão hòa (đỏ)\n'
                 'Nền "Lop 6": như lớp 4 (hiệu chỉnh theo T9)\nRút nước nhanh: áp lực lỗ rỗng\ngiữ như ở mực nước cao')
    KQ['nen_LI'] = sach(kq)
    print('--- độ nhạy: nền lớp 5')
    KQ['nen_L5'] = sach(tinh(nen='L5', mesh=mesh))
    print('--- chương 5: 12 trường hợp, hiện trạng và sau sửa chữa')
    KQ['chuong5'] = {'HT': sach(tinh_c5(False, mesh=mesh))}
    kq_sc = tinh_c5(True)
    KQ['chuong5']['SC'] = sach(kq_sc)
    ve(kq_sc, os.path.join(DIR, '..', 'TAI_VE', 'OnDinh_SuonDoi_SauSC.png'), sau_sc=True)
    json.dump(KQ, open(os.path.join(DIR, 'ket_qua', 'on_dinh_suon_doi.json'), 'w'), ensure_ascii=False, indent=1)
    print('Đã ghi ket_qua/on_dinh_suon_doi.json, ../TAI_VE/OnDinh_MatCatSuonDoi.png, ../TAI_VE/OnDinh_SuonDoi_SauSC.png')
