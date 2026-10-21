# -*- coding: utf-8 -*-
"""
ỔN ĐỊNH TRƯỢT MÁI ĐẬP – MẶT CẮT SƯỜN ĐỒI, 7 TRƯỜNG HỢP NHƯ BẢNG 2-24 BCNCKT (Bảng 2.21 bản thuyết minh SV)
- Bishop đơn giản, mặt trượt trụ tròn, tìm tâm theo lưới tâm × bán kính (on_dinh.py – như SLOPE/W "Grid & Radius").
- Áp lực nước lỗ rỗng lấy từ lời giải thấm phần tử hữu hạn của chính mặt cắt sườn đồi (tham_suon_doi.py).
- Chỉ tiêu tính toán theo TTGH I (c_I, φ_I – TKCS Bảng 2-1, 2-3, 2-4): lõi ĐĐ-I; khối thượng lưu, hạ lưu ĐĐ-II;
  đệm cát; nền "Lop 6" = lớp I (đá phong hóa mạnh). Độ nhạy: nền lớp 5.
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
import matplotlib.pyplot as plt
import numpy as np

import on_dinh as O
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


class MatCatSuon:
    """Bộ tiếp hợp hình học cho on_dinh.LuoiRaster / OnDinh."""
    x_trai, x_phai, Zbot, Zc, Zday = TS.X0, TS.X1, TS.Z_DAY, TS.DINH, 64.6

    def mat_tren(self, x):
        return TS.z_mat(x)


def vat_lieu(nen):
    ma = {1: 'DD-I', 12: 'DD-II', 3: 'DD-II', 4: 'DC', 20: nen}
    return {k: (S.DAT[v]['g'], S.DAT[v]['gs'], S.DAT[v]['cI'], S.DAT[v]['phiI']) for k, v in ma.items()}


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
        if chieu > 0:
            best = od.tim((60.0, 112.0), (79.0, 135.0), nx=26, nz=24, nr=24, x_gioi_han=(40.0, 108.5))
        else:
            best = od.tim((-2.0, 70.0), (79.0, 135.0), nx=28, nz=24, nr=24, x_gioi_han=(1.5, 75.0))
        K, (xc, zc, R) = best
        _, ch = od.K_vong(xc, zc, R, tra_chi_tiet=True)
        kq.append(dict(ma=ma, mo_ta=mota, mai='Hạ lưu' if chieu > 0 else 'Thượng lưu', Zw_tham=Zt, Zw_ngoai=Zng,
                       K=float(K), K_cp=Kcp, dat=bool(K >= Kcp), xc=float(xc), zc=float(zc), R=float(R),
                       xa=float(ch['xa']), xb=float(ch['xb'])))
        if in_ra:
            print('%s %-40s K = %.3f [K] = %.2f  tâm (%.1f; %.1f) R = %.1f  vào–ra %.1f ÷ %.1f  %.0f s'
                  % (ma, mota, K, Kcp, xc, zc, R, ch['xa'], ch['xb'], time.time() - t0))
        kq[-1]['_bt'], kq[-1]['_ras'] = bt, ras
    return kq


def ve(kq, ten_file):
    fig, axs = plt.subplots(4, 2, figsize=(7.4, 5.4))
    mau = {20: '#C9C3E6', 1: '#F7E27C', 12: '#A8D98A', 3: '#A8D98A', 4: '#E879C8'}
    for ax, r in zip(axs.ravel(), kq):
        bt, ras = r['_bt'], r['_ras']
        xy, tri, mat = bt.xy, bt.tri, bt.mat
        for m, c in mau.items():
            sel = mat == m
            if sel.any():
                ax.tripcolor(xy[:, 0], xy[:, 1], tri[sel], facecolors=np.ones(sel.sum()),
                             cmap=matplotlib.colors.ListedColormap([c]), edgecolors='none')
        ax.contour(ras.X, ras.Z, np.nan_to_num(ras.psi, nan=-1), levels=[0.0], colors='#D62728', linewidths=1.0)
        if r['Zw_ngoai'] is not None:
            xn = np.linspace(TS.X0, 62.0, 300)
            zt = TS.z_mat(xn)
            ax.fill_between(xn, r['Zw_ngoai'], zt, where=zt < r['Zw_ngoai'], color='#BFE3FF', alpha=0.7, lw=0)
        ax.plot(*zip(*TS.MAT), color='k', lw=0.7)
        th = np.linspace(0, 2 * np.pi, 600)
        xx, zz = r['xc'] + r['R'] * np.cos(th), r['zc'] + r['R'] * np.sin(th)
        m = (zz <= TS.z_mat(xx)) & (xx >= r['xa'] - 0.3) & (xx <= r['xb'] + 0.3)
        ax.plot(np.where(m, xx, np.nan), np.where(m, zz, np.nan), color='k', lw=1.4)
        ax.set_title('%s – %s: K = %s' % (r['ma'], r['mai'].lower(), ('%.3f' % r['K']).replace('.', ',')),
                     fontsize=8.2, loc='left')
        ax.set_xlim(TS.X0, TS.X1)
        ax.set_ylim(55, 82)
        ax.set_aspect('equal')
        ax.tick_params(labelsize=6.5)
    ax = axs.ravel()[-1]
    ax.axis('off')
    ax.text(0.02, 0.95, 'Mặt cắt sườn đồi – ổn định trượt (Bishop)\n'
                        '— cung trượt nguy hiểm nhất\n— đường bão hòa (đỏ)\nNền "Lop 6" = lớp I (đá phong hóa mạnh)\n'
                        'Rút nước nhanh: áp lực lỗ rỗng\ngiữ như ở mực nước cao',
            transform=ax.transAxes, va='top', fontsize=8)
    fig.tight_layout()
    fig.savefig(ten_file, dpi=220, bbox_inches='tight', facecolor='white')
    plt.close(fig)


if __name__ == '__main__':
    mesh = TS.luoi()
    KQ = {}
    kq = tinh(mesh=mesh)
    ve(kq, os.path.join(DIR, '..', 'TAI_VE', 'OnDinh_MatCatSuonDoi.png'))
    KQ['nen_LI'] = [{k: v for k, v in r.items() if not k.startswith('_')} for r in kq]
    print('--- độ nhạy: nền lớp 5')
    KQ['nen_L5'] = [{k: v for k, v in r.items() if not k.startswith('_')} for r in tinh(nen='L5', mesh=mesh)]
    json.dump(KQ, open(os.path.join(DIR, 'ket_qua', 'on_dinh_suon_doi.json'), 'w'), ensure_ascii=False, indent=1)
    print('Đã ghi ket_qua/on_dinh_suon_doi.json, ../TAI_VE/OnDinh_MatCatSuonDoi.png')
