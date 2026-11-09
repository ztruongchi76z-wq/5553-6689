# -*- coding: utf-8 -*-
"""
HÌNH MINH HỌA (SV nhờ 09/11/2026 – "chưa hình dung được dốc nước"): mặt cắt dọc theo tim tràn chính và mặt cắt ngang dốc nước
tại 3 vị trí. Số liệu: so_lieu.TRAN_CHINH (ngưỡng, bậc hạ, dốc), Bảng 6-7 hồ sơ kiểm định (chiều cao tường x = 0 ÷ 40 m), kết quả
kiem_tra_tran_chinh.tinh() (mặt nước có kể hàm khí, mũi phun – hố xói lũ kiểm tra), bình đồ hiện trạng (TP5, TP7, TP11,
+70,24/+70,25 ở hai đầu tường cuối dốc), ảnh SV chụp (cầu giao thông ≈ +78,5). Các phần chưa có số liệu (trụ pin, cửa van, mũi
phun, bề dày tường, bản đáy) vẽ sơ họa – ghi rõ trên hình.
  python3 ve_mat_cat_doc_nuoc.py      → ../TAI_VE/MinhHoa_MatCatDoc_TranChinh.png, ../TAI_VE/MinhHoa_MatCatNgang_DocNuoc.png
"""
import math
import os

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Polygon, Rectangle

import kiem_tra_tran_chinh as K
import thuy_luc_tran_chinh as T

DIR = os.path.dirname(os.path.abspath(__file__))
RA_DOC = os.path.join(DIR, '..', 'TAI_VE', 'MinhHoa_MatCatDoc_TranChinh.png')
RA_NGANG = os.path.join(DIR, '..', 'TAI_VE', 'MinhHoa_MatCatNgang_DocNuoc.png')
TC = T.TC
# màu: hai chuỗi (đã kiểm tra bằng validate_palette.js – đạt), còn lại trung tính
C_TK, C_KT = '#2a78d6', '#eb6834'
C_NET, C_PHU, C_BT, C_NEN, C_NUOC, C_DAT = '#3d3c39', '#8a8984', '#d9d8d2', '#efeee9', '#dbe9f8', '#e9dfcc'
C_CANH = '#c0392b'                                    # chỉ dùng cho điểm "cần kiểm tra" (kèm chữ)
s2 = lambda x: ('%.2f' % x).replace('.', ',')
s1 = lambda x: ('%.1f' % x).replace('.', ',')

Z_NG = TC['Z_nguong']                                 # +74,70
Z_SAN = Z_NG - TC['P1']                               # +74,00
Z_DD = Z_NG - TC['P2']                                # +73,30 – đầu dốc
I, L = TC['i_doc'], TC['L_doc']
Z_CUOI = Z_DD - I * L                                 # +68,50 theo i = 8 %, L = 60 m
Z_CAU_TREN, Z_CAU_DUOI = 78.55, 78.00                 # cầu giao thông – ước lượng từ ảnh (sơ họa)
T_DAY, T_TUONG = 0.5, 0.5                             # bản đáy, tường – sơ họa (chưa có số liệu)


def tuong_day(x):
    """Cao trình đỉnh tường bên theo Bảng 6-7 (x = 0 ÷ 40 m), sau đó giả thiết 2,5 m."""
    xs = [0, 5, 10, 15, 20, 25, 30, 35, 40]
    if x <= 40:
        return Z_DD - I * x + float(np.interp(x, xs, TC['h_tuong']))
    return Z_DD - I * x + TC['h_tuong'][-1]


def mui_phun_cung(z0, z_moi, alpha):
    """Cung tròn mũi phun (sơ họa): tiếp tuyến với đáy dốc tại x = L, lên tới mũi ở cao trình z_moi, góc hắt alpha."""
    th, al = math.atan(I), math.radians(alpha)
    R = (z_moi - z0) / (math.cos(th) - math.cos(al))
    ts = np.linspace(-th, al, 30)
    xc, zc = L + R * math.sin(th), z0 + R * math.cos(th)          # tâm cung
    return xc + R * np.sin(ts), zc - R * np.cos(ts), R


def mat_cat_doc(kq):
    plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 9})
    fig, ax = plt.subplots(figsize=(11.5, 5.6))
    doc = kq['doc']
    xs = np.array([r['x'] for r in doc], dtype=float)
    day = Z_DD - I * xs
    mp = kq['kt']['mp']
    xb, zb, R = mui_phun_cung(Z_CUOI, K.Z_MUI, K.ALPHA_MP)
    x_moi = xb[-1]
    # ---- nền đá, kết cấu bê tông (sân trước, ngưỡng, bản đáy dốc, mũi phun)
    x_san0 = -16.0
    # mái hạ lưu ngưỡng: SV đo trên bình đồ (lớp "mai doc") dài 3,0 m theo phương dòng chảy (≈ 1:2,1 với P₂ = 1,40 m);
    # dãy trụ pin cách chân ngưỡng ≈ 4,0 m
    L_MAI_HL = 3.0
    x_ng_dinh1 = -L_MAI_HL
    x_ng_dinh0 = x_ng_dinh1 - TC['delta']
    x_ng_tl = x_ng_dinh0 - TC['P1'] * TC['S_TL']
    tren = [(x_san0, Z_SAN), (x_ng_tl, Z_SAN), (x_ng_dinh0, Z_NG), (x_ng_dinh1, Z_NG), (0, Z_DD), (L, Z_CUOI)] + \
           list(zip(xb[1:], zb[1:]))
    z_day_hl = mp['Z_day']
    duoi = [(x_moi, z_day_hl), (L - 2.0, z_day_hl), (L - 2.0, Z_DD - I * (L - 2.0) - T_DAY), (0, Z_DD - T_DAY - 0.4),
            (x_ng_tl, Z_SAN - T_DAY - 0.6), (x_san0, Z_SAN - T_DAY)]
    ax.add_patch(Polygon(tren + duoi, closed=True, fc=C_BT, ec=C_NET, lw=1.0, zorder=3))
    # nền đá phía dưới và lòng dẫn hạ lưu
    x_hl1 = x_moi + mp['x_hl'] + 8
    nen = [(x_san0, Z_SAN - T_DAY), (x_ng_tl, Z_SAN - T_DAY - 0.6), (0, Z_DD - T_DAY - 0.4),
           (L - 2.0, Z_DD - I * (L - 2.0) - T_DAY), (L - 2.0, z_day_hl), (x_hl1, z_day_hl), (x_hl1, 60.5), (x_san0, 60.5)]
    ax.add_patch(Polygon(nen, closed=True, fc=C_NEN, ec='none', zorder=1))
    ax.plot([x_moi, x_moi + mp['x_tl']], [z_day_hl, z_day_hl], color=C_NET, lw=1.0, zorder=4)
    # hố xói lũ kiểm tra (hình thang, mái TL 1:2,2, HL 1:1,5)
    a, b_, hx = x_moi + mp['x_tl'], x_moi + mp['x_hl'], mp['hx']
    ho = [(a, z_day_hl), (a + K.M_HO_TL * hx, z_day_hl - hx), (b_ - K.M_HO_HL * hx, z_day_hl - hx), (b_, z_day_hl)]
    ax.add_patch(Polygon(ho + [(b_, 60.5), (a, 60.5)], closed=True, fc=C_NEN, ec='none', zorder=1))
    ax.add_patch(Polygon(ho, closed=False, fill=False, ec=C_NET, lw=1.0, ls=(0, (4, 2)), zorder=4))
    ax.plot([b_, x_hl1], [z_day_hl, z_day_hl], color=C_NET, lw=1.0, zorder=4)
    # ---- mặt nước hạ lưu (lũ kiểm tra) và luồng phun
    ax.fill_between([x_moi + 1.5, x_hl1], [mp['MNHL']] * 2, [z_day_hl] * 2, color=C_NUOC, lw=0, zorder=2)
    ax.fill_between(np.linspace(a, b_, 60), mp['MNHL'],
                    [max(z_day_hl - min(mp['hx'], (xx - a) / K.M_HO_TL, (b_ - xx) / K.M_HO_HL), z_day_hl - hx) for xx in np.linspace(a, b_, 60)],
                    color=C_NUOC, lw=0, zorder=2)
    ax.plot([x_moi + 1.5, x_hl1], [mp['MNHL']] * 2, color=C_KT, lw=1.0, ls=(0, (6, 3)), zorder=5)
    al = math.radians(K.ALPHA_MP)
    z0j = K.Z_MUI + 0.5 * mp['hH'] * math.cos(al)
    vx, vz = mp['vH'] * math.cos(al), mp['vH'] * math.sin(al)
    t_ve = (vz + math.sqrt(vz ** 2 + 2 * 9.81 * (z0j - mp['MNHL']))) / 9.81
    xj = np.linspace(0, vx * t_ve, 60)
    zj = z0j + xj * math.tan(al) - 9.81 * xj ** 2 / (2 * mp['vH'] ** 2 * math.cos(al) ** 2)
    ax.plot(x_moi + xj, zj, color=C_KT, lw=2.0, zorder=6)
    # ---- mặt nước trên dốc (có kể hàm khí) và trong hồ
    for khoa, c_, mau, nhan, mnl in (('tk', kq['tk'], C_TK, 'Lũ thiết kế P = 1%%, q = %s m³/s' % s2(kq['tk']['Q']), kq['tk']['MNL']),
                                      ('kt', kq['kt'], C_KT, 'Lũ kiểm tra P = 0,2%%, q = %s m³/s' % s2(kq['kt']['Q']), kq['kt']['MNL'])):
        hhk = np.array([r['hhk_' + khoa] for r in doc])
        xx = np.r_[x_san0, x_ng_tl + 0.5, np.linspace(x_ng_dinh0, 0, 12)[1:-1], xs]
        tr = np.linspace(0, 1, 12)[1:-1]
        chuyen = mnl + (Z_DD + hhk[0] - mnl) * (1 - np.cos(np.pi * tr)) / 2
        zz = np.r_[mnl, mnl, chuyen, day + hhk]
        if khoa == 'kt':
            ax.fill_between(xx, zz, np.interp(xx, [p[0] for p in tren[:6]], [p[1] for p in tren[:6]]), color=C_NUOC, lw=0, zorder=2)
        ax.plot(xx, zz, color=mau, lw=2.0, zorder=6, label=nhan + ' – mặt nước có kể hàm khí')
        ax.annotate('%s +%s' % ('MNLTK' if khoa == 'tk' else 'MNLKT', s2(mnl)), xy=(x_san0 + 0.3, mnl), xytext=(0, 3),
                    textcoords='offset points', fontsize=8, color='#0b0b0b')
    # ---- đỉnh tường bên (Bảng 6-7; sau x = 40 m giả thiết) và trụ pin, cầu (sơ họa)
    xt = np.linspace(0, 40, 41)
    ax.plot(xt, [tuong_day(x) for x in xt], color='#0b0b0b', lw=1.2, ls=(0, (1, 1.5)), zorder=5,
            label='Đỉnh tường bên – Bảng 6-7 hồ sơ kiểm định (x = 0 ÷ 40 m)')
    ax.plot([40, L], [tuong_day(40), tuong_day(L)], color=C_PHU, lw=1.2, ls=(0, (4, 3)), zorder=5,
            label='Đỉnh tường bên – giả thiết 2,5 m (hồ sơ không ghi)')
    ax.plot([x_ng_dinh0 - 0.4, 0], [Z_CAU_TREN, tuong_day(0)], color='#0b0b0b', lw=1.2, ls=(0, (1, 1.5)), zorder=5)
    ax.add_patch(Rectangle((x_ng_dinh0 - 1.6, Z_NG), 2.4, Z_CAU_DUOI - Z_NG, fc='#bdbcb6', ec=C_NET, lw=0.8,
                           alpha=0.55, zorder=4))
    ax.add_patch(Rectangle((x_ng_tl - 1.0, Z_CAU_DUOI), x_ng_dinh1 - x_ng_tl + 2.0, Z_CAU_TREN - Z_CAU_DUOI, fc=C_BT, ec=C_NET,
                           lw=0.8, zorder=5))
    # ---- điểm đo bình đồ
    for xp, zp, ten, dx, dy in ((x_ng_dinh0 + 0.25, Z_NG, 'TP5', -16, -16), (0, Z_DD, 'TP7', 4, -16),
                                (x_moi, 68.90, 'TP11', -30, -16)):
        ax.plot(xp, zp, 'o', ms=4.5, mfc='white', mec='#0b0b0b', mew=1.0, zorder=7)
        ax.annotate(ten, xy=(xp, zp), xytext=(dx, dy), textcoords='offset points', fontsize=7.5, color='#0b0b0b',
                    arrowprops=dict(arrowstyle='-', color=C_PHU, lw=0.6))
    ax.plot(L, 70.25, 'X', ms=7, color=C_CANH, zorder=7)
    ax.annotate('+70,24 / +70,25 (bình đồ, hai đầu tường cuối dốc):\nnếu là đỉnh tường thì tường chỉ cao ≈ 1,4 ÷ 1,8 m – cần kiểm tra',
                xy=(L, 70.25), xytext=(L - 33, 74.3), fontsize=7.5, color='#0b0b0b',
                arrowprops=dict(arrowstyle='-', color=C_CANH, lw=0.8))
    # ---- ghi chú các bộ phận
    ax.annotate('Trụ pin, cửa van (sơ họa)', xy=(x_ng_dinh0 - 0.6, 76.2), xytext=(x_san0 + 0.4, 75.35), fontsize=7.5,
                color='#0b0b0b', zorder=8, arrowprops=dict(arrowstyle='-', color=C_PHU, lw=0.6))
    ghi = [((x_san0 + 0.4, 72.25), 'Sân lát đá ≈ +74,0'),
           ((x_ng_dinh1 + 0.3, Z_CAU_TREN + 0.2), 'Cầu giao thông ≈ +78,5 (ước lượng từ ảnh)'),
           ((x_san0 + 0.4, 70.2), 'Ngưỡng +74,70 (TP5), sân trước +74,0 (P₁ = 0,7 m)\nbậc hạ P₂ = 1,40 m, mái hạ lưu dài 3,0 m (bình đồ)\n→ đầu dốc +73,30 (TP7), cách dãy trụ pin ≈ 4,0 m'),
           ((8, 66.9), 'DỐC NƯỚC: L = %s m, i = %s %% (QTVH); bình đồ: chân ngưỡng → cuối mũi phun 61,87 m\nđoạn thu hẹp 31 → 18 m (x = 0 ÷ 20 m), sau đó rộng 18 m' % (s1(L).replace(',0', ''), s1(I * 100).replace(',0', ''))),
           ((x_moi + 0.6, 66.25), 'Mũi phun +%s, α = %d°\n(TP11 +68,90; sơ họa, R ≈ %s m)' % (s2(K.Z_MUI), K.ALPHA_MP, s1(R))),
           ((x_moi + 15.0, mp['MNHL'] + 0.3), 'MNHL lũ kiểm tra +%s' % s2(mp['MNHL'])),
           ((x_moi + 3.0, 61.0), 'Hố xói lũ kiểm tra: sâu %s m, cách mũi phun %s ÷ %s m' % (s2(hx), s2(mp['x_tl']), s2(mp['x_hl']))),
           ((x_moi + 11.5, z_day_hl + 0.25), 'Đáy lòng dẫn ≈ +%s (Bảng 6-8)' % s2(z_day_hl)),
           ((25, 61.6), 'Nền đá lớp II')]
    for (xg, zg), txt in ghi:
        ax.text(xg, zg, txt, fontsize=7.5, color='#0b0b0b', zorder=8, va='bottom')
    ax.set_xlim(x_san0, x_hl1)
    ax.set_ylim(60.5, 80.3)
    ax.set_aspect(2.5)
    ax.set_xlabel('Khoảng cách theo tim tràn, tính từ đầu dốc nước x (m)')
    ax.set_ylabel('Cao trình (m)')
    ax.set_title('Mặt cắt dọc tim tràn chính – sân trước, ngưỡng, dốc nước, mũi phun, hố xói (tỉ lệ đứng phóng 2,5 lần)',
                 fontsize=10, loc='left')
    ax.grid(True, color='#e1e0d9', lw=0.5, zorder=0)
    for s_ in ('top', 'right'):
        ax.spines[s_].set_visible(False)
    ax.legend(loc='upper right', fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(RA_DOC, dpi=220, facecolor='white')
    plt.close(fig)


def mat_cat_ngang(kq):
    plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 9})
    doc = {int(r['x']): r for r in kq['doc']}
    vi_tri = [(0, 'A–A: đầu dốc, x = 0 (B = %s m; tường %s m – Bảng 6-7)'),
              (30, 'B–B: đoạn lăng trụ, x = 30 m (B = %s m; tường %s m – Bảng 6-7)'),
              (60, 'C–C: cuối dốc, x = 60 m (B = %s m; tường %s m – giả thiết)')]
    fig, axs = plt.subplots(3, 1, figsize=(11.5, 9.0))
    for ax, (x, ten) in zip(axs, vi_tri):
        r = doc[x]
        B = r['B']
        z0 = Z_DD - I * x
        ht = tuong_day(x) - z0
        ax.set_title(ten % (s1(B).replace(',0', ''), s2(ht)), fontsize=9.5, loc='left')
        # đất đắp sau tường, nền
        for sgn in (-1, 1):
            xa = sgn * (B / 2 + T_TUONG)
            ax.add_patch(Polygon([(xa, z0 - T_DAY), (sgn * 18.5, z0 - T_DAY), (sgn * 18.5, z0 + ht - 0.15), (xa, z0 + ht - 0.15)],
                                 closed=True, fc=C_DAT, ec='none', zorder=1))
        ax.add_patch(Rectangle((-18.5, z0 - T_DAY - 1.0), 37, 1.0, fc=C_NEN, ec='none', zorder=1))
        # bản đáy và hai tường (sơ họa bề dày)
        ax.add_patch(Rectangle((-B / 2 - T_TUONG, z0 - T_DAY), B + 2 * T_TUONG, T_DAY, fc=C_BT, ec=C_NET, lw=1.0, zorder=3))
        for sgn in (-1, 1):
            x_in = sgn * B / 2
            ls = (0, (4, 3)) if x == 60 else '-'
            ax.add_patch(Rectangle((min(x_in, x_in + sgn * T_TUONG), z0), T_TUONG, ht, fc=C_BT, ec=C_NET, lw=1.0, ls=ls, zorder=3))
        # nước: lũ kiểm tra (tô) và lũ thiết kế (đường), có kể hàm khí
        ax.add_patch(Rectangle((-B / 2, z0), B, r['hhk_kt'], fc=C_NUOC, ec='none', zorder=2))
        ax.plot([-B / 2, B / 2], [z0 + r['hhk_kt']] * 2, color=C_KT, lw=2.0, zorder=4,
                label='Lũ kiểm tra: h = %s m, có hàm khí %s m' % (s2(r['h_kt']), s2(r['hhk_kt'])))
        ax.plot([-B / 2, B / 2], [z0 + r['hhk_tk']] * 2, color=C_TK, lw=2.0, zorder=4,
                label='Lũ thiết kế: h = %s m, có hàm khí %s m' % (s2(r['h_tk']), s2(r['hhk_tk'])))
        # kích thước
        xk = B / 2 + T_TUONG + 0.9
        ax.annotate('', xy=(xk, z0), xytext=(xk, z0 + ht), arrowprops=dict(arrowstyle='<->', color=C_NET, lw=0.8))
        ax.text(xk + 0.3, z0 + ht / 2, 'tường %s m' % s2(ht), fontsize=8, va='center', color='#0b0b0b')
        du = ht - r['hhk_kt']
        ax.annotate('', xy=(B / 2 - 1.0, z0 + r['hhk_kt']), xytext=(B / 2 - 1.0, z0 + ht),
                    arrowprops=dict(arrowstyle='<->', color=C_KT, lw=0.8))
        ax.text(B / 2 - 1.3, z0 + r['hhk_kt'] + du / 2, 'vượt cao %s m' % s2(du), fontsize=8, va='center', ha='right',
                color='#0b0b0b')
        ax.annotate('', xy=(-B / 2, z0 - T_DAY - 0.55), xytext=(B / 2, z0 - T_DAY - 0.55),
                    arrowprops=dict(arrowstyle='<->', color=C_NET, lw=0.8))
        ax.text(0, z0 - T_DAY - 0.5, 'B = %s m' % s1(B).replace(',0', ''), fontsize=8, ha='center', va='bottom', color='#0b0b0b')
        ax.text(-B / 2 + 0.3, z0 + 0.12, 'đáy +%s' % s2(z0), fontsize=7.5, color='#0b0b0b', zorder=5)
        ax.text(-B / 2 + 0.4, z0 + ht - 0.42, 'đỉnh tường +%s' % s2(z0 + ht), fontsize=7.5, ha='left', color='#0b0b0b', zorder=5)
        if x == 0:
            ax.text(-17.8, z0 + ht + 0.75, 'Bề dày tường, bản đáy vẽ sơ họa 0,5 m – đo bề dày tường trong CAD', fontsize=7.5,
                    color='#52514e', style='italic')
        if x == 60:
            ax.plot([-B / 2 - T_TUONG, B / 2 + T_TUONG], [70.25] * 2, color=C_CANH, lw=1.2, ls=(0, (2, 2)), zorder=5)
            ax.text(-18.2, z0 + 3.0, 'Đường đỏ: +70,25 (bình đồ, hai đầu tường cuối dốc). Nếu đây là đỉnh tường thì tường\n'
                    'chỉ cao %s m (%s m nếu đáy là +68,90), thấp hơn mức cần = h_hk lũ kiểm tra + 0,5 m = %s m\n'
                    '→ cần đo kiểm tra tường cuối dốc.' % (s2(70.25 - z0), s2(70.25 - 68.90), s2(r['hhk_kt'] + 0.5)),
                    fontsize=7.5, color='#0b0b0b', va='bottom')
        ax.set_xlim(-18.5, 18.5)
        ax.set_ylim(z0 - T_DAY - 1.0, z0 + 6.4)
        ax.set_aspect(1.0)
        ax.set_ylabel('Cao trình (m)')
        ax.grid(True, color='#e1e0d9', lw=0.5, zorder=0)
        for s_ in ('top', 'right'):
            ax.spines[s_].set_visible(False)
        ax.legend(loc='upper right' if x == 60 else 'upper center', fontsize=7.5, frameon=False)
    axs[-1].set_xlabel('Khoảng cách từ tim tràn (m) – nhìn theo chiều dòng chảy (tỉ lệ ngang = đứng)')
    fig.tight_layout()
    fig.savefig(RA_NGANG, dpi=220, facecolor='white')
    plt.close(fig)


if __name__ == '__main__':
    kq = K.tinh()
    mat_cat_doc(kq)
    mat_cat_ngang(kq)
    print('Đã ghi', os.path.abspath(RA_DOC), os.path.abspath(RA_NGANG))
