# -*- coding: utf-8 -*-
"""
HÌNH BỔ SUNG ĐỢT 6 (12/10/2026):
  Hình 4.2 – Lượng mưa năm (1978 – 2024) và mưa 1 ngày lớn nhất X1max (1978 – 2020) lưu vực hồ Cà Giây, trước và sau
             khi công trình vận hành (2000); số liệu tinh_toan/du_lieu/mua_luu_vuc.json
  Hình 4.3 – Mực nước lớn nhất trong hồ theo mức tăng lưu lượng lũ đến (tràn hiện trạng, vận hành theo QTVH)
  Hình 3.1 – Cấu tạo mặt đập sau sửa chữa (thay hình chi tiết TKCS: đỉnh +78,85 m, dốc một phía 2 %)
  Hình 3.2 – Ảnh bản vẽ TKCS: ghi lại tiêu đề bị lỗi phông VNI bằng chữ Unicode
  python3 ve_hinh_bo_sung_2.py <thư mục tinh_toan> <ảnh Hình 3.2 gốc (jpg)>
"""
import json
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image, ImageDraw, ImageFont

DIR = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(DIR, 'hinh_bo_sung')
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({
    'font.family': 'Liberation Serif', 'font.size': 10, 'axes.titlesize': 11, 'axes.labelsize': 10,
    'legend.fontsize': 9, 'xtick.labelsize': 9, 'ytick.labelsize': 9, 'figure.dpi': 110,
    'savefig.dpi': 220, 'axes.grid': True, 'grid.alpha': 0.3, 'grid.linestyle': '--',
})


def luu(fig, ten, rong=1500):
    png = os.path.join(OUT, ten + '.png')
    fig.savefig(png, bbox_inches='tight')
    plt.close(fig)
    im = Image.open(png).convert('RGB')
    if im.width > rong:
        im = im.resize((rong, int(round(im.height * rong / im.width))), Image.LANCZOS)
    jpg = os.path.join(OUT, ten + '.jpg')
    im.save(jpg, quality=90)
    return jpg


def so(x, n=1):
    return ('%.*f' % (n, x)).replace('.', ',')


def hinh_4_2(dir_tt):
    sys.path.insert(0, dir_tt)
    import kt_mua_lu_bien_doi as KM
    M = KM.M
    nam = np.array([int(y) for y in M['mua_nam']])
    X = np.array([M['mua_nam'][str(y)][12] for y in nam])
    nx = np.array([int(y) for y in M['x1max']])
    x1 = np.array([M['x1max'][str(y)] for y in nx])
    plt.rcParams.update({'font.size': 7.5, 'axes.titlesize': 8, 'axes.labelsize': 7.5, 'legend.fontsize': 6.3, 'xtick.labelsize': 7, 'ytick.labelsize': 7})   # khổ in 15,5 cm
    fig, axs = plt.subplots(2, 1, figsize=(6.3, 5.4), sharex=True)
    for ax, t, v, ten, dv in ((axs[0], nam, X, 'a) Lượng mưa năm lưu vực (QTVH 2026, Bảng 2-9)', 'mm'),
                              (axs[1], nx, x1, 'b) Lượng mưa 1 ngày lớn nhất năm X1max (KTTV 2022, Bảng 3-13)', 'mm')):
        truoc, sau = t < 2000, t >= 2000
        ax.bar(t[truoc], v[truoc], color='#9fb3c8', width=0.75, label='Trước khi vận hành (%d – 1999)' % t.min())
        ax.bar(t[sau], v[sau], color='#e9a66b', width=0.75, label='Sau khi vận hành (2000 – %d)' % t.max())
        mt, ms = v[truoc].mean(), v[sau].mean()
        ax.plot([t.min() - 0.4, 1999.4], [mt, mt], color='#1f4e9c', lw=1.4,
                label='TB %d – 1999: %s %s' % (t.min(), so(mt), dv))
        ax.plot([1999.6, t.max() + 0.4], [ms, ms], color='#b5521b', lw=1.4,
                label='TB 2000 – %d: %s %s (%s%%)' % (t.max(), so(ms), dv, so(100 * (ms / mt - 1)).replace('-', '−')))
        mk = KM.mann_kendall(v)
        b = mk['sen']; a = np.median(v - b * (t - t.min()))
        ax.plot(t, a + b * (t - t.min()), 'k--', lw=1.0,
                label='Xu thế Sen %s mm/năm (Mann – Kendall p = %s)' % (so(b, 2).replace('-', '−'), so(mk['p'], 2)))
        ax.axvline(1999.5, color='0.4', lw=0.8, ls=':')
        ax.set_title(ten, loc='left')
        ax.set_ylabel(dv)
        ax.legend(loc='upper left', fontsize=6.3, ncol=2, framealpha=0.9)
    for yy, vv in ((1982, 245.0), (2003, 280.4), (2010, 234.4)):
        axs[1].annotate('%d: %s' % (yy, so(vv)), (yy, vv), xytext=(0, 3), textcoords='offset points', ha='center',
                        fontsize=6.5)
    axs[0].set_ylim(0, 3300)
    axs[1].set_ylim(0, 420)
    axs[1].set_xlabel('Năm')
    axs[1].text(1999.2, 215, 'Hồ vận hành từ 2000', fontsize=6.5, color='0.3', rotation=90, ha='right', va='center')
    fig.tight_layout()
    return luu(fig, 'hinh_4_2_mua', rong=2000)


def hinh_4_3(dir_tt):
    sys.path.insert(0, dir_tt)
    import so_lieu as S
    import kt_mua_lu_bien_doi as KM
    ks = np.arange(0, 0.301, 0.05)
    plt.rcParams.update({'font.size': 7.5, 'axes.titlesize': 8, 'axes.labelsize': 7.5, 'legend.fontsize': 6.3, 'xtick.labelsize': 7, 'ytick.labelsize': 7})   # khổ in 15,5 cm
    fig, ax = plt.subplots(figsize=(6.3, 3.7))
    for nguon, Qd, mau in (('BCNCKT', S.Q_KTTV, '#c0392b'), ('QTVH', S.Q_QTVH, '#1f4e9c')):
        for p, ls, mk in (('1%', '--', 'o'), ('0.2%', '-', 's')):
            z = [KM.zmax(Qd[p], k)['Zmax'] for k in ks]
            ax.plot(100 * ks, z, ls=ls, marker=mk, color=mau, ms=3, lw=1.1,
                    label='Lũ %s P = %s' % (nguon, p.replace('.', ',')))
    for z, ten, mau in ((77.24, 'MNLTK thiết kế ban đầu (TK 1996) +77,24', '#2e7d32'),
                        (78.35, 'Đỉnh đập hiện trạng +78,35', '0.45'),
                        (78.55, 'Đỉnh đập sau SC − 0,3 m = +78,55 (giới hạn MNLKT)', '#8e44ad'),
                        (78.85, 'Đỉnh đập sau sửa chữa +78,85', 'k'),
                        (79.40, 'Đỉnh tường chắn sóng +79,40', '#6d4c41')):
        ax.axhline(z, color=mau, lw=1.0, ls=':' if z != 78.85 else '-')
        x_, va_ = {78.55: (12.5, 'top'), 78.35: (0.2, 'top')}.get(z, (0.2, 'bottom'))
        ax.text(x_, z + (0.03 if va_ == 'bottom' else -0.03), ten, fontsize=6.5, color=mau, ha='left', va=va_)
    ax.set_xlabel('Mức tăng lưu lượng lũ đến so với lũ thiết kế hiện hành k (%) – tăng cả đỉnh và tổng lượng')
    ax.set_ylabel('Mực nước lớn nhất Zmax (m)')
    ax.set_xlim(-1, 31); ax.set_ylim(77.0, 79.8)
    ax.legend(loc='lower right', fontsize=6.3, ncol=2, framealpha=0.95)
    fig.tight_layout()
    return luu(fig, 'hinh_4_3_lu_tang', rong=2000)


def hinh_3_1(dp=0.13):
    """Mặt cắt đỉnh đập sau sửa chữa – tỷ lệ đứng phóng ×4. Cao trình thiết kế sau lún; nghiệm thu cộng dự phòng lún
    dp (≥ 0,5%H – TCVN 8216:2018 điều 11.4.4). Khổ hình bằng khổ in (≈ 16 cm) để chữ đọc được khi in."""
    fig, ax = plt.subplots(figsize=(7.2, 3.9))
    ex, z0 = 4.0, 77.75
    Z = lambda z: (z - z0) * ex
    fs = 7.0
    hop = dict(boxstyle='round,pad=0.12', fc='white', ec='none', alpha=0.85)
    xL, xR = -3.45, 5.6
    # mái thượng lưu lát đá (giữ nguyên) – tiếp giáp tường ở ≈ +78,35 (chi tiết TKCS)
    xm = np.array([xL, -0.75]); zm = 78.35 - (xm[1] - xm) / 3.0
    ax.fill_between(xm, Z(77.75), Z(zm), color='#e3d2b5', lw=0)
    ax.fill_between(xm, Z(zm), Z(78.54), where=Z(zm) < Z(78.54), color='#cfe6fb', lw=0)
    ax.plot(xm, Z(zm), color='#5d4037', lw=2.0)
    ax.fill_between([-0.75, -0.35], Z(78.35), Z(78.54), color='#cfe6fb', lw=0)
    ax.plot([xL, -0.35], [Z(78.54), Z(78.54)], color='#1f77b4', lw=1.2)
    ax.text(xL + 0.05, Z(78.54) + 0.05, 'MNLKT +78,54', color='#1f77b4', fontsize=fs)
    ax.text(-2.15, Z(77.98), 'Mái thượng lưu lát đá\n(giữ nguyên)', fontsize=fs, ha='center', bbox=hop)
    # tường chắn sóng: thân dày 0,35 m (+78,35 ÷ +79,40), bản móng 1,0 m (+77,85 ÷ +78,35)
    ax.fill([-0.35, 0, 0, -0.35], [Z(78.35), Z(78.35), Z(79.40), Z(79.40)], color='#bdbdbd', ec='k', lw=0.9)
    ax.fill([-0.75, 0.25, 0.25, -0.75], [Z(77.85), Z(77.85), Z(78.35), Z(78.35)], color='#bdbdbd', ec='k', lw=0.9)
    ax.annotate('Tường chắn sóng (giữ nguyên)\nđỉnh +79,40', (-0.35, Z(79.28)), xytext=(xL + 0.05, Z(79.28)),
                fontsize=fs, va='center', arrowprops=dict(arrowstyle='->', lw=0.6))
    # thân đập cũ: mặt sau cào bóc +78,20 tới mép đỉnh cũ (x = 4,7), sau đó mái cũ 1:3
    xo = np.array([0.25, 4.7, xR]); zo = np.array([78.20, 78.20, 78.20 - (xR - 4.7) / 3.0])
    ax.fill_between(xo, Z(77.75), Z(zo), color='#d9b38c', lw=0)
    ax.fill_between([-0.75, 0.25], Z(77.75), Z(77.85), color='#d9b38c', lw=0)
    ax.plot(xo, Z(zo), 'k--', lw=0.6)
    ax.text(2.45, Z(77.90), 'Khối chống thấm thân đập (ĐĐ-I), mặt cào bóc +78,20', ha='center', fontsize=fs)
    # đất đắp bù: +78,60 tại tường, dốc 2 % tới gờ; sau gờ đắp theo mái 1:3
    zd = lambda x: 78.60 - 0.02 * x
    x1 = np.array([0.0, 4.5]); ax.fill_between(x1, Z(78.20), Z(zd(x1)), color='#cfa57a', lw=0)
    zs = lambda x: zd(4.5) - (x - 4.5) / 3.0
    x2 = np.linspace(4.5, xR, 20); ax.fill_between(x2, Z(np.minimum(zo[1], 78.20 - np.maximum(x2 - 4.7, 0) / 3.0)),
                                                   Z(zs(x2)), color='#cfa57a', lw=0)
    ax.plot(x2, Z(zs(x2)), 'k-', lw=0.8)
    ax.text(5.15, Z(78.64), 'Mái HL\nm = 3', fontsize=fs, ha='center')
    ax.text(2.2, Z(78.33), 'Đất đắp bù K ≥ 0,97, k = 3$\\cdot$10$^{-6}$ cm/s:\nđỉnh +78,60 (thi công +%s)'
            % so(78.60 + dp, 2), ha='center', fontsize=fs)
    # đá dăm 5 cm + bê tông 20 cm
    xs = np.array([0.0, 4.3]); zt = lambda x: 78.85 - 0.02 * x
    ax.fill_between(xs, Z(zd(xs)), Z(zd(xs) + 0.05), color='#8c8c8c', lw=0)
    ax.fill_between(xs, Z(zt(xs) - 0.20), Z(zt(xs)), facecolor='#f2f2f2', edgecolor='#555555', lw=0.7, hatch='..')
    ax.text(2.15, Z(78.71), 'BT B20 dày 20 cm (lớp bảo vệ)', ha='center', fontsize=fs, bbox=hop)
    ax.text(2.15, Z(78.565), 'đá dăm 5 cm + nilon', ha='center', fontsize=6.2, color='w')
    # gờ chắn bánh
    ax.fill([4.3, 4.5, 4.5, 4.3], [Z(zt(4.3) - 0.20), Z(zt(4.3) - 0.20), Z(zt(4.3) + 0.25), Z(zt(4.3) + 0.25)],
            color='#bdbdbd', ec='k', lw=0.7)
    ax.text(4.4, Z(zt(4.3) + 0.29), 'Gờ chắn bánh', ha='center', fontsize=fs)
    # đường MNLKT + 0,3
    ax.plot([xL, xR], [Z(78.84), Z(78.84)], color='#8e44ad', lw=0.8, ls=':')
    ax.text(xL + 0.05, Z(78.84) + 0.05, 'MNLKT + 0,3 = +78,84', color='#8e44ad', fontsize=fs)
    # ghi cao trình mặt đập
    ax.annotate('+78,85 thiết kế (sau lún)\nnghiệm thu ≥ +%s' % so(78.85 + dp, 2), (0.03, Z(78.85)),
                xytext=(0.45, Z(79.30)), fontsize=fs, arrowprops=dict(arrowstyle='->', lw=0.6))
    ax.annotate('+78,76 (nghiệm thu ≥ +%s)' % so(78.76 + dp, 2), (4.27, Z(zt(4.27))), xytext=(2.45, Z(79.12)),
                fontsize=fs, arrowprops=dict(arrowstyle='->', lw=0.6))
    ax.text(1.0, Z(78.93), 'dốc ngang 2 % →', fontsize=fs)
    ax.annotate('', (-0.45, Z(79.40)), xytext=(-0.45, Z(78.85)), arrowprops=dict(arrowstyle='<->', lw=0.5))
    ax.text(-0.5, Z(79.05), '0,55 m', fontsize=fs, ha='right')
    ax.annotate('Khe tường – tấm: chèn ma tít;\nđất đầm sát lưng tường', (0.05, Z(78.45)), xytext=(-3.3, Z(78.72)),
                fontsize=fs, bbox=hop, arrowprops=dict(arrowstyle='->', lw=0.5))
    ax.set_xlim(xL, xR); ax.set_ylim(Z(77.75), Z(79.5))
    yt = np.arange(77.8, 79.41, 0.2)
    ax.set_yticks([Z(v) for v in yt]); ax.set_yticklabels([so(v, 1) for v in yt], fontsize=fs)
    ax.tick_params(axis='x', labelsize=fs)
    ax.set_xlabel('Khoảng cách từ mặt hạ lưu tường chắn sóng (m)', fontsize=fs + 0.5)
    ax.set_ylabel('Cao độ (m), tỷ lệ đứng ×4', fontsize=fs + 0.5)
    ax.grid(False)
    fig.tight_layout()
    return luu(fig, 'hinh_3_1_mat_dap', rong=2200)


def hinh_3_2(anh_goc):
    """Che tiêu đề lỗi phông VNI (ảnh chụp bản vẽ TKCS) và ghi lại bằng chữ Unicode."""
    im = Image.open(anh_goc).convert('RGB')
    W, H = im.size
    d = ImageDraw.Draw(im)
    fb = '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'
    fr = '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'
    sx = W / 1500.0
    def ghi(box, dong, co):
        x0, y0, x1, y1 = [int(round(v * sx)) for v in box]
        d.rectangle([x0, y0, x1, y1], fill='white')
        y = y0
        for t, dam in dong:
            f = ImageFont.truetype(fb if dam else fr, int(round(co * sx)))
            w = d.textlength(t, font=f)
            d.text(((x0 + x1) / 2 - w / 2, y), t, fill='black', font=f)
            y += int(round((co + 4) * sx))
    ghi((300, 2, 1180, 64), [('MẶT BẰNG BỐ TRÍ HỆ THỐNG RÃNH THOÁT NƯỚC MÁI HẠ LƯU', True), ('TỶ LỆ 1/200', False)], 22)
    ghi((1300, 108, 1500, 165), [('CẮT II-II', True), ('TỶ LỆ 1/25', False)], 18)
    ghi((1300, 404, 1500, 460), [('CẮT I-I', True), ('TỶ LỆ 1/25', False)], 18)
    ghi((1300, 726, 1500, 782), [('CẮT III-III', True), ('TỶ LỆ 1/50', False)], 18)
    jpg = os.path.join(OUT, 'hinh_3_2_ranh.jpg')
    im.save(jpg, quality=92)
    return jpg


if __name__ == '__main__':
    dir_tt = sys.argv[1]
    print(hinh_4_2(dir_tt))
    print(hinh_4_3(dir_tt))
    print(hinh_3_1())
    if len(sys.argv) > 2:
        print(hinh_3_2(sys.argv[2]))
