# -*- coding: utf-8 -*-
"""
HÌNH BỔ SUNG CHO THUYẾT MINH (tiếp thu góp ý, 10/10/2026)

  python3 ve_hinh_bo_sung.py <thư mục ra>

  Hình 1.2  Trình tự thực hiện và kiểm chứng kết quả tính toán của đồ án (sơ đồ khối).
  Hình 8.6  Phạm vi thay đổi hiệu quả giảm lưu lượng thấm η theo từng thông số của màng (MNLTK) – biểu đồ
            "lốc xoáy" tổng hợp từ Bảng 8.3 ÷ Bảng 8.6 (số liệu: chuyen_de_khoan_phut.py, kt_hai_hang.py).
Kiểu chữ Liberation Serif (cùng cỡ với Times New Roman) như các hình khác của đồ án.
"""
import os
import sys

import matplotlib

matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 10.5, 'savefig.dpi': 220})


def hop(ax, x, y, w, h, chu, mau='#FFFFFF', vien='#1F3864', dam=False, co=10.5):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle='round,pad=0.012,rounding_size=0.018',
                                fc=mau, ec=vien, lw=1.1))
    ax.text(x, y, chu, ha='center', va='center', fontsize=co, fontweight='bold' if dam else 'normal',
            linespacing=1.25)


def mui_ten(ax, x0, y0, x1, y1, net='-', mau='#1F3864'):
    ax.annotate('', xy=(x1, y1), xytext=(x0, y0),
                arrowprops=dict(arrowstyle='-|>', lw=1.1, color=mau, ls=net, shrinkA=0, shrinkB=0,
                                mutation_scale=11))


def hinh_1_2(ra):
    fig, ax = plt.subplots(figsize=(6.3, 5.9))
    ax.set_xlim(0, 1)
    ax.set_ylim(0.09, 1)
    ax.axis('off')
    xm, w, h = 0.31, 0.58, 0.085
    buoc = [
        ('1. Thu thập hồ sơ: BCNCKT, TKCS (2022), QTVH (2026),\nkhảo sát địa hình, địa chất năm 2022', '#F2F2F2'),
        ('2. Rà soát, hiệu chỉnh số liệu cơ bản (Chương 1):\nZ ~ F ~ W, bề rộng tràn 21,6 m, tổng lượng lũ, [K]', '#FFFFFF'),
        ('3. Đánh giá an toàn hiện trạng (Chương 2): cấp nước,\nđiều tiết lũ, cao trình đỉnh, thấm, ổn định', '#FFFFFF'),
        ('4. Lựa chọn giải pháp sửa chữa, nâng cấp\n(Chương 3)', '#FFFFFF'),
        ('5. Thiết kế, kiểm tra các hạng mục: tràn xả lũ\n(Chương 4), đập đất (Chương 5), cống (Chương 6)',
         '#FFFFFF'),
        ('6. Thi công, khối lượng, quan trắc và quản lý\nvận hành (Chương 7)', '#FFFFFF'),
        ('7. Chuyên đề: màng khoan phụt mang tràn sự cố –\nPTHH, độ nhạy, hợp lý hóa (Chương 8)', '#DEEAF6'),
    ]
    ys = [0.935 - i * 0.128 for i in range(len(buoc))]
    for (chu, mau), y in zip(buoc, ys):
        hop(ax, xm, y, w, h, chu, mau=mau, co=9.6)
    for y0, y1 in zip(ys[:-1], ys[1:]):
        mui_ten(ax, xm, y0 - h / 2, xm, y1 + h / 2)
    # cột phải: công cụ và kiểm chứng – đặt ngang các bước 3, 5, 7 mà chúng phục vụ
    xr, wr, hr = 0.815, 0.33, 0.118
    phai = [
        (ys[2], 'Chương trình tính (Python,\nPhụ lục 5): điều tiết, sóng,\nthấm PTHH, Bishop, thủy lực', '#FFF2CC'),
        (ys[4], 'Kiểm chứng (Phụ lục 6): so với\nhồ sơ tư vấn; Spencer; lưới\ntâm mịn; hội tụ lưới PTHH', '#E2EFDA'),
        (ys[6], 'Phân tích độ nhạy giả thiết\n(Phụ lục 6; Bảng 8.3 ÷ 8.6,\nHình 8.6)', '#E2EFDA'),
    ]
    for y, chu, mau in phai:
        hop(ax, xr, y, wr, hr, chu, mau=mau, vien='#7F6000', co=9.0)
        mui_ten(ax, xr - wr / 2, y, xm + w / 2, y, net='--', mau='#7F6000')
    mui_ten(ax, xr, ys[2] - hr / 2, xr, ys[4] + hr / 2, mau='#7F6000')
    mui_ten(ax, xr, ys[4] - hr / 2, xr, ys[6] + hr / 2, mau='#7F6000')
    ax.text(xr, ys[0], 'Nét đứt: chương trình,\nkiểm chứng, độ nhạy dùng\ncho các bước 3, 5, 7', ha='center',
            va='center', fontsize=8.6, style='italic', color='#7F6000')
    fig.tight_layout(pad=0.2)
    for duoi in ('png', 'jpg'):
        fig.savefig(os.path.join(ra, 'hinh_1_2_trinh_tu.' + duoi), bbox_inches='tight',
                    pil_kwargs={'quality': 92} if duoi == 'jpg' else None)
    plt.close(fig)


# Hình 8.6 – phạm vi η (MNLTK +77,63 m) khi thay đổi từng thông số, các thông số khác như màng thiết kế (η = 54 %)
LOC_XOAY = [  # tên, (nhãn cận dưới, η), (nhãn cận trên, η)
    ('Hệ số thấm màng', ('1·10$^{-4}$ cm/s', 11), ('1·10$^{-6}$ cm/s', 71)),
    ('Cao trình đáy màng', ('+74,0 m (trong lớp I)', 7), ('+72,0 ÷ +66,0 m', 54)),
    ('Hệ số thấm lớp đá II', ('1·10$^{-4}$ cm/s', 26), ('1·10$^{-6}$ cm/s', 62)),
    ('Số hàng khoan (chiều dày T)', ('1 hàng, T = 1,8 m', 37), ('3 hàng, T = 4,8 m', 54)),
    ('Cao trình đỉnh màng', ('+77,0 m', 54), ('+78,0 m', 57)),
]


def hinh_8_6(ra):
    fig, ax = plt.subplots(figsize=(6.3, 3.0))
    goc = 54
    n = len(LOC_XOAY)
    for i, (ten, (l0, e0), (l1, e1)) in enumerate(LOC_XOAY):
        y = n - 1 - i
        ax.barh(y, e0 - goc, left=goc, height=0.55, color='#C55A11', alpha=0.85)
        ax.barh(y, e1 - goc, left=goc, height=0.55, color='#2E75B6', alpha=0.85)
        ax.text(min(e0, e1) - 1.2, y, '%s: %d%%' % (l0, e0), ha='right', va='center', fontsize=8.6)
        ax.text(max(e0, e1) + 1.2, y, '%s: %d%%' % (l1, e1), ha='left', va='center', fontsize=8.6)
    ax.axvline(goc, color='k', lw=1.0)
    ax.text(goc + 0.8, -0.55, 'màng thiết kế: η = 54 %', ha='left', va='bottom', fontsize=8.4, style='italic')
    ax.set_yticks(range(n))
    ax.set_yticklabels([t for t, _, _ in LOC_XOAY][::-1], fontsize=9.6)
    ax.set_xlim(-38, 112)
    ax.set_xticks(range(0, 101, 10))
    ax.set_xlabel('Hiệu quả giảm lưu lượng thấm η (%)')
    ax.set_ylim(-0.6, n - 0.1)
    ax.grid(axis='x', ls=':', lw=0.6)
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis='y', length=0)
    fig.tight_layout(pad=0.3)
    for duoi in ('png', 'jpg'):
        fig.savefig(os.path.join(ra, 'hinh_8_6_loc_xoay.' + duoi), bbox_inches='tight',
                    pil_kwargs={'quality': 92} if duoi == 'jpg' else None)
    plt.close(fig)


if __name__ == '__main__':
    ra = sys.argv[1] if len(sys.argv) > 1 else 'hinh_bo_sung'
    os.makedirs(ra, exist_ok=True)
    hinh_1_2(ra)
    hinh_8_6(ra)
    print('Đã vẽ:', sorted(os.listdir(ra)))
