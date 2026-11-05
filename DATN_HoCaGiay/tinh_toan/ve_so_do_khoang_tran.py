# -*- coding: utf-8 -*-
"""
SƠ ĐỒ KHOANG TRÀN CHÍNH – TRƯỜNG HỢP SỰ CỐ 1 CỬA KHÔNG MỞ (điều tiết lũ, xuat_excel_dieu_tiet_lu_potapov.py)
  Nhìn từ thượng lưu về hạ lưu: bờ trái bên trái. Đánh số khoang K1 ÷ K9 từ bờ trái, trụ T1 ÷ T8 (sơ đồ – bản vẽ thi
  công có thể đánh số khác). 9 khoang × 2,4 m; 6 trụ giữa 0,6 m (2 trụ trong mỗi khoang của tràn 3 khoang ban đầu) và
  2 trụ chính ban đầu 1,0 m (T3, T6) – tổng bề rộng 27,2 m.
  (a) cửa khoang biên K1 không mở – trường hợp tính trong file Excel: 8 khoang xả liền nhau K2 ÷ K9, trụ T1 cạnh khoang
      đóng làm việc như tường bên của nhóm → 7 trụ giữa các khoang xả: Σξ = 5 × 0,45 + 2 × 0,45 = 3,15.
  (b) cửa khoang giữa K5 không mở – để so sánh: hai nhóm 4 khoang, Zmax cao hơn (a) 0,3 ÷ 0,6 cm (lũ QTVH 72 giờ).

  python3 ve_so_do_khoang_tran.py [đường dẫn .png ra]
Mặc định: ../TAI_VE/SoDo_KhoangTranChinh_SuCo.png
"""
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch, Rectangle

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'SoDo_KhoangTranChinh_SuCo.png')
plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 8.5, 'savefig.dpi': 220})

B_KHOANG = 2.4
TRU = [0.6, 0.6, 1.0, 0.6, 0.6, 1.0, 0.6, 0.6]          # T1 ÷ T8 từ bờ trái
B_MO, H_TRU, H_NUOC, H_NGUONG = 1.6, 5.4, 3.8, 0.9     # mố bên, chiều cao vẽ trụ, mực nước lũ, chiều dày ngưỡng (sơ đồ)
MAU = dict(nuoc='#cde2fb', vien_nuoc='#2a78d6', tru6='#b8b6ae', tru10='#52514e', dong='#d03b3b', mo='#e7e6e1',
           chu='#0b0b0b', chu2='#52514e')

# vị trí khoang và trụ (m, tính từ mố bên trái)
KHOANG, TRU_X, x = [], [], 0.0
for k in range(9):
    KHOANG.append((x, x + B_KHOANG))
    x += B_KHOANG
    if k < 8:
        TRU_X.append((x, x + TRU[k]))
        x += TRU[k]
B_TONG = x                                              # 27,2 m


def so(v, n=1):
    return ('%.*f' % (n, v)).replace('.', ',')


def ve(ax, dong, tieu_de, ghi_chu):
    ax.add_patch(Rectangle((-B_MO, -H_NGUONG), B_TONG + 2 * B_MO, H_NGUONG, fc=MAU['mo'], ec=MAU['chu2'], lw=0.6,
                           hatch='////'))
    for x0 in (-B_MO, B_TONG):
        ax.add_patch(Rectangle((x0, 0), B_MO, H_TRU + 0.4, fc=MAU['mo'], ec=MAU['chu2'], lw=0.6, hatch='////'))
    for k, (a, b) in enumerate(KHOANG):
        if k + 1 == dong:
            ax.add_patch(Rectangle((a, 0), b - a, H_NUOC + 0.3, fc=MAU['dong'], ec=MAU['chu'], lw=0.6))
            xm, ym = (a + b) / 2, H_NUOC / 2 + 0.75
            for s in (1, -1):
                ax.plot([xm - 0.55, xm + 0.55], [ym - 0.55 * s, ym + 0.55 * s], color='white', lw=2.2,
                        solid_capstyle='round')
            ax.text(xm, H_NUOC / 2 - 0.85, 'cửa\nkhông\nmở', ha='center', va='center', color='white', fontsize=7,
                    linespacing=1.1)
        else:
            ax.add_patch(Rectangle((a, 0), b - a, H_NUOC, fc=MAU['nuoc'], ec='none'))
            ax.plot([a, b], [H_NUOC, H_NUOC], color=MAU['vien_nuoc'], lw=1.2)
            ax.text((a + b) / 2, H_NUOC / 2, 'xả', ha='center', va='center', color=MAU['chu2'], fontsize=7.5)
        ax.text((a + b) / 2, H_TRU - 0.35, 'K%d' % (k + 1), ha='center', va='center', fontsize=9.5, fontweight='bold',
                color=MAU['chu'])
        ax.text((a + b) / 2, -H_NGUONG / 2, so(B_KHOANG), ha='center', va='center', fontsize=7, color=MAU['chu2'],
                bbox=dict(fc='white', ec='none', pad=0.4))
    for k, (a, b) in enumerate(TRU_X):
        day = TRU[k] >= 1.0
        ax.add_patch(Rectangle((a, 0), b - a, H_TRU, fc=MAU['tru10' if day else 'tru6'], ec=MAU['chu'], lw=0.6))
        ax.text((a + b) / 2, H_TRU + 0.3, 'T%d' % (k + 1), ha='center', va='bottom', fontsize=7.5, color=MAU['chu'])
        ax.text((a + b) / 2, H_TRU + 1.05, so(TRU[k]), ha='center', va='bottom', fontsize=6.5, color=MAU['chu2'])
    ax.text(-B_MO / 2, H_TRU + 0.9, 'mố bên\nbờ trái', ha='center', va='bottom', fontsize=7, color=MAU['chu2'])
    ax.text(B_TONG + B_MO / 2, H_TRU + 0.9, 'mố bên\nbờ phải', ha='center', va='bottom', fontsize=7, color=MAU['chu2'])
    ax.text(B_TONG + B_MO + 0.3, H_NUOC, 'mực nước lũ', ha='left', va='center', fontsize=7, color=MAU['vien_nuoc'])
    ax.text(B_TONG + B_MO + 0.3, 0, 'ngưỡng +74,70', ha='left', va='center', fontsize=7, color=MAU['chu2'])
    ax.plot([B_TONG, B_TONG + B_MO + 0.25], [0, 0], color=MAU['chu2'], lw=0.5, ls=':')
    # khoang biên
    for k in (0, 8):
        a, b = KHOANG[k]
        ax.annotate('', xy=(a, H_TRU + 2.35), xytext=(b, H_TRU + 2.35),
                    arrowprops=dict(arrowstyle='|-|', lw=0.7, color=MAU['chu'], mutation_scale=2.5))
        ax.text((a + b) / 2, H_TRU + 2.55, 'khoang biên', ha='center', va='bottom', fontsize=7.5, color=MAU['chu'],
                fontstyle='italic')
    # kích thước tổng
    ax.annotate('', xy=(0, -H_NGUONG - 0.9), xytext=(B_TONG, -H_NGUONG - 0.9),
                arrowprops=dict(arrowstyle='<->', lw=0.6, color=MAU['chu2']))
    ax.text(B_TONG / 2, -H_NGUONG - 0.75, 'B = 9 × 2,4 + 6 × 0,6 + 2 × 1,0 = %s m' % so(B_TONG), ha='center', va='bottom',
            fontsize=7, color=MAU['chu2'], bbox=dict(fc='white', ec='none', pad=0.6))
    ax.set_title(tieu_de, loc='left', fontsize=9.5, fontweight='bold', color=MAU['chu'])
    ax.text(-B_MO, -H_NGUONG - 1.6, ghi_chu, ha='left', va='top', fontsize=8, color=MAU['chu'], linespacing=1.35)
    ax.set_xlim(-B_MO - 0.3, B_TONG + B_MO + 4.2)
    ax.set_ylim(-H_NGUONG - 4.6, H_TRU + 3.5)
    ax.set_aspect('equal')
    ax.axis('off')


fig, axs = plt.subplots(2, 1, figsize=(7.6, 7.3))
ve(axs[0], 1, '(a) Sự cố tính trong file Excel: cửa khoang biên K1 không mở',
   '• Khoang biên = khoang sát mố bên: K1 hoặc K9 (đối xứng – kết quả như nhau).\n'
   '• K1 đóng → còn 8 khoang xả liền nhau K2 ÷ K9. Trụ T1 nằm giữa khoang đóng và khoang xả nên làm việc như\n'
   '   tường bên của nhóm, không tính là trụ giữa hai khoang xả → còn 7 trụ T2 ÷ T8 (5 trụ 0,6 m + 2 trụ 1,0 m):\n'
   '   Σξ = 5 × 0,45 + 2 × 0,45 = 3,15;  ε = 1 − 0,2·(ξmb + Σξ)·H / (8 × 2,4).')
ve(axs[1], 5, '(b) Để so sánh: cửa khoang giữa K5 không mở',
   '• Hai nhóm 4 khoang xả K1 ÷ K4 và K6 ÷ K9; T4 và T5 cạnh khoang đóng làm việc như tường bên của mỗi nhóm.\n'
   '• Thêm một chỗ co hẹp bên nên xả kém hơn (a) một chút: Zmax chỉ cao hơn 0,3 ÷ 0,6 cm (lũ QTVH, P = 1% và 0,2%)\n'
   '   → chọn (a) để tính là đủ, chênh lệch không đáng kể.')
fig.legend(handles=[Patch(fc=MAU['tru6'], ec=MAU['chu'], lw=0.6, label='trụ giữa 0,6 m (mới)'),
                    Patch(fc=MAU['tru10'], ec=MAU['chu'], lw=0.6, label='trụ chính ban đầu 1,0 m'),
                    Patch(fc=MAU['nuoc'], ec=MAU['vien_nuoc'], lw=0.6, label='khoang xả (cửa kéo hết)'),
                    Patch(fc=MAU['dong'], ec=MAU['chu'], lw=0.6, label='cửa không mở được'),
                    Patch(fc=MAU['mo'], ec=MAU['chu2'], lw=0.6, hatch='////', label='mố bên, ngưỡng')],
           loc='lower center', ncol=3, frameon=False, fontsize=7.5, bbox_to_anchor=(0.5, 0.0))
fig.suptitle('Sơ đồ tràn chính nhìn từ thượng lưu: 9 khoang × 2,4 m – đánh số K1 ÷ K9 từ bờ trái', fontsize=10.5,
             fontweight='bold', color=MAU['chu'], y=0.995)
fig.subplots_adjust(left=0.01, right=0.99, top=0.935, bottom=0.06, hspace=0.06)
fig.savefig(RA, bbox_inches='tight', facecolor='white')
print('Đã ghi', os.path.abspath(RA))
