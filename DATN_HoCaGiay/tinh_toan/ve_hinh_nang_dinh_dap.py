# -*- coding: utf-8 -*-
"""
HÌNH MINH HỌA: HAI CÁCH NÂNG CAO TRÌNH ĐỈNH ĐẬP CÀ GIÂY MÀ VẪN GIỮ HỆ SỐ MÁI m = 3
  (a) Nâng trong phạm vi mặt đỉnh (phương án Đ3 của đồ án, chi tiết A – BV-04): bóc lớp mặt đập cũ +78,50 đến +78,35,
      đắp bù K ≥ 0,97 đến +78,60, đá dăm 5 cm + BT B20 20 cm → mặt đập +78,85 (mép thượng lưu, dốc ngang 2 %).
      Tường chắn sóng giữ phía thượng lưu, gờ chắn bánh + tấm lát mái giữ mép hạ lưu: không đắp thêm lên mái.
  (b) Đắp áp mái hạ lưu: muốn giữ cả bề rộng đỉnh và hệ số mái thì mái mới song song với mái cũ, dịch ra
      Δx = m·Δh (phía thượng lưu là tường thẳng đứng); không có tường thì Δx = (m_TL + m_HL)·Δh.
  Hai hình cùng tỷ lệ đứng ×2. Mặt cắt lòng sông (cọc C16): đỉnh rộng 5 m kể cả tường, mái hạ lưu m = 3 đến cơ +68,0
  rộng 14 m (thuyết minh mục 5.2).

  python3 ve_hinh_nang_dinh_dap.py [đường dẫn .png ra]
Mặc định: ../TAI_VE/HinhMinhHoa_NangDinhDap.png
"""
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'HinhMinhHoa_NangDinhDap.png')
plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 8.5, 'savefig.dpi': 220})

EX = 2.0                                  # tỷ lệ đứng
M = 3.0                                   # hệ số mái hạ lưu trên cơ +68,0
Z_CU, Z_BOC, Z_DAT, Z_MOI = 78.50, 78.35, 78.60, 78.85   # đỉnh cũ theo hồ sơ (thống nhất 21/10/2026), bóc 15 cm
DU_PHONG_LUN = 0.13                       # đắp thi công đến +78,73 (TCVN 8216:2018 điều 11.4.4)
X_MEP = 4.70                              # mép hạ lưu mặt đập; gốc x = 0 tại mặt hạ lưu tường chắn sóng
TUONG, MNLKT = 79.40, 78.51
Z_CO, B_CO = 68.0, 14.0
DH = Z_MOI - Z_CU                         # chiều cao nâng so với mặt đập hiện trạng

MAU = dict(dat='#D9B38C', dat_tl='#E3D2B5', moi='#F6CDA9', vien_moi='#EB6834', bt='#BDBDBD', nuoc='#D3E6F8',
           muc='#2A78D6', yc='#8E44AD', chu='#1F1F1F', phu='#52514E')
HOP = dict(boxstyle='round,pad=0.18', fc='white', ec='none', alpha=0.9)


def so(x, n=2):
    return ('%.*f' % (n, x)).replace('.', ',')


def dien_tich(p):
    x, y = np.array(p, float).T
    return 0.5 * abs(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1)))


def ve_vung(ax, p, z0, **kw):
    x, y = np.array(p, float).T
    return ax.fill(x, (y - z0) * EX, **kw)


def ve_duong(ax, p, z0, **kw):
    x, y = np.array(p, float).T
    return ax.plot(x, (y - z0) * EX, **kw)


def mui_ten(ax, a, b, z0, chu, xy_chu, **kw):
    ax.annotate(chu, (a, (b - z0) * EX), xytext=(xy_chu[0], (xy_chu[1] - z0) * EX), fontsize=7.4, color=MAU['chu'],
                bbox=HOP, arrowprops=dict(arrowstyle='->', lw=0.6, color=MAU['chu']), **kw)


def truc(ax, z0, z1, buoc, x0, x1):
    ax.set_xlim(x0, x1)
    ax.set_ylim(0, (z1 - z0) * EX)
    ax.set_aspect('equal')
    zt = np.arange(np.ceil(z0 / buoc) * buoc, z1 + 1e-9, buoc)
    ax.set_yticks((zt - z0) * EX)
    ax.set_yticklabels([so(v, 1 if buoc < 1 else 0) for v in zt])
    ax.tick_params(labelsize=7.4, color=MAU['phu'])
    for s in ax.spines.values():
        s.set_color('#9A9A9A')
    ax.set_xlabel('Khoảng cách từ mặt hạ lưu tường chắn sóng (m)', fontsize=7.8)
    ax.set_ylabel('Cao độ (m), tỷ lệ đứng ×2', fontsize=7.8)


def tuong(ax, z0):
    ve_vung(ax, [(-0.75, 77.85), (0.0, 77.85), (0.0, 78.35), (-0.75, 78.35)], z0, fc=MAU['bt'], ec='k', lw=0.7)
    ve_vung(ax, [(-0.35, 78.35), (0.0, 78.35), (0.0, TUONG), (-0.35, TUONG)], z0, fc=MAU['bt'], ec='k', lw=0.8)


# ================================================================== (a) nâng trong phạm vi mặt đỉnh
fig, (a1, a2) = plt.subplots(2, 1, figsize=(7.4, 8.4), gridspec_kw=dict(height_ratios=[1.0, 1.05]))
z0 = 77.55
xL, xR = -2.2, 7.6
z_mai = lambda x: Z_CU - (x - X_MEP) / M            # mái hạ lưu cũ
x_day = X_MEP + M * (Z_CU - z0)
# thượng lưu: mái lát đá 1:3 tới chân móng tường, nước tới MNLKT
x_tl = np.array([xL, -0.75])
z_tl = 78.35 - (-0.75 - x_tl) / 3.0
a1.fill_between(x_tl, 0, (z_tl - z0) * EX, color=MAU['dat_tl'], lw=0)
a1.fill_between(np.r_[x_tl, -0.35], (np.r_[z_tl, 78.35] - z0) * EX, (MNLKT - z0) * EX, color=MAU['nuoc'], lw=0)
ve_duong(a1, [(xL, MNLKT), (-0.35, MNLKT)], z0, color=MAU['muc'], lw=1.1)
a1.text(xL + 0.08, (MNLKT - z0) * EX + 0.05, 'MNLKT +%s' % so(MNLKT), color=MAU['muc'], fontsize=7.4)
# thân đập cũ (sau khi bóc lớp mặt đến +78,35; mép đỉnh cũ +78,50)
cu = [(-0.75, z0), (-0.75, 77.85), (0.0, 77.85), (0.0, Z_BOC), (X_MEP, Z_BOC), (X_MEP, Z_CU), (x_day, z0)]
ve_vung(a1, cu, z0, fc=MAU['dat'], ec='none')
a1.text(2.35, (77.82 - z0) * EX, 'Thân đập cũ (khối chống thấm ĐĐ-I)', ha='center', fontsize=7.4, color=MAU['chu'])
# đất đắp bù K ≥ 0,97 (cao trình sau lún) – mép hạ lưu bám theo đường mái cũ kéo dài
z_dat = lambda x: Z_DAT - 0.02 * x
dap_a = [(0.0, Z_BOC), (X_MEP, Z_BOC), (X_MEP, z_dat(X_MEP)), (4.50, z_dat(4.50)), (4.30, z_dat(4.30)), (0.0, Z_DAT)]
ve_vung(a1, dap_a, z0, fc=MAU['moi'], ec=MAU['vien_moi'], lw=0.8, hatch='////')
a1.text(2.2, (78.465 - z0) * EX, 'Đất đắp bù K ≥ 0,97: +%s → +%s (thi công +%s)'
        % (so(Z_BOC), so(Z_DAT), so(Z_DAT + DU_PHONG_LUN)), ha='center', va='center', fontsize=7.2, bbox=HOP)
# đá dăm 5 cm + BT B20 20 cm, dốc ngang 2 %
z_mat = lambda x: Z_MOI - 0.02 * x
ve_vung(a1, [(0.0, Z_DAT), (4.30, z_dat(4.30)), (4.30, z_mat(4.30)), (0.0, Z_MOI)], z0, fc='#F2F2F2', ec='#555555',
        lw=0.7, hatch='..')
a1.text(2.15, (78.715 - z0) * EX, 'BT B20 dày 20 cm + đá dăm 5 cm', ha='center', va='center', fontsize=7.2, bbox=HOP)
# gờ chắn bánh, tấm lát mái 10 cm (mép trên tấm bám đường mái cũ kéo dài lên)
ve_vung(a1, [(4.30, z_dat(4.50)), (4.50, z_dat(4.50)), (4.50, z_mat(4.30) + 0.25), (4.30, z_mat(4.30) + 0.25)], z0,
        fc=MAU['bt'], ec='k', lw=0.7)
ve_vung(a1, [(4.50, z_dat(4.50)), (5.25, z_dat(4.50) - 0.75 / M), (5.25, z_dat(4.50) - 0.75 / M - 0.10),
             (4.50, z_dat(4.50) - 0.10)], z0, fc=MAU['bt'], ec='k', lw=0.6)
tuong(a1, z0)
# mặt đập cũ và mái hạ lưu (không đổi)
ve_duong(a1, [(0.0, Z_CU), (X_MEP, Z_CU)], z0, color=MAU['phu'], lw=0.9, ls='--')
ve_duong(a1, [(X_MEP, Z_CU), (x_day, z0)], z0, color=MAU['chu'], lw=1.3)
# đường yêu cầu MNLKT + 0,3
ve_duong(a1, [(xL, MNLKT + 0.3), (xR, MNLKT + 0.3)], z0, color=MAU['yc'], lw=0.8, ls=':')
a1.text(4.62, (MNLKT + 0.3 - z0) * EX + 0.05, 'Yêu cầu: MNLKT + 0,3 = +%s' % so(MNLKT + 0.3), color=MAU['yc'],
        fontsize=7.4)
# ghi chú
mui_ten(a1, -0.36, 79.30, z0, 'Tường chắn sóng\n(giữ nguyên)\nđỉnh +%s' % so(TUONG), (xL + 0.08, 79.22), va='center')
mui_ten(a1, 0.04, Z_MOI, z0, 'Mặt đập mới +%s\n(mép thượng lưu, sau lún)' % so(Z_MOI), (0.9, 79.22), va='center')
mui_ten(a1, 3.4, Z_CU, z0, 'Mặt đập cũ +%s (nét đứt): bóc lớp mặt hư hỏng đến +%s' % (so(Z_CU), so(Z_BOC)),
        (2.35, 78.05), va='center', ha='center')
a1.text(4.40, (z_mat(4.30) + 0.30 - z0) * EX, 'Gờ chắn bánh', ha='center', fontsize=7.4)
mui_ten(a1, 6.5, z_mai(6.5), z0, 'Mái hạ lưu giữ m = 3:\nkhông đắp thêm lên mái', (5.45, 78.47), va='center')
a1.annotate('', (-0.35, (79.52 - z0) * EX), xytext=(X_MEP, (79.52 - z0) * EX),
            arrowprops=dict(arrowstyle='<->', lw=0.6, color=MAU['chu']))
a1.text(2.2, (79.55 - z0) * EX, 'Bề rộng đỉnh ≈ 5,0 m (kể cả tường) – giữ nguyên', ha='center', va='bottom',
        fontsize=7.4)
V_A = dien_tich(dap_a) + DU_PHONG_LUN * 4.5
a1.text(6.4, (77.60 - z0) * EX, 'Đắp ≈ %s m³/m dài (kể dự phòng lún)' % so(V_A, 1), ha='right', va='bottom',
        fontsize=7.6, fontweight='bold', bbox=HOP)
truc(a1, z0, 79.75, 0.2, xL, xR)
a1.set_title('(a) Cách 1 – nâng trong phạm vi mặt đỉnh (đồ án chọn, phương án Đ3): mái không thay đổi',
             loc='left', fontsize=8.6, fontweight='bold')

# ================================================================== (b) đắp áp mái hạ lưu
z0 = 66.9
xL, xR = -3.0, 46.0
x_co_cu = X_MEP + M * (Z_CU - Z_CO)
x_co_moi = X_MEP + M * (Z_MOI - Z_CO)
DX = x_co_moi - x_co_cu
# bậc đánh cấp trên mái cũ: cao 0,5 m, rộng m·0,5 m (mũi bậc nằm trên đường mái cũ)
bac = [(X_MEP, Z_CU)]
x, z = X_MEP, Z_CU
while z - 0.5 > Z_CO + 1e-9:
    z -= 0.5
    bac.append((x, z))
    x += 0.5 * M
    bac.append((x, z))
bac += [(x, Z_CO), (x_co_cu, Z_CO)]
x_tl = -0.75 - 3.0 * (78.35 - 77.30)
cu = [(xL, z0), (xL, 77.30), (x_tl, 77.30), (-0.75, 78.35), (0.0, 78.35), (0.0, Z_CU)] + bac + [(xR, Z_CO), (xR, z0)]
ve_vung(a2, cu, z0, fc=MAU['dat'], ec='none')
ve_vung(a2, [(xL, 77.30), (x_tl, 77.30), (-0.75, 78.35), (-0.35, 78.35), (-0.35, MNLKT), (xL, MNLKT)], z0,
        fc=MAU['nuoc'], ec='none')
ve_duong(a2, [(xL, MNLKT), (-0.35, MNLKT)], z0, color=MAU['muc'], lw=1.0)
dap_b = [(0.0, Z_CU)] + bac + [(x_co_moi, Z_CO), (X_MEP, Z_MOI), (0.0, Z_MOI)]
ve_vung(a2, dap_b, z0, fc=MAU['moi'], ec=MAU['vien_moi'], lw=0.6, hatch='////')
ve_duong(a2, bac, z0, color=MAU['vien_moi'], lw=0.5)
tuong(a2, z0)
# mặt cắt cũ (nét đứt) và mới
ve_duong(a2, [(0.0, Z_CU), (X_MEP, Z_CU), (x_co_cu, Z_CO)], z0, color=MAU['phu'], lw=0.9, ls='--')
ve_duong(a2, [(0.0, Z_MOI), (X_MEP, Z_MOI), (x_co_moi, Z_CO), (xR, Z_CO)], z0, color=MAU['chu'], lw=1.3)
# kích thước Δh, Δx
a2.annotate('', (1.2, (Z_CU - z0) * EX), xytext=(1.2, (Z_MOI - z0) * EX),
            arrowprops=dict(arrowstyle='<->', lw=0.6, shrinkA=0, shrinkB=0))
a2.text(1.6, (Z_MOI + 0.12 - z0) * EX, 'Δh = %s m' % so(DH), fontsize=7.6, ha='left', va='bottom')
z_dx = 73.0
x1, x2 = X_MEP + M * (Z_CU - z_dx), X_MEP + M * (Z_MOI - z_dx)
a2.annotate('', (x1, (z_dx - z0) * EX), xytext=(x2, (z_dx - z0) * EX),
            arrowprops=dict(arrowstyle='<->', lw=0.7, shrinkA=0, shrinkB=0))
a2.annotate('Δx = m·Δh = 3 × %s = %s m\n(mái mới song song với mái cũ)' % (so(DH), so(DX)),
            ((x1 + x2) / 2, (z_dx - z0) * EX), xytext=(x2 + 2.2, (z_dx + 2.2 - z0) * EX), fontsize=7.6, bbox=HOP,
            arrowprops=dict(arrowstyle='->', lw=0.6))
mui_ten(a2, 12.2, Z_CU - (12.2 - X_MEP) / M - 0.25, z0, 'Đánh cấp mái cũ (bậc cao 0,5 m)\nrồi đắp từng lớp, đầm chặt',
        (12.6, 70.6), va='center')
mui_ten(a2, 27.5, Z_CU - (27.5 - X_MEP) / M, z0, 'Mái cũ m = 3 (nét đứt)', (23.0, 68.55), va='center')
mui_ten(a2, 31.0, Z_MOI - (31.0 - X_MEP) / M, z0, 'Mái mới m = 3', (34.0, 72.6), va='center')
a2.text(xR - 0.4, (Z_CO + 0.3 - z0) * EX, 'Cơ +68,0\nrộng %s → %s m' % (so(B_CO, 1), so(B_CO - DX, 1)),
        ha='right', va='bottom', fontsize=7.4, bbox=HOP)
a2.annotate('', (-0.35, (79.62 - z0) * EX), xytext=(X_MEP, (79.62 - z0) * EX),
            arrowprops=dict(arrowstyle='<->', lw=0.6, shrinkA=0, shrinkB=0))
a2.text(X_MEP + 0.6, (79.62 - z0) * EX, 'Bề rộng đỉnh giữ 5,0 m', ha='left', va='center', fontsize=7.4)
V_B = dien_tich([(0.0, Z_CU), (X_MEP, Z_CU), (x_co_cu, Z_CO), (x_co_moi, Z_CO), (X_MEP, Z_MOI), (0.0, Z_MOI)])
a2.text(xR - 0.3, (Z_CO - 0.75 - z0) * EX, 'Đắp ≈ %s m³/m dài (≈ %s lần cách 1, chưa kể đào – đắp lại bậc cấp)'
        % (so(V_B, 1), so(V_B / V_A, 0)), ha='right', va='center', fontsize=7.6,
        fontweight='bold', bbox=HOP)
a2.text(xL + 0.6, (69.6 - z0) * EX,
        'Khi đắp áp mái hạ lưu:\n'
        '• bóc cỏ, rễ, đất tơi; đánh cấp mái cũ\n'
        '• đắp lớp 20 ÷ 25 cm, đầm K ≥ 0,95 ÷ 0,97\n'
        '• đất đắp thấm bằng hoặc tốt hơn đất thân\n'
        '  đập (hoặc có lớp lọc) để không dâng\n'
        '  đường bão hòa; nối dài thoát nước chân đập\n'
        '• không có tường chắn sóng (mái thượng lưu\n'
        '  kéo dài lên): Δx = (m_TL + m_HL)·Δh',
        fontsize=7.2, va='center', ha='left', bbox=dict(boxstyle='round,pad=0.35', fc='white', ec='#B8BEC6', lw=0.6))
truc(a2, z0, 80.1, 1.0, xL, xR)
a2.set_title('(b) Cách 2 – giữ cả bề rộng đỉnh và m: đắp áp mái hạ lưu, mái dịch song song Δx = m·Δh',
             loc='left', fontsize=8.6, fontweight='bold')

fig.tight_layout(h_pad=1.6)
fig.savefig(RA, bbox_inches='tight', facecolor='white')
print('Đã ghi', os.path.abspath(RA))
print('Cách 1: %.2f m3/m (dự phòng lún %.2f); cách 2: %.2f m3/m; Δx = %.2f m; tỷ số %.1f'
      % (V_A, DU_PHONG_LUN * 4.5, V_B, DX, V_B / V_A))
