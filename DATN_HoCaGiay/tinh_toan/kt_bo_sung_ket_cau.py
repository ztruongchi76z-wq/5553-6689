# -*- coding: utf-8 -*-
"""
KIỂM TRA BỔ SUNG (tiếp thu góp ý 12/10/2026)
 (1) Ổn định trục vít khi đóng cưỡng bức có giá đỡ trung gian: gối trung gian chỉ giữ chuyển vị ngang (không giữ xoay)
     -> đoạn dưới (gối trung gian – liên kết khớp với cửa van) làm việc như thanh hai đầu khớp μ = 1,0; đoạn trên
     (đai ốc máy – gối trung gian) μ = 0,7. So sánh với cách tính cũ μ = 0,7 cho cả hai đoạn; vị trí gối hợp lý.
 (2) Neo ngưỡng tràn D25: hệ số ổn định với khả năng chịu kéo ĐÃ KIỂM CHỨNG bằng thí nghiệm (100 kN/thanh);
     kiểm tra nhổ khối đá quanh neo (bỏ qua sức kháng cắt của đá, chỉ kể trọng lượng đẩy nổi) và chiều dài neo phía
     bê tông.
 (3) Dòng chảy tràn trên mái hạ lưu trong ô rãnh 4 × 4 m đặt xiên 45°: đường chảy dài nhất theo hướng dốc là đường
     chéo ô 4·√2 ≈ 5,66 m (≈ 5,96 m theo mặt mái nếu 4 m là kích thước hình chiếu bằng).
"""
import json
import math
import os

DIR = os.path.dirname(os.path.abspath(__file__))
out = {}

# (1) trục vít D85, d1 = 72 mm, E = 2,1e5 MPa; lực đóng theo Bảng 6.3
E = 2.1e11
d1 = 0.072
J = math.pi * d1 ** 4 / 64
P_dong = {'MNDBT f=0,3': 95.7, 'MNDBT f=0,4': 130.0, 'MNDBT f=0,5': 164.3,
          'MNLTK f=0,3': 136.8, 'MNLTK f=0,4': 183.5, 'MNLTK f=0,5': 230.3}   # kN


def P_th(mu, L):
    return math.pi ** 2 * E * J / (mu * L) ** 2 / 1e3     # kN


L_tong = 6.0
ph = {}
ph['khong_goi_mu07_L6'] = P_th(0.7, 6.0)
ph['goi_giua_cu_mu07_L3'] = P_th(0.7, 3.0)
ph['goi_giua_mu10_L3'] = min(P_th(1.0, 3.0), P_th(0.7, 3.0))
# vị trí gối tối ưu: 0,7·L1 = 1,0·L2 , L1 + L2 = 6 -> L2 = 6·0,7/1,7
L2 = L_tong * 0.7 / 1.7
L1 = L_tong - L2
ph['goi_toi_uu_L1'] = L1; ph['goi_toi_uu_L2'] = L2
ph['goi_toi_uu'] = min(P_th(0.7, L1), P_th(1.0, L2))
ph['hai_goi_mu10_L2'] = P_th(1.0, 2.0)
lam = {k: v for k, v in (('L6_mu07', 0.7 * 6 / (d1 / 4)), ('L3_mu10', 3.0 / (d1 / 4)), ('L2_mu10', 2.0 / (d1 / 4)),
                         ('L247_mu10', L2 / (d1 / 4)))}
out['truc_vit'] = dict(P_th=ph, do_manh=lam, Kod={})
print('Trục vít: Pth (kN)', {k: round(v, 1) for k, v in ph.items()}, '| độ mảnh', {k: round(v) for k, v in lam.items()})
for k, P in P_dong.items():
    r = {kk: ph[kk] / P for kk in ('khong_goi_mu07_L6', 'goi_giua_cu_mu07_L3', 'goi_giua_mu10_L3', 'goi_toi_uu',
                                   'hai_goi_mu10_L2')}
    out['truc_vit']['Kod'][k] = r
    print('  %-12s Pđóng %6.1f  K: không gối %.2f | gối giữa (μ=0,7 cũ) %.2f | gối giữa μ=1,0 %.2f | gối cách cửa %.2f m %.2f'
          ' | 2 gối %.2f' % (k, P, r['khong_goi_mu07_L6'], r['goi_giua_cu_mu07_L3'], r['goi_giua_mu10_L3'], L2,
                            r['goi_toi_uu'], r['hai_goi_mu10_L2']))

# (2) neo ngưỡng tràn
n = 2 / 1.5                       # thanh/m dài ngưỡng
f = 0.60
# [K] = γn·γlc/γc – QCVN 04-05:2022 điều 2.5.3, Bảng B.1, B.2 (γc = 0,95: mặt trượt qua tiếp xúc bê tông – đá)
TH = {'MNLTK': dict(V=132.5, H=97.2, Kcp=1.15 * 1.00 / 0.95), 'MNLKT': dict(V=128.2, H=118.7, Kcp=1.15 * 0.90 / 0.95)}
N_tt, N_thu = 171.8, 100.0
neo = {}
for k, d in TH.items():
    neo[k] = dict(K_khong_neo=f * d['V'] / d['H'], K_Ntt=f * (d['V'] + n * N_tt) / d['H'],
                  K_Nthu=f * (d['V'] + n * N_thu) / d['H'], N_can=(d['Kcp'] * d['H'] / f - d['V']) / n)
    print('Neo %s: K không neo %.2f | K với N_tt=171,8 %.2f | K với lực thử 100 kN %.2f | N cần %.1f kN ([K]=%.3f)'
          % (k, neo[k]['K_khong_neo'], neo[k]['K_Ntt'], neo[k]['K_Nthu'], neo[k]['N_can'], d['Kcp']))
# nhổ khối đá: mỗi neo chịu một lăng trụ đá diện tích 1,5 m (dọc hàng) × 2,5 m (nửa bề rộng bản đáy 5,05 m) cao 2,0 m
gam_da_dn = 26.0 - 9.81            # kN/m3 trọng lượng riêng đẩy nổi đá phong hóa vừa (giả thiết γ = 26 kN/m3)
W_da = 1.5 * 2.5 * 2.0 * gam_da_dn
neo['khoi_da'] = dict(W=W_da, K_nho_N_can_max=W_da / max(neo[k]['N_can'] for k in TH), K_nho_N_thu=W_da / N_thu)
# chiều dài neo trong bê tông (bản đáy dày 1,0 m – vữa bơm kín lỗ D76 xuyên qua ngưỡng)
tau_bs, tau_vb = 2.0e3, 0.5e3     # kPa: thép gân – vữa; vữa – bê tông (thận trọng)
L_bt = 1.0
neo['phia_be_tong'] = dict(N_thep_vua=math.pi * 0.025 * L_bt * tau_bs, N_vua_bt=math.pi * 0.076 * L_bt * tau_vb)
print('Nhổ khối đá: W = %.0f kN/neo -> hệ số %.2f với N cần lớn nhất, %.2f với lực thử 100 kN' % (
    W_da, neo['khoi_da']['K_nho_N_can_max'], neo['khoi_da']['K_nho_N_thu']))
print('Phía bê tông (L = 1,0 m): thép – vữa %.0f kN, vữa – bê tông %.0f kN' % (
    neo['phia_be_tong']['N_thep_vua'], neo['phia_be_tong']['N_vua_bt']))
out['neo'] = neo

# (3) dòng chảy tràn trên mái hạ lưu
C_, n_m, beta = 0.6, 0.03, math.atan(1 / 3)
xoi = {}
for nhan, ty_le in (('20%', 0.20), ('30%', 0.30), ('40%', 0.40)):
    i = 191.8 * ty_le / 0.5 / 1000 / 3600          # m/s (mưa 30 phút = tỷ lệ × X10%)
    for L in (4.0, 4 * math.sqrt(2), 4 * math.sqrt(2) * math.sqrt(1 + 1 / 9), 34.0):
        q = C_ * i * L
        h = (q * n_m / math.sqrt(math.sin(beta))) ** 0.6
        tau = 9810 * h * math.sin(beta)
        xoi['%s_L%.2f' % (nhan, L)] = dict(i_mmh=i * 3.6e6, L=L, q=q, h_mm=1000 * h, v=q / h, tau=tau)
for k, v in xoi.items():
    print('Xói mái %-10s i=%5.1f mm/h L=%5.2f m: q=%.2e h=%.2f mm v=%.2f m/s τ=%.2f Pa' % (
        k, v['i_mmh'], v['L'], v['q'], v['h_mm'], v['v'], v['tau']))
out['xoi_mai'] = xoi
json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kt_bo_sung_ket_cau.json'), 'w'), ensure_ascii=False, indent=1)
