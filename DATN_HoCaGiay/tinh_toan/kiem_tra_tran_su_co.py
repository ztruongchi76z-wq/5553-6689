# -*- coding: utf-8 -*-
"""
KIỂM TRA TRÀN SỰ CỐ HIỆN TRẠNG (ĐỒ ÁN SỬA CHỮA) – cùng cách làm với tràn chính (kiem_tra_tran_chinh.py).
Tràn sự cố: 1 khoang b = 5 m, cửa van phẳng, ngưỡng +71,00 m, sau ngưỡng không có bậc hạ (P = 0); dốc nước lăng trụ rộng
5 m, dài 32 m, i = 8 % (QTVH Bảng 13); n = 0,015. Mố bên vuông góc (bình đồ cửa vào, SV chốt 07/11/2026): ξmb = 1,00.
  1. Khả năng xả (1 khoang, công thức (2') hồ sơ kiểm định): ε = 1 − 0,2·ξmb·(H/b), lấy H/b ≤ 1 → ε·m. Ngưỡng không có bậc
     hạ nên chảy phân giới ở đầu dốc: giới hạn lý thuyết m_max = φ·(2/3)·√(1/3) = 0,366 (φ = 0,95).
  2. Đầu dốc là mặt cắt khống chế (i > i_k): h = h_k. Đường mặt nước bII bằng sai phân ΔE/Δl = i − J_tb, chọn trước Δl = 2 m,
     thử dần h, bắt đầu từ 0,99·h_k (như thuyết minh mục 4.3); hàm khí h_hk = h(1 + v/100); chiều cao tường cần h_hk + a
     (a = 0,5 m); lưu tốc so với [v]kx = 17,6 m/s (bê tông M200 – như dốc tràn chính).
  3. Thông số dòng chảy cuối dốc (h, v, Fr = v²/gh, E, q) – căn cứ kiểm tra tiêu năng khi có hồ sơ (hồ sơ hiện có chưa ghi
     hình thức, kích thước công trình tiêu năng cuối dốc tràn sự cố).
Lưu lượng tràn sự cố: kết quả điều tiết lũ (file f1f2) – 4 trường hợp như tràn chính.
  python3 kiem_tra_tran_su_co.py      (in kết quả)
"""
import math

from scipy.optimize import brentq

import so_lieu as S
import thuy_luc_tran_chinh as T

G = 9.81
TSC = S.TRAN_SU_CO                      # Z_nguong = 71,00; b = 5; L_doc = 32; i_doc = 0,08; b_doc = 5
DL = 2.0                                # chiều dài đoạn chọn trước (m) – 16 đoạn
K0 = 0.99                               # độ sâu bắt đầu = 0,99·h_k (dòng chảy ở đầu dốc sát trạng thái phân giới)
V_KX = 17.6                             # như kiem_tra_tran_chinh.V_KX (dốc BT M200, Bảng 6-7 hồ sơ kiểm định)
M_MAX = T.PHI * (2 / 3) * math.sqrt(1 / 3)
CA = [('Lũ thiết kế P = 1%', 'P1_BinhThuong'), ('P = 1%, 1 cửa tràn chính không mở', 'P1_SuCo'),
      ('Lũ kiểm tra P = 0,2%', 'P02_BinhThuong'), ('P = 0,2%, 1 cửa tràn chính không mở', 'P02_SuCo')]


def eps_sc(H, ts):
    """ε tràn sự cố (như file điều tiết lũ): 1 − 0,2·[ξmb + (n − 1)ξmt]·(H/b)/n, H/b ≤ 1; n = 1 → 1 − 0,2·ξmb·H/b."""
    return 1 - 0.2 * (ts['SC_xmb'] + (ts['SC_n'] - 1) * ts['SC_xmt']) * min(H / ts['SC_b'], 1.0) / ts['SC_n']


def q_sc(H, ts):
    if H <= 0:
        return 0.0
    return ts['SC_s'] * eps_sc(H, ts) * ts['SC_m'] * ts['SC_n'] * ts['SC_b'] * math.sqrt(2 * G) * H ** 1.5


def buoc(Q, B, h1, dL, i):
    """Chọn trước Δl, tìm độ sâu cuối đoạn (nhánh dòng xiết): E₁ + i·Δl − J₁·Δl/2 = E₂ + J₂·Δl/2."""
    a = T.yeu_to(Q, B, h1)
    K = a['E'] + i * dL - 0.5 * a['J'] * dL

    def f(h):
        y = T.yeu_to(Q, B, h)
        return y['E'] + 0.5 * y['J'] * dL - K
    return brentq(f, 0.02, T.h_k(Q, B) * 0.999999, xtol=1e-12)


def duong_mat_nuoc(Q, B=TSC['b_doc'], L=TSC['L_doc'], i=TSC['i_doc'], dL=DL):
    h = K0 * T.h_k(Q, B)
    hang = [dict(x=0.0, **T.yeu_to(Q, B, h))]
    x = 0.0
    for _ in range(int(round(L / dL))):
        h = buoc(Q, B, h, dL, i)
        x = round(x + dL, 9)
        hang.append(dict(x=x, **T.yeu_to(Q, B, h)))
    for a, b in zip(hang, hang[1:]):
        b['dE'] = b['E'] - a['E']
        b['Jtb'] = 0.5 * (a['J'] + b['J'])
        b['i_J'] = i - b['Jtb']
        b['dL'] = b['dE'] / b['i_J']            # Δl tính lại – bằng Δl chọn khi h thử đúng
    for r in hang:
        r['hhk'] = r['h'] * (1 + r['v'] / 100)
        r['ht_can'] = r['hhk'] + T.A_TUONG
        r['Fr'] = r['v'] ** 2 / (G * r['h'])
    return hang


def tinh(ts=None):
    ts = ts or T.doc_excel()
    B = ts['SC_b']
    ds = []
    for ten, s in CA:
        r = ts[s]
        Q, Z = r['qsc'], r['Z']
        H = Z - ts['SC_Z']
        e = eps_sc(H, ts)
        bang = duong_mat_nuoc(Q, B)
        ds.append(dict(ten=ten, s=s, Q=Q, Z=Z, H=H, H_b=H / B, eps=e, m0=e * ts['SC_m'], Q_ct=q_sc(H, ts), q=Q / B,
                       q_tong=r['q'], ty_le=Q / r['q'], ik=T.i_k(Q, B), h0=T.h_0(Q, B, i=TSC['i_doc']), bang=bang,
                       dau=bang[0], cuoi=bang[-1], v_max=max(x['v'] for x in bang)))
    return dict(ts=ts, ca=ds, m_max=M_MAX, B=B, L=TSC['L_doc'], i=TSC['i_doc'])


if __name__ == '__main__':
    kq = tinh()
    print('m_max (φ = %.2f) = %.4f' % (T.PHI, kq['m_max']))
    for c in kq['ca']:
        d, e = c['dau'], c['cuoi']
        print('%-38s Q %7.2f (%.1f%% q tổng) Z %.3f H %.3f H/b %.3f ε %.4f εm %.4f Q_ct %.2f | hk %.3f ik %.5f h0 %.3f | '
              'đầu h %.3f v %.2f hhk %.3f tường %.2f | cuối h %.3f v %.2f Fr %.2f E %.2f hhk %.3f tường %.2f | vmax %.2f'
              % (c['ten'], c['Q'], 100 * c['ty_le'], c['Z'], c['H'], c['H_b'], c['eps'], c['m0'], c['Q_ct'], c['ik']['hk'],
                 c['ik']['ik'], c['h0']['h0'], d['h'], d['v'], d['hhk'], d['ht_can'], e['h'], e['v'], e['Fr'], e['E'],
                 e['hhk'], e['ht_can'], c['v_max']))
    for r in kq['ca'][2]['bang']:
        print('  x %4.1f h %.4f v %.3f E %.4f J %.5f hhk %.3f tường %.2f Δl %s' % (r['x'], r['h'], r['v'], r['E'], r['J'], r['hhk'],
                                                                              r['ht_can'], '%.4f' % r['dL'] if 'dL' in r else ''))
