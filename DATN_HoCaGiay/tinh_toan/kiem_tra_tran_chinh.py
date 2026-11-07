# -*- coding: utf-8 -*-
"""
KIỂM TRA TRÀN CHÍNH HIỆN TRẠNG (ĐỒ ÁN SỬA CHỮA) THEO CÁCH LÀM CỦA HỒ SƠ KIỂM ĐỊNH AN TOÀN ĐẬP (QTVH, chương 6 – SV gửi trích
05/11/2026), thay cho trình tự file mẫu thiết kế tràn mới:
  1. Khả năng xả lũ của ngưỡng: Q = σn·ε·m·Σb·√(2g)·H₀^1,5; bỏ qua lưu tốc tới gần (hồ sơ kiểm định: Ω_T ≫ 4·H·b) nên H₀ = H
     – cùng công thức với điều tiết lũ (file DieuTietLu_Potapov_f1f2); quan hệ Q ~ Z, kết quả điều tiết.
  2. Thủy lực dốc nước: độ sâu co hẹp đầu dốc (E₀ = P₂ + H), đường mặt nước bằng phương pháp sai phân, dốc có bề rộng thay
     đổi – chọn trước ΔL, thử dần h (như hồ sơ kiểm định); hàm khí h_hk = h(1 + v/100); độ vượt cao tường bên; không xói.
  3. Tiêu năng mũi phun – hố xói theo TCVN 8420:2010 (công thức (7) ÷ (10) của hồ sơ kiểm định): chiều dài phóng xa L, góc
     đổ α₁, chiều sâu xói t, h_x = t − h + c, khoảng cách tới điểm sâu nhất L₁, chiều dài hố b_x = 2,5h + 3,7h_x; mép hố
     xói phía mũi phun x = L₁ − 1,25h − 2,2h_x (mái thượng lưu hố 1:2,2) phải > 0.
Số liệu hạ lưu suy từ Bảng 6-8 hồ sơ kiểm định (tính ngược khớp bảng): cao trình mũi phun +69,0 m, góc hắt 25°, kích thước
khối đá d = c = 0,8 m, quan hệ Q ~ h, Q ~ MNHL (nội suy). L tính theo lưu tốc thực ở mũi phun (φ²·Z_H = v_H²/2g).
  python3 kiem_tra_tran_chinh.py      (in kết quả)
"""
import math

import numpy as np

import kiem_tra_tran_su_co as SC
import thuy_luc_tran_chinh as T

G = 9.81
TC = T.TC
Z_MUI = 69.0              # cao trình mũi phun (m) – Z_H của Bảng 6-8: +69,0 ÷ +69,07
ALPHA_MP = 25.0           # góc hắt mũi phun (độ) – tính ngược Bảng 6-8, như thuyết minh 4.2.4
K_MP = 1.0                # hệ số kể hàm khí, tách dòng (Fr_H < 30 ÷ 35)
D_DA = 0.80               # kích thước khối đá nứt nẻ tương đương d = c (m) – tính ngược Bảng 6-8
M_HO_TL, M_HO_HL = 2.2, 1.5   # hệ số mái thượng lưu, hạ lưu hố xói (hồ sơ kiểm định)
V_KX = 17.6               # lưu tốc không xói nhỏ nhất của dốc BT M200 trong Bảng 6-7 hồ sơ kiểm định (m/s)
# Bảng 6-8 hồ sơ kiểm định: Q (m³/s), độ sâu nước hạ lưu h (m), mực nước hạ lưu MNHL = 74,70 + H_tr − Z₀ (m)
KD_Q = [50.0, 100.0, 150.0, 200.0, 242.0, 400.0]
KD_h = [1.12, 1.73, 2.23, 2.67, 3.01, 4.11]
KD_MNHL = [66.91, 67.36, 67.70, 68.09, 68.37, 69.13]
KD_BANG68 = [(50, 1.11, 0.37, 7.60, 8.9, 6.8, 1.12, 17.20, 1.28, 7.53), (100, 1.76, 0.59, 9.4, 9.1, 7.4, 1.73, 17.80, 1.83, 11.09),
             (150, 2.30, 0.82, 10.2, 9.3, 8.0, 2.23, 19.30, 2.19, 13.70), (200, 2.79, 1.04, 10.7, 9.4, 8.5, 2.67, 21.01, 2.48, 15.86),
             (242, 3.17, 1.22, 11.0, 9.5, 8.8, 3.01, 22.58, 2.68, 17.44), (400, 4.43, 1.88, 11.8, 10.0, 10.1, 4.11, 30.69, 3.27, 22.39)]
Q_MUC = [50.0, 100.0, 150.0, 200.0]   # các cấp lưu lượng thêm để vẽ đường bao hố xói (như hồ sơ kiểm định)


def mui_phun(Q, MNL, hH, vH):
    """Tiêu năng mũi phun theo TCVN 8420:2010 (hồ sơ kiểm định, công thức (7) ÷ (10))."""
    a = math.radians(ALPHA_MP)
    h = float(np.interp(Q, KD_Q, KD_h))
    MNHL = float(np.interp(Q, KD_Q, KD_MNHL))
    ZH, Z0 = MNL - Z_MUI, MNL - MNHL
    phi = vH / math.sqrt(2 * G * ZH)
    L = K_MP * phi ** 2 * ZH * math.sin(2 * a) * (1 + math.sqrt(1 + (Z0 / ZH - 1) / (phi ** 2 * math.sin(a) ** 2)))
    tg1 = math.sqrt(math.tan(a) ** 2 + 2 * G * (0.5 * hH * math.cos(a) + Z_MUI - MNHL) / (vH ** 2 * math.cos(a) ** 2))
    Fr = vH ** 2 / (G * hH)
    t = (0.1 + 0.45 * Fr ** 0.25) * math.sqrt(Fr) * (h / D_DA) ** 0.25 * hH
    hx = t - h + D_DA
    L1 = L + t / tg1
    bx = 2.5 * h + 3.7 * hx
    return dict(Q=Q, MNL=MNL, hH=hH, vH=vH, h=h, MNHL=MNHL, ZH=ZH, Z0=Z0, phi=phi, L=L, tg1=tg1, a1=math.degrees(math.atan(tg1)),
                Fr=Fr, t=t, hx=hx, L1=L1, bx=bx, x_tl=L1 - 1.25 * h - M_HO_TL * hx, x_hl=L1 + 1.25 * h + M_HO_HL * hx,
                Z_day=MNHL - h)


def quan_he_Q_Z(ts, Z_ds):
    """Quan hệ Q ~ Z: tràn chính 9 / 8 khoang, tràn sự cố (như file điều tiết lũ)."""
    def q_sc(z):
        return SC.q_sc(z - ts['SC_Z'], ts)
    kq = []
    for z in Z_ds:
        H = max(z - ts['TC_Z'], 0.0)
        kq.append(dict(Z=z, H=H, eps=T.eps(H, ts), eps_sc=T.eps(H, ts, su_co=True), q9=T.q_tran(H, ts) if H > 0 else 0.0,
                       q8=T.q_tran(H, ts, su_co=True) if H > 0 else 0.0, qsc=q_sc(z),
                       eps_tsc=SC.eps_sc(max(z - ts['SC_Z'], 0.0), ts)))
    for r in kq:
        r['q_bt'], r['q_su_co'] = r['q9'] + r['qsc'], r['q8'] + r['qsc']
    return kq


def do_nhay_doc(ts, i):
    """Độ nhạy theo độ dốc dốc nước (hồ sơ kiểm định ghi i = 10 %): tính lại đường mặt nước lũ thiết kế, lũ kiểm tra với độ
    dốc i (P₂ giữ nguyên) – lưu tốc cuối dốc lớn nhất, độ vượt cao tường bên nhỏ nhất."""
    i0 = TC['i_doc']
    TC['i_doc'] = i
    try:
        du_min, v_max = float('inf'), 0.0
        for s in ('P1_BinhThuong', 'P02_BinhThuong'):
            Q = ts[s]['qtc']
            hc, _ = T.h_co_hep(Q, T.H0_theo_Q(Q, ts))
            bang = T.duong_mat_nuoc(Q, hc)
            for j, x in enumerate(range(0, int(TC['L_doc']) + 1, 5)):
                r = [b for b in bang if abs(b['x'] - x) < 1e-6][0]
                ht = TC['h_tuong'][j] if j < len(TC['h_tuong']) else TC['h_tuong'][-1]
                du_min = min(du_min, ht - r['h'] * (1 + r['v'] / 100))
            v_max = max(v_max, bang[-1]['v'])
    finally:
        TC['i_doc'] = i0
    return dict(i=i, v_max=v_max, du_min=du_min)


def tinh(ts=None):
    ts = ts or T.doc_excel()
    Q_tk, Q_kt = ts['P1_BinhThuong']['qtc'], ts['P02_BinhThuong']['qtc']
    cap = sorted([(q, 'Q = %g m³/s' % q) for q in Q_MUC] + [(Q_tk, 'Lũ thiết kế P = 1%'), (Q_kt, 'Lũ kiểm tra P = 0,2%')])
    ds = []
    for Q, ten in cap:
        H = T.H0_theo_Q(Q, ts)
        hc, E0 = T.h_co_hep(Q, H)
        bang = T.duong_mat_nuoc(Q, hc)
        c = bang[-1]
        MNL = ts['TC_Z'] + H
        ds.append(dict(ten=ten, Q=Q, H=H, MNL=MNL, E0=E0, hc=hc, hk_dau=T.h_k(Q, TC['b_doc_dau']), bang=bang,
                       mp=mui_phun(Q, MNL, c['h'], c['v'])))
    tk = [c for c in ds if c['ten'].startswith('Lũ thiết kế')][0]
    kt = [c for c in ds if c['ten'].startswith('Lũ kiểm tra')][0]
    doc = []
    for j, x in enumerate(range(0, int(TC['L_doc']) + 1, 5)):
        a = [r for r in tk['bang'] if abs(r['x'] - x) < 1e-6][0]
        b = [r for r in kt['bang'] if abs(r['x'] - x) < 1e-6][0]
        ht = TC['h_tuong'][j] if j < len(TC['h_tuong']) else TC['h_tuong'][-1]
        hhk_a, hhk_b = a['h'] * (1 + a['v'] / 100), b['h'] * (1 + b['v'] / 100)
        doc.append(dict(x=x, B=a['B'], h_tk=a['h'], hhk_tk=hhk_a, v_tk=a['v'], h_kt=b['h'], hhk_kt=hhk_b, v_kt=b['v'], ht=ht,
                        du_tk=ht - hhk_a, du_kt=ht - hhk_b, ht_can=max(hhk_a, hhk_b) + T.A_TUONG,
                        gia_thiet=j >= len(TC['h_tuong'])))
    Z_ds = sorted(set([round(74.70 + 0.25 * k, 2) for k in range(0, 20)] + [round(tk['MNL'], 2), round(kt['MNL'], 2)]))
    return dict(ts=ts, cap=ds, tk=tk, kt=kt, doc=doc, qz=quan_he_Q_Z(ts, Z_ds))


if __name__ == '__main__':
    kq = tinh()
    for c in kq['cap']:
        m = c['mp']
        print('%-22s Q %7.2f H %.3f MNL %.3f hc %.3f | cuối dốc h %.3f v %.2f | h_HL %.2f MNHL %.2f ZH %.2f Z0 %.2f φ %.3f '
              'L %.2f α1 %.1f° t %.2f hx %.2f L1 %.2f bx %.2f mép TL %.2f'
              % (c['ten'], c['Q'], c['H'], c['MNL'], c['hc'], m['hH'], m['vH'], m['h'], m['MNHL'], m['ZH'], m['Z0'], m['phi'],
                 m['L'], m['a1'], m['t'], m['hx'], m['L1'], m['bx'], m['x_tl']))
    for r in kq['doc']:
        print('  x %2d B %.2f TK h %.3f hhk %.3f v %.2f | KT h %.3f hhk %.3f v %.2f | tường %.2f dư %.2f / %.2f'
              % (r['x'], r['B'], r['h_tk'], r['hhk_tk'], r['v_tk'], r['h_kt'], r['hhk_kt'], r['v_kt'], r['ht'], r['du_tk'], r['du_kt']))
    # kiểm tra lại với số liệu Bảng 6-8 hồ sơ kiểm định (cùng công thức, φ² theo v_H)
    for Q, Htr, hH, vH, Z0, ZH, h, L1, hx, bx in KD_BANG68:
        m = mui_phun(Q, 74.70 + Htr, hH, vH)
        print('  KĐ Q %3d: hx %.2f (bảng %.2f) bx %.2f (bảng %.2f) L1 %.2f (bảng %.2f – bảng dùng (7) không nhân φ²)'
              % (Q, m['hx'], hx, m['bx'], bx, m['L1'], L1))
