# -*- coding: utf-8 -*-
"""
THỦY LỰC TRÀN CHÍNH HỒ CÀ GIÂY THEO TRÌNH TỰ FILE MẪU “THIẾT KẾ ĐƯỜNG TRÀN THÁO LŨ” (SV gửi 05/11/2026)
  1. Kiểm tra khả năng tháo: xét cột nước lưu tốc tới gần (Ω_TL so với 4·Σb·H), hệ số co hẹp bên ε (TCVN 9147:2012, như
     file điều tiết lũ), m, σn; so Q_tt ở MNLTK, MNLKT với lưu lượng tràn chính trong kết quả điều tiết lũ, sai số δ.
  2. Thủy lực dốc nước: độ sâu đầu dốc (độ sâu co hẹp sau bậc hạ P = 1,40 m – khác mẫu tràn đỉnh rộng h_đd = h_k), độ sâu
     phân giới h_k, độ dốc phân giới i_k, độ sâu dòng đều h_0 (Agơrốtskin), đường mặt nước (cộng trực tiếp; đoạn thu hẹp
     chọn trước ΔL = 2,5 / 5 m, thử dần độ sâu), kiểm tra không xói, chiều cao tường bên (hàm khí), chiều dày bản đáy
     (Đômbrốpxki).
Thông số tràn, kết quả điều tiết đọc từ file Excel điều tiết lũ (TAI_VE/DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx – LibreOffice
tính lại trước khi đọc); hình học dốc nước, tường bên theo so_lieu.TRAN_CHINH.
  python3 thuy_luc_tran_chinh.py      (in kết quả)
"""
import math
import os
import subprocess
import tempfile
import warnings

from openpyxl import load_workbook
from scipy.optimize import brentq

import so_lieu as S

warnings.filterwarnings('ignore')
DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(DIR, '..', 'TAI_VE', 'DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx')
G = 9.81
ALPHA_TL = 1.10          # hệ số động năng dòng tới gần (mẫu: α = 1 ÷ 1,1)
ALPHA = 1.00             # hệ số động năng, động lượng trên dốc (mẫu: α = 1)
GOC_TUONG = 14.0         # góc mở rộng tường cánh thượng lưu (độ) – như mẫu
L_S = 12.0               # chiều dài tường cánh thượng lưu giả thiết (m): (3 ÷ 5)·H với cả hai mực nước lũ
PHI = 0.95               # hệ số lưu tốc chảy qua ngưỡng xuống chân bậc
N_NHAM = 0.015           # hệ số nhám dốc nước bê tông (như thuyết minh mục 4.2.3)
A_TUONG = 0.50           # độ vượt cao an toàn tường bên (m) – mẫu chọn a = 0,5 m
ALPHA_NEN = 1.0          # hệ số nền trong công thức Đômbrốpxki (mẫu: nền sét 0,8 ÷ 1,0)
V_KX = (19.2, 21.1)      # lưu tốc không xói bê tông cốt thép M200, h = 0,5 / 1 m (mẫu: TCVN 4118:2012, phụ lục I, Bảng I5)
TC = S.TRAN_CHINH
N_BUOC_LT = 8            # số bước trên mỗi đoạn dốc: thu hẹp 20/8 = 2,5 m; lăng trụ 40/8 = 5 m


# ------------------------------------------------------------------ số liệu từ file Excel điều tiết lũ
def doc_excel(xlsx=XLSX):
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', tmp, xlsx], check=True,
                       capture_output=True, timeout=900)
        wb = load_workbook(os.path.join(tmp, os.path.basename(xlsx)), data_only=True)
        o = lambda ten: (lambda ws, dc: wb[ws][dc.replace('$', '')].value)(*next(wb.defined_names[ten].destinations))
        ts = {k: o(k) for k in ('TC_Z', 'TC_n', 'TC_b', 'TC_n1', 'TC_d1', 'TC_x1', 'TC_n2', 'TC_d2', 'TC_x2', 'TC_Btong',
                                'TC_m', 'TC_xmb', 'TC_sx', 'TC_s', 'TC_n_sc', 'TC_sx_sc', 'SC_Z', 'Z0', 'Z_dinh_SC', 'Z_tuong')}
        for s in ('P1_BinhThuong', 'P02_BinhThuong', 'P1_SuCo', 'P02_SuCo'):
            ws = wb[s]
            ts[s] = dict(Z=ws['F6'].value, t=ws['F7'].value, q=ws['F8'].value, qtc=ws['F9'].value, qsc=ws['G9'].value,
                         H=ws['F10'].value, Qden=max(ws['B%d' % r].value for r in range(16, 89)))
    return ts


# ------------------------------------------------------------------ khả năng tháo
def eps(H0, ts, su_co=False):
    n, sx = (ts['TC_n_sc'], ts['TC_sx_sc']) if su_co else (ts['TC_n'], ts['TC_sx'])
    return 1 - 0.2 * (ts['TC_xmb'] + sx) * H0 / (n * ts['TC_b'])


def q_tran(H0, ts, su_co=False):
    n = ts['TC_n_sc'] if su_co else ts['TC_n']
    return ts['TC_s'] * eps(H0, ts, su_co) * ts['TC_m'] * n * ts['TC_b'] * math.sqrt(2 * G) * H0 ** 1.5


def H0_theo_Q(Q, ts):
    return brentq(lambda h: q_tran(h, ts) - Q, 1e-4, 8.0)


def kha_nang_thao(ts):
    """Kiểm tra khả năng tháo ở MNLTK, MNLKT (vận hành bình thường) theo trình tự mẫu."""
    B = ts['TC_Btong']                                 # bề rộng tràn kể trụ pin (cửa vào)
    Z_day = ts['TC_Z'] - TC['P1']                      # cao trình đáy kênh dẫn (sân trước)
    B_TL = B + 2 * L_S * math.tan(math.radians(GOC_TUONG))
    Sb = ts['TC_n'] * ts['TC_b']
    kq = {}
    for k, s in (('1%', 'P1_BinhThuong'), ('0.2%', 'P02_BinhThuong')):
        r = ts[s]
        H, q = r['H'], r['qtc']
        H_TL = r['Z'] - Z_day
        Om = B_TL * H_TL
        V0 = q / Om
        H0 = H + ALPHA_TL * V0 ** 2 / (2 * G)
        e = eps(H0, ts)
        Qtt = ts['TC_s'] * e * ts['TC_m'] * Sb * math.sqrt(2 * G) * H0 ** 1.5
        kq[k] = dict(Z=r['Z'], H=H, q=q, q_tong=r['q'], H_TL=H_TL, Om=Om, gh=4 * Sb * H, V0=V0, hv=H0 - H, H0=H0, eps=e,
                     Qtt=Qtt, sai_so=abs(Qtt - q) / Qtt * 100, Ls=((3 * H, 5 * H)))
    return dict(B=B, Z_day=Z_day, B_TL=B_TL, Sb=Sb, delta=TC['delta'], kq=kq)


# ------------------------------------------------------------------ thủy lực dốc nước
def B_doc(x):
    if x <= TC['L_chuyen_tiep']:
        return TC['b_doc_dau'] + (TC['b_doc_cuoi'] - TC['b_doc_dau']) * x / TC['L_chuyen_tiep']
    return TC['b_doc_cuoi']


def yeu_to(Q, B, h, n=N_NHAM):
    w = B * h
    chi = B + 2 * h
    R = w / chi
    v = Q / w
    C = R ** (1 / 6) / n
    J = v ** 2 / (C ** 2 * R)
    return dict(B=B, h=h, w=w, chi=chi, R=R, v=v, C=C, J=J, E=h + ALPHA * v ** 2 / (2 * G))


def h_k(Q, B):
    return (ALPHA * (Q / B) ** 2 / G) ** (1 / 3)


def i_k(Q, B, n=N_NHAM):
    hk = h_k(Q, B)
    y = yeu_to(Q, B, hk, n)
    return dict(Q=Q, q=Q / B, hk=hk, w=y['w'], chi=y['chi'], R=y['R'], C=y['C'], ik=Q ** 2 / (y['w'] ** 2 * y['C'] ** 2 * y['R']))


def h_0(Q, B, i=TC['i_doc'], n=N_NHAM):
    """Độ sâu dòng đều; kèm các đại lượng của phương pháp Agơrốtskin (m = 0 → 4m₀ = 8)."""
    h0 = brentq(lambda h: B * h * (B * h / (B + 2 * h)) ** (2 / 3) * math.sqrt(i) / n - Q, 1e-4, 20.0)
    f = 8 * math.sqrt(i) / Q
    R_ln = (n / f) ** (3 / 8)                   # mặt cắt lợi nhất: Q = 4m₀·√i·R^(8/3)/n
    return dict(Q=Q, f=f, R_ln=R_ln, B=B, B_R=B / R_ln, h_R=h0 / R_ln, h0=h0)


def h_co_hep(Q, H0, B=TC['b_doc_dau'], P=TC['P2']):
    """Độ sâu co hẹp ở chân bậc (đầu dốc): Q = φ·B·h·√(2g(E₀ − h)), E₀ = P + H₀ – lấy nghiệm nhỏ (dòng xiết)."""
    E0 = P + H0
    return brentq(lambda h: PHI * B * h * math.sqrt(2 * G * (E0 - h)) - Q, 1e-4, 2 * E0 / 3), E0


def buoc_co_dinh(Q, x1, h1, dL):
    """Chọn trước ΔL, tìm độ sâu cuối đoạn (nhánh dòng xiết): E₁ + i·ΔL = E₂ + J̄·ΔL."""
    a = yeu_to(Q, B_doc(x1), h1)
    B2 = B_doc(x1 + dL)
    g_ = lambda h: a['E'] + TC['i_doc'] * dL - (yeu_to(Q, B2, h)['E'] + 0.5 * (a['J'] + yeu_to(Q, B2, h)['J']) * dL)
    return brentq(g_, 0.02, h_k(Q, B2))


def duong_mat_nuoc(Q, h_dau):
    """Bảng cộng trực tiếp, chọn trước ΔL (đoạn thu hẹp L_tt/8 = 2,5 m; đoạn lăng trụ (L_dốc − L_tt)/8 = 5 m) và thử dần độ
    sâu cuối đoạn – có chiều sâu tại các mặt cắt kiểm tra tường bên x = 0, 5, …, 60 m (như file Excel)."""
    L_tt, L_d = TC['L_chuyen_tiep'], TC['L_doc']
    hang = [dict(x=0.0, **yeu_to(Q, B_doc(0), h_dau))]
    x, h = 0.0, h_dau
    for k in range(2 * N_BUOC_LT):
        dL = L_tt / N_BUOC_LT if k < N_BUOC_LT else (L_d - L_tt) / N_BUOC_LT
        h = buoc_co_dinh(Q, x, h, dL)
        x = round(x + dL, 9)
        hang.append(dict(x=x, **yeu_to(Q, B_doc(x), h)))
    for a, b in zip(hang, hang[1:]):
        b['dE'] = b['E'] - a['E']
        b['Jtb'] = 0.5 * (a['J'] + b['J'])
        b['i_J'] = TC['i_doc'] - b['Jtb']
        b['dL'] = b['x'] - a['x']
    return hang


def thuy_luc_doc(ts):
    Q_tk, Q_kt = ts['P1_BinhThuong']['qtc'], ts['P02_BinhThuong']['qtc']
    cap = [('0,4Q', 0.4 * Q_tk), ('0,6Q', 0.6 * Q_tk), ('0,8Q', 0.8 * Q_tk), ('Q (P = 1%)', Q_tk), ('Q (P = 0,2%)', Q_kt)]
    kq = []
    for ten, Q in cap:
        H0 = H0_theo_Q(Q, ts)
        hc, E0 = h_co_hep(Q, H0)
        bang = duong_mat_nuoc(Q, hc)
        kq.append(dict(ten=ten, Q=Q, H0=H0, E0=E0, hc=hc, hk_dau=h_k(Q, TC['b_doc_dau']), ik=i_k(Q, TC['b_doc_cuoi']),
                       ik_dau=i_k(Q, TC['b_doc_dau']), h0=h_0(Q, TC['b_doc_cuoi']), bang=bang))
    # tường bên với lưu lượng lớn nhất qua tràn chính (lũ kiểm tra)
    lon = kq[-1]
    mn = {r['x']: r['h'] for r in lon['bang']}
    tuong = []
    for j, x in enumerate(range(0, int(TC['L_doc']) + 1, 5)):
        y = yeu_to(lon['Q'], B_doc(x), mn[float(x)])
        hhk = y['h'] * (1 + y['v'] / 100)
        ht_ht = TC['h_tuong'][j] if j < len(TC['h_tuong']) else TC['h_tuong'][-1]
        tuong.append(dict(x=x, B=y['B'], h=y['h'], v=y['v'], hhk=hhk, ht_yc=hhk + A_TUONG, ht_ht=ht_ht,
                          du=ht_ht - hhk, gia_thiet=j >= len(TC['h_tuong'])))
    # chiều dày bản đáy (Đômbrốpxki) với lưu lượng lớn nhất
    d, c = lon['bang'][0], lon['bang'][-1]
    v_tb, h_tb = (d['v'] + c['v']) / 2, (d['h'] + c['h']) / 2
    ban_day = dict(Q=lon['Q'], v_dau=d['v'], v_cuoi=c['v'], v_tb=v_tb, h_dau=d['h'], h_cuoi=c['h'], h_tb=h_tb,
                   t=(0.03 * ALPHA_NEN * v_tb * math.sqrt(h_tb), 0.035 * ALPHA_NEN * v_tb * math.sqrt(h_tb)))
    return dict(cap=kq, tuong=tuong, ban_day=ban_day)


if __name__ == '__main__':
    ts = doc_excel()
    k = kha_nang_thao(ts)
    print('B = %.2f, Z đáy kênh %.2f, B_TL = %.2f, Σb = %.2f' % (k['B'], k['Z_day'], k['B_TL'], k['Sb']))
    for p, r in k['kq'].items():
        print('P=%-4s Z %.3f H %.3f q %.2f | H_TL %.3f Ω %.2f 4ΣbH %.2f V0 %.3f hv %.4f H0 %.3f ε %.4f Qtt %.2f δ %.2f%%'
              % (p, r['Z'], r['H'], r['q'], r['H_TL'], r['Om'], r['gh'], r['V0'], r['hv'], r['H0'], r['eps'], r['Qtt'],
                 r['sai_so']))
    tl = thuy_luc_doc(ts)
    for c in tl['cap']:
        b = c['bang']
        print('%-13s Q %7.2f H0 %.3f E0 %.3f hc %.3f hk31 %.3f | B18: hk %.3f ik %.5f h0 %.3f (f %.4f Rln %.3f h/R %.3f) | '
              'h20 %.3f h60 %.3f v60 %.2f | %d dòng' % (c['ten'], c['Q'], c['H0'], c['E0'], c['hc'], c['hk_dau'], c['ik']['hk'],
                                                     c['ik']['ik'], c['h0']['h0'], c['h0']['f'], c['h0']['R_ln'],
                                                     c['h0']['h_R'], [r for r in b if abs(r['x'] - 20) < 1e-6][0]['h'],
                                                     b[-1]['h'], b[-1]['v'], len(b)))
    for r in tl['tuong']:
        print('  x %2d B %5.2f h %.3f v %.2f hhk %.3f yêu cầu %.2f hiện trạng %.2f dư %.2f%s'
              % (r['x'], r['B'], r['h'], r['v'], r['hhk'], r['ht_yc'], r['ht_ht'], r['du'], ' (giả thiết)' if r['gia_thiet'] else ''))
    print('bản đáy', {k_: (round(v, 3) if isinstance(v, float) else tuple(round(x, 3) for x in v)) for k_, v in tl['ban_day'].items()})
