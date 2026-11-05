# -*- coding: utf-8 -*-
"""Kiểm tra lại thủy lực tràn sự cố (góp ý 11/10).
Tràn sự cố b = 5 m, ngưỡng +71,00, không có bậc hạ sau ngưỡng (P = 0): mặt cắt khống chế ở đầu dốc nước, độ sâu đầu
dốc lấy bằng độ sâu phân giới hk (dốc i = 0,08 > ik). Đường mặt nước trên dốc (L = 32 m, n = 0,015) tính bằng phương
pháp cộng trực tiếp như thuy_luc_tran.py, bắt đầu từ 0,99·hk (đường nước hạ bII).
Ghi thêm: lưu lượng lớn nhất theo phương trình (4-1) với E0 = H0, φ = 0,95 (hc = 2E0/3) và nghiệm dòng xiết nếu có."""
import json
import os
import numpy as np

G = 9.81
B, L, I, N = 5.0, 32.0, 0.08, 0.015
KQ = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ket_qua')


def hk(Q, b=B):
    return (Q ** 2 / (G * b ** 2)) ** (1 / 3)


def dt(h, Q, b=B, n=N):
    A = b * h; R = A / (b + 2 * h); v = Q / A
    return dict(v=v, J=v ** 2 * n ** 2 / R ** (4 / 3), E=h + v ** 2 / (2 * G))


def duong_mat_nuoc(Q, h0, dL=2.0):
    xs, hs, x, h = [0.0], [h0], 0.0, h0
    while x < L - 1e-9:
        dx = min(dL, L - x); d1 = dt(h, Q)
        lo, hi = 0.05 * h, 0.999 * hk(Q)
        for _ in range(80):
            h2 = 0.5 * (lo + hi); d2 = dt(h2, Q)
            f = d2['E'] + 0.5 * (d1['J'] + d2['J']) * dx - d1['E'] - I * dx
            lo, hi = (h2, hi) if f > 0 else (lo, h2)
        h = 0.5 * (lo + hi); x += dx; xs.append(x); hs.append(h)
    return [dict(x=x, h=h, v=dt(h, Q)['v'], h_hk=h * (1 + dt(h, Q)['v'] / 100)) for x, h in zip(xs, hs)]


out = {}
for ten, Q, H0 in (('P=1%', 138.6, 6.63), ('P=0,2%', 162.7, 7.54)):
    phi = 0.95
    hm = 2 * H0 / 3
    Qmax = phi * B * hm * np.sqrt(2 * G * (H0 - hm))
    nghiem = None
    if Qmax > Q:
        lo, hi = 1e-4, hm
        for _ in range(100):
            h = 0.5 * (lo + hi)
            lo, hi = (lo, h) if phi * B * h * np.sqrt(2 * G * (H0 - h)) > Q else (h, hi)
        nghiem = 0.5 * (lo + hi)
    hK = hk(Q)
    prof = duong_mat_nuoc(Q, 0.99 * hK)
    cuoi = prof[-1]
    out[ten] = dict(Q=Q, H0=H0, Qmax_41=Qmax, nghiem_41=nghiem, hk=hK, h_dau=prof[0]['h'], v_dau=prof[0]['v'],
                    tuong_dau=prof[0]['h_hk'] + 0.5, h_cuoi=cuoi['h'], v_cuoi=cuoi['v'], tuong_cuoi=cuoi['h_hk'] + 0.5,
                    m_tuong_duong=Q / (B * np.sqrt(2 * G) * H0 ** 1.5), m_max=phi * (2 / 3) * np.sqrt(1 / 3),
                    prof=prof)
    print('%-7s Q=%.1f H0=%.2f  Qmax(4-1)=%.2f  nghiệm=%s  hk=%.3f  m=%.4f (m_max %.4f)' % (
        ten, Q, H0, Qmax, None if nghiem is None else round(nghiem, 3), hK, out[ten]['m_tuong_duong'], out[ten]['m_max']))
    for p in prof[::4]:
        print('    x=%4.1f h=%.3f v=%.2f hhk=%.3f tường yêu cầu=%.2f' % (p['x'], p['h'], p['v'], p['h_hk'], p['h_hk'] + 0.5))
    print('    cuối: h=%.3f v=%.2f tường=%.2f' % (cuoi['h'], cuoi['v'], cuoi['h_hk'] + 0.5))
json.dump(out, open(os.path.join(KQ, 'kt_tran_su_co.json'), 'w'), indent=1, default=float)
