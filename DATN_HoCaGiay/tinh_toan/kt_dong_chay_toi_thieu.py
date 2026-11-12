# -*- coding: utf-8 -*-
"""
KIỂM TRA ĐIỀU TIẾT TOÀN CHUỖI CÓ KỂ DÒNG CHẢY TỐI THIỂU (mục 2.2.3) – TCVN 10778:2024, công thức (6), (9), (14)
  Q_ri = Q_yi + Q_tti + Q_bi + Q_ti + Q_xi : nhu cầu QTVH (mục 4.4.3, chưa gồm Q_tt) + Q_tt = 0,027 m³/s [QTVH mục 3.4.2]
  Thuật toán, số liệu như can_bang_nuoc.py (chuỗi tháng 44 năm, ΔZ, k_t = 1 %/tháng, P = m/(n + 1)).
  Điều kiện (9): m ≥ (n + 1)·P_CN/100 = 45 × 0,85 = 38,25 → ít nhất 39 năm đủ nước.

  python3 kt_dong_chay_toi_thieu.py [du_lieu/can_bang_nuoc.json]
Dữ liệu du_lieu/can_bang_nuoc.json đi kèm can_bang_nuoc.py (thư mục tinh_toan trên Drive).
"""
import json
import os
import sys

import numpy as np
import so_lieu as S

DIR = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, 'du_lieu', 'can_bang_nuoc.json'),
                   encoding='utf-8'))
DZ = np.array([68.4, 64.4, 76.9, 66.6, 62.3, 55.1, 49.8, 51.8, 43.9, 40.8, 52.2, 66.7]) / 1000.0   # m [QTVH Bảng 3-24]
W_DEN = np.array([r[1:13] for r in D['dong_chay']])                                                # 10^6 m³
NHU_CAU = {}
for k, v in D['nhu_cau'].items():
    ten = 'TH1' if 'TH1 ' in v['nhan'][0] else ('TH1A' if 'TH1A' in v['nhan'][0] else 'TH2')
    NHU_CAU[ten] = np.array(v['tong'][:12])
NGAY = np.array([31, 28.25, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
Q_TT = 0.027                                        # m³/s – dòng chảy tối thiểu tuyến đập [QTVH Bảng 3-26]
W_TT = Q_TT * 86400 * NGAY / 1e6                    # 10^6 m³/tháng


def dieu_tiet_chuoi(q, V_c, V_bt, k_t=0.01):
    n = W_DEN.shape[0]
    V = V_bt
    thieu = np.zeros((n, 12))
    for j in range(n):
        for i in range(12):
            Vtb = V
            for _ in range(3):
                Z = S.Z_of_V(Vtb)
                Wbh = DZ[i] * S.F_of_Z(Z)
                Wt = k_t * Vtb
                V1 = V + W_DEN[j, i] - q[i] - Wbh - Wt
                Vtb = 0.5 * (V + min(max(V1, V_c), V_bt))
            if V1 > V_bt:
                V1 = V_bt
            if V1 < V_c:
                thieu[j, i] = V_c - V1
                V1 = V_c
            V = V1
    nam_thieu = (thieu > 1e-6).any(1)
    m = int((~nam_thieu).sum())
    return dict(n=n, m=m, P=100.0 * m / (n + 1), P_mn=100.0 * m / n, thang_thieu=int((thieu > 1e-6).sum()),
                W_thieu=float(thieu.sum()), W_dung=float(q.sum()))


def mndbt_can(q, V_c, P_yc=85.0):
    lo, hi = 70.0, 80.0
    for _ in range(30):
        mid = 0.5 * (lo + hi)
        lo, hi = (lo, mid) if dieu_tiet_chuoi(q, V_c, float(S.V_of_Z(mid)))['P'] >= P_yc else (mid, hi)
    return hi


if __name__ == '__main__':
    V_c, V_bt = float(S.V_of_Z(S.MNC)), float(S.V_of_Z(S.MNDBT))
    out = dict(Q_tt=Q_TT, W_tt_nam=float(W_TT.sum()), m_yeu_cau=45 * 0.85)
    for ten, q in NHU_CAU.items():
        for nhan, qq in (('khong_Qtt', q), ('co_Qtt', q + W_TT)):
            r = dieu_tiet_chuoi(qq, V_c, V_bt)
            r['MNDBT_can'] = mndbt_can(qq, V_c)
            out['%s|%s' % (ten, nhan)] = r
            print('%-5s %-9s Wdùng %6.2f  m = %d/%d  P = %.1f%%  tháng thiếu %d  ΣWthiếu %.2f  MNDBT cần %.2f m' % (
                ten, nhan, r['W_dung'], r['m'], r['n'], r['P'], r['thang_thieu'], r['W_thieu'], r['MNDBT_can']))
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kt_dong_chay_toi_thieu.json'), 'w'), ensure_ascii=False,
              indent=1, default=float)
