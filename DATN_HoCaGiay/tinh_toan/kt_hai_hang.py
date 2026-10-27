# -*- coding: utf-8 -*-
"""Kiểm tra (theo góp ý): màng 2 hàng giữ hàng A, B (cách 1,5 m, T = 3,3 m) so với màng 2 hàng đối xứng qua tim;
đáy màng theo TCVN 8216:2018 điều 9.2.2.9 (mặt lớp II − 3 m) cho màng 3 hàng."""
import json, numpy as np
import chuyen_de_khoan_phut as K
XA, XB, XC = 16.31, 17.81, 19.31
out = {}
for Zw in (77.63, 78.54):
    q0 = K.giai(Zw, None)[2]['q_in']
    cas = {'2h_doi_xung': dict(xa=K.X_TIM - 1.65, xb=K.X_TIM + 1.65, ztop=77.0, zbot=68.3),
           '2h_A_B': dict(xa=XA - 0.9, xb=XB + 0.9, ztop=77.0, zbot=68.3),
           '2h_B_C': dict(xa=XB - 0.9, xb=XC + 0.9, ztop=77.0, zbot=68.3),
           '3h_TK': dict(xa=K.X_TIM - 2.4, xb=K.X_TIM + 2.4, ztop=77.0, zbot=68.3),
           '3h_day_69_4': dict(xa=K.X_TIM - 2.4, xb=K.X_TIM + 2.4, ztop=77.0, zbot=69.4),
           '2h_A_B_day_69_4': dict(xa=XA - 0.9, xb=XB + 0.9, ztop=77.0, zbot=69.4)}
    for ten, man in cas.items():
        r = K.giai(Zw, man)[2]
        out['%.2f_%s' % (Zw, ten)] = dict(q=r['q_ngay'], eta=1 - r['q_in'] / q0, Jra=r['J_ra'], Jm=r.get('J_man_tb', 0),
                                          dH=r.get('dH_man_max', 0), zseep=r.get('z_seep_max'))
        print(Zw, ten, 'q=%.4f eta=%.1f%% Jra=%.2f Jm=%.2f dHm=%.2f seep=%s' % (
            r['q_ngay'], 100 * (1 - r['q_in'] / q0), r['J_ra'], r.get('J_man_tb', 0), r.get('dH_man_max', 0), r.get('z_seep_max')))
json.dump(out, open('../ket_qua/kt_hai_hang.json', 'w'), indent=1, default=float)
