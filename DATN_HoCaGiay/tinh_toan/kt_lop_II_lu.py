# -*- coding: utf-8 -*-
"""Kiểm tra (góp ý 11/10): hiệu quả màng khi lớp II thấm hơn giả thiết 1e-5 cm/s.
Tiêu chí hiện trường q ≤ 5 Lu (TCVN 8216:2018 điều 9.2.2.9) tương ứng K lớp II tới ≈ 6,5e-5 cm/s (1 Lu ≈ 1,3e-5 cm/s),
nên tính lại η cho K_II = 1e-5; 3e-5; 6,5e-5; 1e-4 cm/s với màng TKCS (3 hàng, đáy +68,3) và các phương án hợp lý hóa
(đáy +69,4 = mặt lớp II − 3 m; 2 hàng A, B) tại mặt cắt đại diện, MNLTK và MNLKT."""
import json
import os
import chuyen_de_khoan_phut as K

XA, XB = 16.31, 17.81
PA = {'3h_TK': dict(xa=K.X_TIM - 2.4, xb=K.X_TIM + 2.4, ztop=77.0, zbot=68.3),
      '3h_day_69_4': dict(xa=K.X_TIM - 2.4, xb=K.X_TIM + 2.4, ztop=77.0, zbot=69.4),
      '2h_A_B': dict(xa=XA - 0.9, xb=XB + 0.9, ztop=77.0, zbot=68.3),
      '2h_A_B_day_69_4': dict(xa=XA - 0.9, xb=XB + 0.9, ztop=77.0, zbot=69.4)}
out = {}
msh0 = K.luoi(None)
for KII in (1e-5, 3e-5, 6.5e-5, 1e-4):
    for Zw in (77.63, 78.54):
        q0 = K.giai(Zw, None, K_II=KII, msh=msh0)[2]
        for ten, man in PA.items():
            r = K.giai(Zw, man, K_II=KII)[2]
            k = '%.1e_%.2f_%s' % (KII, Zw, ten)
            out[k] = dict(K_II=KII, Zw=Zw, pa=ten, q0=q0['q_ngay'], q=r['q_ngay'], eta=1 - r['q_in'] / q0['q_in'],
                          Jra=r['J_ra'], Jra0=q0['J_ra'], zseep=r.get('z_seep_max'), zseep0=q0.get('z_seep_max'))
            print('%-34s η = %5.1f%%  Jra = %.2f (chưa màng %.2f)  điểm ra %s / %s' % (
                k, 100 * out[k]['eta'], r['J_ra'], q0['J_ra'], r.get('z_seep_max'), q0.get('z_seep_max')), flush=True)
json.dump(out, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ket_qua', 'kt_lop_II_lu.json'), 'w'), indent=1,
          default=float)
