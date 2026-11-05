# Kiem tra anh huong cao trinh mat dap (+78,75 -> +78,85) trong mo hinh chuyen de
import numpy as np, importlib, sys
import chuyen_de_khoan_phut as K
def chay(zc):
    K.CD[3] = zc; K.CD[4] = zc
    K.MAT_Z = np.concatenate([[70.0, 70.0], K.CD, [73.70]])
    MAN = dict(xa=K.X_TIM - 2.4, xb=K.X_TIM + 2.4, ztop=77.00, zbot=68.30)
    m0 = K.luoi(None); m1 = K.luoi(MAN)
    out = {}
    for Zw in (74.70, 77.63, 78.54):
        _, _, r0 = K.giai(Zw, None, msh=m0); _, _, r1 = K.giai(Zw, MAN, msh=m1)
        out[Zw] = (r0['q_ngay'], r1['q_ngay'], 1 - r1['q_in']/r0['q_in'], r0['J_ra'], r1['J_ra'], r0.get('z_seep_max'), r1.get('z_seep_max'), r1.get('dH_man_max',0), r1.get('J_man_tb',0))
        print(zc, Zw, 'q0=%.4f q1=%.4f eta=%.0f%% Jra %.2f/%.2f zseep %s/%s dHm=%.2f Jm=%.2f' % out[Zw])
    return out
a = chay(78.75); b = chay(78.85)
