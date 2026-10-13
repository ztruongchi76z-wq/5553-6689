# -*- coding: utf-8 -*-
"""
KIỂM TRA BỔ SUNG CẤP NƯỚC (mục 2.2.3, 6.3.1 – tiếp thu góp ý 12/10/2026)
 (1) Hồ đa mục tiêu, mức bảo đảm khác nhau (tưới 85 %, sinh hoạt 90 %) – TCVN 10778:2024 mục 8.2.2.2:
     - Cách 1: P_CN = Σ(P_j·W_qj)/ΣW_qj (công thức 10), sau đó kiểm tra như hồ đơn mục tiêu (công thức 9, 14);
     - Kiểm tra theo thứ tự ưu tiên (tinh thần cách 2 – hồ ảo): tháng thiếu nước cắt giảm tưới trước, sinh hoạt
       chỉ thiếu khi lượng cấp được trong tháng nhỏ hơn nhu cầu sinh hoạt -> mức bảo đảm riêng của sinh hoạt.
 (2) Mực nước hồ cuối tháng thấp hơn +67,94 m (cống chỉ tháo < 10 m³/s, mục 6.3.1) và lưu lượng bình quân tháng
     theo nhu cầu.
 Thuật toán điều tiết như kt_dong_chay_toi_thieu.py (k_t = 1 %/tháng, ΔZ bốc hơi phụ thêm, P = m/(n + 1)).
"""
import json
import os

import numpy as np
import so_lieu as S
import kt_dong_chay_toi_thieu as K

DIR = os.path.dirname(os.path.abspath(__file__))
D = K.D
NN, SH = {}, {}
for k, v in D['nhu_cau'].items():
    ten = 'TH1' if 'TH1 ' in v['nhan'][0] else ('TH1A' if 'TH1A' in v['nhan'][0] else 'TH2')
    NN[ten] = np.array(v['nn'][:12])
    SH[ten] = np.array(v['sh'][:12])
Z_CONG_10 = 65.60 + (10.0 / (0.35 * 1.8 * np.sqrt(2 * S.G))) ** (2 / 3)   # +67,94 m (công thức 6-1)


def mo_phong(q, V_c, V_bt, k_t=0.01):
    """Như K.dieu_tiet_chuoi nhưng trả về quá trình: V cuối tháng, lượng thiếu, lượng cấp được."""
    n = K.W_DEN.shape[0]
    V = V_bt
    thieu = np.zeros((n, 12))
    Vc = np.zeros((n, 12))
    for j in range(n):
        for i in range(12):
            Vtb = V
            for _ in range(3):
                Z = S.Z_of_V(Vtb)
                Wbh = K.DZ[i] * S.F_of_Z(Z)
                Wt = k_t * Vtb
                V1 = V + K.W_DEN[j, i] - q[i] - Wbh - Wt
                Vtb = 0.5 * (V + min(max(V1, V_c), V_bt))
            if V1 > V_bt:
                V1 = V_bt
            if V1 < V_c:
                thieu[j, i] = V_c - V1
                V1 = V_c
            V = V1
            Vc[j, i] = V
    cap = q[None, :] - thieu
    return thieu, cap, Vc


if __name__ == '__main__':
    V_c, V_bt = float(S.V_of_Z(S.MNC)), float(S.V_of_Z(S.MNDBT))
    out = dict(Z_cong_10=float(Z_CONG_10))
    n = K.W_DEN.shape[0]
    for ten in ('TH1', 'TH1A', 'TH2'):
        for nhan, them in (('khong_Qtt', 0.0), ('co_Qtt', K.W_TT)):
            q = NN[ten] + SH[ten] + them
            thieu, cap, Vc = mo_phong(q, V_c, V_bt)
            W_nn, W_sh = NN[ten].sum(), SH[ten].sum()
            P_chung = (85.0 * W_nn + 90.0 * W_sh) / (W_nn + W_sh)          # công thức (10)
            m_yc = (n + 1) * P_chung / 100.0                                  # công thức (9)
            nam_du = ~(thieu > 1e-6).any(1)
            m = int(nam_du.sum())
            # ưu tiên sinh hoạt (và dòng chảy tối thiểu nếu có): sinh hoạt thiếu khi lượng cấp được < nhu cầu SH
            sh_thang = SH[ten][None, :] + (them if np.isscalar(them) else them[None, :]) * 0 
            thieu_sh = cap < (SH[ten][None, :] - 1e-9)
            m_sh = int((~thieu_sh.any(1)).sum())
            ty_le_cat_nn = np.where(thieu > 1e-6, thieu / NN[ten][None, :], 0.0)
            Zc = S.Z_of_V(Vc)
            mua_kho = [0, 1, 2, 3]   # tháng 1 – 4 (mùa khô, nhu cầu tưới lớn)
            r = dict(P_chung=P_chung, m_yeu_cau=m_yc, m=m, P=100.0 * m / (n + 1),
                     m_sinh_hoat=m_sh, P_sinh_hoat=100.0 * m_sh / (n + 1),
                     cap_nho_nhat_thang_thieu=float(cap[thieu > 1e-6].min()) if (thieu > 1e-6).any() else None,
                     cat_giam_nn_lon_nhat=float(ty_le_cat_nn.max()),
                     thang_Z_duoi_6794=int((Zc < Z_CONG_10).sum()),
                     thang_Z_duoi_6794_mua_kho=int((Zc[:, mua_kho] < Z_CONG_10).sum()),
                     nam_co_thang_Z_duoi_6794=int((Zc < Z_CONG_10).any(1).sum()),
                     Z_cuoi_thang_nho_nhat=float(Zc.min()),
                     Q_tb_thang_lon_nhat=float((q / (K.NGAY * 86400 / 1e6)).max()))
            out['%s|%s' % (ten, nhan)] = r
            print('%-4s %-9s P_chung %.2f%% (m ≥ %.2f) m=%d P=%.1f%% | SH ưu tiên: m=%d P=%.1f%% | cấp nhỏ nhất tháng thiếu %.2f'
                  ' | cắt tưới max %.0f%% | tháng Z<%.2f: %d (mùa khô %d; %d năm) | Zmin %.2f | Qtb tháng max %.2f m3/s'
                  % (ten, nhan, P_chung, m_yc, m, r['P'], m_sh, r['P_sinh_hoat'], r['cap_nho_nhat_thang_thieu'] or 0,
                     100 * r['cat_giam_nn_lon_nhat'], Z_CONG_10, r['thang_Z_duoi_6794'], r['thang_Z_duoi_6794_mua_kho'],
                     r['nam_co_thang_Z_duoi_6794'], r['Z_cuoi_thang_nho_nhat'], r['Q_tb_thang_lon_nhat']))
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kt_cap_nuoc_bo_sung.json'), 'w'), ensure_ascii=False,
              indent=1, default=float)
