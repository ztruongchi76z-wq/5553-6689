# -*- coding: utf-8 -*-
"""
XÁC ĐỊNH CAO TRÌNH ĐỈNH ĐẬP (mục 5.1 của đồ án)
TCVN 8216:2018 điều 10.1:   Zđ = Zh + Δh + hsl + a
  Z1 = MNDBT + Δh + hsl + a      (gió P = 4 %)
  Z2 = MNLTK + Δh' + hsl' + a'   (gió P = 50 %)
  Z3 = MNLKT + a''
  a, a', a'' – chiều cao an toàn, Bảng 3 TCVN 8216:2018, đập cấp II: 0,7; 0,5; 0,2 m
Sóng do gió và sóng leo: TCVN 8421:2010 (phương pháp của SNiP 2.06.04-82*):
  - Nước dềnh do gió  Δh = kw·V²·D/(g·H)·cosα ,  kw = 3·10⁻⁷·(1 + 0,3V)
  - Sóng nước sâu (H > 0,5λ): gh/V² và gτ/V tra theo đường bao trên của đồ thị theo gt/V và gD/V²
    (dùng công thức xấp xỉ của Krưlốp: gh/V² = 0,16·{1 − [1 + 0,006·(gD/V²)^0,5]^−2};
     gτ/V = 19,1·(gh/V²)^0,625), lấy cặp giá trị nhỏ hơn
  - h1% = K1%·h ; sóng leo h_sl1% = kr·kp·ksp·krun·kα·h1% (công thức 25; kα theo Bảng 9 – góc giữa front sóng
    và pháp tuyến mái)
  - Vận tốc gió tính toán (Phụ lục A, A.3.3): Vw = kfl·kl·Vl; Vl – vận tốc gió trạm (Bảng 1.4, trạm Phan Thiết);
    kl theo Bảng A3, địa hình dạng A (vùng ven biển trống trải): 1,10 (Vl ≤ 15 m/s) ÷ 1,09 (20 ÷ 35 m/s);
    kfl = 0,675 + 4,5/Vl ≤ 1 (quy đổi số liệu phong tiêu) – lấy bằng 1 (thiên an toàn, vì chưa rõ dụng cụ đo).
  Sửa 12/10/2026 khi đối chiếu toàn văn TCVN 8421:2010: kα trước đây nội suy sai bảng (thiên lớn 5 ÷ 11 %),
  trước đây chưa quy đổi Vl → Vw (thiên nhỏ ≈ 10 %).
"""
import numpy as np
import so_lieu as S

G = S.G
A_AN_TOAN = {'II': (0.70, 0.50, 0.20)}   # TCVN 8216:2018 Bảng 3


def kw(V):
    return 3e-7 * (1 + 0.3 * V)      # Bảng A2: 2,1·10⁻⁶ (20 m/s) … 4,8·10⁻⁶ (50 m/s)


KW_V, KW_K = [20, 30, 40, 50], [2.1e-6, 3.0e-6, 3.9e-6, 4.8e-6]     # TCVN 8421:2010 Bảng A2


def kw_bang(V):
    """kw tra Bảng A2 (nội suy; V ≤ 20 m/s lấy 2,1·10⁻⁶ – bảng không có giá trị nhỏ hơn)."""
    return float(np.interp(V, KW_V, KW_K))


def V_tinh_toan(Vl, dung_kfl=False):
    """Vận tốc gió tính toán Vw = kfl·kl·Vl (TCVN 8421:2010, A.3.3), địa hình dạng A (Bảng A3)."""
    kl = float(np.interp(Vl, [10, 15, 20, 25, 30, 35, 40], [1.10, 1.10, 1.09, 1.09, 1.09, 1.09, 1.08]))
    kfl = min(0.675 + 4.5 / Vl, 1.0) if dung_kfl else 1.0
    return kfl * kl * Vl


def k_alpha(goc):
    """Hệ số giảm sóng leo theo góc tới của front sóng (TCVN 8421:2010, Bảng 9)."""
    return float(np.interp(goc, [0, 10, 20, 30, 40, 50, 60], [1.0, 0.98, 0.96, 0.92, 0.87, 0.82, 0.76]))


def song_nuoc_sau(V, D, t_h):
    """Trả về (h_tb, tau_tb) theo hai điều kiện thời gian gió thổi và đà gió, lấy giá trị nhỏ hơn."""
    gD = G * D / V ** 2
    a_D = 0.16 * (1 - (1 + 0.006 * np.sqrt(gD)) ** -2)
    # theo thời gian: dùng quan hệ đà gió tương đương (đường bao trên đồ thị) gt/V -> gD/V² tương đương
    gt = G * t_h * 3600 / V
    # quan hệ đường bao đồ thị theo thời gian (xấp xỉ): gh/V² = 0,16·{1 − [1 + 0,006·(gDt/V²)^0,5]^−2}
    # với gDt/V² = (gt/V / 4,03)^(1,6) (đà gió tương đương với thời gian gió thổi t)
    gDt = (gt / 4.03) ** 1.6
    a_t = 0.16 * (1 - (1 + 0.006 * np.sqrt(gDt)) ** -2)
    a = min(a_D, a_t)
    h = a * V ** 2 / G
    tau = 19.1 * a ** 0.625 * V / G
    return h, tau, gD, gt, a_D, a_t


def K1(gD):
    """Hệ số K1% (đồ thị Hình 2 TCVN 8421:2010) – nội suy theo log(gD/V²)."""
    xs = np.log10([20, 100, 153, 514, 1000, 5000])
    ys = [2.00, 2.08, 2.10, 2.13, 2.15, 2.18]
    return float(np.interp(np.log10(gD), xs, ys))


def kr_kp(rel):
    """Hệ số độ nhám kr và thấm kp của lớp gia cố mái theo độ nhám tương đối Δ/h1% (bảng TCVN 8421:2010)."""
    xs = [0.0, 0.002, 0.005, 0.01, 0.05, 0.1, 0.2]
    kr = [1.0, 1.0, 0.95, 0.90, 0.80, 0.75, 0.70]
    kp = [0.9, 0.9, 0.85, 0.80, 0.70, 0.60, 0.50]
    # bảng gốc: <0,002: 1/0,9 ; 0,005–0,01: 0,95/0,85 ; 0,02: 0,9/0,8 ; 0,05: 0,8/0,7 ; 0,1: 0,75/0,6 ; >0,2: 0,7/0,5
    xs = [0.002, 0.0075, 0.02, 0.05, 0.1, 0.2]
    kr = [1.00, 0.95, 0.90, 0.80, 0.75, 0.70]
    kp = [0.90, 0.85, 0.80, 0.70, 0.60, 0.50]
    return float(np.interp(rel, xs, kr)), float(np.interp(rel, xs, kp))


def ksp(V, m):
    """Hệ số ksp theo vận tốc gió và hệ số mái (bảng TCVN 8421:2010), hàng m = 3÷5."""
    Vs = [5, 10, 20]
    if m < 3:
        k = [1.1, 1.1, 1.4]
    elif m <= 5:
        k = [0.8, 1.1, 1.5]
    else:
        k = [0.6, 1.2, 1.6]
    return float(np.interp(V, Vs, k))


def krun(lam_h, m):
    """Hệ số sóng leo k_run theo λ/h1% với mái m = 3 (tra đồ thị Hình 10 TCVN 8421:2010, nội suy)."""
    xs = [5, 8, 10, 12, 15, 20, 25]
    ys = [1.20, 1.40, 1.48, 1.57, 1.70, 1.84, 1.95]
    return float(np.interp(lam_h, xs, ys))


def tinh_truong_hop(Z, V, D, t_h=6.0, Z_day=59.0, m=3.0, delta=0.02, alpha=0.0, quy_doi_gio=True, kw_tra_bang=False):
    """V – vận tốc gió trạm Vl; khi quy_doi_gio = True tính với Vw = kfl·kl·Vl (A.3.3).
    kw_tra_bang = True: kw tra Bảng A2 (bản 27/10 cho SV); False: đường thẳng qua các điểm của bảng (bản cũ)."""
    V_tram = V
    if quy_doi_gio:
        V = V_tinh_toan(V)
    H = Z - Z_day
    k_w = kw_bang(V) if kw_tra_bang else kw(V)
    dh = k_w * V ** 2 * D / (G * H) * np.cos(np.radians(alpha))
    h, tau, gD, gt, aD, at = song_nuoc_sau(V, D, t_h)
    lam = G * tau ** 2 / (2 * np.pi)
    sau = H > 0.5 * lam
    k1 = K1(gD)
    h1 = k1 * h
    kr, kp = kr_kp(delta / h1)
    ks = ksp(V, m)
    kru = krun(lam / h1, m)
    ka = k_alpha(alpha)
    hsl = kr * kp * ks * kru * ka * h1
    return dict(Z=Z, H=H, V_tram=V_tram, V=V, D=D, kw=k_w, dh=dh, gt=gt, gD=gD, a_D=aD, a_t=at, h=h, tau=tau,
                lam=lam, sau=sau, K1=k1, h1=h1, kr=kr, kp=kp, ksp=ks, krun=kru, ka=ka, hsl=hsl)


def cao_trinh(MNDBT, MNLTK, MNLKT, V4, V50, D_dbt, D_tk, D_kt, cap='II'):
    a, a1, a2 = A_AN_TOAN[cap]
    r1 = tinh_truong_hop(MNDBT, V4, D_dbt)
    r2 = tinh_truong_hop(MNLTK, V50, D_tk)
    Z1 = MNDBT + r1['dh'] + r1['hsl'] + a
    Z2 = MNLTK + r2['dh'] + r2['hsl'] + a1
    Z3 = MNLKT + a2
    return dict(r1=r1, r2=r2, Z1=Z1, Z2=Z2, Z3=Z3, Zmax=max(Z1, Z2, Z3), a=(a, a1, a2),
                Z_tcs_dat=MNLKT + 0.30)


if __name__ == '__main__':
    V4 = S.GIO[4][S.GIO_HUONG.index('N')]
    V50 = S.GIO[50][S.GIO_HUONG.index('N')]
    for lab, (z1, z2) in {'Hồ sơ BCNCKT': (77.43, 78.25), 'Đồ án (KTTV)': (77.63, 78.54),
                          'Đồ án (QTVH)': (77.50, 78.44)}.items():
        r = cao_trinh(S.MNDBT, z1, z2, V4, V50, 2970, 3025, 3047)
        print(lab, 'Z1 %.2f Z2 %.2f Z3 %.2f -> %.2f ; MNLKT+0.3 = %.2f' % (r['Z1'], r['Z2'], r['Z3'], r['Zmax'], r['Z_tcs_dat']))
        for k in ['r1', 'r2']:
            x = r[k]
            print('   ', k, ' '.join('%s=%.4g' % (kk, vv) for kk, vv in x.items() if not isinstance(vv, (bool, np.bool_))))
