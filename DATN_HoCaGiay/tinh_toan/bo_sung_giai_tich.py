# -*- coding: utf-8 -*-
"""
CÁC TÍNH TOÁN BỔ SUNG (giải tích) – hoàn thiện thuyết minh theo góp ý (10/10/2026)
 (1) Cao trình đỉnh đập theo các hướng gió thổi vào mái thượng lưu (tuyến đập gần Đông – Tây).
 (2) Độ nhạy MNLKT theo hệ số lưu lượng tràn m (0,45 đồ án / 0,467 BCNCKT).
 (3) Cống lấy nước khi độ dốc đáy i = 3 ‰: độ sâu dòng đều, lưu lượng lớn nhất còn chảy không áp.
 (4) Neo thép ngưỡng tràn chính: lực neo yêu cầu khi bỏ qua lực dính (c = 0), kiểm tra D25 a = 1,5 m.
 (5) Khung kín cống hộp (tham khảo): nội lực, cốt thép yêu cầu – tiêu chí đối chiếu khi kiểm tra cống.
 (6) Độ nhạy ứng suất tiếp dòng chảy tràn trên mái theo tỷ lệ mưa 30 phút / mưa ngày; lưu tốc rãnh dọc.
 (7) Mức bảo đảm cấp nước theo m/(n+1) và m/n.
"""
import json
import numpy as np
import so_lieu as S
import cao_trinh_dinh as CT
import dieu_tiet_lu as DT
import cong_va_tran_ket_cau as CV

G = S.G
KQ = {}

# ---------------------------------------------------------------- (1) hướng gió
# Tuyến đập (bình đồ TKCS, hệ VN-2000): từ (777,4; 213,3) đến (1870,9; 94,2) -> phương N96,2°E;
# pháp tuyến mái thượng lưu hướng về lòng hồ: phương vị 6,2° (lòng hồ ở phía Bắc đập).
p1, p2 = np.array([777.39, 213.26]), np.array([1870.92, 94.18])
d = p2 - p1
phuong_tuyen = (np.degrees(np.arctan2(d[0], d[1])) + 360) % 360          # phương vị tuyến đập
phap_tuyen = (phuong_tuyen - 90.0) % 360                                  # phương vị pháp tuyến TL
huong = dict(N=0, NE=45, E=90, SE=135, S=180, SW=225, W=270, NW=315)
gio = []
for h, az in huong.items():
    goc = abs((az - phap_tuyen + 180) % 360 - 180)                       # góc giữa hướng gió tới và pháp tuyến
    k = S.GIO_HUONG.index(h)
    V4, V50 = S.GIO[4][k], S.GIO[50][k]
    row = dict(huong=h, goc=round(goc, 1), V4=V4, V50=V50)
    if goc <= 60.0:            # sóng truyền vào mái (góc tới ≤ 60°): tính với đà gió lớn nhất (thiên lớn)
        r1 = CT.tinh_truong_hop(S.MNDBT, V4, 2970.0, alpha=goc)
        r2 = CT.tinh_truong_hop(77.63, V50, 3025.0, alpha=goc)
        row.update(Z1=S.MNDBT + r1['dh'] + r1['hsl'] + 0.70, Z2=77.63 + r2['dh'] + r2['hsl'] + 0.50,
                   hsl1=r1['hsl'], hsl2=r2['hsl'], ka=r1['ka'])
    gio.append(row)
KQ['gio'] = dict(phuong_tuyen=phuong_tuyen, phap_tuyen=phap_tuyen, bang=gio)
print('Tuyến đập phương vị %.1f°, pháp tuyến mái TL %.1f°' % (phuong_tuyen, phap_tuyen))
for r in gio:
    print(r)

# ---------------------------------------------------------------- (2) độ nhạy m tràn
tc0, sc0 = DT.tran_hien_trang()
out = {}
for m in (0.45, 0.467):
    tc = DT.Tran('TC', 74.70, 9, 2.40, m, 0.70, 0.45)
    sc = DT.Tran('SC', 71.00, 1, 5.00, m, 0.70, 0.45)
    r1 = DT.dieu_tiet(S.Q_KTTV['1%'], tc, sc)
    r2 = DT.dieu_tiet(S.Q_KTTV['0.2%'], tc, sc)
    r3 = DT.dieu_tiet(S.Q_KTTV['0.2%'], tc, sc, ket_van=True)
    out[m] = dict(MNLTK=r1['Zmax'], MNLKT=r2['Zmax'], ket=r3['Zmax'])
    print('m = %.3f: MNLTK %.3f MNLKT %.3f (kẹt van %.3f)' % (m, r1['Zmax'], r2['Zmax'], r3['Zmax']))
KQ['do_nhay_m'] = out

# ---------------------------------------------------------------- (3) cống i = 3 ‰
cong = {}
for i in (0.03, 0.003):
    rr = {}
    for Q in (6.54, 7.31, 10.0):
        h0 = CV.h_chuan(Q, 1.8, i); hk = (Q ** 2 / (G * 1.8 ** 2)) ** (1 / 3)
        rr[Q] = dict(h0=h0, hk=hk, day=h0 / 2.0, luu_khong=2.0 - h0)
    # lưu lượng lớn nhất còn thỏa h ≤ 0,75H (mức đầy 75 %) và lưu không ≥ 0,4 m
    def Q_h(h):
        A = 1.8 * h; R = A / (1.8 + 2 * h)
        return A * R ** (2 / 3) * np.sqrt(i) / 0.014
    rr['Qmax_75'] = Q_h(1.50)
    cong[i] = rr
    print('i = %.3f' % i, {k: (round(v['h0'], 3), round(v['hk'], 3)) for k, v in rr.items() if k != 'Qmax_75'},
          'Qmax(h=1,5 m) = %.2f' % rr['Qmax_75'])
KQ['cong_do_doc'] = {str(k): v for k, v in cong.items()}

# ---------------------------------------------------------------- (4) neo ngưỡng tràn
f = 0.60
neo = {}
for ten, Zw, Kcp in (('MNLTK', 77.63, 1.15), ('MNLKT', 78.54, 1.035)):
    hch = 1.0 if ten == 'MNLTK' else 1.2          # như bài tính Bảng 4.4
    r = CV.on_dinh_nguong(Zw, h_chan=hch, c=0.0, ten=ten)
    can = Kcp * r['Px'] - f * r['SV']               # lực chống trượt bổ sung cần thiết (kN/m)
    neo[ten] = dict(SV=r['SV'], SH=r['Px'], K_c0=r['K_truot'], R_can=can)
As = np.pi * 0.025 ** 2 / 4 * 1e6       # mm²
Rs = 350.0                                # MPa – CB400-V
N1 = As * Rs / 1000.0                     # kN – lực kéo tính toán 1 thanh
n_m = 2 / 1.5                             # 2 hàng, a = 1,5 m -> thanh/m dài ngưỡng
R_neo = f * N1 * n_m                      # kN/m – đóng góp chống trượt (ma sát do lực kẹp của neo)
for ten in neo:
    v = neo[ten]
    v['K_neo'] = (f * v['SV'] + R_neo) / v['SH']
    v['luc_keo_can_1_thanh'] = v['R_can'] / f / n_m
L_ngam, D_lo, tau = 2.0, 0.076, 0.40      # m, m, MPa (bám dính vữa – đá phong hóa vừa, lấy thận trọng)
N_bam = np.pi * D_lo * L_ngam * tau * 1000.0    # kN
KQ['neo'] = dict(As_mm2=As, Rs=Rs, N1=N1, n_m=n_m, R_neo=R_neo, N_bam_da=N_bam, chi_tiet=neo)
print('Neo: As=%.0f mm² N1=%.1f kN, %.2f thanh/m -> R_neo=%.1f kN/m; bám dính đá L=2 m D76: %.0f kN' % (
    As, N1, n_m, R_neo, N_bam))
for ten, v in neo.items():
    print(' ', ten, 'ΣV=%.1f ΣH=%.1f K(c=0)=%.2f R cần=%.1f kN/m -> K có neo=%.2f; lực kéo cần 1 thanh %.1f kN' % (
        v['SV'], v['SH'], v['K_c0'], v['R_can'], v['K_neo'], v['luc_keo_can_1_thanh']))


# ---------------------------------------------------------------- (5) khung kín cống hộp (tham khảo)
def khung_kin(L, Hc, qt, qb, e1, e2, EI_t=1.0, EI_b=1.0, EI_w=1.0, n=400):
    """Khung chữ nhật kín đối xứng (tim: nhịp L, chiều cao Hc), tải: nắp qt (kN/m), đáy qb (phản lực nền, kN/m),
    tường: áp lực ngang hình thang e1 (đỉnh) → e2 (đáy). Giải bằng phương pháp lực với ẩn mô men góc (đối xứng):
    M_A (góc trên), M_B (góc dưới) – điều kiện liên tục góc xoay tại hai góc."""
    # góc xoay đầu thanh do tải (dầm đơn giản) và do mô men đầu thanh: dùng tích phân số (phương pháp đơn vị)
    def dam_don(Lb, w, EI):
        x = np.linspace(0, Lb, n + 1); wx = w(x)
        # phản lực, mô men dầm đơn giản
        RA = np.trapezoid(wx * (Lb - x), x) / Lb
        M0 = np.array([RA * xi - np.trapezoid(wx[:k + 1] * (xi - x[:k + 1]), x[:k + 1]) for k, xi in enumerate(x)])
        return x, M0
    # thanh nắp, đáy: tải đều; tường: hình thang theo chiều sâu (x đo từ góc trên)
    xt, M0t = dam_don(L, lambda x: np.full_like(x, qt), EI_t)
    xb, M0b = dam_don(L, lambda x: np.full_like(x, qb), EI_b)
    xw, M0w = dam_don(Hc, lambda x: e1 + (e2 - e1) * x / Hc, EI_w)
    # mô men dương: căng thớ trong. Mô men tại góc: MA (trên), MB (dưới) – âm (căng thớ ngoài).
    # Ma trận mềm cho ẩn X1 = MA, X2 = MB (đặt đối xứng), điều kiện góc xoay liên tục tại góc.
    # Góc xoay đầu dầm đơn giản do tải: θ = ∫ M0 · (1 − x/L) dx / EI (đầu trái).
    th_t = np.trapezoid(M0t * (1 - xt / L), xt) / EI_t
    th_b = np.trapezoid(M0b * (1 - xb / L), xb) / EI_b
    th_w_A = np.trapezoid(M0w * (1 - xw / Hc), xw) / EI_w      # tại đầu trên tường
    th_w_B = np.trapezoid(M0w * (xw / Hc), xw) / EI_w          # tại đầu dưới tường
    # góc xoay do mô men đầu: dầm đơn giản chịu M ở hai đầu bằng nhau (đối xứng) -> θ = M·L/(2EI);
    # tường chịu MA ở đầu trên, MB ở đầu dưới -> θ_A = (MA·Hc/3 + MB·Hc/6)/EI, θ_B = (MA·Hc/6 + MB·Hc/3)/EI
    # liên tục tại góc A: θ_nắp(A) = θ_tường(A); tại B: θ_đáy(B) = θ_tường(B) (cùng quy ước chiều)
    A = np.array([[L / (2 * EI_t) + Hc / (3 * EI_w), Hc / (6 * EI_w)],
                  [Hc / (6 * EI_w), L / (2 * EI_b) + Hc / (3 * EI_w)]])
    bvec = np.array([th_t + th_w_A, th_b + th_w_B])
    MA, MB = np.linalg.solve(A, bvec)
    Mt_giua = M0t[n // 2] - MA
    Mb_giua = M0b[n // 2] - MB
    Mw = M0w - MA * (1 - xw / Hc) - MB * (xw / Hc)
    return dict(MA=MA, MB=MB, Mt_giua=Mt_giua, Mb_giua=Mb_giua, Mw_max=float(Mw.max()), Mw_min=float(Mw.min()))


def cot_thep(M, t, Rs=260.0, Kn=1.15, nc=1.0, a=0.05, b=1.0):
    """As (cm²/m) ≈ Kn·nc·M / (Rs·0,9·h0) – ước tính tham khảo (cấu kiện chịu uốn, bỏ qua lực dọc)."""
    h0 = t - a
    return Kn * nc * abs(M) * 1e3 / (Rs * 1e6 * 0.9 * h0) * 1e4


gam = 20.6
t_n, t_w, t_d = 0.40, 0.40, 0.45                  # giả thiết chiều dày nắp, tường, đáy (m)
Lc, Hc = 1.8 + t_w, 2.0 + 0.5 * (t_n + t_d)
khung = {}
for ten, Hdat in (('truoc_SC', 78.35 - 67.10), ('sau_SC', 78.85 - 67.10)):
    for Kd, K0f in ((1.0, 1.0), (1.0, 0.5), (1.3, 0.5)):
        qv = Kd * gam * Hdat
        qt = qv + 25.0 * t_n
        e1 = K0f * 0.674 * gam * (Hdat + 0.5 * t_n)
        e2 = K0f * 0.674 * gam * (Hdat + t_n + 2.0 + 0.5 * t_d)
        W_tuong = 2 * 25.0 * t_w * Hc
        qb = qt + W_tuong / Lc            # phản lực nền trừ trọng lượng bản đáy (tải uốn bản đáy)
        r = khung_kin(Lc, Hc, qt, qb, e1, e2)
        r.update(qv=qv, e1=e1, e2=e2, qb=qb, As_nap_giua=cot_thep(r['Mt_giua'], t_n), As_goc=cot_thep(r['MA'], t_w),
                 As_day_giua=cot_thep(r['Mb_giua'], t_d))
        khung['%s_Kd%.1f_K0x%.1f' % (ten, Kd, K0f)] = r
        print(ten, 'Kđ=%.1f K0×%.1f' % (Kd, K0f), 'qv=%.0f e=%.0f/%.0f MA=%.1f MB=%.1f M nắp=%.1f M đáy=%.1f kNm/m'
              % (qv, e1, e2, r['MA'], r['MB'], r['Mt_giua'], r['Mb_giua']),
              '-> As nắp %.1f, góc %.1f, đáy %.1f cm²/m' % (r['As_nap_giua'], r['As_goc'], r['As_day_giua']))
KQ['khung_cong'] = dict(t=(t_n, t_w, t_d), Lc=Lc, Hc=Hc, kq=khung)

# ---------------------------------------------------------------- (6) mưa, rãnh dọc
mua = {}
n_mat, C_dc = 0.03, 0.6
beta = np.arctan(1 / 3.0)
for ty in (0.20, 0.30, 0.40):
    i_mm_h = ty * 191.8 / 0.5
    i_ms = i_mm_h / 1000 / 3600
    for L in (4.0, 34.0):
        q = C_dc * i_ms * L
        h = (q * n_mat / np.sqrt(np.sin(beta))) ** 0.6
        tau = 9810 * h * np.sin(beta)
        mua['%.0f%%_L%.0f' % (ty * 100, L)] = dict(i=i_mm_h, tau=tau)
print('Ứng suất tiếp (Pa):', {k: round(v['tau'], 2) for k, v in mua.items()})
# rãnh dọc 40×40 theo mái 1:3 ở lưu lượng thiết kế 0,028 m³/s
Qd, b_r, n_r = 0.028, 0.40, 0.015
S_r = np.sin(beta)
lo, hi = 1e-4, 0.4
for _ in range(100):
    h = 0.5 * (lo + hi); A = b_r * h; R = A / (b_r + 2 * h)
    if A * R ** (2 / 3) * np.sqrt(S_r) / n_r > Qd:
        hi = h
    else:
        lo = h
v_tk = Qd / (b_r * h)
KQ['mua'] = mua; KQ['ranh_doc'] = dict(h=h, v=v_tk)
print('Rãnh dọc ở Q = 0,028 m³/s: h = %.3f m, v = %.2f m/s' % (h, v_tk))

# ---------------------------------------------------------------- (7) mức bảo đảm
for ten, m in (('TH1', 18), ('TH1A', 39), ('TH2', 37), ('TH1_tk', 18), ('TH1A_tk', 37), ('TH2_tk', 35)):
    print(ten, 'm/(n+1) = %.1f%%  m/n = %.1f%%' % (100 * m / 45, 100 * m / 44))

json.dump(KQ, open('../ket_qua/bo_sung_giai_tich.json', 'w'), ensure_ascii=False, indent=1, default=float)
