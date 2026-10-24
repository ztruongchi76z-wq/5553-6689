# -*- coding: utf-8 -*-
"""
(1) Ổn định trượt, lật, ứng suất đáy ngưỡng tràn chính (mục 4.4)
(2) Thủy lực cống lấy nước (mục 6.3): khả năng tháo ở MNC, trạng thái chảy, đường mặt nước, tiêu năng
(3) Lực đóng mở cửa van cống (mục 6.4) – kiểm tra máy VĐ20
"""
import numpy as np
import so_lieu as S

G, GW = S.G, S.GAMMA_W
GBT = 24.0   # kN/m3 bê tông


# ---------------------------------------------------------------------------
# (1) NGƯỠNG TRÀN CHÍNH – mặt cắt hình thang bạt góc [BCC mục 2.4.4]: P1 = 0,7 m; P = 1,4 m; δ = 0,5 m; S = 3,0; S' = 1,75
# ---------------------------------------------------------------------------
def on_dinh_nguong(Zw, h_chan=1.0, f=0.60, c=0.0, t_ban=1.0, ten=''):
    """Khối ngưỡng + bản đáy (dải rộng 1 m). t_ban: chiều dày bản đáy dưới cao trình đáy dốc (giả thiết
    theo cấu tạo thông thường – cần đối chiếu hồ sơ hoàn công). Áp lực đẩy ngược phân bố tuyến tính từ
    γw·(Zw − Zb) ở gót đến γw·(Zday + h_chan − Zb) ở mũi (không kể màng chống thấm, thiên an toàn)."""
    Zn, P1, P2, d, S1, S2 = 74.70, 0.70, 1.40, 0.50, 3.00, 1.75
    Z_tl, Z_day = Zn - P1, Zn - P2            # sân thượng lưu 74,00 ; đáy dốc 73,30
    Zb = Z_day - t_ban                         # đáy khối (mặt tiếp xúc với nền đá lớp II)
    xA, xB, xC, xD = -S1 * P1, 0.0, d, d + S2 * P2
    poly = np.array([[xA, Zb], [xA, Z_tl], [xB, Zn], [xC, Zn], [xD, Z_day], [xD, Zb]])
    x, y = poly[:, 0], poly[:, 1]
    cr = x * np.roll(y, -1) - np.roll(x, -1) * y
    A2 = np.sum(cr)
    cx = np.sum((x + np.roll(x, -1)) * cr) / (3 * A2)
    A = abs(A2) / 2
    W = GBT * A
    B = xD - xA
    H1 = Zw - Zb; H2 = Zw - Z_tl; H3 = Zw - Zn
    Px_dung = 0.5 * GW * (H1 ** 2 - H2 ** 2)
    Px_nghieng = 0.5 * GW * (H2 ** 2 - H3 ** 2)
    def tam(Ht, Hd):
        h = Ht - Hd
        return h * (2 * Hd + Ht) / (3 * (Hd + Ht))
    y_dung = tam(H1, H2)
    y_nghieng = (Z_tl - Zb) + tam(H2, H3)
    Wn_mai = GW * 0.5 * (H2 + H3) * (xB - xA); x_mai = xA + (xB - xA) * (H2 + 2 * H3) / (3 * (H2 + H3))
    Wn_dinh = GW * H3 * d; x_dinh = 0.5 * (xB + xC)
    u1, u2 = GW * H1, GW * (Z_day + h_chan - Zb)
    U = 0.5 * (u1 + u2) * B
    xU = xA + B * (u1 + 2 * u2) / (3 * (u1 + u2))
    SV = W + Wn_mai + Wn_dinh - U
    SH = Px_dung + Px_nghieng
    K_truot = (f * SV + c * B) / SH
    M_on = W * (xD - cx) + Wn_mai * (xD - x_mai) + Wn_dinh * (xD - x_dinh)
    M_lat = Px_dung * y_dung + Px_nghieng * y_nghieng + U * (xD - xU)
    K_lat = M_on / M_lat
    xo = 0.5 * (xA + xD)
    # mô men quanh tâm đáy (dương: gây nén mép hạ lưu)
    M0 = (W * (cx - xo) + Wn_mai * (x_mai - xo) + Wn_dinh * (x_dinh - xo) - U * (xU - xo)
          + Px_dung * y_dung + Px_nghieng * y_nghieng)
    s_tb, ds = SV / B, 6 * M0 / B ** 2
    return dict(ten=ten, Zw=Zw, A=A, W=W, B=B, Px=SH, Wn=Wn_mai + Wn_dinh, U=U, SV=SV, K_truot=K_truot,
                K_lat=K_lat, s_hl=s_tb + ds, s_tl=s_tb - ds, f=f, c=c, t_ban=t_ban, h_chan=h_chan)


# ---------------------------------------------------------------------------
# (2) CỐNG LẤY NƯỚC: hộp 1,8 x 2,0 m, L = 67 m, ngưỡng +65,60, i theo hồ sơ = 3 %
# ---------------------------------------------------------------------------
def kha_nang_qua_cong(Z_ho, b=1.8, Zn=65.60, m=0.35, phi_n=1.0):
    """Cửa vào cống không áp làm việc như đập tràn đỉnh rộng chảy tự do: Q = m·b·√(2g)·H0^1,5."""
    H = Z_ho - Zn
    return m * b * np.sqrt(2 * G) * max(H, 0) ** 1.5


def h_chuan(Q, b, i, n=0.014):
    lo, hi = 1e-3, 5.0
    for _ in range(100):
        h = 0.5 * (lo + hi)
        A = b * h; R = A / (b + 2 * h)
        Qc = A * R ** (2 / 3) * np.sqrt(i) / n
        if Qc > Q:
            hi = h
        else:
            lo = h
    return 0.5 * (lo + hi)


def duong_mat_nuoc_cong(Q, b, i, L, h_dau, n=0.014, N=67):
    """Đường mặt nước trong cống từ cửa vào (dòng xiết khi i > ik) – cộng trực tiếp."""
    xs = np.linspace(0, L, N + 1); hs = [h_dau]
    hkk = (Q ** 2 / (G * b ** 2)) ** (1 / 3)
    for k in range(N):
        dx = xs[k + 1] - xs[k]; h1 = hs[-1]
        def E(h): return h + (Q / (b * h)) ** 2 / (2 * G)
        def J(h):
            A = b * h; R = A / (b + 2 * h); v = Q / A
            return v ** 2 * n ** 2 / R ** (4 / 3)
        lo, hi = 0.05, (0.999 * hkk if h1 < hkk else 5.0)
        for _ in range(80):
            h2 = 0.5 * (lo + hi)
            f = E(h2) + 0.5 * (J(h1) + J(h2)) * dx - E(h1) - i * dx
            if h1 < hkk:
                lo, hi = (h2, hi) if f > 0 else (lo, h2)
            else:
                lo, hi = (lo, h2) if f > 0 else (h2, hi)
        hs.append(0.5 * (lo + hi))
    hs = np.array(hs)
    v = Q / (b * hs)
    return xs, hs, v


def be_tieu_nang(Q, b, h1, Z_h, h_h):
    """Kiểm tra nước nhảy sau cống: độ sâu liên hiệp h'' = h1/2·(√(1+8Fr²) − 1), chiều dài nước nhảy L = 4,5·h''."""
    v1 = Q / (b * h1); Fr = v1 / np.sqrt(G * h1)
    h2 = 0.5 * h1 * (np.sqrt(1 + 8 * Fr ** 2) - 1)
    Ln = 4.5 * h2
    return dict(v1=v1, Fr1=Fr, h2=h2, L_nhay=Ln, sigma=h_h / h2)


# ---------------------------------------------------------------------------
# (3) LỰC ĐÓNG MỞ CỬA VAN PHẲNG CỐNG LẤY NƯỚC (cửa van trượt, gioăng cao su)
# ---------------------------------------------------------------------------
def luc_dong_mo(Z_ho=74.70, b=1.80, h=2.00, Z_day=65.60, G_van=20.0, f_truot=0.40, f_gioang=0.70,
                b_gioang=0.04, n_G=1.1, n_ms=1.2, W_hut=5.0, ten=''):
    """Lực nâng cửa: P = nG·G + n_ms·(T1 + T2) + Wh ; T1 = f_trượt·W (ma sát gối trượt do áp lực nước),
    T2 = ma sát gioăng cạnh bên; kN."""
    hc = Z_ho - (Z_day + h / 2)
    W = GW * hc * b * h                           # tổng áp lực nước lên cửa
    T1 = f_truot * W
    # gioăng hai cạnh bên: áp lực trung bình trên dải gioăng rộng b_gioang dọc chiều cao h
    T2 = f_gioang * GW * hc * b_gioang * h * 2
    P_mo = n_G * G_van + n_ms * (T1 + T2) + W_hut
    P_dong = n_ms * (T1 + T2) - 0.9 * G_van        # lực đóng cần thêm (âm: tự đóng được bằng trọng lượng)
    return dict(ten=ten, hc=hc, W=W, T1=T1, T2=T2, P_mo=P_mo, P_mo_T=P_mo / 9.81, P_dong=P_dong,
                P_dong_T=P_dong / 9.81)



def on_dinh_truc_vit(P_nen, d_chan=0.072, L=6.0, mu=0.7, E=2.1e8):
    """Ổn định (Euler) trục vít chịu nén khi đóng cưỡng bức: Pth = π²·E·J/(μL)², J = π·d⁴/64 (d: đường kính chân ren)."""
    J = np.pi * d_chan ** 4 / 64
    Pth = np.pi ** 2 * E * J / (mu * L) ** 2
    return dict(J=J, Pth=Pth, K=Pth / P_nen if P_nen > 0 else np.inf, mu=mu, L=L, d=d_chan)


if __name__ == '__main__':
    for Zw, lab in [(74.70, 'MNDBT'), (77.63, 'MNLTK 77,63'), (78.54, 'MNLKT 78,54')]:
        for c in [0.0, 50.0]:
            r = on_dinh_nguong(Zw, c=c, ten=lab)
            print(lab, 'c=%.0f' % c, 'A=%.2f W=%.1f Px=%.1f Wn=%.1f U=%.1f SV=%.1f Kt=%.2f Kl=%.2f s_hl=%.1f s_tl=%.1f' %
                  (r['A'], r['W'], r['Px'], r['Wn'], r['U'], r['SV'], r['K_truot'], r['K_lat'], r['s_hl'], r['s_tl']))
    print('Q qua cống ở MNC', round(kha_nang_qua_cong(S.MNC), 2), 'ở MNDBT', round(kha_nang_qua_cong(S.MNDBT), 2))
    for i in [0.03, 0.003]:
        for Q in [6.54, 10.0]:
            h0 = h_chuan(Q, 1.8, i); hkk = (Q ** 2 / (G * 1.8 ** 2)) ** (1 / 3)
            print('i=%.3f Q=%.2f h0=%.3f hk=%.3f' % (i, Q, h0, hkk), 'xiết' if h0 < hkk else 'êm')
    xs, hs, v = duong_mat_nuoc_cong(6.54, 1.8, 0.03, 67.0, 0.99 * (6.54 ** 2 / (G * 1.8 ** 2)) ** (1 / 3))
    print('cuối cống h=%.3f v=%.2f' % (hs[-1], v[-1]), 'khoảng lưu không min %.2f m' % (2.0 - hs.max()))
    print(be_tieu_nang(6.54, 1.8, hs[-1], 0, 1.2))
    for f in [0.30, 0.40, 0.50]:
        r = luc_dong_mo(f_truot=f)
        print('f=%.2f' % f, 'W=%.1f kN P_mo=%.1f kN = %.2f T' % (r['W'], r['P_mo'], r['P_mo_T']))
