# -*- coding: utf-8 -*-
"""
TÍNH TOÁN ĐIỀU TIẾT LŨ HỒ CÀ GIÂY (mục 4.1 của đồ án)
Phương pháp: giải hệ phương trình cân bằng nước + phương trình xả qua tràn theo từng thời đoạn
            (Q1+Q2)/2·Δt − (q1+q2)/2·Δt = V2 − V1 ,  q = f(Z)
Lưu lượng qua tràn (đập tràn thực dụng chảy tự do, TCVN 9147:2012):
            Q = σn·ε·m·Σb·√(2g)·H0^1,5
            ε = 1 − 0,2·[ξmb + (n−1)·ξmt]·H0/(n·b)
"""
import numpy as np
import so_lieu as S

SQ2G = np.sqrt(2 * S.G)


# ---------------------------------------------------------------------------
# Quan hệ lưu lượng – mực nước của hai tràn
# ---------------------------------------------------------------------------
class Tran:
    def __init__(self, ten, Z_nguong, n, b, m, xi_mb, xi_mt, em_const=None):
        self.ten, self.Zn, self.n, self.b, self.m = ten, Z_nguong, n, b, m
        self.xi_mb, self.xi_mt, self.em_const = xi_mb, xi_mt, em_const

    def eps(self, H):
        if self.em_const is not None:
            return 1.0
        H = np.maximum(H, 0.0)
        return 1.0 - 0.2 * (self.xi_mb + (self.n - 1) * self.xi_mt) * H / (self.n * self.b)

    def Q(self, Z):
        H = np.maximum(np.asarray(Z, float) - self.Zn, 0.0)
        em = self.em_const if self.em_const is not None else self.eps(H) * self.m
        return em * self.n * self.b * SQ2G * H ** 1.5


def tran_hien_trang():
    """Phương án tính chính của đồ án (hiện trạng 2026): tràn chính 9 khoang x 2,4 m, tràn sự cố 1 khoang 5 m.
    m = 0,45 (đập tràn thực dụng thấp, P1 = 0,7 m < 0,5H – lấy thiên an toàn so với m = 0,467 hồ sơ BCNCKT tra);
    ξmb = 0,70 (mố bên lượn tròn), ξmt = 0,45 (mố trụ đầu tròn)."""
    tc = Tran('Tràn chính', 74.70, 9, 2.40, 0.45, 0.70, 0.45)
    sc = Tran('Tràn sự cố', 71.00, 1, 5.00, 0.45, 0.70, 0.45)
    return tc, sc


def tran_bcnckt():
    """Thông số tư vấn dùng trong BCNCKT 2022 (PL1): B = 25 m, εm = 0,42 cho cả hai tràn."""
    return (Tran('Tràn chính (BCNCKT)', 74.70, 1, 25.0, None, 0, 0, em_const=0.42),
            Tran('Tràn sự cố (BCNCKT)', 71.00, 1, 5.0, None, 0, 0, em_const=0.42))


def tran_qtvh():
    """Thông số QTVH 2026: B = 21,6 m với εm ≈ 0,40; tràn sự cố εm ≈ 0,37."""
    return (Tran('Tràn chính (QTVH)', 74.70, 9, 2.40, None, 0, 0, em_const=0.40),
            Tran('Tràn sự cố (QTVH)', 71.00, 1, 5.0, None, 0, 0, em_const=0.37))


# ---------------------------------------------------------------------------
# Thuật toán điều tiết
# ---------------------------------------------------------------------------
def dieu_tiet(Qden, tc, sc, Z0=S.MNDBT, dt_h=1.0, sub=10, ket_van=False, giu_mndbt=True,
              V_of_Z=S.V_of_Z, t_them=48):
    """Qden: list lưu lượng đến tại t = 1, 2, ... giờ (bảng hồ sơ); Q(0) lấy = Q(1).
    Chia mỗi giờ thành 'sub' bước con (nội suy tuyến tính Q đến) để giảm sai số.
    Quy tắc vận hành (theo QTVH): đầu trận lũ hồ ở MNDBT; khi lũ về, mở dần cửa van tràn sự cố để
    xả bằng lưu lượng đến (giữ MNDBT) cho tới khi cửa mở hết; sau đó hai tràn cùng chảy tự do.
    ket_van=True: cửa van tràn sự cố bị kẹt đóng hoàn toàn (Qsc = 0)."""
    Qh = np.array([Qden[0]] + list(Qden) + [Qden[-1]] * 0, float)
    # kéo dài nhánh xuống bằng 0 để tính hết quá trình xả (không có thêm nước đến)
    Qh = np.concatenate([Qh, np.zeros(t_them)])
    T = np.arange(len(Qh)) * dt_h
    dts = dt_h * 3600.0 / sub
    tt = [0.0]; ZZ = [Z0]; q1s = [0.0]; q2s = [0.0]; Qin = [Qh[0]]
    Z = Z0
    V = V_of_Z(Z)

    def qtot(z):
        a = tc.Q(z)
        b = 0.0 if ket_van else sc.Q(z)
        return a, b

    a0, b0 = qtot(Z)
    # trạng thái ban đầu: hồ ở MNDBT, tràn chính chưa chảy (H=0), cửa tràn sự cố điều tiết để q = Q
    q_main, q_sc = a0, (0.0 if ket_van else min(Qh[0], b0))
    q1s[0], q2s[0] = q_main, q_sc
    # các giai đoạn vận hành: 'giu' (giữ MNDBT bằng van sự cố), 'tu_do' (hai tràn chảy tự do),
    # 'dong' (cuối lũ: đóng dần van sự cố để giữ mực nước, q = Q đến)
    mode = 'giu' if (giu_mndbt and not ket_van) else 'tu_do'
    for k in range(len(Qh) - 1):
        for j in range(sub):
            t0 = T[k] + j / sub * dt_h
            t1 = T[k] + (j + 1) / sub * dt_h
            Qa = np.interp(t0, T, Qh); Qb = np.interp(t1, T, Qh)
            q_old = q_main + q_sc
            if mode in ('giu', 'dong'):
                cap = sc.Q(Z) + tc.Q(Z)
                if Qb <= cap:
                    q_main = float(tc.Q(Z)); q_sc = Qb - q_main
                    tt.append(t1); ZZ.append(Z); q1s.append(q_main); q2s.append(q_sc); Qin.append(Qb)
                    continue
                mode = 'tu_do'
            # giải V2 + q(Z2)·dt/2 = V1 − q1·dt/2 + (Qa+Qb)/2·dt  (đơn vị 10^6 m3)
            rhs = V + ((Qa + Qb) / 2.0 - q_old / 2.0) * dts / 1e6
            lo, hi = 59.0, 85.0
            for _ in range(60):
                mid = 0.5 * (lo + hi)
                a, b = qtot(mid)
                f = V_of_Z(mid) + (a + b) / 2.0 * dts / 1e6 - rhs
                if f > 0:
                    hi = mid
                else:
                    lo = mid
            Z = 0.5 * (lo + hi)
            V = V_of_Z(Z)
            q_main, q_sc = qtot(Z)
            # cuối lũ: mực nước rút xuống MNDBT -> chuyển sang đóng dần van giữ mực nước
            if giu_mndbt and (not ket_van) and Z <= Z0 and Qb < q_old:
                mode = 'dong'
            tt.append(t1); ZZ.append(Z); q1s.append(q_main); q2s.append(q_sc); Qin.append(Qb)
    r = dict(t=np.array(tt), Z=np.array(ZZ), q_tc=np.array(q1s), q_sc=np.array(q2s), Q=np.array(Qin))
    r['q'] = r['q_tc'] + r['q_sc']
    i = int(np.argmax(r['Z']))
    r['Zmax'] = r['Z'][i]; r['t_Zmax'] = r['t'][i]
    r['qmax'] = r['q'].max(); r['q_tc_max'] = r['q_tc'][i]; r['q_sc_max'] = r['q_sc'][i]
    r['Vpl'] = float(V_of_Z(r['Zmax']) - V_of_Z(Z0))
    r['Htc'] = r['Zmax'] - tc.Zn; r['Hsc'] = r['Zmax'] - sc.Zn
    r['Qmax_den'] = max(Qden); r['t_Qmax'] = float(np.argmax(Qden) + 1)
    # kiểm tra cân bằng nước toàn trận
    dt = np.diff(r['t']) * 3600
    W_den = np.sum((r['Q'][1:] + r['Q'][:-1]) / 2 * dt) / 1e6
    W_xa = np.sum((r['q'][1:] + r['q'][:-1]) / 2 * dt) / 1e6
    r['W_den'], r['W_xa'] = W_den, W_xa
    r['dV'] = float(V_of_Z(r['Z'][-1]) - V_of_Z(Z0))
    r['sai_so_CBN'] = W_den - W_xa - r['dV']
    return r


def bang_gio(r, Qden):
    """Trích kết quả tại các mốc giờ tròn để đưa vào phụ lục."""
    out = []
    for h in range(0, len(Qden) + 30):
        i = int(np.argmin(np.abs(r['t'] - h)))
        out.append((h, r['Q'][i], r['Z'][i], r['q_tc'][i], r['q_sc'][i], r['q'][i],
                    float(S.V_of_Z(r['Z'][i]))))
    return out


if __name__ == '__main__':
    tc, sc = tran_hien_trang()
    for H in [1, 2, 2.73, 3, 3.61, 4]:
        print('H', H, 'eps tc', round(tc.eps(H), 3), 'em', round(tc.eps(H) * tc.m, 3),
              'Qtc', round(float(tc.Q(74.7 + H)), 1))
    for H in [3.7, 6.43, 7.31]:
        print('H', H, 'eps sc', round(sc.eps(H), 3), 'em', round(sc.eps(H) * sc.m, 3),
              'Qsc', round(float(sc.Q(71 + H)), 1))
    for nguon, Qd in [('KTTV', S.Q_KTTV), ('QTVH', S.Q_QTVH)]:
        for p in ['1%', '0.2%', '0.1%']:
            if p not in Qd:
                continue
            for kv in [False, True]:
                r = dieu_tiet(Qd[p], tc, sc, ket_van=kv)
                print(nguon, p, 'kẹt' if kv else 'BT', 'Zmax %.2f t=%.1f qmax %.1f (tc %.1f sc %.1f) Vpl %.2f'
                      % (r['Zmax'], r['t_Zmax'], r['qmax'], r['q_tc_max'], r['q_sc_max'], r['Vpl']),
                      'CBN err %.4f' % r['sai_so_CBN'])
    # đối chiếu với hồ sơ: dùng thông số của tư vấn
    for lab, fn in [('BCNCKT-param', tran_bcnckt), ('QTVH-param', tran_qtvh)]:
        a, b = fn()
        for nguon, Qd in [('KTTV', S.Q_KTTV), ('QTVH', S.Q_QTVH)]:
            for p in ['1%', '0.2%']:
                r = dieu_tiet(Qd[p], a, b)
                print(lab, nguon, p, 'Zmax %.2f qmax %.1f' % (r['Zmax'], r['qmax']))
