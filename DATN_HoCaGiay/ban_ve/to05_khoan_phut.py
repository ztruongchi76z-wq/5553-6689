# -*- coding: utf-8 -*-
"""BV-05: MÀNG KHOAN PHỤT CHỐNG THẤM MANG TRÁI TRÀN SỰ CỐ – CHUYÊN ĐỀ (Chương 8).

Gồm: mặt bằng bố trí hố khoan theo phương án hợp lý hóa (TL 1/200), cắt dọc tuyến màng (ngang 1/200, đứng 1/100),
mặt cắt ngang đại diện với kết quả tính thấm (TL 1/250), sơ đồ cấu tạo hố khoan và phân đoạn phụt, các bảng kết quả,
độ nhạy, so sánh phương án và quan trắc.
Số liệu: tinh_toan/chuyen_de_khoan_phut.py, ket_qua/khoan_phut.json; địa tầng dọc tuyến số hóa từ bản vẽ TKCS
"Cắt dọc bố trí khoan phụt" (No-CG-22TK-TKKP-01).
"""
import math

import du_lieu_ban_ve as D
from bv_lib import To, KhungNhin, so

S_AC = [3.0 * i for i in range(17)]                 # lý trình hố hàng A, C: A01..A17, C01..C17 (s = 0, 3, …, 48)
S_B_TK = [1.5 + 3.0 * i for i in range(17)]         # hàng B: B01..B17 (s = 1,5; 4,5; …; 49,5) – so le với A, C
L3 = 15.0                                           # đoạn 3 hàng sát tường bên tràn sự cố
S_C = [s for s in S_AC if s <= L3]                  # hàng C giữ lại (đoạn 3 hàng): C01..C06
# Phương án hợp lý hóa: hai hàng KỀ NHAU A, B (cách 1,5 m, so le) suốt 50 m + hàng C trong đoạn 15 m sát tràn.
# (Không dùng A, C: hai hàng cách 3,0 m để hở dải giữa không được phụt; T = Bb + 0,6A chỉ đúng với hai hàng kề nhau.)
L_MAN = 50.0
VUNG_TN = (30.6, 41.5)                              # vùng phụt thí nghiệm theo TKCS
KT = [('KT1', 36.0, D.KP_HANG['B']), ('KT2', 8.25, 0.5 * (D.KP_HANG['A'] + D.KP_HANG['B'])),
      ('KT3', 24.0, D.KP_HANG['B'])]                # KT1, KT3: giữa hai hố hàng B, cách hố hàng A 1,5 m (xa hố phụt nhất)
# giếng đo áp: 5 m về hạ lưu hàng khoan hạ lưu nhất của đoạn (P1: sau hàng C – đoạn 3 hàng; P2: sau hàng B – đoạn 2 hàng)
GIENG = [('P1', 9.0, D.KP_HANG['C'] + 5.0), ('P2', 30.0, D.KP_HANG['B'] + 5.0)]


def noi_suy(pts, s):
    """Nội suy tuyến tính trên đường gấp khúc [(s, z)]; ngoại suy theo đoạn đầu/cuối."""
    if s <= pts[0][0]:
        (a, za), (b, zb) = pts[0], pts[1]
    elif s >= pts[-1][0]:
        (a, za), (b, zb) = pts[-2], pts[-1]
    else:
        for (a, za), (b, zb) in zip(pts[:-1], pts[1:]):
            if a <= s <= b:
                break
    return za + (zb - za) * (s - a) / (b - a)


def z_mat(s):
    return noi_suy(D.KPD_MAT, s)


def z_II(s):
    return noi_suy(D.KPD_I_II, s)


def day_tk(s):
    return noi_suy(D.KPD_DAY_TK, s)


def day_dx(s):
    """Đáy màng đề xuất theo TCVN 8216:2018 điều 9.2.2.9 (đập cấp II: màng không sâu quá cao trình nền có lượng mất
    nước đơn vị (5 ÷ 7) Lu cộng thêm 3 m): thấp hơn mặt lớp II 3,0 m khi ép nước lớp II cho q ≤ 5 Lu; không sâu hơn
    đáy TKCS. Lớp II nứt nẻ (q > 5 Lu): đáy = cao trình đạt q ≤ 5 Lu trừ 3 m (theo kết quả ép nước)."""
    return max(z_II(s) - 3.0, day_tk(s))


def doan(pts, s0, s1, buoc=0.5):
    """Cắt đường gấp khúc trong khoảng [s0, s1]."""
    out = [(s0, noi_suy(pts, s0))]
    out += [p for p in pts if s0 < p[0] < s1]
    out.append((s1, noi_suy(pts, s1)))
    return out


def khoi_luong():
    """So sánh khối lượng khoan, phụt theo hình học bản vẽ (cùng một cách tính cho các phương án):
    TKCS – 3 hàng A, B, C dài 50 m, đáy TKCS: 51 hố, khoan 506,0 m, phụt 430,9 m (trong lớp II 176,2 m);
    hợp lý hóa – A, B dài 50 m + C đoạn 15 m, đáy TKCS: 40 hố, 391,4 m, 332,3 m (142,2 m);
    hợp lý hóa + đáy theo điều 9.2.2.9 (mặt lớp II − 3 m): 40 hố, 365,1 m, 305,9 m (115,8 m)."""
    def tinh(ds, day):
        kh = ph = ph2 = 0.0
        for s in ds:
            zb = day(s)
            kh += z_mat(s) - zb
            ph += D.KPD_DINH_MAN - zb
            ph2 += max(0.0, min(z_II(s), D.KPD_DINH_MAN) - zb)
        return len(ds), kh, ph, ph2
    tk = tinh(S_AC * 2 + S_B_TK, day_tk)
    a = tinh(S_AC + S_B_TK + S_C, day_tk)
    b = tinh(S_AC + S_B_TK + S_C, day_dx)
    return tk, a, b


# ====================================================================== ký hiệu
def ho(to, X, Y, thu_tu, r=0.9, lop='NET_CHINH'):
    """Hố khoan phụt: thứ tự I – tròn đặc; thứ tự II – tròn có chữ thập."""
    to.circle((X, Y), r, lop)
    if thu_tu == 1:
        to.hatch([(X + r * math.cos(a / 12 * math.pi), Y + r * math.sin(a / 12 * math.pi)) for a in range(24)],
                 'SOLID', lop='NET_CHINH', mau_to=7)
    else:
        to.line((X - r * 1.5, Y), (X + r * 1.5, Y), lop)
        to.line((X, Y - r * 1.5), (X, Y + r * 1.5), lop)


def ho_kt(to, X, Y, r=1.3):
    p = [(X, Y + r), (X - r * 0.87, Y - r * 0.5), (X + r * 0.87, Y - r * 0.5)]
    to.pline(p, 'SUA_CHUA', dong=True)
    to.hatch(p, 'SOLID', lop='SUA_CHUA', mau_to=1)


def gieng(to, X, Y, r=1.3):
    to.circle((X, Y), r, 'NUOC')
    to.circle((X, Y), r * 0.45, 'NUOC')
    to.hatch([(X + r * 0.45 * math.cos(a / 8 * math.pi), Y + r * 0.45 * math.sin(a / 8 * math.pi)) for a in range(16)],
             'SOLID', lop='NUOC', mau_to=5)


def ho_ks(to, X, Y, r=1.5):
    """Hố khoan khảo sát: tròn chia 4, tô hai góc phần tư."""
    to.circle((X, Y), r, 'NET_CHINH')
    to.line((X - r, Y), (X + r, Y), 'NET_CHINH')
    to.line((X, Y - r), (X, Y + r), 'NET_CHINH')
    for a0 in (0, 180):
        p = [(X, Y)] + [(X + r * math.cos(math.radians(a0 + a)), Y + r * math.sin(math.radians(a0 + a)))
                        for a in range(0, 91, 15)]
        to.hatch(p, 'SOLID', lop='NET_CHINH', mau_to=7)


def mui_ten(to, p0, p1, lop='NET_CHINH', dai=2.5):
    to.line(p0, p1, lop)
    a = math.atan2(p1[1] - p0[1], p1[0] - p0[0])
    for da in (2.7, -2.7):
        to.line(p1, (p1[0] + dai * math.cos(a + da), p1[1] + dai * math.sin(a + da)), lop)


def ky_hieu_mai(to, pts_mep, huong, dai=6.0, buoc=6.0):
    """Ký hiệu mái dốc: các vạch dài/ngắn vuông góc với mép, hướng 'huong' (vector đơn vị)."""
    (x0, y0), (x1, y1) = pts_mep
    L = math.hypot(x1 - x0, y1 - y0)
    n = int(L / buoc)
    for k in range(n + 1):
        t = k * buoc / L
        x, y = x0 + (x1 - x0) * t, y0 + (y1 - y0) * t
        d = dai if k % 2 == 0 else dai * 0.5
        to.line((x, y), (x + huong[0] * d, y + huong[1] * d), 'NET_MANH', color=7)


# ====================================================================== mặt bằng
PX = lambda s: 40.0 + (s + 10.0) * 5.0              # 1/200, s = lý trình dọc màng
PY = lambda x: 455.0 + (29.0 - x) * 5.0              # x = tọa độ ngang đập (như mặt cắt đại diện), thượng lưu ở trên


def mat_bang(to):
    to.ten_hinh('MẶT BẰNG BỐ TRÍ HỐ KHOAN PHỤT – PHƯƠNG ÁN HỢP LÝ HÓA', 40, 576, '1/200', h=4.5, can='LEFT')
    sA, sB = -10.0, 52.0
    xt0, xt1 = D.KP_X_TUONG, D.KP_X_TUONG + 0.5
    tsc0, tsc1 = D.KPD_TSC['ngoai'][0], D.KPD_TSC['ngoai'][1]
    for a, b, lop in ((sA, tsc0, 'NET_MANH'), (tsc1, sB, 'NET_CHINH')):
        # tường chắn sóng (2 nét + gạch), mép hạ lưu mặt đập
        to.line((PX(a), PY(xt0)), (PX(b), PY(xt0)), lop)
        to.line((PX(a), PY(xt1)), (PX(b), PY(xt1)), lop)
        to.line((PX(a), PY(D.KP_X_MEP_HL)), (PX(b), PY(D.KP_X_MEP_HL)), lop)
    to.hatch([(PX(tsc1), PY(xt0)), (PX(sB), PY(xt0)), (PX(sB), PY(xt1)), (PX(tsc1), PY(xt1))], 'ANSI31', 0.35)
    ky_hieu_mai(to, ((PX(tsc1 + 1.0), PY(xt0)), (PX(sB), PY(xt0))), (0, 1), dai=7.0)
    ky_hieu_mai(to, ((PX(tsc1 + 1.0), PY(D.KP_X_MEP_HL)), (PX(sB), PY(D.KP_X_MEP_HL))), (0, -1), dai=7.0)
    # đường cắt đứt hai đầu
    for s in (sA, sB):
        to.line((PX(s), PY(7.2)), (PX(s), PY(28.8)), 'NET_MANH')
    # tràn sự cố: hai tường bên, kênh 5 m
    for a, b in ((tsc0, D.KPD_TSC['trong'][0]), (D.KPD_TSC['trong'][1], tsc1)):
        p = [(PX(a), PY(7.2)), (PX(b), PY(7.2)), (PX(b), PY(28.8)), (PX(a), PY(28.8))]
        to.hatch(p, 'AR-CONC', 0.03)
        to.pline(p, 'NET_CHINH', dong=True)
    sm = 0.5 * (D.KPD_TSC['trong'][0] + D.KPD_TSC['trong'][1])
    to.text('TRÀN SỰ CỐ (b = 5 m)', (PX(sm), PY(20.5)), 2.2, 'MC', rot=90)
    mui_ten(to, (PX(sm), PY(8.0)), (PX(sm), PY(13.5)), 'NUOC')
    # tim các hàng khoan
    yA, yB, yC = PY(D.KP_HANG['A']), PY(D.KP_HANG['B']), PY(D.KP_HANG['C'])
    to.line((PX(-0.3), yA), (PX(L_MAN + 0.6), yA), 'TRUC')
    to.line((PX(-0.3), yB), (PX(L_MAN + 0.6), yB), 'TRUC')
    to.line((PX(-0.3), yC), (PX(L3 + 0.4), yC), 'TRUC')
    to.line((PX(L3 + 0.4), yC), (PX(L_MAN + 0.6), yC), 'NET_MANH', linetype='DASHED', ltscale=0.3)
    # hố khoan: hàng A, B đủ 50 m (so le), hàng C chỉ trong đoạn 3 hàng (C07 ÷ C17 của TKCS bỏ)
    for i, s in enumerate(S_AC):
        ho(to, PX(s), yA, 1 if i % 2 == 0 else 2)
        to.text('A%02d' % (i + 1), (PX(s) - 0.8, yA + 2.0), 1.7, 'LEFT')
        if s <= L3:
            ho(to, PX(s), yC, 1 if i % 2 == 0 else 2)
            to.text('C%02d' % (i + 1), (PX(s) - 0.8, yC - 3.9), 1.7, 'LEFT')
        else:
            to.circle((PX(s), yC), 0.7, 'NET_MANH')
    for i, s in enumerate(S_B_TK):
        ho(to, PX(s), yB, 1 if i % 2 == 0 else 2)
        to.text('B%02d' % (i + 1), (PX(s) + 1.3, yB - 2.9), 1.7)
    # ranh giới đoạn 3 hàng / 2 hàng
    to.line((PX(L3), PY(xt1)), (PX(L3), PY(D.KP_X_MEP_HL)), 'SUA_CHUA', linetype='DASHED', ltscale=0.4)
    yk = PY(22.4)
    to.dim((PX(0), yk), (PX(L3), yk), 0, 200)
    to.dim((PX(L3), yk), (PX(L_MAN), yk), 0, 200)
    to.text('ĐOẠN 3 HÀNG A, B, C (T = 4,8 m)', (PX(0.5 * L3), yk - 4.2), 2.0, 'CENTER', 'SUA_CHUA')
    to.text('ĐOẠN 2 HÀNG A, B (T = 3,3 m) – áp dụng khi phụt thí nghiệm đạt', (PX(0.5 * (L3 + L_MAN)), yk - 4.2),
            2.0, 'CENTER', 'SUA_CHUA')
    to.dim((PX(0), PY(27.6)), (PX(L_MAN), PY(27.6)), 0, 200, text='L = 5000')
    # khoảng cách hố, hàng
    to.dim((PX(S_AC[1]), yA + 6.0), (PX(S_AC[2]), yA + 6.0), 0, 200)
    to.dim((PX(L_MAN + 1.6), yA), (PX(L_MAN + 1.6), yB), 0, 200, goc=90)
    to.dim((PX(L_MAN + 1.6), yB), (PX(L_MAN + 1.6), yC), 0, 200, goc=90)
    for ten, y in (('A', yA), ('B', yB), ('C', yC)):
        to.text('HÀNG ' + ten + (' (đoạn 15 m)' if ten == 'C' else ''), (PX(L_MAN + 3.2), y), 2.0, 'ML')
    # vùng phụt thí nghiệm
    a, b = VUNG_TN
    to.pline([(PX(a), PY(15.2)), (PX(b), PY(15.2)), (PX(b), PY(20.4)), (PX(a), PY(20.4))], 'SUA_CHUA', dong=True,
             linetype='DASHED', ltscale=0.35)
    to.text('VÙNG PHỤT THÍ NGHIỆM (TKCS)', (PX(0.5 * (a + b)), PY(15.2) + 1.0), 1.9, 'CENTER', 'SUA_CHUA')
    # hố kiểm tra, giếng đo áp, hố khảo sát CG7
    for ten, s, x in KT:
        ho_kt(to, PX(s), PY(x))
        to.text(ten, (PX(s) + 1.8, PY(x) - 0.9), 1.8, 'LEFT', 'SUA_CHUA')
    for ten, s, xg in GIENG:
        gieng(to, PX(s), PY(xg))
        to.text(ten, (PX(s) + 2.0, PY(xg) - 0.9), 1.9, 'LEFT', 'NUOC')
    ho_ks(to, PX(D.KPD_CG7['s']), PY(18.6))
    to.text('CG7', (PX(D.KPD_CG7['s']) + 2.0, PY(18.6) - 2.6), 1.8)
    # phụt tiếp xúc dọc lưng tường bên
    to.chu_thich_mui_ten((PX(0.15), PY(12.4)), (PX(4.0), PY(7.9)),
                         'Phụt tiếp xúc dọc lưng tường bên tràn bằng hố khoan nghiêng (mục 8.6)', h=1.9, rong=72)
    # cọc, hướng dòng chảy, ghi chú
    for ten, s in (('C23', 0.0), ('C22', D.KPD_C22), ('C22+13,97', L_MAN)):
        to.line((PX(s), PY(7.0)), (PX(s), PY(7.0) + 3.0), 'NET_CHINH')
        to.text(ten, (PX(s), PY(7.0) + 3.6), 2.0, 'CENTER')
    to.text('MÁI THƯỢNG LƯU', (PX(30.0), PY(8.4)), 2.0, 'MC')
    to.text('MÁI HẠ LƯU', (PX(30.0), PY(27.9)), 2.0, 'MC')
    to.text('TƯỜNG CHẮN SÓNG', (PX(43.0), PY(xt1) - 3.2), 1.9, 'MC')
    mui_ten(to, (PX(-8.4), PY(12.0)), (PX(-8.4), PY(24.0)), 'NUOC')
    to.text('Hướng dòng chảy', (PX(-8.4) + 1.5, PY(18.0)), 1.8, 'MC', rot=90)


# ====================================================================== cắt dọc
QY = lambda z: 295.0 + (z - 65.5) * 10.0             # đứng 1/100


def cat_doc(to):
    to.ten_hinh('CẮT DỌC TUYẾN MÀNG KHOAN PHỤT (hố các hàng A, B, C chiếu lên tim màng)', 40, 446, None, h=4.5,
                can='LEFT')
    to.text('TỶ LỆ: NGANG 1/200, ĐỨNG 1/100', (40, 439.5), 3.0)
    s0, s1 = -10.0, 52.0
    # thước cao độ
    xr = 33.0
    to.line((xr, QY(65.5)), (xr, QY(79.5)), 'NET_CHINH')
    for z in range(66, 80):
        to.line((xr, QY(z)), (xr + 1.5, QY(z)), 'NET_CHINH')
        to.text('%d,0' % z, (xr - 0.8, QY(z)), 1.7, 'MR')
    to.text('Cao độ (m)', (xr - 2.0, QY(79.5) + 2.5), 1.8, 'CENTER')
    # mặt đất / mặt đập hiện trạng (bỏ đoạn tràn sự cố)
    tsc0, tsc1, zt0, zt1 = D.KPD_TSC['ngoai']
    to.pline([(PX(s), QY(z)) for s, z in doan(D.KPD_MAT, s0, tsc0)], 'NET_CHINH')
    to.pline([(PX(s), QY(z)) for s, z in doan(D.KPD_MAT, tsc1, s1)], 'NET_CHINH')
    to.pline([(PX(s), QY(z)) for s, z in doan(D.KPD_MAT_DUOI_MD, tsc1, s1)], 'NET_MANH')
    # ranh giới lớp
    for pts, a, b in ((D.KPD_5_I, s0, tsc0), (D.KPD_5_I, tsc1, s1), (D.KPD_I_II, s0, tsc0), (D.KPD_I_II, tsc1, s1),
                      (D.KPD_DD_5, D.KPD_DD_5[0][0], s1)):
        to.pline([(PX(s), QY(z)) for s, z in doan(pts, a, b)], 'DIA_CHAT', linetype='DASHED', ltscale=0.5)
    # tràn sự cố (bê tông)
    ti0, ti1, zi0, zi1 = D.KPD_TSC['trong']
    for a, b, c, d in ((tsc0, ti0, zt0, zt1), (ti1, tsc1, zt0, zt1), (ti0, ti1, zt0, zi0), (ti0, ti1, zi1, zt1)):
        p = [(PX(a), QY(c)), (PX(b), QY(c)), (PX(b), QY(d)), (PX(a), QY(d))]
        to.hatch(p, 'AR-CONC', 0.03)
        to.pline(p, 'NET_CHINH', dong=True)
    to.text('TRÀN', (PX(0.5 * (ti0 + ti1)), QY(75.2)), 2.0, 'MC')
    to.text('SỰ CỐ', (PX(0.5 * (ti0 + ti1)), QY(74.6)), 2.0, 'MC')
    # màng đề xuất (gạch chéo) và đáy màng TKCS
    ss = [i * 0.5 for i in range(int(L_MAN / 0.5) + 1)]
    vung_man = [(PX(0), QY(D.KPD_DINH_MAN)), (PX(L_MAN), QY(D.KPD_DINH_MAN))] + \
               [(PX(s), QY(day_dx(s))) for s in reversed(ss)]
    to.hatch(vung_man, 'ANSI31', 1.1, lop='SUA_CHUA', mau_to=1)
    to.pline(vung_man, 'SUA_CHUA', dong=True)
    to.pline([(PX(s), QY(z)) for s, z in D.KPD_DAY_TK] + [(PX(L_MAN), QY(D.KPD_DINH_MAN))], 'SUA_CHUA',
             linetype='DASHED', ltscale=0.4)
    to.line((PX(L3), QY(80.2)), (PX(L3), QY(66.0)), 'SUA_CHUA', linetype='DASHDOT', ltscale=0.4)
    to.text('ĐOẠN 3 HÀNG', (PX(L3) - 1.5, QY(80.0)), 2.0, 'RIGHT', 'SUA_CHUA')
    to.text('ĐOẠN 2 HÀNG (A, B)', (PX(L3) + 1.5, QY(80.0)), 2.0, 'LEFT', 'SUA_CHUA')
    # hố khoan chiếu lên tim màng: liền nét tới đáy đề xuất, nét đứt tới đáy TKCS; hàng A (và hàng C trong đoạn
    # 3 hàng) tại s = 0, 3, 6…; hàng B so le tại s = 1,5; 4,5…
    for hang, ds in (('A', S_AC), ('B', S_B_TK)):
        for i, s in enumerate(ds):
            zm, zd, zk = z_mat(s), day_dx(s), day_tk(s)
            to.line((PX(s), QY(zm)), (PX(s), QY(zd)), 'NET_CHINH')
            if zd - zk > 0.05:
                to.line((PX(s), QY(zd)), (PX(s), QY(zk)), 'NET_MANH', linetype='DASHED', ltscale=0.2)
            if hang == 'A':
                nhan = 'A%02d/C%02d' % (i + 1, i + 1) if s <= L3 else 'A%02d' % (i + 1)
            else:
                nhan = 'B%02d' % (i + 1)
            to.text(nhan, (PX(s), QY(zm) + 1.2), 1.7, 'CENTER')
    # hố khảo sát CG7
    c = D.KPD_CG7
    xc = PX(c['s'])
    to.rect(xc - 0.8, QY(c['day']), 1.6, QY(c['mat']) - QY(c['day']), 'NET_CHINH')
    for z, sau in c['ranh']:
        to.line((xc - 0.8, QY(z)), (xc + 2.5, QY(z)), 'NET_MANH')
        to.text(sau, (xc + 2.7, QY(z)), 1.6, 'ML')
    to.text('10,0', (xc + 2.7, QY(c['day'])), 1.6, 'ML')
    to.text('CG7', (xc, QY(c['mat']) + 4.5), 2.0, 'CENTER')
    # mực nước ngầm khi khoan
    to.pline([(PX(s), QY(z)) for s, z in doan(D.KPD_MNN, s0, tsc0)], 'NUOC', linetype='DASHED', ltscale=0.4)
    to.pline([(PX(s), QY(z)) for s, z in doan(D.KPD_MNN, tsc1, D.KPD_MNN[-1][0])], 'NUOC', linetype='DASHED',
             ltscale=0.4)
    to.cao_trinh(xc - 3.0, QY(D.KPD_MNN[-1][1]), 0, ben='L', h=1.7, nuoc=True, chu='MNN khi khoan +73,77')
    # nhãn lớp, hệ số thấm
    for s, z, ten in ((29.5, 76.4, '5'), (31.0, 73.0, 'I'), (40.0, 68.6, 'II'), (47.0, 77.6, 'ĐĐ-I'),
                      (-8.6, 77.9, '5'), (-8.6, 76.2, 'I'), (-8.6, 72.0, 'II')):
        r = 1.6 if len(ten) < 3 else 2.4
        to.circle((PX(s), QY(z)), r, 'NET_MANH')
        to.text(ten, (PX(s), QY(z)), 1.7, 'MC')
    to.text('K = 6,54E-05 cm/s', (PX(22.6), QY(76.1)), 1.6, 'LEFT')
    to.text('K = 4,35E-04 cm/s', (PX(22.6), QY(73.9)), 1.6, 'LEFT')
    to.text('MĐ', (PX(51.0), QY(78.65)), 1.6, 'MC')
    # cao trình
    to.cao_trinh(PX(L_MAN) + 1.0, QY(D.KPD_DINH_MAN), D.KPD_DINH_MAN, h=1.8, chu='+77,00 đỉnh màng')
    to.cao_trinh(PX(0) + 1.0, QY(70.04), 0, h=1.7, chu='+70,04')
    to.cao_trinh(PX(L_MAN) + 1.0, QY(67.0), 0, h=1.7, chu='+67,00')
    to.chu_thich_mui_ten((PX(9.0), QY(day_tk(9.0))), (PX(13.0), QY(66.6)), 'Đáy màng TKCS (giới hạn sâu nhất)',
                         h=1.8, rong=52)
    to.chu_thich_mui_ten((PX(27.0), QY(day_dx(27.0))), (PX(29.0), QY(66.6)),
                         'Đáy màng đề xuất = mặt lớp II − 3,0 m (TCVN 8216:2018, điều 9.2.2.9)', h=1.8, rong=84)
    for ten, s in (('C23', 0.0), ('C22', D.KPD_C22), ('C22+13,97', L_MAN)):
        to.text(ten, (PX(s), QY(79.4)), 1.9, 'CENTER')
    # bảng số liệu dưới cắt dọc
    y0, hd = QY(65.5) - 1.0, 8.0
    hang = [('Lý trình màng (m)', lambda s: so(s, 1)), ('Mặt lớp II (m)', lambda s: so(z_II(s), 2)),
            ('Đáy màng TKCS (m)', lambda s: so(day_tk(s), 2)), ('Đáy màng đề xuất (m)', lambda s: so(day_dx(s), 2))]
    xs = S_AC + [L_MAN]
    for k, (ten, f) in enumerate(hang):
        yt, yb = y0 - k * hd, y0 - (k + 1) * hd
        to.line((25.0, yt), (PX(s1), yt), 'BANG')
        to.text(ten, (26.0, 0.5 * (yt + yb)), 2.0, 'ML', 'CHU_DAM' if k == 0 else 'CHU',
                'VN_DAM' if k == 0 else 'VN')
        for s in xs:
            to.text(f(s), (PX(s) + 0.9, 0.5 * (yt + yb)), 1.7, 'MC', rot=90)
    yb = y0 - len(hang) * hd
    to.line((25.0, yb), (PX(s1), yb), 'BANG')
    for x in (25.0, PX(-2.0), PX(s1)):
        to.line((x, y0), (x, yb), 'BANG')


# ====================================================================== mặt cắt ngang đại diện
def mat_cat(to):
    kn = KhungNhin(45.0, 168.0, -8.0, 66.0, 200)
    to.ten_hinh('MẶT CẮT NGANG ĐẠI DIỆN – SƠ ĐỒ VÀ KẾT QUẢ TÍNH THẤM (MN hồ = MNLKT +78,54)', 40, 252, '1/200',
                h=4.5, can='LEFT')
    x0, x1, zd = -8.0, 62.0, 66.0
    mat = doan(D.KP_MAT, x0, x1)
    g = lambda x: noi_suy(D.KP_MAT, x)
    l5 = lambda x: noi_suy(D.KP_L5_I, x)
    l2 = lambda x: noi_suy(D.KP_LI_II, x)
    man = D.KP_MAN

    def vung(xa, xb, tren, duoi, mau, tl, buoc=0.5):
        n = max(2, int((xb - xa) / buoc) + 1)
        xs = [xa + (xb - xa) * k / (n - 1) for k in range(n)]
        p = [kn(x, tren(x)) for x in xs] + [kn(x, duoi(x)) for x in reversed(xs)]
        if max(abs(tren(x) - duoi(x)) for x in xs) > 0.02:
            to.hatch(p, mau, tl)
    t5 = g
    d5 = lambda x: min(l5(x), g(x))
    dI = lambda x: min(l2(x), g(x))
    for a, b in ((x0, man['xa']), (man['xb'], x1)):
        vung(a, b, t5, d5, 'AR-SAND', 0.12)
        vung(a, b, d5, dI, 'ANSI37', 1.4)
        vung(a, b, dI, lambda x: zd, 'ANSI33', 1.8)
    vung(man['xa'], man['xb'], t5, lambda x: man['ztop'], 'AR-SAND', 0.12)
    vung(man['xa'], man['xb'], lambda x: man['zbot'], lambda x: zd, 'ANSI33', 1.8)
    # màng thiết kế (T = 4,8 m), đáy đề xuất theo điều 9.2.2.9 (mặt lớp II − 3 m), biên màng 2 hàng A, B (T = 3,3 m)
    pm = [kn(man['xa'], man['ztop']), kn(man['xb'], man['ztop']), kn(man['xb'], man['zbot']), kn(man['xa'], man['zbot'])]
    to.hatch(pm, 'ANSI31', 0.6, lop='SUA_CHUA', mau_to=1)
    to.pline(pm, 'SUA_CHUA', dong=True)
    z_dx = round(noi_suy(D.KP_LI_II, D.KP_HANG['B']) - 3.0, 2)            # +69,40 tại mặt cắt đại diện
    to.line(kn(man['xa'], z_dx), kn(man['xb'], z_dx), 'SUA_CHUA', linetype='DASHED', ltscale=0.3)
    xb2 = D.KP_HANG['B'] + 0.9                                           # biên hạ lưu màng 2 hàng A, B
    to.line(kn(xb2, man['ztop']), kn(xb2, z_dx), 'SUA_CHUA', linetype='DASHED', ltscale=0.3)
    # mặt đất, ranh giới lớp, biên cắt
    to.pline([kn(*p) for p in mat], 'NET_CHINH')
    for L in (D.KP_L5_I, D.KP_LI_II):
        xs = [x0 + k * 0.25 for k in range(int((x1 - x0) / 0.25) + 1)]
        seg = []
        for x in xs:
            z = noi_suy(L, x)
            if z < g(x) - 0.02 and not (man['xa'] < x < man['xb']):
                seg.append(kn(x, z))
            else:
                if len(seg) > 1:
                    to.pline(seg, 'DIA_CHAT', linetype='DASHED', ltscale=0.4)
                seg = []
        if len(seg) > 1:
            to.pline(seg, 'DIA_CHAT', linetype='DASHED', ltscale=0.4)
    to.line(kn(x0, zd), kn(x1, zd), 'NET_MANH')
    for x in (x0, x1):
        to.line(kn(x, zd), kn(x, g(x)), 'NET_MANH')
    # tường chắn sóng (sơ đồ)
    pt = [kn(10.56, 78.0), kn(11.06, 78.0), kn(11.06, D.TUONG_CS), kn(10.56, D.TUONG_CS)]
    to.hatch(pt, 'AR-CONC', 0.02)
    to.pline(pt, 'NET_CHINH', dong=True)
    # mực nước
    xw = noi_suy([(z, x) for x, z in D.KP_MAT[1:6]], D.MN['MNLKT'])
    to.muc_nuoc(kn(x0, 0)[0], kn(xw, 0)[0], kn(0, D.MN['MNLKT'])[1], 'MNLKT', D.MN['MNLKT'], nhan_x=kn(-6.0, 0)[0])
    for ten in ('MNLTK', 'MNDBT'):
        xw2 = noi_suy([(z, x) for x, z in D.KP_MAT[1:6]], D.MN[ten])
        y = kn(0, D.MN[ten])[1]
        to.line((kn(x0, 0)[0], y), (kn(xw2, 0)[0], y), 'NUOC', linetype='DASHED', ltscale=0.3)
        to.text('%s %s' % (ten, so(D.MN[ten], 2, True)), (kn(x0, 0)[0] + 1.0, y + 0.6), 1.8, 'LEFT', 'NUOC')
    # đường bão hòa MNLKT: chưa có màng (tím, nét đứt) và có màng thiết kế (đỏ, liền)
    for tt, kieu, mau in (('HT', 'DASHED', 6), ('TK', 'CONTINUOUS', 1)):
        zb = D.KP_ZBH[tt][78.54]
        xr, zr = D.KP_DIEM_RA[(tt, 78.54)]
        pts = [(xw, D.MN['MNLKT'])] + [(x, z) for x, z in sorted(zb.items()) if x < xr] + [(xr, zr)]
        pts += [p for p in D.KP_MAT if xr < p[0] <= 45.0] + [(45.0, zb[45.0])]
        to.pline([kn(*p) for p in pts], 'THAM', linetype=kieu, ltscale=0.25, color=mau)
        X, Y = kn(xr, zr)
        to.pline([(X, Y), (X - 1.2, Y + 2.0), (X + 1.2, Y + 2.0)], 'THAM', dong=True, color=mau)
    to.chu_thich_mui_ten(kn(36.02, 75.145), kn(42.0, 78.9), 'Điểm ra +75,14 (chưa có màng)', h=1.9, rong=46)
    to.chu_thich_mui_ten(kn(37.48, 74.75), kn(45.0, 77.4), 'Điểm ra +74,75 (có màng, tại chân mái)', h=1.9, rong=58)
    to.chu_thich_mui_ten(kn(29.0, 76.0), kn(33.0, 80.0), 'Đường bão hòa chưa có màng', h=1.9, rong=42)
    to.chu_thich_mui_ten(kn(28.0, 75.25), kn(30.0, 72.4), 'Đường bão hòa có màng thiết kế', h=1.9, rong=48)
    # hàng khoan, kích thước
    for ten, x in D.KP_HANG.items():
        to.line(kn(x, 78.75), kn(x, man['zbot']), 'NET_CHINH')
        to.text(ten, (kn(x, 0)[0], kn(0, 78.75)[1] + 1.2), 2.2, 'CENTER')
    yk = kn(0, 78.75)[1] + 6.0
    to.dim((kn(D.KP_HANG['A'], 0)[0], yk), (kn(D.KP_HANG['B'], 0)[0], yk), 0, 200)
    to.dim((kn(D.KP_HANG['B'], 0)[0], yk), (kn(D.KP_HANG['C'], 0)[0], yk), 0, 200)
    yT = kn(0, man['zbot'])[1] - 5.0
    to.dim((kn(man['xa'], 0)[0], yT), (kn(man['xb'], 0)[0], yT), 0, 200, text='T = <>')
    # giếng đo áp: P2 (đoạn 2 hàng, 5 m sau hàng B), P1 (đoạn 3 hàng, 5 m sau hàng C) – chiếu lên mặt cắt
    for ten, _, xg in GIENG:
        gieng(to, *kn(xg, 78.75 + 0.6))
        to.text(ten, (kn(xg, 0)[0], kn(0, 78.75 + 0.6)[1] + 2.0), 1.8, 'CENTER', 'NUOC')
        to.line(kn(xg, 78.75), kn(xg, 72.6), 'NUOC')
        to.rect(kn(xg, 0)[0] - 0.4, kn(0, 72.6)[1], 0.8, kn(0, 74.2)[1] - kn(0, 72.6)[1], 'NUOC')
    to.text('Giếng P2, P1', (kn(max(g[2] for g in GIENG), 0)[0] + 1.6, kn(0, 73.6)[1]), 1.9, 'LEFT', 'NUOC')
    to.chu_thich_mui_ten(kn(xb2, 71.8), kn(xb2 + 4.0, 71.0), 'Biên màng 2 hàng A, B (T = 3,3 m)', h=1.8, rong=50)
    # cao trình
    to.cao_trinh(*kn(man['xb'] + 0.3, man['ztop']), 0, h=1.8, chu='+77,00')
    to.cao_trinh(*kn(man['xb'] + 0.3, man['zbot']), 0, h=1.8, chu='+68,30 (đáy TKCS)')
    to.cao_trinh(*kn(man['xb'] + 0.3, z_dx), 0, h=1.8, chu='%s (đáy đề xuất: mặt lớp II − 3 m)' % so(z_dx, 2, True))
    to.cao_trinh(*kn(10.81, D.TUONG_CS), 0, ben='L', h=1.8, chu='+79,40')
    to.cao_trinh(*kn(13.6, 78.75), 0, h=1.8, chu='+78,75')
    to.cao_trinh(*kn(D.KP_X_CHAN, 74.75), 0, h=1.8, chu='+74,75')
    for x, z, ten in ((4.0, 76.0, '5'), (2.0, 73.4, 'I'), (6.0, 69.0, 'II'), (55.0, 73.0, 'I'), (55.0, 70.0, 'II'),
                      (30.0, 74.6, '5')):
        X, Y = kn(x, z)
        to.circle((X, Y), 1.7, 'NET_MANH')
        to.text(ten, (X, Y), 1.8, 'MC')
    X, Y = kn(17.81, 66.0)
    to.text('MÀNG KHOAN PHỤT', (X, Y - 3.2), 2.0, 'CENTER', 'SUA_CHUA')


# ====================================================================== biểu đồ độ nhạy
def bieu_do(to, x0, y0, rong=88.0, cao=46.0):
    """Hình 8.5: η theo K màng (trục log) và theo cao trình đáy màng."""
    to.text('HIỆU QUẢ GIẢM LƯU LƯỢNG THẤM η – MN hồ MNLTK (Hình 8.5)', (x0 + rong + 9.0, y0 + cao + 14.0), 3.3,
            'CENTER', 'CHU_DAM', 'VN_DAM')

    def khung(xa, ten_x, ticks_x, lab_x, f_x, du_lieu, tieu_de, ghi=None):
        to.rect(xa, y0, rong, cao, 'NET_CHINH')
        to.text(tieu_de, (xa + rong / 2, y0 + cao + 4.0), 2.2, 'CENTER', 'CHU_DAM', 'VN_DAM')
        for v in range(0, 81, 20):
            y = y0 + cao * v / 80.0
            to.line((xa, y), (xa + rong, y), 'NET_MANH')
            to.text('%d' % v, (xa - 1.2, y), 1.8, 'MR')
        to.text('η (%)', (xa - 7.0, y0 + cao / 2), 1.9, 'MC', rot=90)
        for t, lab in zip(ticks_x, lab_x):
            x = xa + rong * f_x(t)
            to.line((x, y0), (x, y0 + cao), 'NET_MANH')
            to.text(lab, (x, y0 - 2.8), 1.8, 'CENTER')
        to.text(ten_x, (xa + rong / 2, y0 - 7.2), 1.9, 'CENTER')
        pts = [(xa + rong * f_x(a), y0 + cao * b / 80.0) for a, b in du_lieu]
        to.pline(pts, 'SUA_CHUA')
        for p in pts:
            to.circle(p, 0.8, 'SUA_CHUA')
            to.hatch([(p[0] + 0.8 * math.cos(k * math.pi / 6), p[1] + 0.8 * math.sin(k * math.pi / 6)) for k in range(12)],
                     'SOLID', lop='SUA_CHUA', mau_to=1)
        for a, b in du_lieu:
            to.text('%d%%' % b, (xa + rong * f_x(a) + 1.2, y0 + cao * b / 80.0 + 1.2), 1.7, 'LEFT', 'SUA_CHUA')
        if ghi:
            for (gx, nhan) in ghi:
                x = xa + rong * f_x(gx)
                to.line((x, y0), (x, y0 + cao), 'DIA_CHAT', linetype='DASHED', ltscale=0.3)
                to.text(nhan, (x - 0.8, y0 + 2.0), 1.7, 'LEFT', rot=90)
    fK = lambda k: (math.log10(k) + 6.0) / 2.0
    khung(x0, 'Hệ số thấm màng K màng (cm/s)', [1e-6, 1e-5, 1e-4], ['1E-06', '1E-05', '1E-04'], fK,
          [(1e-6, 71), (5e-6, 62), (1e-5, 54), (5e-5, 23), (1e-4, 11)], 'a) Ảnh hưởng chất lượng màng')
    fz = lambda z: (74.5 - z) / (74.5 - 65.5)
    khung(x0 + rong + 18.0, 'Cao trình đáy màng (m)', [74, 72, 70, 68, 66], ['+74', '+72', '+70', '+68', '+66'], fz,
          [(74.0, 7), (72.0, 54), (70.0, 54), (68.3, 54), (66.0, 54)], 'b) Ảnh hưởng chiều sâu màng',
          ghi=[(72.4, 'mặt lớp II')])


# ====================================================================== sơ đồ hố khoan
def so_do_ho_khoan(to, x0, y0):
    """Sơ đồ cấu tạo hố khoan và phân đoạn phụt (không tỷ lệ); y0 = mặt đập."""
    to.ten_hinh('SƠ ĐỒ CẤU TẠO HỐ KHOAN, PHÂN ĐOẠN PHỤT', x0, y0 + 30, 'không tỷ lệ', h=3.8, can='LEFT')
    xl, xc, xr = x0 + 4, x0 + 40, x0 + 118          # mép trái khối, tim hố, mép phải
    ys = y0
    lop = [('Mặt đập + đất đắp (tầng phản áp, không phụt)', ys, ys - 14, 'GRAVEL', 0.08),
           ('Lớp 5', ys - 14, ys - 40, 'AR-SAND', 0.1),
           ('Lớp I', ys - 40, ys - 66, 'ANSI37', 0.9),
           ('Lớp II', ys - 66, ys - 108, 'ANSI33', 1.2)]
    for ten, a, b, mau, tl in lop:
        to.hatch([(xl, a), (xc - 3, a), (xc - 3, b), (xl, b)], mau, tl)
        to.line((xl, b), (xr, b), 'DIA_CHAT', linetype='DASHED', ltscale=0.4)
        to.text(ten, (xr + 1.5, 0.5 * (a + b)), 2.0, 'ML')
    to.line((xl, ys), (xr, ys), 'NET_CHINH')
    to.line((xl, ys), (xl, ys - 108), 'NET_MANH')
    # ống chèn D110 trong tầng phản áp, hố khoan D91
    to.rect(xc - 3, ys - 14, 6, 14 + 6, 'NET_CHINH')
    to.line((xc - 1.8, ys - 14), (xc - 1.8, ys - 108), 'NET_CHINH')
    to.line((xc + 1.8, ys - 14), (xc + 1.8, ys - 108), 'NET_CHINH')
    to.line((xc - 1.8, ys - 108), (xc + 1.8, ys - 108), 'NET_CHINH')
    to.chu_thich_mui_ten((xc + 3, ys - 5), (xc + 26, ys - 4), 'Ống chèn D110', h=1.9, rong=24)
    to.chu_thich_mui_ten((xc + 1.8, ys - 100), (xc + 26, ys - 104), 'Hố khoan D91', h=1.9, rong=24)
    # phân đoạn phụt từ trên xuống
    dd = [(ys - 14, ys - 40, 'Đoạn 1'), (ys - 40, ys - 66, 'Đoạn 2'), (ys - 66, ys - 87, 'Đoạn 3'),
          (ys - 87, ys - 108, 'Đoạn 4')]
    for a, b, ten in dd:
        to.line((xc + 5.0, a), (xc + 5.0, b), 'SUA_CHUA')
        to.line((xc + 4.0, a), (xc + 6.0, a), 'SUA_CHUA')
        to.text(ten, (xc + 7.0, 0.5 * (a + b)), 1.9, 'ML', 'SUA_CHUA')
        to.rect(xc - 2.4, a - 1.6, 4.8, 1.6, 'NET_CHINH')            # nút phụt (packer) đầu đoạn
    to.line((xc + 4.0, ys - 108), (xc + 6.0, ys - 108), 'SUA_CHUA')
    to.chu_thich_mui_ten((xc + 2.4, ys - 41.0), (xc + 26, ys - 46.0), 'Nút phụt đầu đoạn', h=1.9, rong=28)
    to.text('Vữa xi măng – bentonite', (xc + 26.0, ys - 25.0), 1.9, 'LEFT', 'SUA_CHUA')
    to.text('(đất đắp, lớp 5, lớp I)', (xc + 26.0, ys - 28.2), 1.9, 'LEFT', 'SUA_CHUA')
    to.text('Vữa xi măng (lớp II)', (xc + 26.0, ys - 86.0), 1.9, 'LEFT', 'SUA_CHUA')
    to.text('Mỗi đoạn 2 ÷ 3 m; khoan – phụt –', (xc + 26.0, ys - 58.0), 1.9, 'LEFT')
    to.text('chờ ninh kết – khoan đoạn tiếp theo', (xc + 26.0, ys - 61.2), 1.9, 'LEFT')
    # đường ống phụt, hồi vữa, thiết bị
    yt = ys + 16
    to.line((xc - 0.6, ys + 6), (xc - 0.6, yt), 'NET_CHINH')
    to.line((xc + 0.6, ys + 6), (xc + 0.6, yt - 5), 'NET_CHINH')
    to.line((xc - 0.6, yt), (xc + 44, yt), 'NET_CHINH')
    to.line((xc + 0.6, yt - 5), (xc + 44, yt - 5), 'NET_CHINH')
    to.circle((xc + 10, yt), 1.6, 'NET_CHINH')
    to.text('Đồng hồ áp lực', (xc + 10, yt + 2.6), 1.8, 'CENTER')
    mui_ten(to, (xc + 36, yt), (xc + 28, yt), 'NET_CHINH', 1.5)
    mui_ten(to, (xc + 28, yt - 5), (xc + 36, yt - 5), 'NET_CHINH', 1.5)
    to.text('ống phụt', (xc + 32, yt + 0.8), 1.7, 'CENTER')
    to.text('ống hồi vữa', (xc + 32, yt - 8.0), 1.7, 'CENTER')
    to.rect(xc + 44, yt - 8, 18, 11, 'NET_CHINH')
    to.text('Bơm vữa', (xc + 53, yt - 1.6), 1.8, 'MC')
    to.text('cao áp', (xc + 53, yt - 4.6), 1.8, 'MC')
    to.rect(xc + 68, yt - 8, 22, 11, 'NET_CHINH')
    to.text('Máy trộn +', (xc + 79, yt - 1.6), 1.8, 'MC')
    to.text('thùng đo vữa', (xc + 79, yt - 4.6), 1.8, 'MC')
    to.line((xc + 62, yt - 2.5), (xc + 68, yt - 2.5), 'NET_CHINH')


# ====================================================================== bảng
def cac_bang(to):
    x0 = 418.0
    cot = [17, 21, 21, 11, 23, 33, 15, 12]
    hang = [['MN hồ', 'q chưa có\nmàng', 'q có\nmàng', 'η', 'J ra chưa\n/ có màng', 'Điểm thấm ra\nchưa / có màng',
             'ΔHm max\n(m)', 'Jm TB']] + [list(r) for r in D.KP_KQ]
    y = to.bang(x0, 576, cot, hang, h=2.1, cao_dong=5.0,
                tieu_de='KẾT QUẢ TÍNH THẤM MẶT CẮT ĐẠI DIỆN (Bảng 8.2)')
    to.text('q: m³/ngđ/m dài; η = 1 − q có màng / q chưa có màng; "chân xa": thấm ra ở vùng trũng x > 60 m.',
            (x0, y - 3.5), 1.8)
    # dự báo mực nước giếng đo áp: P1 – màng 3 hàng (TK), P2 – màng 2 hàng A, B (nội suy giữa x = 20,8 và 25,9 m)
    def zg(tt, z, xg):
        b = D.KP_ZBH[tt][z]
        return b[20.8] + (b[25.9] - b[20.8]) * (xg - 20.8) / (25.9 - 20.8)
    (_, _, x1), (_, _, x2) = GIENG
    cot2 = [20, 32, 32]
    hang2 = [['MN hồ', 'P1 (đoạn 3 hàng)\nchưa / có màng', 'P2 (đoạn 2 hàng)\nchưa / có màng']] + \
            [[so(z, 2, True), '%s / %s' % (so(zg('HT', z, x1), 2, True), so(zg('TK', z, x1), 2, True)),
              '%s / %s' % (so(zg('HT', z, x2), 2, True), so(zg('2H', z, x2), 2, True))] for z in (74.70, 77.63, 78.54)]
    y2 = to.bang(x0 + 172, 576, cot2, hang2, h=2.1, cao_dong=5.0, tieu_de='MỰC NƯỚC DỰ BÁO TẠI GIẾNG P1, P2')
    to.text('P1 cách hàng C 5 m (x = %s m), P2 cách hàng B 5 m (x = %s m) về hạ lưu.' % (so(x1, 2), so(x2, 2)),
            (x0 + 172, y2 - 3.5), 1.8)
    to.text('Mực nước đo ≈ "chưa có màng": màng chưa hiệu quả.', (x0 + 172, y2 - 6.5), 1.8)

    # độ nhạy
    yt = 520.0
    to.text('PHÂN TÍCH ĐỘ NHẠY – MN hồ MNLTK +77,63 (Bảng 8.3 ÷ 8.6)', (x0 + 135, yt + 6.0), 3.3, 'CENTER',
            'CHU_DAM', 'VN_DAM')
    bangs = [('a) Hệ số thấm màng', [22, 12, 12], [['K màng (cm/s)', 'η', 'J ra']] +
              [[a.replace('·10⁻', 'E-0'), b, c] for a, b, c in D.KP_NHAY_K]),
             ('b) Cao trình đáy màng', [22, 12, 12], [['Đáy màng (m)', 'η', 'J ra']] + [list(r) for r in D.KP_NHAY_DAY]),
             ('c) Hệ số thấm lớp II', [22, 12, 22], [['K lớp II (cm/s)', 'η', 'J ra chưa / có']] +
              [[a.replace('·10⁻', 'E-0'), b, c] for a, b, c in D.KP_NHAY_KII]),
             ('d) Số hàng khoan, đỉnh màng', [36, 24, 24], [['Bố trí', 'η MNLTK\nđỉnh +77 / +78', 'η MNLKT\nđỉnh +77 / +78']] +
              [list(r) for r in D.KP_NHAY_T])]
    xx = x0
    for ten, c, h in bangs:
        to.text(ten, (xx, yt - 4.0), 2.2, 'LEFT', 'CHU_DAM', 'VN_DAM')
        yd = to.bang(xx, yt - 6.0, c, h, h=2.0, cao_dong=4.6)
        xx += sum(c) + 6.0
    xd = xx - sum(bangs[-1][1]) - 6.0
    to.text('Vị trí dải 2 hàng (A–B, B–C, đối xứng qua tim) hầu như', (xd, yd - 3.2), 1.8)
    to.text('không ảnh hưởng: η = 47,2 ÷ 47,4 % (MNLTK).', (xd, yd - 6.0), 1.8)
    to.text('Kết luận: hiệu quả màng nhạy nhất với chất lượng phụt (K màng) và việc cắt hết lớp I; '
            'chiều sâu cắm vào lớp II phụ thuộc K lớp II.', (x0, yt - 41.0), 2.0)

    # so sánh phương án (η, J ra tại mặt cắt đại diện – thuộc đoạn 2 hàng)
    tk, a, b = khoi_luong()
    pt = lambda v, g: ('%+.0f%%' % (100 * (v / g - 1))).replace('-', '−')
    pc = lambda v, g: '%s (%s)' % (so(v, 1), pt(v, g))
    E = D.KP_ETA_PA
    eta = lambda pa, k: '%s / %s' % E[pa][k]
    cot3 = [46, 46, 64, 64]
    hang3 = [['Chỉ tiêu', 'TKCS', 'Hợp lý hóa – đáy TKCS', 'Hợp lý hóa – đáy theo điều 9.2.2.9'],
             ['Bố trí hàng khoan', '3 hàng A, B, C; L = 50 m', '3 hàng 15 m + 2 hàng A, B 35 m',
              '3 hàng 15 m + 2 hàng A, B 35 m'],
             ['Chiều dày hiệu quả T', '4,8 m', '4,8 / 3,3 m', '4,8 / 3,3 m'],
             ['Đỉnh / đáy màng (m)', '+77,00 / +70,0 ÷ +67,0', '+77,00 / như TKCS',
              '+77,00 / mặt lớp II − 3 m (≥ đáy TKCS)'],
             ['Số hố khoan phụt', '%d' % tk[0], '%d (%s)' % (a[0], pt(a[0], tk[0])), '%d (%s)' % (b[0], pt(b[0], tk[0]))],
             ['Tổng chiều dài khoan (m)', so(tk[1], 1), pc(a[1], tk[1]), pc(b[1], tk[1])],
             ['Tổng chiều dài phụt (m)', so(tk[2], 1), pc(a[2], tk[2]), pc(b[2], tk[2])],
             ['– trong đó trong lớp II (m)', so(tk[3], 1), pc(a[3], tk[3]), pc(b[3], tk[3])],
             ['η / J ra – MNLTK +77,63*', eta('TK', 0), eta('HLH_TK', 0), eta('HLH_9229', 0)],
             ['η / J ra – MNLKT +78,54*', eta('TK', 1), eta('HLH_TK', 1), eta('HLH_9229', 1)],
             ['Điều kiện áp dụng', '–', 'Phụt thí nghiệm 2 hàng A, B đạt', 'Như cột trước; đáy màng xác định'],
             ['', '', 'K ≤ 1E-05 cm/s; lớp II q ≤ 1 Lu', 'bằng ép nước lớp II (tiêu chí 5 Lu)']]
    y3 = 466.0
    yb = to.bang(x0, y3, cot3, hang3, h=2.0, cao_dong=4.6,
                 tieu_de='SO SÁNH PHƯƠNG ÁN MÀNG TKCS VÀ PHƯƠNG ÁN HỢP LÝ HÓA (mục 8.6)')
    to.text('Chiều dài khoan, phụt tính theo hình học cắt dọc (đỉnh phụt +77,00; khoan từ mặt đập hiện trạng) để so sánh; '
            'khối lượng thanh toán theo TKCS (Bảng 7.1).', (x0, yb - 3.5), 1.8)
    to.text('(*) η, J ra tại mặt cắt đại diện (thuộc đoạn 2 hàng); đoạn 3 hàng sát tràn giữ η = 54% (MNLTK), 52% (MNLKT).',
            (x0, yb - 6.5), 1.8)
    to.text('η ứng với K lớp II = 1E-05 cm/s (≈ 0,8 Lu); nếu lớp II thấm tới 5 Lu (K ≈ 6,5E-05) η mọi phương án chỉ còn '
            '26 ÷ 32% (Bảng 8.8), ở MNLKT dòng thấm vẫn ra mái hạ lưu (+74,95).', (x0, yb - 9.5), 1.8)
    # khối lượng TKCS
    cot4 = [72, 16, 30]
    hang4 = [['Hạng mục (TKCS, Bảng 7.1)', 'Đơn vị', 'Khối lượng'],
             ['Khoan trong đất đắp, lớp phong hóa / trong đá lớp II', 'm', '333,25 / 146,69'],
             ['Phụt vữa xi măng – sét / vữa xi măng', 'm', '256,75 / 146,69'],
             ['Xi măng PC40 / bentonite', 'kg', '14.853 / 15.405'],
             ['Hố khoan kiểm tra; đoạn thí nghiệm đổ, ép nước', 'hố; đoạn', '3 ; 9']]
    to.bang(x0 + 228, y3, cot4, hang4, h=2.0, cao_dong=4.6, tieu_de='KHỐI LƯỢNG KHOAN PHỤT (TKCS)')
    return yb


def ghi_chu(to, x, y):
    gc = ('GHI CHÚ:\\P'
          '1. Cao trình ghi bằng m, kích thước ghi bằng cm. Lý trình màng tính từ mặt ngoài tường bên trái tràn sự cố '
          '(cọc C23) về phía đập. \\P'
          '2. Màng theo TKCS: dài 50 m, 3 hàng A, B, C cách nhau 1,5 m, hố cách 3,0 m, T = Bb + 0,6A = 3,0 + 1,8 = 4,8 m; '
          'đỉnh +77,00, đáy +70,0 ÷ +67,0; K màng ≤ 1E-05 cm/s. \\P'
          '3. Phương án hợp lý hóa (Chương 8): giữ 3 hàng A, B, C trên đoạn 15 m sát tràn (C01 ÷ C06); đoạn 15 ÷ 50 m '
          'chỉ phụt hai hàng kề nhau A, B (cách 1,5 m, hố so le; T = Bb + 0,6A = 1,5 + 1,8 = 3,3 m), bỏ hàng C. Chỉ áp '
          'dụng khi phụt thí nghiệm 2 hàng A, B trong vùng phụt thử cho K ≤ 1E-05 cm/s (≈ 0,8 Lu) ở hố kiểm tra KT1 và '
          'ép nước lớp II cho q ≤ 1 Lu (quy đổi 1 Lu ≈ 1,3E-05 cm/s); nếu 1 < q ≤ 5 Lu phải tính lại η theo K đo được, '
          'cân nhắc giữ 3 hàng trên toàn tuyến; nếu màng thử không đạt thì phụt đủ 3 hàng như TKCS. \\P'
          '4. Đáy màng theo TCVN 8216:2018 điều 9.2.2.9 (đập cấp II: màng không sâu quá cao trình nền có lượng mất nước '
          '(5 ÷ 7) Lu cộng thêm 3 m): đáy = mặt lớp II − 3,0 m khi ép nước lớp II cho q ≤ 5 Lu, không sâu hơn đáy '
          'TKCS; nếu lớp II nứt nẻ (q > 5 Lu) thì đáy = cao trình tại đó q giảm còn 5 Lu, trừ 3 m. Hiệu quả η trên bản vẽ '
          'ứng với K lớp II = 1E-05 cm/s (≈ 0,8 Lu); với q tới 5 Lu η chỉ còn 26 ÷ 32% (Bảng 8.8) – tính lại theo K đo '
          'được. \\P'
          '5. Trình tự: phụt thí nghiệm → đoạn 3 hàng: hàng C → A → B; đoạn 2 hàng: hàng B → A (hàng hạ lưu trước); '
          'trong mỗi hàng phụt hố thứ tự I trước, hố thứ tự II sau; phân đoạn phụt từ trên xuống, mỗi đoạn 2 ÷ 3 m '
          '(TCVN 8645:2011, TCVN 8644:2011). \\P'
          '6. Vữa xi măng – bentonite trong đất, lớp 5, lớp I; vữa xi măng trong lớp II; vữa loãng trước, đặc dần. '
          'Áp lực phụt tăng dần theo độ sâu, không gây nứt tách đất đắp (TKCS: 0,75 ÷ 5,0 kG/cm²); theo dõi '
          'đẩy trồi mặt đập và chuyển vị tường bên trong khi phụt. \\P'
          '7. Kết thúc đoạn phụt khi lượng tiêu hao vữa nhỏ hơn quy định trong thời gian quy định dưới áp lực thiết '
          'kế; lấp hố bằng vữa xi măng – sét. \\P'
          '8. Kiểm tra chất lượng: 3 hố KT ở vị trí xa hố phụt nhất (giữa hai hố cùng hàng; KT2, KT3 điều chỉnh theo '
          'đoạn tiêu hao vữa lớn), thí nghiệm đổ nước, ép nước 9 đoạn; màng đạt khi K ≤ 1E-05 cm/s (≈ 0,8 Lu). \\P'
          '9. Khoan phụt thi công trước khi đắp bù, đổ bê tông mặt đập (BV-04) và trước khi lát lại tấm BT mái tràn '
          'sự cố. \\P'
          '10. Quan trắc sau xử lý: giếng đo áp P1 (đoạn 3 hàng, cách hàng C 5 m), P2 (đoạn 2 hàng, cách hàng B 5 m), '
          'đầu đo trong lớp I; so sánh với mực nước dự báo; theo dõi thấm ở mái hạ lưu khi mực nước hồ vượt +74,0 m.')
    to.mtext(gc, (x, y), 2.1, rong=205)


def chu_giai(to, x, y):
    to.text('KÝ HIỆU', (x, y), 3.0, 'LEFT', 'CHU_DAM', 'VN_DAM')
    y -= 6.0
    muc = [('ho1', 'Hố khoan phụt thứ tự I'), ('ho2', 'Hố khoan phụt thứ tự II'),
           ('hoB', 'Hố hàng C theo TKCS (bỏ khi áp dụng 2 hàng A, B)'), ('kt', 'Hố khoan kiểm tra (KT)'),
           ('gieng', 'Giếng đo áp (P)'), ('ks', 'Hố khoan khảo sát CG7 (TKCS)'),
           ('man', 'Màng khoan phụt'), ('bh0', 'Đường bão hòa chưa có màng (MNLKT)'),
           ('bh1', 'Đường bão hòa có màng thiết kế (MNLKT)'), ('dc', 'Ranh giới lớp địa chất'),
           ('5', 'Lớp 5: đá cát bột kết phong hóa hoàn toàn thành sét pha lẫn dăm sạn, K = 6,54E-05 cm/s'),
           ('I', 'Lớp I: đá cát bột kết phong hóa mạnh, nứt nẻ mạnh, K = 4,35E-04 cm/s'),
           ('II', 'Lớp II: đá cát bột kết phong hóa vừa, nứt nẻ trung bình, K giả thiết 1E-05 cm/s')]
    for ma, ten in muc:
        xc, yc = x + 5.0, y
        if ma == 'ho1':
            ho(to, xc, yc, 1)
        elif ma == 'ho2':
            ho(to, xc, yc, 2)
        elif ma == 'hoB':
            to.circle((xc, yc), 0.7, 'NET_MANH')
        elif ma == 'kt':
            ho_kt(to, xc, yc)
        elif ma == 'gieng':
            gieng(to, xc, yc)
        elif ma == 'ks':
            ho_ks(to, xc, yc)
        elif ma == 'man':
            p = [(x + 1, yc - 1.8), (x + 9, yc - 1.8), (x + 9, yc + 1.8), (x + 1, yc + 1.8)]
            to.hatch(p, 'ANSI31', 0.45, lop='SUA_CHUA')
            to.pline(p, 'SUA_CHUA', dong=True)
        elif ma in ('bh0', 'bh1'):
            to.line((x + 1, yc), (x + 9, yc), 'THAM', linetype='DASHED' if ma == 'bh0' else 'CONTINUOUS', ltscale=0.25,
                    color=6 if ma == 'bh0' else 1)
        elif ma == 'dc':
            to.line((x + 1, yc), (x + 9, yc), 'DIA_CHAT', linetype='DASHED', ltscale=0.4)
        else:
            mau = {'5': ('AR-SAND', 0.12), 'I': ('ANSI37', 1.2), 'II': ('ANSI33', 1.6)}[ma]
            p = [(x + 1, yc - 1.8), (x + 9, yc - 1.8), (x + 9, yc + 1.8), (x + 1, yc + 1.8)]
            to.hatch(p, mau[0], mau[1])
            to.pline(p, 'NET_MANH', dong=True)
        to.text(ten, (x + 12.0, yc), 2.0, 'ML')
        y -= 5.6


def ve(thu_muc):
    to = To(5, 'MÀNG KHOAN PHỤT CHỐNG THẤM MANG TRÁI\nTRÀN SỰ CỐ – CHUYÊN ĐỀ', 'Như hình vẽ', 'HCG-BV05')
    # cột trái: cắt dọc (ngang 1/200, đứng 1/100 -> 1/100, 1/50), mặt cắt (1/200 -> 1/100), sơ đồ hố khoan, ký hiệu
    a = to.dat(cat_doc, f=2.0, k_chu=0.625, neo='TL', xy=(30.0, 818.0))
    b = to.dat(mat_cat, f=2.0, k_chu=0.625, neo='TL', xy=(30.0, a[1] - 22.0))
    c = to.dat(so_do_ho_khoan, 26.0, 128.0, f=1.25, neo='TL', xy=(30.0, b[1] - 22.0))
    to.dat(chu_giai, 236.0, 152.0, f=1.25, neo='TL', xy=(c[2] + 35.0, b[1] - 22.0))
    # cột phải: mặt bằng (1/200), bảng, biểu đồ, ghi chú
    xr = max(a[2], b[2]) + 25.0
    m = to.dat(mat_bang, f=1.0, k_chu=1.2, neo='TL', xy=(xr, 818.0))
    t = to.dat(cac_bang, f=1.15, neo='TL', xy=(xr, m[1] - 20.0))
    d = to.dat(bieu_do, 436.0, 255.0, f=1.25, neo='TL', xy=(xr, t[1] - 18.0))
    to.dat(ghi_chu, 418.0, 400.0, f=1.25, neo='TL', xy=(xr, d[1] - 18.0))
    to.luu(thu_muc + '/BV05_KhoanPhut_ChuyenDe.dxf', thu_muc + '/BV05_KhoanPhut_ChuyenDe.pdf',
           thu_muc + '/BV05_KhoanPhut_ChuyenDe.png')
    return to


if __name__ == '__main__':
    import sys
    tk, a, b = khoi_luong()
    print('TKCS', tk, '\nHLH-a', a, '\nHLH-b', b)
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
