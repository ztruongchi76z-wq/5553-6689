# -*- coding: utf-8 -*-
"""BV-02: CẮT DỌC TIM ĐẬP – HIỆN TRẠNG VÀ SAU SỬA CHỮA.

Số liệu: du_lieu_cat_doc_dap.py (số hóa từ bản vẽ TKCS "cắt dọc đỉnh đập": cọc, cao độ mặt đập hiện trạng, đường chân
khay, ranh giới địa tầng, nước ngầm, hố khoan). Ngang 1/2000, đứng 1/200.
"""
import du_lieu_ban_ve as D
import du_lieu_cat_doc_dap as C
from bv_lib import To, so

X0 = 98.0
SX = lambda s: X0 + s * 0.5                       # 1/2000
ZMIN, ZMAX = 50.0, 82.0
SY = lambda z: 200.0 + (min(max(z, ZMIN), ZMAX) - ZMIN) * 10.0    # 1/100
S_DAP = C.CD_DAP


def mui_ten(to, p0, p1, lop='NET_CHINH', dai=2.0, **kw):
    import math
    to.line(p0, p1, lop, **kw)
    a = math.atan2(p1[1] - p0[1], p1[0] - p0[0])
    for da in (2.7, -2.7):
        to.line(p1, (p1[0] + dai * math.cos(a + da), p1[1] + dai * math.sin(a + da)), lop, **kw)


def cat_doc(to):
    to.ten_hinh('CẮT DỌC TIM ĐẬP – HIỆN TRẠNG VÀ SAU SỬA CHỮA', 40, 576, None, h=4.5, can='LEFT')
    to.text('TỶ LỆ: NGANG 1/2000, ĐỨNG 1/100', (40, 569.5), 3.0)
    # thước cao độ
    xr = X0 - 6.0
    to.line((xr, SY(ZMIN)), (xr, SY(ZMAX)), 'NET_CHINH')
    for z in range(int(ZMIN), int(ZMAX) + 1, 1):
        to.line((xr, SY(z)), (xr + 1.5, SY(z)), 'NET_CHINH')
        if z % 2 == 0:
            to.text('%d' % z, (xr - 0.8, SY(z)), 1.7, 'MR')
    to.text('Cao độ (m)', (xr, SY(ZMAX) + 3.0), 1.8, 'CENTER')
    # mực nước hồ
    for ten, z in (('MNLKT', D.MN['MNLKT']), ('MNLTK', D.MN['MNLTK']), ('MNDBT', D.MN['MNDBT'])):
        to.line((SX(S_DAP[0]), SY(z)), (SX(S_DAP[1]), SY(z)), 'NUOC', linetype='DASHED', ltscale=0.4)
        to.text('%s %s' % (ten, so(z, 2, True)), (SX(1052.0) + 1.0, SY(z) - (2.4 if ten == 'MNLTK' else 0.0)),
                1.8, 'LEFT', 'NUOC')
    # địa tầng, chân khay, nước ngầm
    for pl in C.CD_DIA_CHAT:
        to.pline([(SX(s), SY(z)) for s, z in pl], 'DIA_CHAT', linetype='DASHED', ltscale=0.4)
    to.pline([(SX(s), SY(z)) for s, z in C.CD_CHAN_KHAY], 'NET_CHINH', linetype='DASHED', ltscale=0.5)
    for pl in C.CD_NUOC_NGAM:
        to.pline([(SX(s), SY(z)) for s, z in pl], 'NUOC', linetype='DASHDOT', ltscale=0.4)
    for ten, s, z in C.CD_TEN_LOP:
        if ten == 'KT':
            continue
        r = 1.6 if len(ten) < 3 else 2.4
        to.circle((SX(s), SY(z)), r, 'NET_MANH')
        to.text(ten, (SX(s), SY(z)), 1.6, 'MC')
    to.chu_thich_mui_ten((SX(400.0), SY(noi_suy(C.CD_CHAN_KHAY, 400.0))), (SX(420.0), SY(52.0)),
                         'Đáy chân khay đập (TKCS)', h=1.8, rong=40)
    to.chu_thich_mui_ten((SX(560.0), SY(noi_suy(C.CD_NUOC_NGAM[0], 560.0))), (SX(540.0), SY(66.0)),
                         'Mực nước ngầm khi khảo sát', h=1.8, rong=44)
    # mặt đập / mặt đất hiện trạng
    to.pline([(SX(s), SY(z)) for _, s, z in C.CD_COC], 'NET_CHINH')
    # tràn sự cố
    s0, s1, z0, z1 = C.CD_TSC
    p = [(SX(s0), SY(z0)), (SX(s1), SY(z0)), (SX(s1), SY(z1)), (SX(s0), SY(z1))]
    to.hatch(p, 'AR-CONC', 0.02)
    to.pline(p, 'NET_CHINH', dong=True)
    # màng khoan phụt C23 → C22 + 13,97 (đáy TKCS)
    pm = [(SX(63.98), SY(77.0)), (SX(113.98), SY(77.0)), (SX(113.98), SY(67.0)), (SX(63.98), SY(70.04))]
    to.hatch(pm, 'ANSI31', 0.5, lop='SUA_CHUA', mau_to=1)
    to.pline(pm, 'SUA_CHUA', dong=True)
    # đỉnh đập: TKCS +78,75 (nét đứt), đồ án +78,85 (đỏ), tường chắn sóng +79,40
    to.line((SX(S_DAP[0]), SY(D.DINH_SC)), (SX(S_DAP[1]), SY(D.DINH_SC)), 'SUA_CHUA')
    to.line((SX(S_DAP[0]), SY(78.75)), (SX(S_DAP[1]), SY(78.75)), 'NET_MANH', linetype='DASHED', ltscale=0.3, color=3)
    to.line((SX(S_DAP[0]), SY(D.TUONG_CS)), (SX(S_DAP[1]), SY(D.TUONG_CS)), 'NET_CHINH', linetype='DASHDOT',
            ltscale=0.4)
    to.chu_thich_mui_ten((SX(250.0), SY(D.TUONG_CS)), (SX(262.0), SY(81.4)),
                         'Đỉnh tường chắn sóng +79,40 (giữ nguyên)', h=1.9, rong=58)
    to.chu_thich_mui_ten((SX(660.0), SY(D.DINH_SC)), (SX(700.0), SY(81.4)),
                         'Đỉnh đập sau sửa chữa +78,85 (đồ án); TKCS +78,75', h=1.9, rong=72)
    to.chu_thich_mui_ten((SX(420.0), SY(78.36)), (SX(442.0), SY(76.6)),
                         'Mặt đập hiện trạng +78,31 ÷ +78,50', h=1.9, rong=52)
    # hố khoan
    for ten, s, zm, zd in C.CD_HO_KHOAN:
        to.rect(SX(s) - 0.6, SY(zd), 1.2, SY(zm) - SY(zd), 'NET_CHINH')
        to.text(ten, (SX(s), SY(zm) + 14.0), 1.9, 'CENTER')
        to.text('%s m' % so(zm - zd, 1), (SX(s) + 1.3, SY(zd) + 1.0), 1.6, 'LEFT')
    # cống lấy nước (vị trí, sơ đồ)
    sc = 1030.0
    pc = [(SX(sc) - 0.7, SY(64.3)), (SX(sc) + 0.7, SY(64.3)), (SX(sc) + 0.7, SY(67.1)), (SX(sc) - 0.7, SY(67.1))]
    to.hatch(pc, 'SOLID', lop='NET_CHINH', mau_to=8)
    to.pline(pc, 'NET_CHINH', dong=True)
    to.chu_thich_mui_ten((SX(sc), SY(65.7)), (SX(990.0), SY(62.0)), 'Cống lấy nước 1,8 × 2,0 m (vị trí theo TKCS)',
                         h=1.8, rong=66)
    # nhãn công trình, tuyến quan trắc
    for ten, s in C.CD_NHAN:
        if ten.startswith('Tuyến'):
            to.line((SX(s), SY(79.6)), (SX(s), SY(80.6)), 'NUOC')
            to.text(ten.replace('Tuyến quan trắc ', ''), (SX(s), SY(80.8)), 1.9, 'CENTER', 'NUOC')
    to.line((SX(187.7), SY(79.6)), (SX(187.7), SY(80.6)), 'NUOC')
    to.text('T9', (SX(187.7), SY(80.8)), 1.9, 'CENTER', 'NUOC')
    to.text('Tuyến quan trắc thấm (giếng tự động)', (SX(680.0), SY(80.8)), 1.8, 'CENTER', 'NUOC')
    to.chu_thich_mui_ten((SX(60.7), SY(76.0)), (SX(36.0), SY(73.0)), 'Tràn sự cố', h=1.9, rong=18)
    to.chu_thich_mui_ten((SX(88.0), SY(72.0)), (SX(130.0), SY(66.0)),
                         'Màng khoan phụt L = 50 m, đỉnh +77,00 (BV-05)', h=1.8, rong=68)
    to.text('Nhà quản lý', (SX(46.4), SY(80.2)), 1.8, 'CENTER')
    # kích thước chiều dài đập
    yk = SY(ZMAX) + 4.0
    to.dim((SX(S_DAP[0]), yk), (SX(S_DAP[1]), yk), 0, 2000, text='Đập đất C23 – C2: L = 986,75 m')
    # cao trình
    to.cao_trinh(SX(1060.0), SY(D.DINH_SC), 0, h=1.8, chu='+78,85')


def noi_suy(pts, s):
    pts = sorted(pts)
    for (a, za), (b, zb) in zip(pts[:-1], pts[1:]):
        if a <= s <= b:
            return za + (zb - za) * (s - a) / (b - a) if b > a else za
    return pts[-1][1] if s > pts[-1][0] else pts[0][1]


def bang_so_lieu(to):
    y0, hd = SY(ZMIN) - 3.0, 8.5
    coc = C.CD_COC
    hang = [('Tên cọc', lambda i: coc[i][0]),
            ('Khoảng cách lẻ (m)', lambda i: so(coc[i][1] - coc[i - 1][1], 2) if i > 0 else ''),
            ('Khoảng cách cộng dồn (m)', lambda i: so(coc[i][1], 2)),
            ('Cao độ mặt đập / mặt đất hiện trạng (m)', lambda i: so(coc[i][2], 2)),
            ('Cao độ đỉnh đập TKCS (m)', lambda i: '78,75' if S_DAP[0] - 0.1 <= coc[i][1] <= S_DAP[1] + 0.1 else ''),
            ('Cao độ đỉnh đập đồ án (m)', lambda i: '78,85' if S_DAP[0] - 0.1 <= coc[i][1] <= S_DAP[1] + 0.1 else ''),
            ('Nâng đỉnh so với hiện trạng (m)',
             lambda i: so(D.DINH_SC - coc[i][2], 2) if S_DAP[0] - 0.1 <= coc[i][1] <= S_DAP[1] - 1.0 else '')]
    xl = 25.0
    xs = [SX(c[1]) for c in coc]
    for i in range(1, len(xs)):
        if xs[i] - xs[i - 1] < 2.6:
            xs[i - 1] -= 1.3
            xs[i] = xs[i - 1] + 2.6
    for k, (ten, f) in enumerate(hang):
        yt, yb = y0 - k * hd, y0 - (k + 1) * hd
        to.line((xl, yt), (SX(1100.0) + 4.0, yt), 'BANG')
        to.text(ten, (xl + 1.0, 0.5 * (yt + yb)), 1.8, 'ML', 'CHU_DAM' if k == 0 else 'CHU', 'VN_DAM' if k == 0 else 'VN')
        for i in range(len(coc)):
            v = f(i)
            if not v:
                continue
            x = xs[i]
            if k == 1:
                x = 0.5 * (SX(coc[i - 1][1]) + SX(coc[i][1]))
            to.text(v, (x, 0.5 * (yt + yb)), 1.6, 'MC', rot=90)
    yb = y0 - len(hang) * hd
    to.line((xl, yb), (SX(1100.0) + 4.0, yb), 'BANG')
    for x in (xl, SX(-3.0), SX(1100.0) + 4.0):
        to.line((x, y0), (x, yb), 'BANG')
    for _, s, _ in C.CD_COC:
        to.line((SX(s), SY(ZMIN)), (SX(s), y0), 'NET_MANH', linetype='DASHED', ltscale=0.2)


def chu_giai(to, x, y):
    to.text('KÝ HIỆU', (x, y), 3.0, 'LEFT', 'CHU_DAM', 'VN_DAM')
    y -= 6.0
    muc = [('ht', 'Mặt đập, mặt đất hiện trạng'), ('da', 'Đỉnh đập sau sửa chữa +78,85 (đồ án)'),
           ('tk', 'Đỉnh đập TKCS +78,75'), ('tcs', 'Đỉnh tường chắn sóng +79,40'),
           ('ck', 'Đáy chân khay đập (TKCS)'), ('dc', 'Ranh giới địa tầng'), ('nn', 'Mực nước ngầm'),
           ('mn', 'Mực nước hồ đặc trưng'), ('kp', 'Màng khoan phụt chống thấm (BV-05)')]
    for ma, ten in muc:
        xa, xb = x + 1.0, x + 11.0
        if ma == 'ht':
            to.line((xa, y), (xb, y), 'NET_CHINH')
        elif ma == 'da':
            to.line((xa, y), (xb, y), 'SUA_CHUA')
        elif ma == 'tk':
            to.line((xa, y), (xb, y), 'NET_MANH', linetype='DASHED', ltscale=0.3, color=3)
        elif ma == 'tcs':
            to.line((xa, y), (xb, y), 'NET_CHINH', linetype='DASHDOT', ltscale=0.4)
        elif ma == 'ck':
            to.line((xa, y), (xb, y), 'NET_CHINH', linetype='DASHED', ltscale=0.5)
        elif ma == 'dc':
            to.line((xa, y), (xb, y), 'DIA_CHAT', linetype='DASHED', ltscale=0.4)
        elif ma == 'nn':
            to.line((xa, y), (xb, y), 'NUOC', linetype='DASHDOT', ltscale=0.4)
        elif ma == 'mn':
            to.line((xa, y), (xb, y), 'NUOC', linetype='DASHED', ltscale=0.4)
        else:
            p = [(xa, y - 1.6), (xb, y - 1.6), (xb, y + 1.6), (xa, y + 1.6)]
            to.hatch(p, 'ANSI31', 0.5, lop='SUA_CHUA', mau_to=1)
            to.pline(p, 'SUA_CHUA', dong=True)
        to.text(ten, (x + 14.0, y), 2.0, 'ML')
        y -= 5.4
    y -= 2.0
    to.text('ĐỊA TẦNG (theo TKCS)', (x, y), 2.4, 'LEFT', 'CHU_DAM', 'VN_DAM')
    for ten, mo_ta in (('MĐ', 'Mặt đập: hỗn hợp dăm, sạn, cuội'), ('ĐĐ-I', 'Đất đắp: sét pha lẫn sạn sỏi, nửa cứng'),
                       ('ĐĐ-II', 'Đất đắp: sét pha, dẻo cứng'), ('4', 'Lớp 4 (nền)'),
                       ('5', 'Đá cát bột kết phong hóa hoàn toàn (sét pha lẫn dăm sạn)'),
                       ('I', 'Đá cát bột kết phong hóa mạnh, nứt nẻ mạnh'),
                       ('II', 'Đá cát bột kết phong hóa vừa, nứt nẻ trung bình')):
        y -= 5.2
        r = 1.6 if len(ten) < 3 else 2.4
        to.circle((x + 5.0, y), r, 'NET_MANH')
        to.text(ten, (x + 5.0, y), 1.6, 'MC')
        to.text(mo_ta, (x + 14.0, y), 1.9, 'ML')


def ghi_chu(to, x, y):
    gc = ('GHI CHÚ: \\P'
          '1. Cao trình ghi bằng m. Cọc, cao độ mặt đập hiện trạng, đường chân khay, địa tầng, hố khoan số hóa từ bản vẽ '
          'TKCS (cắt dọc đỉnh đập, ngang 1/500, đứng 1/200). \\P'
          '2. Đỉnh đập sau sửa chữa +78,85 (mép thượng lưu) – cao hơn TKCS 0,10 m theo điều kiện MNLKT + 0,3 m khi có '
          'tường chắn sóng (TCVN 8216:2018, điều 10.1.5). Mặt đập hiện trạng +78,31 ÷ +78,50: phải nâng 0,35 ÷ 0,54 m. \\P'
          '3. Trình tự: bóc lớp mặt đá dăm láng nhựa hư hỏng → đắp bù đất đầm K ≥ 0,97 đến +78,60 → đá dăm đệm 5 cm, '
          'nilon → BT M250 dày 20 cm, dốc ngang 2% (BV-04). Khoan phụt mang tràn sự cố thi công trước khi đắp bù. \\P'
          '4. Lớp 5, lớp I tại mang trái tràn sự cố là đường dẫn nước chính (K = 6,54E-05 và 4,35E-04 cm/s – CG7).')
    to.mtext(gc, (x, y), 2.0, rong=170)


def bang_cao_trinh(to):
    cot = [62, 34, 34, 34]
    hang = [['Đại lượng', 'Z1: MNDBT', 'Z2: MNLTK', 'Z3: MNLKT'],
            ['Mực nước tính toán (m)', '+74,70', '+77,63', '+78,54'],
            ['Gió V (m/s), hướng Bắc', '13,8 (P = 4%)', '7,6 (P = 50%)', '–'],
            ['Đà gió D (m)', '2.970', '3.025', '–'],
            ['Sóng leo hsl1% (m)', '1,153', '0,398', '–'],
            ['Độ dềnh Δh (m)', '0,006', '0,001', '–'],
            ['Độ vượt cao an toàn a (m)', '0,70', '0,50', '0,20'],
            ['Cao trình đỉnh yêu cầu (m)', '+76,56', '+78,53', '+78,74']]
    to.bang(30.0, 118.0, cot, hang, h=1.9, cao_dong=4.8, tieu_de='CAO TRÌNH ĐỈNH ĐẬP YÊU CẦU (Bảng 2.4)')
    cot2 = [78, 18, 18, 18, 22, 26]
    hang2 = [['Trường hợp', 'Z1', 'Z2', 'Z3', 'Zđ = max', 'MNLKT + 0,3'],
             ['Hồ sơ BCNCKT (MNLTK +77,43; MNLKT +78,25)', '76,56', '78,33', '78,45', '78,45', '78,55'],
             ['Đồ án – lũ BCNCKT (+77,63; +78,54)', '76,56', '78,53', '78,74', '78,74', '78,84'],
             ['Đồ án – lũ QTVH (+77,50; +78,44)', '76,56', '78,40', '78,64', '78,64', '78,74']]
    y = to.bang(210.0, 118.0, cot2, hang2, h=1.9, cao_dong=4.8, tieu_de='TỔNG HỢP CAO TRÌNH ĐỈNH ĐẬP YÊU CẦU (Bảng 2.5)')
    to.text('Chọn đỉnh đập sau sửa chữa +78,85 m (mép thượng lưu) ≥ max(78,74; 78,84). Đỉnh hiện trạng +78,31 ÷ +78,50 '
            '< MNLKT +78,54 tại nhiều đoạn.', (210.0, y - 4.0), 1.9)
    to.text('TCVN 8216:2018: Bảng 3 (a = 0,7 / 0,5 / 0,2 m – cấp II), Bảng 4 (gió P = 4% / 50%), điều 10.1.5 '
            '(có tường chắn sóng: đỉnh đập ≥ MNLKT + 0,3 m).', (210.0, y - 8.0), 1.9)


def ve(thu_muc):
    to = To(2, 'CẮT DỌC TIM ĐẬP\nHIỆN TRẠNG VÀ SAU SỬA CHỮA', 'Như hình vẽ', 'HCG-BV02')
    cat_doc(to)
    bang_so_lieu(to)
    chu_giai(to, 662.0, 560.0)
    ghi_chu(to, 655.0, 345.0)
    bang_cao_trinh(to)
    to.luu(thu_muc + '/BV02_CatDocDap.dxf', thu_muc + '/BV02_CatDocDap.pdf', thu_muc + '/BV02_CatDocDap.png')
    return to


if __name__ == '__main__':
    import sys
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
