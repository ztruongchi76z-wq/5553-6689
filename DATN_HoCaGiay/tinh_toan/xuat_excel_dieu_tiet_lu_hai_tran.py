# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL: ĐIỀU TIẾT LŨ HỒ CÀ GIÂY VỚI HAI TRÀN CÓ CAO TRÌNH NGƯỠNG KHÁC NHAU – PHƯƠNG PHÁP PÔ-TA-PỐP
Giả thiết (theo yêu cầu SV): trước lũ hồ ở MNDBT; cửa van tràn sự cố mở hoàn toàn suốt trận lũ (chảy tự do như
đập tràn); trường hợp kẹt van: chỉ tràn chính làm việc. Hai tràn cùng một mực nước hồ: q(Z) = q_tc(Z) + q_sc(Z),
điều tiết MỘT lần với q tổng. Bước Δt = 1 giờ; f = V/Δt + q/2; f2 = f1 + Q̄ − q1.

  python3 xuat_excel_dieu_tiet_lu_hai_tran.py [đường dẫn .xlsx ra]
Mặc định: ../TAI_VE/DieuTietLu_HaiTran_HoCaGiay.xlsx
"""
import math
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import so_lieu as S

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'DieuTietLu_HaiTran_HoCaGiay.xlsx')
T_MAX = 48                          # số giờ tính
Z_DAU, Z_CUOI, BUOC = 71.0, 81.0, 0.01
N_QH = int(round((Z_CUOI - Z_DAU) / BUOC)) + 1

FONT = 'Times New Roman'
f_thuong = Font(name=FONT, size=11)
f_dam = Font(name=FONT, size=11, bold=True)
f_tieu_de = Font(name=FONT, size=14, bold=True, color='12324A')
f_muc = Font(name=FONT, size=12, bold=True, color='A65A10')
f_nhap = Font(name=FONT, size=11, color='0000FF')
f_nhap_dam = Font(name=FONT, size=11, color='0000FF', bold=True)
f_trang = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_ghi_chu = Font(name=FONT, size=10, italic=True, color='52514E')
nen_nhap = PatternFill('solid', fgColor='FFF2CC')
nen_tieu_de = PatternFill('solid', fgColor='12324A')
nen_kq = PatternFill('solid', fgColor='FBF3E8')
nen_tc = PatternFill('solid', fgColor='EEF3FA')
nen_sc = PatternFill('solid', fgColor='F4F1FA')
nen_dinh = PatternFill('solid', fgColor='F8D7D7')
vien = Border(*(Side(style='thin', color='B8BEC6'),) * 4)
giua = Alignment(horizontal='center', vertical='center', wrap_text=True)
trai = Alignment(horizontal='left', vertical='top', wrap_text=True)
cg = Alignment(horizontal='center')

wb = Workbook()


def ghi(ws, o, gt, font=f_thuong, fmt=None, fill=None, can=None, border=True):
    c = ws[o]
    c.value = gt
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if can:
        c.alignment = can
    if border:
        c.border = vien
    return c


def ten(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def tieu_de_cot(ws, hang, cot0, nhan, cao=None):
    for k, t in enumerate(nhan):
        ghi(ws, '%s%d' % (get_column_letter(cot0 + k), hang), t, f_trang, fill=nen_tieu_de, can=giua)
    if cao:
        ws.row_dimensions[hang].height = cao


def in_an(ws, ngang=True):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if ngang else 'portrait'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5


def q_tran(z, p):
    """Lưu lượng qua tràn (m³/s) theo TCVN 9147:2012, p = tiền tố tên thông số ('TC' hoặc 'SC')."""
    H = 'MAX({z}-{p}_Z,0)'.format(z=z, p=p)
    return ('{p}_s*(1-0.2*({p}_xmb+({p}_n-1)*{p}_xmt)*{H}/({p}_n*{p}_b))*{p}_m*{p}_n*{p}_b*SQRT(2*g_)*{H}^1.5'
            .format(p=p, H=H))


# ================================================================== HƯỚNG DẪN
hd = wb.active
hd.title = 'HuongDan'
hd.column_dimensions['A'].width = 3
hd.column_dimensions['B'].width = 125
hd.sheet_view.showGridLines = False
ND = [
    ('T', 'ĐIỀU TIẾT LŨ HỒ CÀ GIÂY VỚI HAI TRÀN CÓ NGƯỠNG KHÁC NHAU – PHƯƠNG PHÁP PÔ-TA-PỐP'),
    ('g', 'Mọi ô kết quả là công thức. Đổi thông số tràn, mực nước trước lũ ở sheet SoLieu; đổi tần suất, trường hợp ở ô vàng đầu mỗi sheet điều tiết.'),
    ('M', '1. Cộng dồn hay tính song song?'),
    ('-', 'Hai tràn cùng lấy nước từ MỘT hồ nên ở mỗi thời điểm chỉ có MỘT mực nước Z. Tràn nào có ngưỡng thấp hơn Z thì tràn đó xả: tràn sự cố (ngưỡng +71,00 m) xả khi Z > +71,00; tràn chính (ngưỡng +74,70 m) xả khi Z > +74,70.'),
    ('-', 'Lưu lượng xả của hồ là TỔNG hai tràn tại CÙNG mực nước: q(Z) = q tràn chính(Z) + q tràn sự cố(Z). Vậy “cộng” là đúng ở bước lập quan hệ q ~ Z (sheet QuanHe, cột q hai tràn), sau đó điều tiết MỘT lần với q tổng – đó chính là tính song song (hai tràn cùng làm việc).'),
    ('-', 'KHÔNG được điều tiết riêng từng tràn rồi cộng kết quả: mỗi lần tính riêng đã cho TOÀN BỘ trận lũ đi qua một tràn, nên mực nước cao hơn thực tế; cộng hai lưu lượng xả của hai lần tính riêng thì tính trùng lượng nước lũ.'),
    ('-', 'Kết quả “tính độc lập” trong bảng của bạn chỉ có ý nghĩa là trường hợp MỘT tràn không làm việc: “chỉ tràn chính” chính là trường hợp cửa van tràn sự cố bị kẹt (P = 1%: +78,21 m ở file này, bạn tính +78,28 m); “chỉ tràn sự cố” là trường hợp cả 9 cửa tràn chính không mở – không phải trường hợp thiết kế.'),
    ('M', '2. Giả thiết tính toán'),
    ('-', 'Khi lũ về, mực nước hồ ở MNDBT = +74,70 m (ô SoLieu!B5). Tràn chính 9 khoang × 2,4 m, cửa van mở hết, chảy tự do.'),
    ('-', 'Tràn sự cố có cửa van phẳng: giả thiết cửa van MỞ HOÀN TOÀN suốt trận lũ (kéo lên khỏi mặt nước) nên tràn làm việc như đập tràn chảy tự do, tính bằng cùng công thức với tràn chính.'),
    ('-', 'Trường hợp kẹt van: cửa tràn sự cố đóng, chỉ tràn chính xả (chọn “Kẹt van” ở ô F3 của sheet điều tiết).'),
    ('-', 'Bỏ qua lưu tốc tới gần (H₀ ≈ H), chảy không ngập (σn = 1). Dung tích theo bình đồ 2022 (thuyết minh mục 1.2.3).'),
    ('-', 'Lưu ý: thuyết minh dùng quy tắc vận hành QTVH (đầu lũ dùng cửa van tràn sự cố giữ MNDBT, khi xả không kịp mới mở hết) nên MNLTK, MNLKT cao hơn file này khoảng 2 cm (+77,63 / +78,54 so với +77,61 / +78,51 m). Khi mở hết cửa ngay từ đầu, hồ rút xuống khoảng +74,64 m trước khi lũ lớn về.'),
    ('M', '3. Công thức lưu lượng qua tràn (TCVN 9147:2012)'),
    ('c', 'q = σn·ε·m·n·b·√(2g)·H^1,5        ε = 1 − 0,2·[ξmb + (n − 1)·ξmt]·H / (n·b)'),
    ('-', 'H = Z − Z ngưỡng (lấy 0 khi Z thấp hơn ngưỡng); n – số khoang; b – bề rộng một khoang; m – hệ số lưu lượng; ξmb, ξmt – hệ số hình dạng mố bên, mố trụ. Thông số từng tràn ở sheet SoLieu (giống chương trình của đồ án).'),
    ('M', '4. Phương pháp Pô-ta-pốp'),
    ('c', '(Q₁ + Q₂)/2·Δt − (q₁ + q₂)/2·Δt = V₂ − V₁    ⇒    V₂/Δt + q₂/2 = V₁/Δt + q₁/2 + Q̄ − q₁    ⇒    f₂ = f₁ + Q̄ − q₁'),
    ('-', 'f = V/Δt + q/2 (đơn vị m³/s) chỉ phụ thuộc mực nước Z, nên lập sẵn bảng Z ~ V ~ q ~ f (sheet QuanHe, bước 0,01 m; một cột f cho trường hợp hai tràn, một cột cho trường hợp kẹt van).'),
    ('-', 'Mỗi giờ: (1) Q̄ = (Q₁ + Q₂)/2; (2) f₂ = f₁ + Q̄ − q₁; (3) tra ngược bảng f ~ Z được Z₂ (nội suy tuyến tính); (4) tính q₂ = q tràn chính(Z₂) + q tràn sự cố(Z₂) và V₂; (5) sang giờ sau với f₁ = f₂, q₁ = q₂.'),
    ('-', 'Giờ 0: Z = MNDBT, q = q(MNDBT). Trường hợp hai tràn: tràn sự cố đã xả khoảng 64 m³/s (H = 3,70 m), tràn chính chưa xả. Trường hợp kẹt van: q = 0.'),
    ('-', 'Kết quả: Zmax = MNLTK (P = 1%) hoặc MNLKT (P = 0,2%), lưu lượng xả lớn nhất qmax, phân chia cho từng tràn. Kiểm tra: W đến − W xả = ΔV (sai số ≈ 0).'),
    ('M', '5. Trình tự trong file'),
    ('-', 'SoLieu: thông số hai tràn, mực nước trước lũ, đường quá trình lũ P = 1% và 0,2% (BCNCKT – mô hình Gurdich 36 giờ), bảng Z ~ F ~ W 2022.'),
    ('-', 'QuanHe: bảng Z ~ F ~ V ~ q tràn chính ~ q tràn sự cố ~ q tổng ~ f (1.001 dòng, +71,00 ÷ +81,00 m).'),
    ('-', 'P1_HaiTran, P1_KetVan, P02_HaiTran, P02_KetVan: bốn bảng điều tiết giống nhau, chỉ khác ô tần suất (C3) và trường hợp (F3).'),
    ('-', 'KetQua: tổng hợp 4 trường hợp, so sánh với thuyết minh và với bảng tính độc lập của bạn, biểu đồ lưu lượng và mực nước.'),
    ('M', '6. Chú giải màu ô'),
    ('-', 'Nền vàng chữ xanh: ô được sửa. Chữ xanh: số liệu gốc hồ sơ. Chữ đen: công thức. Nền đỏ nhạt: dòng có mực nước lớn nhất.'),
]
for h, (loai, txt) in enumerate(ND, 1):
    o = hd['B%d' % h]
    o.value = ('•  ' + txt) if loai == '-' else txt
    o.alignment = trai
    if loai == 'T':
        o.font, hd.row_dimensions[h].height = f_tieu_de, 24
    elif loai == 'M':
        o.font, hd.row_dimensions[h].height = f_muc, 22
    elif loai == 'c':
        o.font = Font(name=FONT, size=12, bold=True, color='12324A')
        o.alignment = Alignment(horizontal='center', vertical='center')
        hd.row_dimensions[h].height = 24
    elif loai == 'g':
        o.font, hd.row_dimensions[h].height = f_ghi_chu, 30
    else:
        o.font = f_thuong
        hd.row_dimensions[h].height = 15.5 * max(1, math.ceil(len(txt) / 145))
in_an(hd)

# ================================================================== SỐ LIỆU
sl = wb.create_sheet('SoLieu')
for c, w in zip('ABCDEFGH', (44, 11, 10, 14, 14, 14, 14, 30)):
    sl.column_dimensions[c].width = w
ghi(sl, 'A1', 'SỐ LIỆU ĐẦU VÀO – ĐIỀU TIẾT LŨ HAI TRÀN', f_tieu_de, border=False)
ghi(sl, 'A2', 'Ô nền vàng chữ xanh: được sửa. Chữ xanh: số liệu gốc hồ sơ. Chữ đen: công thức.', f_ghi_chu, border=False)


def khoi_thong_so(h0, tieu_de, dong):
    ghi(sl, 'A%d' % h0, tieu_de, f_muc, border=False)
    tieu_de_cot(sl, h0 + 1, 1, ['Thông số', 'Giá trị', 'Đơn vị'])
    ghi(sl, 'D%d' % (h0 + 1), 'Ghi chú, nguồn', f_trang, fill=nen_tieu_de, can=giua)
    sl.merge_cells('D%d:H%d' % (h0 + 1, h0 + 1))
    for k, (nhan, gt, dv, ten_o, gc, nhap) in enumerate(dong):
        r = h0 + 2 + k
        ghi(sl, 'A%d' % r, nhan, f_dam)
        ghi(sl, 'B%d' % r, gt, f_nhap_dam if nhap else f_dam, '0.00' if isinstance(gt, float) else None,
            nen_nhap if nhap else None, cg)
        ghi(sl, 'C%d' % r, dv, can=cg)
        ghi(sl, 'D%d' % r, gc, f_ghi_chu)
        sl.merge_cells('D%d:H%d' % (r, r))
        if ten_o:
            ten(ten_o, 'SoLieu!$B$%d' % r)
    return h0 + 2 + len(dong)


h = khoi_thong_so(4, 'A. THÔNG SỐ CHUNG', [
    ('Mực nước hồ khi lũ đến (MNDBT)', 74.70, 'm', 'Z0', 'Giả thiết bất lợi: lũ đến khi hồ đầy ở MNDBT', True),
    ('Bước thời gian Δt', 1, 'giờ', 'dt_h', 'Bằng bước của đường quá trình lũ (cố định 1 giờ)', False),
    ('Δt tính bằng giây', '=dt_h*3600', 's', 'dt_s', '', False),
    ('Gia tốc trọng trường g', 9.81, 'm/s²', 'g_', '', False),
])
sl['B7'].number_format = '0'
sl['B8'].number_format = '0'
TC = S.TRAN_CHINH
h = khoi_thong_so(h + 1, 'B. TRÀN CHÍNH (9 khoang có cửa van, mở hết – chảy tự do)', [
    ('Cao trình ngưỡng', 74.70, 'm', 'TC_Z', 'QTVH 2026; thuyết minh mục 1.6.2', True),
    ('Số khoang n', 9, 'khoang', 'TC_n', 'Đơn vị quản lý chia 9 khoang × 2,4 m (QTVH mục 6)', True),
    ('Bề rộng một khoang b', 2.40, 'm', 'TC_b', 'Σb = 21,6 m', True),
    ('Hệ số lưu lượng m', 0.45, '', 'TC_m', 'Đồ án lấy 0,45 (thuyết minh 2.3.2); BCNCKT 0,467; QTVH 0,40', True),
    ('Hệ số hình dạng mố bên ξmb', 0.70, '', 'TC_xmb', 'Mố bên lượn tròn', True),
    ('Hệ số hình dạng mố trụ ξmt', 0.45, '', 'TC_xmt', 'Mố trụ đầu tròn', True),
    ('Hệ số chảy ngập σn', 1.0, '', 'TC_s', 'Chảy không ngập', True),
])
sl['B%d' % (h - 6)].number_format = '0'
h = khoi_thong_so(h + 1, 'C. TRÀN SỰ CỐ (cửa van phẳng – giả thiết mở hoàn toàn)', [
    ('Cao trình ngưỡng', 71.00, 'm', 'SC_Z', 'KTTV Bảng 1-2', True),
    ('Số khoang n', 1, 'khoang', 'SC_n', '', True),
    ('Bề rộng khoang b', 5.00, 'm', 'SC_b', 'KTTV Bảng 1-2', True),
    ('Hệ số lưu lượng m', 0.45, '', 'SC_m', 'Như chương trình của đồ án', True),
    ('Hệ số hình dạng mố bên ξmb', 0.70, '', 'SC_xmb', '', True),
    ('Hệ số hình dạng mố trụ ξmt', 0.45, '', 'SC_xmt', 'Không dùng khi n = 1', True),
    ('Hệ số chảy ngập σn', 1.0, '', 'SC_s', 'Chảy không ngập', True),
])
sl['B%d' % (h - 6)].number_format = '0'
# D. đường quá trình lũ
HL = h + 1
ghi(sl, 'A%d' % HL, 'D. ĐƯỜNG QUÁ TRÌNH LŨ ĐẾN (m³/s) – BCNCKT 2022, mô hình Gurdich 36 giờ (KTTV Bảng 4-5)', f_muc, border=False)
ghi(sl, 'A%d' % (HL + 1), 'Tần suất P (%) →', f_dam, can=Alignment(horizontal='right'))
ghi(sl, 'B%d' % (HL + 1), 1, f_nhap_dam, '0.0', can=cg)
ghi(sl, 'C%d' % (HL + 1), 0.2, f_nhap_dam, '0.0', can=cg)
tieu_de_cot(sl, HL + 2, 1, ['t (giờ)', 'P = 1%', 'P = 0,2%'])
for t in range(0, 37):
    r = HL + 3 + t
    ghi(sl, 'A%d' % r, t, f_nhap, '0', can=cg)
    for k, p in enumerate(('1%', '0.2%')):
        ghi(sl, '%s%d' % ('BC'[k], r), S.Q_KTTV[p][max(t, 1) - 1], f_nhap, '0', can=cg)
RL0, RL1 = HL + 3, HL + 3 + 36
ten('Lu_t', 'SoLieu!$A$%d:$A$%d' % (RL0, RL1))
ten('Lu_Q', 'SoLieu!$B$%d:$C$%d' % (RL0, RL1))
ten('Lu_P', 'SoLieu!$B$%d:$C$%d' % (HL + 1, HL + 1))
ghi(sl, 'D%d' % (HL + 3), 'Giờ 0 lấy bằng giờ 1 (như chương trình của đồ án); sau giờ 36 lưu lượng đến = 0.', f_ghi_chu, border=False)
ghi(sl, 'D%d' % (HL + 5), 'Tổng lượng lũ tính từ bảng (10⁶ m³):', f_dam, border=False)
for k, cl in enumerate('BC'):
    ghi(sl, '%s%d' % ('EF'[k], HL + 5),
        '=SUMPRODUCT(({c}{a}:{c}{b}+{c}{a1}:{c}{b1})/2)*dt_s/1000000+{c}{b1}/2*dt_s/1000000'.format(
            c=cl, a=RL0, b=RL1 - 1, a1=RL0 + 1, b1=RL1), f_dam, '0.00', nen_kq, cg)
ghi(sl, 'D%d' % (HL + 6), 'Tổng lượng lũ ghi trong KTTV Bảng 4-4 (10⁶ m³):', f_dam, border=False)
ghi(sl, 'E%d' % (HL + 6), S.W_KTTV_GHI['1%'], f_nhap, '0.00', can=cg)
ghi(sl, 'F%d' % (HL + 6), S.W_KTTV_GHI['0.2%'], f_nhap, '0.00', can=cg)
ghi(sl, 'D%d' % (HL + 7), 'Thuyết minh mục 1.4.3: dùng trực tiếp tung độ đường quá trình (tổng lượng 36,27 / 47,69), '
    'không dùng tổng lượng ghi – thiên an toàn.', f_ghi_chu, border=False)
ghi(sl, 'E%d' % (HL + 4), 'P = 1%', f_trang, fill=nen_tieu_de, can=giua)
ghi(sl, 'F%d' % (HL + 4), 'P = 0,2%', f_trang, fill=nen_tieu_de, can=giua)
# E. Z ~ F ~ W
HZ = RL1 + 3
ghi(sl, 'A%d' % HZ, 'E. QUAN HỆ Z ~ F ~ W – BÌNH ĐỒ 2022 (QTVH Bảng 1-2)', f_muc, border=False)
tieu_de_cot(sl, HZ + 1, 1, ['Z (m)', 'F (km²)', 'W (10⁶ m³)'])
ghi(sl, 'D%d' % (HZ + 1), 'Ghi chú', f_trang, fill=nen_tieu_de, can=giua)
for k, (z, f) in enumerate(zip(S.Z_TAB, S.F_TAB)):
    r = HZ + 2 + k
    ghi(sl, 'A%d' % r, float(z), f_nhap, '0.00', can=cg)
    ghi(sl, 'B%d' % r, float(f), f_nhap, '0.00', can=cg)
    ghi(sl, 'C%d' % r, 0 if k == 0 else '=C{p}+(A{h}-A{p})*(B{p}+B{h})/2'.format(p=r - 1, h=r), fmt='0.000', can=cg)
RZ1 = HZ + 1 + len(S.Z_TAB)
ghi(sl, 'D%d' % (HZ + 2), 'W cộng dồn từ F (F biến đổi tuyến tính giữa hai cao trình): dùng cho mọi tính toán – '
    'V(+74,70) = 36,41·10⁶ m³ như thuyết minh. Điều tiết lũ chỉ dùng chênh lệch dung tích nên chỉ phần trên MNDBT quyết định kết quả.',
    f_ghi_chu, border=False)
ten('Bang_Z', 'SoLieu!$A$%d:$A$%d' % (HZ + 2, RZ1))
ten('Bang_F', 'SoLieu!$B$%d:$B$%d' % (HZ + 2, RZ1))
ten('Bang_W', 'SoLieu!$C$%d:$C$%d' % (HZ + 2, RZ1))
sl.freeze_panes = 'A4'
in_an(sl)

# ================================================================== QUAN HỆ Z ~ V ~ q ~ f
qh = wb.create_sheet('QuanHe')
RONG = (8, 8, 10, 7, 7, 9, 7, 7, 9, 10, 10, 11, 11)
for k, w in enumerate(RONG):
    qh.column_dimensions[get_column_letter(k + 1)].width = w
ghi(qh, 'A1', 'BẢNG QUAN HỆ Z ~ V ~ q ~ f (bước 0,01 m) – dùng để tra trong các sheet điều tiết', f_tieu_de, border=False)
ghi(qh, 'A2', 'f = V·10⁶/Δt + q/2 (m³/s). Cột q hai tràn = q tràn chính + q tràn sự cố tại CÙNG mực nước Z.', f_ghi_chu, border=False)
tieu_de_cot(qh, 4, 1, ['Z (m)', 'F (km²)', 'V (10⁶ m³)', 'H tràn chính (m)', 'ε tràn chính', 'q tràn chính (m³/s)',
                       'H tràn sự cố (m)', 'ε tràn sự cố', 'q tràn sự cố (m³/s)', 'q hai tràn (m³/s)',
                       'q kẹt van (m³/s)', 'f hai tràn (m³/s)', 'f kẹt van (m³/s)'], cao=45)
for i in range(N_QH):
    r = 5 + i
    k = 'MATCH(A{r},Bang_Z,1)'.format(r=r)
    ct = {
        'A': Z_DAU if i == 0 else '=ROUND(A%d+%s,2)' % (r - 1, BUOC),
        'B': '=INDEX(Bang_F,{k})+(A{r}-INDEX(Bang_Z,{k}))*(INDEX(Bang_F,{k}+1)-INDEX(Bang_F,{k}))'.format(k=k, r=r),
        'C': '=INDEX(Bang_W,{k})+(A{r}-INDEX(Bang_Z,{k}))*(INDEX(Bang_F,{k})+B{r})/2'.format(k=k, r=r),
        'D': '=MAX(A%d-TC_Z,0)' % r,
        'E': '=1-0.2*(TC_xmb+(TC_n-1)*TC_xmt)*D%d/(TC_n*TC_b)' % r,
        'F': '=TC_s*E{r}*TC_m*TC_n*TC_b*SQRT(2*g_)*D{r}^1.5'.format(r=r),
        'G': '=MAX(A%d-SC_Z,0)' % r,
        'H': '=1-0.2*(SC_xmb+(SC_n-1)*SC_xmt)*G%d/(SC_n*SC_b)' % r,
        'I': '=SC_s*H{r}*SC_m*SC_n*SC_b*SQRT(2*g_)*G{r}^1.5'.format(r=r),
        'J': '=F%d+I%d' % (r, r),
        'K': '=F%d' % r,
        'L': '=C{r}*1000000/dt_s+J{r}/2'.format(r=r),
        'M': '=C{r}*1000000/dt_s+K{r}/2'.format(r=r),
    }
    fmts = {'A': '0.00', 'B': '0.000', 'C': '0.000', 'D': '0.00', 'E': '0.000', 'F': '0.0', 'G': '0.00', 'H': '0.000',
            'I': '0.0', 'J': '0.0', 'K': '0.0', 'L': '0.0', 'M': '0.0'}
    for col, f in ct.items():
        c = qh['%s%d' % (col, r)]
        c.value = f
        c.font = f_thuong
        c.number_format = fmts[col]
        c.alignment = cg
        if col in 'DEF':
            c.fill = nen_tc
        elif col in 'GHI':
            c.fill = nen_sc
R_QH = 4 + N_QH
ten('QH_Z', 'QuanHe!$A$5:$A$%d' % R_QH)
ten('QH_V', 'QuanHe!$C$5:$C$%d' % R_QH)
ten('QH_f', 'QuanHe!$L$5:$M$%d' % R_QH)
qh.freeze_panes = 'B5'
in_an(qh, ngang=False)

# ================================================================== CÁC SHEET ĐIỀU TIẾT
R0 = 16                          # dòng giờ 0
R1 = R0 + T_MAX
DT_HDR = ['t (giờ)', 'Q đến (m³/s)', 'Q̄ = (Q₁+Q₂)/2', 'f₁ (giờ trước)', 'q₁ (giờ trước)', 'f₂ = f₁ + Q̄ − q₁',
          'Z₂ (m) tra ngược f ~ Z', 'q tràn chính (m³/s)', 'q tràn sự cố (m³/s)', 'q xả = tổng (m³/s)',
          'V₂ (10⁶ m³)', 'f₂ = V₂/Δt + q₂/2']


def sheet_dieu_tiet(ten_ws, P, th):
    ws = wb.create_sheet(ten_ws)
    for k, w in enumerate((8, 10, 11, 12, 11, 12, 12, 12, 12, 12, 11, 12)):
        ws.column_dimensions[get_column_letter(k + 1)].width = w
    ghi(ws, 'A1', 'ĐIỀU TIẾT LŨ – PHƯƠNG PHÁP PÔ-TA-PỐP (Δt = 1 giờ)', f_tieu_de, border=False)
    ghi(ws, 'A3', 'Tần suất P (%)', f_dam)
    ws.merge_cells('A3:B3')
    ghi(ws, 'C3', P, f_nhap_dam, '0.0', nen_nhap, cg)
    ghi(ws, 'D3', 'Trường hợp', f_dam)
    ws.merge_cells('D3:E3')
    ghi(ws, 'F3', th, f_nhap_dam, None, nen_nhap, cg)
    ws.merge_cells('F3:G3')
    ghi(ws, 'A4', 'Cột lũ', f_ghi_chu)
    ws.merge_cells('A4:B4')
    ghi(ws, 'C4', '=MATCH(C3,Lu_P,0)', fmt='0', can=cg)
    ghi(ws, 'D4', 'Cột f', f_ghi_chu)
    ws.merge_cells('D4:E4')
    ghi(ws, 'F4', '=IF(F3="Kẹt van",2,1)', fmt='0', can=cg)
    ghi(ws, 'H3', '← sửa ô vàng: P = 1 hoặc 0,2; trường hợp “Hai tràn” hoặc “Kẹt van”', f_ghi_chu, border=False)
    dv1 = DataValidation(type='list', formula1='"1,0.2"', allow_blank=False)
    dv2 = DataValidation(type='list', formula1='"Hai tràn,Kẹt van"', allow_blank=False)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add('C3')
    dv2.add('F3')
    # tổng hợp
    rg = lambda c: '%s%d:%s%d' % (c, R0, c, R1)
    vt = 'MATCH(MAX({g}),{g},0)'.format(g=rg('G'))
    TH = [
        (6, 'Mực nước lớn nhất Zmax (m)', '=MAX(%s)' % rg('G'), '0.00'),
        (7, 'Thời điểm xuất hiện (giờ)', '=INDEX(%s,%s)' % (rg('A'), vt), '0'),
        (8, 'Lưu lượng xả lớn nhất qmax (m³/s)', '=MAX(%s)' % rg('J'), '0.0'),
        (9, 'Lúc Zmax: q tràn chính / q tràn sự cố (m³/s)', '=INDEX(%s,%s)' % (rg('H'), vt), '0.0'),
        (10, 'Cột nước trên ngưỡng lúc Zmax: tràn chính / tràn sự cố (m)', '=F6-TC_Z', '0.00'),
        (11, 'Dung tích phòng lũ V(Zmax) − V(MNDBT) (10⁶ m³)', '=INDEX(%s,%s)-K%d' % (rg('K'), vt, R0), '0.000'),
        (12, 'Cân bằng nước: W đến − W xả − ΔV (10⁶ m³, phải ≈ 0)',
         '=SUMPRODUCT((B{a}:B{b}+B{a1}:B{b1})/2)*dt_s/1000000-SUMPRODUCT((J{a}:J{b}+J{a1}:J{b1})/2)*dt_s/1000000-(K{b1}-K{a})'
         .format(a=R0, b=R1 - 1, a1=R0 + 1, b1=R1), '0.0000'),
    ]
    for r, nhan, ct, fmt in TH:
        ghi(ws, 'A%d' % r, nhan, f_dam)
        ws.merge_cells('A%d:E%d' % (r, r))
        ghi(ws, 'F%d' % r, ct, f_dam, fmt, nen_kq, cg)
    ghi(ws, 'G9', '=INDEX(%s,%s)' % (rg('I'), vt), f_dam, '0.0', nen_kq, cg)
    ghi(ws, 'G10', '=F6-SC_Z', f_dam, '0.00', nen_kq, cg)
    tieu_de_cot(ws, 15, 1, DT_HDR, cao=45)
    QH_COT = 'INDEX(QH_f,0,$F$4)'
    for t in range(0, T_MAX + 1):
        r = R0 + t
        k = 'MATCH(F{r},{c},1)'.format(r=r, c=QH_COT)
        kz = 'MATCH(G{r},QH_Z,1)'.format(r=r)
        ct = {
            'A': t,
            'B': '=IFERROR(INDEX(Lu_Q,MATCH(A{r},Lu_t,0),$C$4),0)'.format(r=r),
            'H': '=' + q_tran('G%d' % r, 'TC'),
            'I': '=IF($F$4=2,0,%s)' % q_tran('G%d' % r, 'SC'),
            'J': '=H%d+I%d' % (r, r),
            'K': ('=INDEX(QH_V,{k})+(G{r}-INDEX(QH_Z,{k}))*(INDEX(QH_V,{k}+1)-INDEX(QH_V,{k}))'
                  '/(INDEX(QH_Z,{k}+1)-INDEX(QH_Z,{k}))').format(k=kz, r=r),
            'L': '=K{r}*1000000/dt_s+J{r}/2'.format(r=r),
        }
        if t == 0:
            ct.update({'C': '', 'D': '', 'E': '', 'F': '', 'G': '=Z0'})
        else:
            ct.update({
                'C': '=(B%d+B%d)/2' % (r - 1, r),
                'D': '=L%d' % (r - 1),
                'E': '=J%d' % (r - 1),
                'F': '=D%d+C%d-E%d' % (r, r, r),
                'G': ('=INDEX(QH_Z,{k})+(F{r}-INDEX(QH_f,{k},$F$4))*(INDEX(QH_Z,{k}+1)-INDEX(QH_Z,{k}))'
                      '/(INDEX(QH_f,{k}+1,$F$4)-INDEX(QH_f,{k},$F$4))').format(k=k, r=r),
            })
        fm = {'A': '0', 'B': '0.0', 'C': '0.0', 'D': '0.0', 'E': '0.0', 'F': '0.0', 'G': '0.000', 'H': '0.0', 'I': '0.0',
              'J': '0.0', 'K': '0.000', 'L': '0.0'}
        for col, f in ct.items():
            c = ws['%s%d' % (col, r)]
            c.value = f
            c.font = f_dam if col in 'AG' else f_thuong
            c.number_format = fm[col]
            c.alignment = cg
            c.border = vien
            if col == 'H':
                c.fill = nen_tc
            elif col == 'I':
                c.fill = nen_sc
    ws.conditional_formatting.add('A%d:L%d' % (R0, R1), FormulaRule(formula=['$G%d=MAX($G$%d:$G$%d)' % (R0, R0, R1)],
                                                                   fill=nen_dinh))
    ghi(ws, 'H5', 'Giờ 0: Z = MNDBT, q = q(MNDBT) theo trường hợp chọn. Từ giờ 1: f₂ = f₁ + Q̄ − q₁ → tra Z₂ → q₂, V₂.',
        f_ghi_chu, border=False)
    ws.freeze_panes = 'B16'
    in_an(ws)
    return ws


CA = [('P1_HaiTran', 1, 'Hai tràn'), ('P1_KetVan', 1, 'Kẹt van'), ('P02_HaiTran', 0.2, 'Hai tràn'),
      ('P02_KetVan', 0.2, 'Kẹt van')]
for ten_ws, P, th in CA:
    sheet_dieu_tiet(ten_ws, P, th)

# ================================================================== KẾT QUẢ
kq = wb.create_sheet('KetQua')
for k, w in enumerate((34, 14, 11, 9, 11, 12, 12, 12, 14, 30)):
    kq.column_dimensions[get_column_letter(k + 1)].width = w
ghi(kq, 'A1', 'KẾT QUẢ ĐIỀU TIẾT LŨ – HỒ CÀ GIÂY (hai tràn, cửa van tràn sự cố mở hoàn toàn)', f_tieu_de, border=False)
tieu_de_cot(kq, 3, 1, ['Sheet', 'P (%) – trường hợp', 'Zmax (m)', 't (giờ)', 'qmax (m³/s)', 'q tràn chính lúc Zmax',
                       'q tràn sự cố lúc Zmax', 'Dung tích phòng lũ (10⁶ m³)', 'Thuyết minh (quy tắc QTVH)', 'Ý nghĩa'],
            cao=45)
Y_NGHIA = {'P1_HaiTran': 'MNLTK – trường hợp thiết kế', 'P1_KetVan': 'Kẹt van tràn sự cố – tổ hợp đặc biệt',
           'P02_HaiTran': 'MNLKT – trường hợp kiểm tra', 'P02_KetVan': 'Kẹt van + lũ kiểm tra – đánh giá rủi ro'}
TM = {'P1_HaiTran': 77.63, 'P1_KetVan': 78.21, 'P02_HaiTran': 78.54, 'P02_KetVan': 79.14}
for k, (ten_ws, P, th) in enumerate(CA):
    r = 4 + k
    ghi(kq, 'A%d' % r, ten_ws, f_dam)
    ghi(kq, 'B%d' % r, "='%s'!C3&\" – \"&'%s'!F3" % (ten_ws, ten_ws), can=cg)
    for col, src, fmt in (('C', 'F6', '0.00'), ('D', 'F7', '0'), ('E', 'F8', '0.0'), ('F', 'F9', '0.0'),
                          ('G', 'G9', '0.0'), ('H', 'F11', '0.00')):
        ghi(kq, '%s%d' % (col, r), "='%s'!%s" % (ten_ws, src), f_dam if col == 'C' else f_thuong, fmt, nen_kq, cg)
    ghi(kq, 'I%d' % r, TM[ten_ws], f_nhap, '0.00', can=cg)
    ghi(kq, 'J%d' % r, Y_NGHIA[ten_ws], f_ghi_chu)
ghi(kq, 'A9', 'Cột “Thuyết minh”: chương trình của đồ án, đầu lũ dùng cửa van tràn sự cố giữ MNDBT rồi mới mở hết (QTVH) – '
              'cao hơn giả thiết mở hoàn toàn từ đầu khoảng 2 cm; trường hợp kẹt van trùng nhau.', f_ghi_chu, border=False)
ghi(kq, 'A11', 'SO SÁNH VỚI BẢNG TÍNH HAI TRÀN ĐỘC LẬP (bảng của bạn)', f_muc, border=False)
tieu_de_cot(kq, 12, 1, ['Cách tính', 'Zmax P = 1% (m)', 'Zmax P = 0,2% (m)', '', '', 'Nhận xét'], cao=30)
kq.merge_cells('F12:J12')
SS = [('Chỉ tràn chính (bảng của bạn)', 78.28, 79.15, 'Thực chất là trường hợp kẹt van tràn sự cố – gần kết quả kẹt van của file này'),
      ('Chỉ tràn sự cố (bảng của bạn)', 78.37, 79.39, 'Trường hợp cả 9 cửa tràn chính không mở – không phải trường hợp thiết kế'),
      ('Kẹt van – file này', "=C5", "=C7", 'Chỉ tràn chính làm việc'),
      ('Hai tràn cùng làm việc – file này', "=C4", "=C6", 'ĐÚNG cho MNLTK, MNLKT: q(Z) = q tràn chính(Z) + q tràn sự cố(Z), điều tiết một lần')]
for k, (nhan, a, b, nx) in enumerate(SS):
    r = 13 + k
    ghi(kq, 'A%d' % r, nhan, f_dam)
    ghi(kq, 'B%d' % r, a, f_nhap if not isinstance(a, str) else f_dam, '0.00', None if not isinstance(a, str) else nen_kq, cg)
    ghi(kq, 'C%d' % r, b, f_nhap if not isinstance(b, str) else f_dam, '0.00', None if not isinstance(b, str) else nen_kq, cg)
    ghi(kq, 'F%d' % r, nx, f_ghi_chu)
    kq.merge_cells('F%d:J%d' % (r, r))
ghi(kq, 'A18', 'Kết luận: MNLTK = Zmax (P = 1%, hai tràn) =', f_dam, border=False)
ghi(kq, 'D18', '=C4', f_dam, '0.00', nen_kq, cg)
ghi(kq, 'A19', '             MNLKT = Zmax (P = 0,2%, hai tràn) =', f_dam, border=False)
ghi(kq, 'D19', '=C6', f_dam, '0.00', nen_kq, cg)
# số liệu biểu đồ
HB = 22
ghi(kq, 'A%d' % (HB - 1), 'SỐ LIỆU VẼ BIỂU ĐỒ (lấy từ các sheet điều tiết)', f_muc, border=False)
tieu_de_cot(kq, HB, 1, ['t (giờ)', 'Q đến P = 1%', 'q xả – hai tràn', 'q xả – kẹt van', 'Z P = 1% – hai tràn',
                        'Z P = 1% – kẹt van', 'Z P = 0,2% – hai tràn', 'Z P = 0,2% – kẹt van'], cao=45)
for t in range(0, T_MAX + 1):
    r = HB + 1 + t
    src = R0 + t
    for col, f in (('A', '=P1_HaiTran!A%d' % src), ('B', '=P1_HaiTran!B%d' % src), ('C', '=P1_HaiTran!J%d' % src),
                   ('D', '=P1_KetVan!J%d' % src), ('E', '=P1_HaiTran!G%d' % src), ('F', '=P1_KetVan!G%d' % src),
                   ('G', '=P02_HaiTran!G%d' % src), ('H', '=P02_KetVan!G%d' % src)):
        ghi(kq, '%s%d' % (col, r), f, fmt='0' if col == 'A' else ('0.0' if col in 'BCD' else '0.00'), can=cg)
HB1 = HB + 1 + T_MAX


def bieu_do(tieu_de, truc_y, cot0, cot1, mau, ymin, ymax, buoc, vi_tri):
    ch = LineChart()
    ch.title = tieu_de
    ch.y_axis.title = truc_y
    ch.x_axis.title = 't (giờ)'
    ch.height, ch.width = 9, 17
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = ymin, ymax
    ch.y_axis.majorUnit = buoc
    ch.y_axis.number_format = '0'
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblSkip = 6
    ch.legend.position = 'b'
    ch.add_data(Reference(kq, min_col=cot0, max_col=cot1, min_row=HB, max_row=HB1), titles_from_data=True)
    ch.set_categories(Reference(kq, min_col=1, min_row=HB + 1, max_row=HB1))
    for s, (m, net) in zip(ch.series, mau):
        s.smooth = False
        s.marker.symbol = 'none'
        s.graphicalProperties.line.solidFill = m
        s.graphicalProperties.line.width = 22225
        if net:
            s.graphicalProperties.line.dashStyle = net
    kq.add_chart(ch, vi_tri)


bieu_do('Lưu lượng đến và xả – P = 1%', 'Lưu lượng (m³/s)', 2, 4,
        [('52514E', None), ('2A78D6', None), ('2A78D6', 'dash')], 0, 1000, 100, 'L3')
bieu_do('Mực nước hồ theo thời gian', 'Mực nước (m)', 5, 8,
        [('2A78D6', None), ('2A78D6', 'dash'), ('EB6834', None), ('EB6834', 'dash')], 74, 80, 1, 'L22')
in_an(kq)

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'QuanHe': '8A8984', 'KetQua': '1BAF7A'
                                    }.get(ws.title, '2A78D6')
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
