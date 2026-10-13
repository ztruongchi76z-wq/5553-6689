# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL: ĐIỀU TIẾT NĂM THEO PHƯƠNG PHÁP LẬP BẢNG CÓ KỂ TỔN THẤT (năm thiết kế P = 85%) – HỒ CÀ GIÂY
Toàn bộ kết quả là công thức. Trình tự như giáo trình: điều tiết không kể tổn thất → tính tổn thất theo quá trình
dung tích → điều tiết có kể tổn thất → lặp đến khi dung tích hữu ích hội tụ. Tích sớm, hồ làm việc một lần trong năm.

  python3 xuat_excel_dieu_tiet_nam_lap_bang.py [đường dẫn .xlsx ra]
Mặc định: ../TAI_VE/DieuTietNam_LapBang_CoTonThat.xlsx
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
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'DieuTietNam_LapBang_CoTonThat.xlsx')

# ------------------------------------------------------------------ số liệu gốc (QTVH 2026)
W_DEN_85 = [9.30, 4.77, 4.19, 9.21, 15.62, 24.62, 14.01, 19.03, 20.02, 31.97, 12.96, 10.04]      # Bảng 3-19, tháng 1 ÷ 12
W_141 = [0.584, 0.322, 0.177, 0.096, 0.054, 0.031, 0.016, 6.359, 11.747, 15.744, 8.779, 2.344]   # Bảng 3-15
W_244 = [1.83, 0.77, 0.43, 1.48, 12.64, 20.18, 11.52, 11.18, 8.28, 16.23, 1.15, 0.33]          # Bảng 3-17
W_SL = [6.66, 3.80, 3.59, 7.39, 2.93, 4.26, 2.47, 1.44, 0.0, 0.0, 3.03, 7.13]                  # Bảng 3-18
NHU_CAU = {
    'TH1': [9.80, 11.93, 11.22, 17.62, 7.04, 10.13, 5.73, 9.91, 3.62, 3.20, 11.75, 20.19],      # Bảng 4-20
    'TH1A': [13.27, 10.48, 9.37, 14.20, 5.92, 8.31, 5.59, 8.90, 2.57, 2.99, 9.50, 15.54],     # Bảng 4-21
    'TH2': [13.48, 10.69, 9.61, 14.44, 5.94, 8.31, 5.59, 8.90, 2.57, 2.99, 9.50, 15.54],      # Bảng 4-22
}
MO_TA_TH = {
    'TH1': 'Tưới 8.965 ha, 100% cây trồng cạn tưới truyền thống + sinh hoạt (QTVH Bảng 4-20)',
    'TH1A': 'Tưới 8.965 ha, 50% cây trồng cạn tưới tiết kiệm + sinh hoạt (QTVH Bảng 4-21)',
    'TH2': 'Tưới tiềm năng 9.540 ha, 50% tưới tiết kiệm + sinh hoạt (QTVH Bảng 4-22)',
}
DZ_MM = [68.4, 64.4, 76.9, 66.6, 62.3, 55.1, 49.8, 51.8, 43.9, 40.8, 52.2, 66.7]                # Bảng 3-24

# ------------------------------------------------------------------ định dạng
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
nen_tt = PatternFill('solid', fgColor='EEF3FA')
nen_do = PatternFill('solid', fgColor='F8D7D7')
nen_xanh = PatternFill('solid', fgColor='E3F2EA')
vien = Border(*(Side(style='thin', color='B8BEC6'),) * 4)
giua = Alignment(horizontal='center', vertical='center', wrap_text=True)
trai = Alignment(horizontal='left', vertical='top', wrap_text=True)
can_giua = Alignment(horizontal='center')

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


def noi_suy(x, bx, by):
    """Nội suy tuyến tính trong bảng (bx tăng dần): tìm dòng k có bx_k ≤ x (MATCH) rồi nội suy."""
    k = 'MATCH(%s,%s,1)' % (x, bx)
    return ('INDEX({by},{k})+({x}-INDEX({bx},{k}))*(INDEX({by},{k}+1)-INDEX({by},{k}))'
            '/(INDEX({bx},{k}+1)-INDEX({bx},{k}))').format(x=x, bx=bx, by=by, k=k)


def in_an(ws, ngang=True):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if ngang else 'portrait'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5


# ================================================================== HƯỚNG DẪN
hd = wb.active
hd.title = 'HuongDan'
hd.column_dimensions['A'].width = 3
hd.column_dimensions['B'].width = 125
hd.sheet_view.showGridLines = False
ND = [
    ('T', 'ĐIỀU TIẾT NĂM – PHƯƠNG PHÁP LẬP BẢNG CÓ KỂ TỔN THẤT (năm thiết kế P = 85%) – HỒ CHỨA NƯỚC CÀ GIÂY'),
    ('g', 'Mọi ô kết quả là công thức: đổi số liệu ở sheet SoLieu (trường hợp dùng nước, MNC, MNDBT, k_t, tháng bắt đầu năm thủy lợi) thì toàn bộ file tự tính lại.'),
    ('M', '1. Vì sao chọn phương pháp lập bảng cho đồ án'),
    ('-', 'Là phương pháp chuẩn của giáo trình Thủy văn công trình, Quy hoạch và thiết kế hệ thống thủy lợi – thầy, cô và hội đồng quen; mỗi con số tính tay được và giải thích được.'),
    ('-', 'Khối lượng vừa với đồ án: 1 năm thiết kế × 12 tháng, 4 bảng (không kể tổn thất, có kể tổn thất lần 1, 2, 3), in vào thuyết minh 2 ÷ 3 trang. Phương pháp toàn chuỗi cần 44 năm × 12 tháng = 528 bước tính – phải lập trình hoặc bảng tính lớn.'),
    ('-', 'Đủ trả lời câu hỏi của đồ án sửa chữa: giữ MNC +67,50 m, MNDBT +74,70 m thì dung tích hữu ích hiện có (bình đồ 2022) có đủ cấp nước cho năm thiết kế không, MNDBT cần thiết là bao nhiêu.'),
    ('-', 'Hạn chế: chỉ xét một năm đại biểu nên không cho mức bảo đảm thực của cả chuỗi; trường hợp sát ngưỡng có thể kết luận khác toàn chuỗi (TH2 ở mục 5). Nếu được hỏi, nêu kết quả kiểm tra toàn chuỗi đã có trong thuyết minh: TH1A 39/44 năm (86,7%), TH2 37/44 năm (82,2%).'),
    ('M', '2. Số liệu (sheet SoLieu, có ghi nguồn)'),
    ('-', 'Dòng chảy năm thiết kế P = 85% đến hồ: QTVH 2026 Bảng 3-19 = nội lưu vực 141 km² (Bảng 3-15, mô hình năm 2023 có phân phối bất lợi) + khu giữa 244 km² đến đập dâng Sông Lũy đã trừ dòng chảy tối thiểu (Bảng 3-17) + nước hồ Sông Lũy cấp qua kênh tiếp nước (Bảng 3-18).'),
    ('-', 'Nhu cầu dùng nước theo tháng: QTVH Bảng 4-20 (TH1), 4-21 (TH1A), 4-22 (TH2). Tổn thất bốc hơi phụ thêm ΔZ: QTVH Bảng 3-24. Tổn thất thấm k_t = 1%/tháng dung tích bình quân (TCVN 10778:2024 Bảng 1).'),
    ('-', 'Quan hệ Z ~ F ~ W: bình đồ khảo sát 2022 (QTVH Bảng 1-2), dung tích tính lại từ diện tích; sheet BangTra lập bảng chi tiết 0,1 m để nội suy (Vc = V(+67,50) = 6,82; V(+74,70) = 36,41 triệu m³ – trùng thuyết minh).'),
    ('M', '3. Hình thức điều tiết (sheet B1, phần đầu)'),
    ('-', 'Năm thiết kế: W đến = 175,75·10⁶ m³ > W dùng (TH1A) = 106,65·10⁶ m³ → lượng nước đến trong năm đủ cho nhu cầu cả năm, hồ chỉ cần trữ nước thừa mùa mưa bù cho mùa khô cùng năm → ĐIỀU TIẾT NĂM. (Lưu ý: W là tổng lượng, đơn vị 10⁶ m³, không phải m³/s.)'),
    ('-', 'Năm thủy lợi bắt đầu từ tháng 5 – tháng đầu tiên lượng nước đến lớn hơn lượng nước dùng sau mùa khô: đầu tháng 5 hồ ở MNC; tích nước các tháng thừa (5 ÷ 11), đầy thì xả thừa; cấp nước các tháng thiếu (12 ÷ 4); cuối tháng 4 hồ về MNC.'),
    ('-', 'Kết luận này dựa vào nguồn nước bổ sung từ hệ thống Sông Lũy: riêng dòng chảy nội lưu vực 141 km² (46,25·10⁶ m³) nhỏ hơn nhu cầu.'),
    ('M', '4. Nguyên lý và các bước tính'),
    ('c', 'V cuối = V đầu + W đến − (W dùng + W tổn thất) − W xả,     Vc ≤ V ≤ Vc + Vhi'),
    ('-', 'Bước 1 – sheet B1_KhongTonThat: ΔV = W đến − W dùng; tháng thừa (ΔV > 0) hồ tích nước, tháng thiếu (ΔV < 0) hồ cấp nước. Hồ làm việc một lần trong năm (một thời kỳ cấp nước liên tục tháng 12 ÷ 4) nên dung tích hữu ích cần Vhi = tổng lượng thiếu. Quá trình dung tích tính theo cách tích sớm: thừa đến đâu tích đến đó, đầy (Vc + Vhi) thì xả.'),
    ('-', 'Bước 2 – sheet B2_Lan1: tính tổn thất từ quá trình dung tích của bước 1: V̄ = (V đầu + V cuối)/2 → tra Z̄ (bảng V ~ Z) → F̄ (bảng Z ~ F); W bốc hơi = ΔZ × F̄ (mm ÷ 1000 × km² = 10⁶ m³); W thấm = k_t × V̄; W tổn thất = W bốc hơi + W thấm. Điều tiết lại với W dùng + W tổn thất → Vhi lần 1.'),
    ('-', 'Bước 3, 4 – sheet B3_Lan2, B4_Lan3: tính lại tổn thất theo quá trình dung tích của lần trước và điều tiết lại; dừng khi sai số Vhi giữa hai lần liên tiếp ≤ 1% (thường 2 lần là đạt).'),
    ('-', 'Bước 5 – sheet KetLuan: Vbt cần = Vc + Vhi → MNDBT cần = Z(Vbt cần); so với MNDBT +74,70 m và dung tích hữu ích hiện có 29,60·10⁶ m³; biểu đồ quá trình dung tích hồ trong năm.'),
    ('-', 'Kiểm tra trong mỗi bảng: (1) số đợt thiếu nước trong năm phải bằng 1 (nếu > 1 thì hồ làm việc nhiều lần, không lấy Vhi bằng tổng lượng thiếu); (2) cuối năm thủy lợi hồ về đúng Vc.'),
    ('M', '5. Kết quả (đối chiếu chương trình Python)'),
    ('-', 'TH1A: Vhi không kể tổn thất 25,35 → lần 1: 27,76 → lần 2: 27,85 → lần 3: 27,85·10⁶ m³ (hội tụ). MNDBT cần ≈ +74,42 m < +74,70 m; Vhi hiện có 29,60 > 27,85 → ĐẠT (dư khoảng 1,75·10⁶ m³).'),
    ('-', 'TH2: Vhi = 28,81·10⁶ m³, MNDBT cần ≈ +74,57 m → đạt theo năm thiết kế, nhưng toàn chuỗi 44 năm chỉ 37/44 năm (82,2%) – trường hợp sát ngưỡng, phải nêu rõ.   TH1: Vhi = 36,19·10⁶ m³, MNDBT cần ≈ +75,71 m > +74,70 m → KHÔNG ĐẠT.'),
    ('M', '6. Chú giải màu ô và bài tập'),
    ('-', 'Nền vàng chữ xanh: ô được sửa để thử. Chữ xanh nền trắng: số liệu gốc hồ sơ (có ghi nguồn). Chữ đen: công thức – không gõ đè. Nền xanh lá: tháng thừa nước; nền đỏ nhạt: tháng thiếu nước.'),
    ('-', 'Bài tập: chọn TH1, TH2 ở ô SoLieu!C5 rồi xem KetLuan; đổi k_t = 0,5% xem Vhi giảm bao nhiêu; tự bấm máy tính tay tháng 12 của bảng B2 rồi so với Excel.'),
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
        hd.row_dimensions[h].height = 22
    elif loai == 'g':
        o.font, hd.row_dimensions[h].height = f_ghi_chu, 30
    else:
        o.font = f_thuong
        hd.row_dimensions[h].height = 15.5 * max(1, math.ceil(len(txt) / 145))

# ================================================================== SỐ LIỆU
sl = wb.create_sheet('SoLieu')
for c, w in zip('ABCDEFGHIJKLMNOP', (46, 9, 10, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 11, 55)):
    sl.column_dimensions[c].width = w
ghi(sl, 'A1', 'SỐ LIỆU ĐẦU VÀO – ĐIỀU TIẾT NĂM THIẾT KẾ P = 85%', f_tieu_de, border=False)
ghi(sl, 'A2', 'Ô nền vàng chữ xanh: được sửa. Chữ xanh: số liệu gốc hồ sơ. Chữ đen: công thức. Dung tích: 10⁶ m³.',
    f_ghi_chu, border=False)
ghi(sl, 'A4', 'A. THÔNG SỐ TÍNH TOÁN', f_muc, border=False)
tieu_de_cot(sl, 4, 3, ['Giá trị'])
ghi(sl, 'D4', 'Đơn vị', f_trang, fill=nen_tieu_de, can=giua)
ghi(sl, 'E4', 'Ghi chú, nguồn', f_trang, fill=nen_tieu_de, can=giua)
sl.merge_cells('E4:P4')
TS = [
    (5, 'Trường hợp dùng nước (chọn TH1 / TH1A / TH2)', 'TH1A', None, '', 'Mặc định TH1A – trường hợp thuyết minh kết luận đạt'),
    (6, 'Mực nước chết MNC', 67.5, '0.00', 'm', 'Giữ như thiết kế (thuyết minh 2.2.1)'),
    (7, 'Mực nước dâng bình thường MNDBT hiện có', 74.7, '0.00', 'm', 'Bằng ngưỡng tràn chính'),
    (8, 'Hệ số tổn thất thấm k_t', 0.01, '0.0%', '/tháng', 'TCVN 10778:2024 Bảng 1: 0,5 ÷ 1,0%/tháng (địa chất lòng hồ tốt); đồ án lấy 1%'),
    (9, 'Sai số cho phép của Vhi giữa hai lần lặp', 0.01, '0.0%', '', 'Dừng lặp khi |Vhi lần k − Vhi lần k−1| / Vhi lần k ≤ giá trị này'),
    (10, 'Tháng bắt đầu năm thủy lợi', 5, '0', 'tháng', 'Tháng đầu tiên W đến > W dùng sau mùa khô (xem B1)'),
]
for hang, nhan, gt, fmt, dv, gc in TS:
    ghi(sl, 'A%d' % hang, nhan, f_dam)
    sl.merge_cells('A%d:B%d' % (hang, hang))
    ghi(sl, 'C%d' % hang, gt, f_nhap_dam, fmt, nen_nhap, can_giua)
    ghi(sl, 'D%d' % hang, dv, can=can_giua)
    ghi(sl, 'E%d' % hang, gc, f_ghi_chu)
    sl.merge_cells('E%d:P%d' % (hang, hang))
dv_th = DataValidation(type='list', formula1='"TH1,TH1A,TH2"', allow_blank=False)
sl.add_data_validation(dv_th)
dv_th.add('C5')
dv_m = DataValidation(type='whole', operator='between', formula1='1', formula2='12')
sl.add_data_validation(dv_m)
dv_m.add('C10')
for nm, o in (('TH_chon', 'C5'), ('Z_MNC', 'C6'), ('Z_MNDBT', 'C7'), ('k_tham', 'C8'), ('SaiSo_cp', 'C9'),
              ('Thang_dau', 'C10')):
    ten(nm, 'SoLieu!$%s$%s' % (o[0], o[1:]))
V_Z = lambda z: noi_suy(z, 'BT_Z', 'BT_V')
KQ = [(11, 'Dung tích chết Vc = V(MNC)', '=' + V_Z('Z_MNC'), 'Tra bảng 0,1 m (sheet BangTra) – bình đồ 2022. Thiết kế ban đầu 8,41 (chỉ so sánh)'),
      (12, 'Dung tích ứng MNDBT hiện có Vbt = V(MNDBT)', '=' + V_Z('Z_MNDBT'), 'Thiết kế ban đầu 37,24'),
      (13, 'Dung tích hữu ích hiện có Vhi = Vbt − Vc', '=C12-C11', 'Thiết kế ban đầu 28,83')]
for hang, nhan, ct, gc in KQ:
    ghi(sl, 'A%d' % hang, nhan, f_dam)
    sl.merge_cells('A%d:B%d' % (hang, hang))
    ghi(sl, 'C%d' % hang, ct, f_dam, '0.000', can=can_giua)
    ghi(sl, 'D%d' % hang, '10⁶ m³', can=can_giua)
    ghi(sl, 'E%d' % hang, gc, f_ghi_chu)
    sl.merge_cells('E%d:P%d' % (hang, hang))
ten('V_chet', 'SoLieu!$C$11')
ten('V_bt_co', 'SoLieu!$C$12')
ten('Vhi_co', 'SoLieu!$C$13')

THANG_HDR = ['Tháng %d' % m for m in range(1, 13)]
# B. dòng chảy năm thiết kế
ghi(sl, 'A15', 'B. DÒNG CHẢY NĂM THIẾT KẾ P = 85% ĐẾN HỒ CÀ GIÂY (10⁶ m³/tháng) – QTVH 2026', f_muc, border=False)
tieu_de_cot(sl, 16, 1, ['Thành phần'])
sl.merge_cells('A16:B16')
tieu_de_cot(sl, 16, 3, THANG_HDR + ['Cả năm', 'Nguồn'], cao=30)
DONG = [(17, 'Nội lưu vực hồ 141 km² (mô hình năm 2023)', W_141, 'Bảng 3-15', f_nhap),
        (18, 'Khu giữa 244 km² đến đập dâng Sông Lũy (đã trừ dòng chảy tối thiểu)', W_244, 'Bảng 3-17', f_nhap),
        (19, 'Hồ Sông Lũy cấp bổ sung qua kênh tiếp nước', W_SL, 'Bảng 3-18', f_nhap),
        (20, 'TỔNG LƯỢNG NƯỚC ĐẾN HỒ – dùng tính toán', W_DEN_85, 'Bảng 3-19 (tổng QTVH; cộng 3 dòng trên lệch < 0,5%)', f_nhap_dam)]
for hang, nhan, vals, nguon, ft in DONG:
    ghi(sl, 'A%d' % hang, nhan, f_dam if hang == 20 else f_thuong)
    sl.merge_cells('A%d:B%d' % (hang, hang))
    for m in range(12):
        ghi(sl, '%s%d' % (get_column_letter(3 + m), hang), vals[m], ft, '0.00', nen_kq if hang == 20 else None, can_giua)
    ghi(sl, 'O%d' % hang, '=SUM(C%d:N%d)' % (hang, hang), f_dam, '0.00', nen_kq if hang == 20 else None, can_giua)
    ghi(sl, 'P%d' % hang, nguon, f_ghi_chu)
ten('DongChay85', 'SoLieu!$C$20:$N$20')
# C. nhu cầu
ghi(sl, 'A22', 'C. NHU CẦU DÙNG NƯỚC TỔNG (tưới + sinh hoạt), 10⁶ m³/tháng – QTVH Bảng 4-20 ÷ 4-22', f_muc, border=False)
tieu_de_cot(sl, 23, 1, ['Trường hợp'])
sl.merge_cells('A23:B23')
tieu_de_cot(sl, 23, 3, THANG_HDR + ['Cả năm', 'Nội dung'], cao=30)
for k, th in enumerate(('TH1', 'TH1A', 'TH2')):
    hang = 24 + k
    ghi(sl, 'A%d' % hang, th, f_dam, can=can_giua)
    sl.merge_cells('A%d:B%d' % (hang, hang))
    for m in range(12):
        ghi(sl, '%s%d' % (get_column_letter(3 + m), hang), NHU_CAU[th][m], f_nhap, '0.00', can=can_giua)
    ghi(sl, 'O%d' % hang, '=SUM(C%d:N%d)' % (hang, hang), f_dam, '0.00', can=can_giua)
    ghi(sl, 'P%d' % hang, MO_TA_TH[th], f_ghi_chu)
ghi(sl, 'A27', '→ Trường hợp đang chọn', f_dam, fill=nen_kq)
ghi(sl, 'B27', '=TH_chon', f_dam, fill=nen_kq, can=can_giua)
for m in range(12):
    cl = get_column_letter(3 + m)
    ghi(sl, '%s27' % cl, '=INDEX(%s$24:%s$26,MATCH(TH_chon,$A$24:$A$26,0))' % (cl, cl), f_dam, '0.00', nen_kq, can_giua)
ghi(sl, 'O27', '=SUM(C27:N27)', f_dam, '0.00', nen_kq, can_giua)
ten('NhuCau_chon', 'SoLieu!$C$27:$N$27')
# D. bốc hơi
ghi(sl, 'A29', 'D. TỔN THẤT BỐC HƠI PHỤ THÊM ΔZ (mm/tháng) – QTVH Bảng 3-24', f_muc, border=False)
tieu_de_cot(sl, 30, 1, ['Đại lượng'])
sl.merge_cells('A30:B30')
tieu_de_cot(sl, 30, 3, THANG_HDR + ['Cả năm'], cao=30)
ghi(sl, 'A31', 'ΔZ (mm)', f_dam)
sl.merge_cells('A31:B31')
for m in range(12):
    ghi(sl, '%s31' % get_column_letter(3 + m), DZ_MM[m], f_nhap, '0.0', can=can_giua)
ghi(sl, 'O31', '=SUM(C31:N31)', f_dam, '0.0', can=can_giua)
ten('BocHoi', 'SoLieu!$C$31:$N$31')
# E. Z ~ F ~ W
ghi(sl, 'A33', 'E. QUAN HỆ Z ~ F ~ W LÒNG HỒ – BÌNH ĐỒ 2022 (QTVH Bảng 1-2)', f_muc, border=False)
tieu_de_cot(sl, 34, 1, ['Z (m)', 'F (km²)', 'W hồ sơ', 'W tính lại'], cao=30)
for k, (z, f, w) in enumerate(zip(S.Z_TAB, S.F_TAB, S.W_TAB_HOSO)):
    hang = 35 + k
    ghi(sl, 'A%d' % hang, float(z), f_nhap, '0.00', can=can_giua)
    ghi(sl, 'B%d' % hang, float(f), f_nhap, '0.00', can=can_giua)
    ghi(sl, 'C%d' % hang, float(w), f_nhap, '0.00', can=can_giua)
    ct = '=0' if k == 0 else '=D{p}+(A{h}-A{p})/3*(B{p}+B{h}+SQRT(B{p}*B{h}))'.format(p=hang - 1, h=hang)
    ghi(sl, 'D%d' % hang, ct, fmt='0.00', can=can_giua)
H_Z1 = 34 + len(S.Z_TAB)
for k, t in enumerate(['F ghi “ha” trong hồ sơ nhưng đúng là km² (TKCS Hình 1-1).',
                       'W tính lại theo chóp cụt ΔW = ΔZ/3·(F₁ + F₂ + √(F₁F₂)); dưới +69 m khớp cột W hồ sơ, từ +70 m cột W hồ sơ không khớp F.',
                       'Khi tính dùng bảng chi tiết 0,1 m (sheet BangTra) lập từ cột F.']):
    ghi(sl, 'F%d' % (35 + k), t, f_ghi_chu)
    sl.merge_cells('F%d:P%d' % (35 + k, 35 + k))
ten('Bang_Z', 'SoLieu!$A$35:$A$%d' % H_Z1)
ten('Bang_F', 'SoLieu!$B$35:$B$%d' % H_Z1)
sl.freeze_panes = 'C5'

# ================================================================== BẢNG TRA 0,1 m
bt = wb.create_sheet('BangTra')
for c, w in zip('ABCDE', (9, 11, 12, 3, 80)):
    bt.column_dimensions[c].width = w
ghi(bt, 'A1', 'BẢNG TRA Z ~ F ~ V CHI TIẾT (bước 0,1 m)', f_tieu_de, border=False)
ghi(bt, 'A2', 'F nội suy tuyến tính từ bảng 1 m (SoLieu mục E); V cộng dồn chóp cụt ΔV = ΔZ/3·(F₁ + F₂ + √(F₁F₂)).',
    f_ghi_chu, border=False)
tieu_de_cot(bt, 4, 1, ['Z (m)', 'F (km²)', 'V (10⁶ m³)'])
N_BT = 261
for k in range(N_BT):
    hang = 5 + k
    ghi(bt, 'A%d' % hang, 59.0 if k == 0 else '=ROUND(A%d+0.1,1)' % (hang - 1), fmt='0.0', can=can_giua)
    ghi(bt, 'B%d' % hang, '=IF(A{h}>=MAX(Bang_Z),INDEX(Bang_F,ROWS(Bang_F)),{f})'.format(
        h=hang, f=noi_suy('A%d' % hang, 'Bang_Z', 'Bang_F')), fmt='0.000', can=can_giua)
    ghi(bt, 'C%d' % hang, 0 if k == 0 else '=C{p}+(A{h}-A{p})/3*(B{p}+B{h}+SQRT(B{p}*B{h}))'.format(p=hang - 1, h=hang),
        fmt='0.000', can=can_giua)
ten('BT_Z', 'BangTra!$A$5:$A$%d' % (4 + N_BT))
ten('BT_F', 'BangTra!$B$5:$B$%d' % (4 + N_BT))
ten('BT_V', 'BangTra!$C$5:$C$%d' % (4 + N_BT))
for k, t in enumerate(['Biết Z tìm V (hoặc ngược lại): tìm dòng kẹp giá trị cần tra rồi nội suy tuyến tính.',
                       'Ví dụ: V(+67,50) = 6,816; V(+74,70) = 36,412 triệu m³.']):
    ghi(bt, 'E%d' % (5 + k), t, f_ghi_chu, border=False)
bt.freeze_panes = 'A5'

# ================================================================== CÁC BẢNG ĐIỀU TIẾT
R0, R1 = 6, 17                 # 12 tháng của năm thủy lợi
REG_HDR = ['Tháng', 'W đến (10⁶ m³)', 'W dùng (10⁶ m³)', 'W tổn thất (10⁶ m³)', 'W dùng + W tổn thất',
           'Thừa ΔV+ = W đến − (W dùng + W tt)', 'Thiếu ΔV− = (W dùng + W tt) − W đến', 'V đầu tháng (10⁶ m³)',
           'V cuối tháng (10⁶ m³)', 'W xả thừa (10⁶ m³)', 'Z cuối tháng (m)']
TT_HDR = ['Tháng', 'V đầu tháng (lần trước)', 'V cuối tháng (lần trước)', 'V̄ = (V đầu + V cuối)/2', 'Z̄ (m) tra V ~ Z',
          'F̄ (km²) tra Z ~ F', 'ΔZ bốc hơi (mm)', 'W bốc hơi = ΔZ·F̄', 'W thấm = k_t·V̄', 'W tổn thất']


def bang_dieu_tiet(ws, ten_ws, tieu_de, co_ton_that, ws_truoc=None):
    """Bảng điều tiết ở A5:K18 (giống nhau ở mọi sheet) + phần tổng hợp A20:H27; bảng tổn thất ở M5:V18."""
    for c, w in zip('ABCDEFGHIJK', (8, 10, 10, 10, 11, 13, 13, 11, 11, 10, 10)):
        ws.column_dimensions[c].width = w
    ws.column_dimensions['L'].width = 3
    for c, w in zip('MNOPQRSTUV', (8, 11, 11, 11, 10, 10, 10, 11, 11, 11)):
        ws.column_dimensions[c].width = w
    ghi(ws, 'A1', tieu_de, f_tieu_de, border=False)
    for o, t, nm, fmt in (('A2', 'TH', 'TH_chon', None), ('D2', 'Vc =', 'V_chet', '0.000'),
                          ('G2', 'k_t =', 'k_tham', '0.0%')):
        ghi(ws, o, t, f_dam, fill=nen_kq)
        c2 = '%s2' % get_column_letter(ws[o].column + 1)
        ghi(ws, c2, '=' + nm, f_dam, fmt, nen_kq, can_giua)
    ghi(ws, 'A4', 'BẢNG ĐIỀU TIẾT NĂM %s' % ('CÓ KỂ TỔN THẤT' if co_ton_that else 'KHÔNG KỂ TỔN THẤT'), f_muc, border=False)
    tieu_de_cot(ws, 5, 1, REG_HDR, cao=58)
    for i in range(12):
        r = R0 + i
        cts = {
            'A': '=Thang_dau' if i == 0 else '=MOD(A%d,12)+1' % (r - 1),
            'B': '=INDEX(DongChay85,1,A%d)' % r,
            'C': '=INDEX(NhuCau_chon,1,A%d)' % r,
            'D': ('=V%d' % r) if co_ton_that else 0,
            'E': '=C%d+D%d' % (r, r),
            'F': '=MAX(B%d-E%d,0)' % (r, r),
            'G': '=MAX(E%d-B%d,0)' % (r, r),
            'H': '=V_chet' if i == 0 else '=I%d' % (r - 1),
            'I': '=IF(F{r}>0,MIN(H{r}+F{r},V_chet+$H$20),H{r}-G{r})'.format(r=r),
            'J': '=IF(F{r}>0,MAX(H{r}+F{r}-(V_chet+$H$20),0),0)'.format(r=r),
            'K': '=' + noi_suy('I%d' % r, 'BT_V', 'BT_Z'),
        }
        for col, ct in cts.items():
            ghi(ws, '%s%d' % (col, r), ct, f_dam if col == 'A' else f_thuong,
                '0' if col == 'A' else ('0.00' if col == 'K' else '0.00'), can=can_giua)
    ghi(ws, 'A18', 'Cộng', f_dam, fill=nen_kq, can=can_giua)
    for col in 'BCDEFGJ':
        ghi(ws, '%s18' % col, '=SUM(%s%d:%s%d)' % (col, R0, col, R1), f_dam, '0.00', nen_kq, can_giua)
    for col in 'HIK':
        ghi(ws, '%s18' % col, '', fill=nen_kq)
    ws.conditional_formatting.add('F%d:F%d' % (R0, R1), FormulaRule(formula=['F%d>0' % R0], fill=nen_xanh))
    ws.conditional_formatting.add('G%d:G%d' % (R0, R1), FormulaRule(formula=['G%d>0' % R0], fill=nen_do))
    # phần tổng hợp
    TH = [
        (20, 'Dung tích hữu ích cần Vhi = Σ ΔV− (tổng lượng thiếu của thời kỳ cấp nước)', '=G18', '0.00'),
        (21, 'Dung tích ứng MNDBT cần Vbt = Vc + Vhi', '=V_chet+H20', '0.00'),
        (22, 'MNDBT cần = Z(Vbt cần) (m)', '=' + noi_suy('H21', 'BT_V', 'BT_Z'), '0.00'),
        (23, 'Kiểm tra: V cuối năm thủy lợi − Vc (phải = 0)', '=I17-V_chet', '0.000'),
        (24, 'Kiểm tra: số đợt thiếu nước trong năm (phải = 1)',
         '=SUMPRODUCT((G%d:G%d>0)*(G%d:G%d=0))+IF(G%d>0,1,0)' % (R0 + 1, R1, R0, R1 - 1, R0), '0'),
    ]
    if ws_truoc:
        TH += [(25, 'Vhi lần trước (sheet %s)' % ws_truoc, "='%s'!H20" % ws_truoc, '0.00'),
               (26, 'Sai số |Vhi − Vhi lần trước| / Vhi', "=ABS(H20-H25)/H20", '0.00%'),
               (27, 'Đánh giá', '=IF(H26<=SaiSo_cp,"Hội tụ – dừng lặp","Chưa hội tụ – lặp tiếp")', None)]
    for hang, nhan, ct, fmt in TH:
        ws.merge_cells('A%d:G%d' % (hang, hang))
        ghi(ws, 'A%d' % hang, nhan, f_dam)
        ghi(ws, 'H%d' % hang, ct, f_dam, fmt, nen_kq, can_giua)
        if hang == 27:
            ws.merge_cells('H27:K27')
    ghi(ws, 'A29', 'Cách tích nước: tích sớm – tháng thừa thì tích đến khi đầy (Vc + Vhi), phần thừa còn lại xả qua tràn; '
                   'tháng thiếu thì hồ cấp nước (dung tích giảm đúng bằng lượng thiếu).', f_ghi_chu, border=False)
    if co_ton_that:
        ghi(ws, 'M4', 'BẢNG TÍNH TỔN THẤT (theo quá trình dung tích của sheet %s)' % ws_truoc, f_muc, border=False)
        tieu_de_cot(ws, 5, 13, TT_HDR, cao=58)
        for i in range(12):
            r = R0 + i
            cts = {
                'M': '=A%d' % r,
                'N': "='%s'!H%d" % (ws_truoc, r),
                'O': "='%s'!I%d" % (ws_truoc, r),
                'P': '=(N%d+O%d)/2' % (r, r),
                'Q': '=' + noi_suy('P%d' % r, 'BT_V', 'BT_Z'),
                'R': '=' + noi_suy('Q%d' % r, 'Bang_Z', 'Bang_F'),
                'S': '=INDEX(BocHoi,1,M%d)' % r,
                'T': '=S%d/1000*R%d' % (r, r),
                'U': '=k_tham*P%d' % r,
                'V': '=T%d+U%d' % (r, r),
            }
            for col, ct in cts.items():
                fmt = {'M': '0', 'Q': '0.00', 'R': '0.000', 'S': '0.0'}.get(col, '0.000')
                ghi(ws, '%s%d' % (col, r), ct, f_dam if col == 'M' else f_thuong, fmt, nen_tt, can_giua)
        ghi(ws, 'M18', 'Cộng', f_dam, fill=nen_kq, can=can_giua)
        for col in 'NOPQRS':
            ghi(ws, '%s18' % col, '', fill=nen_kq)
        for col in 'TUV':
            ghi(ws, '%s18' % col, '=SUM(%s%d:%s%d)' % (col, R0, col, R1), f_dam, '0.000', nen_kq, can_giua)
        ghi(ws, 'M20', 'Cột D của bảng điều tiết (W tổn thất) lấy từ cột “W tổn thất” của bảng này.', f_ghi_chu, border=False)
    ws.freeze_panes = 'B6'
    in_an(ws)


b1 = wb.create_sheet('B1_KhongTonThat')
bang_dieu_tiet(b1, 'B1_KhongTonThat', 'BƯỚC 1 – XÁC ĐỊNH HÌNH THỨC ĐIỀU TIẾT VÀ ĐIỀU TIẾT NĂM KHÔNG KỂ TỔN THẤT', False)
# hình thức điều tiết (đặt ở cột M:V của B1)
b1.column_dimensions['M'].width = 56
b1.column_dimensions['N'].width = 14
ghi(b1, 'M6', 'XÁC ĐỊNH HÌNH THỨC ĐIỀU TIẾT', f_muc, border=False)
HT = [
    (7, 'Tổng lượng nước đến năm thiết kế P = 85%, W đến (10⁶ m³)', '=SUM(DongChay85)', '0.00'),
    (8, 'Tổng nhu cầu dùng nước trong năm, W dùng (10⁶ m³)', '=SUM(NhuCau_chon)', '0.00'),
    (9, 'So sánh', '=IF(N7>N8,"W đến > W dùng → ĐIỀU TIẾT NĂM","W đến ≤ W dùng → cần ĐIỀU TIẾT NHIỀU NĂM")', None),
    (10, 'Riêng nội lưu vực 141 km² (10⁶ m³)', '=SoLieu!O17', '0.00'),
    (11, 'Tổng lượng thừa trong năm Σ ΔV+ (10⁶ m³)', '=F18', '0.00'),
    (12, 'Tổng lượng thiếu trong năm Σ ΔV− (10⁶ m³)', '=G18', '0.00'),
]
for hang, nhan, ct, fmt in HT:
    ghi(b1, 'M%d' % hang, nhan, f_dam)
    ghi(b1, 'N%d' % hang, ct, f_dam, fmt, nen_kq, can_giua)
    if hang == 9:
        b1.merge_cells('N9:S9')
for k, t in enumerate(['W là tổng lượng nước, đơn vị 10⁶ m³ (không phải m³/s).',
                       'Năm thủy lợi bắt đầu từ tháng có W đến > W dùng đầu tiên sau mùa khô (tháng 5): đầu tháng 5 hồ ở MNC,',
                       'tích nước các tháng thừa, cuối mùa lũ (tháng 11) hồ đầy MNDBT, cấp nước tháng 12 ÷ 4, cuối tháng 4 về MNC.',
                       'Kết luận điều tiết năm dựa vào nước bổ sung từ hệ thống Sông Lũy (riêng nội lưu vực nhỏ hơn nhu cầu).']):
    ghi(b1, 'M%d' % (14 + k), t, f_ghi_chu, border=False)

b2 = wb.create_sheet('B2_Lan1')
bang_dieu_tiet(b2, 'B2_Lan1', 'BƯỚC 2 – ĐIỀU TIẾT NĂM CÓ KỂ TỔN THẤT – LẦN 1', True, 'B1_KhongTonThat')
b3 = wb.create_sheet('B3_Lan2')
bang_dieu_tiet(b3, 'B3_Lan2', 'BƯỚC 3 – ĐIỀU TIẾT NĂM CÓ KỂ TỔN THẤT – LẦN 2', True, 'B2_Lan1')
b4 = wb.create_sheet('B4_Lan3')
bang_dieu_tiet(b4, 'B4_Lan3', 'BƯỚC 4 – ĐIỀU TIẾT NĂM CÓ KỂ TỔN THẤT – LẦN 3 (kiểm tra hội tụ)', True, 'B3_Lan2')

# ================================================================== KẾT LUẬN
kl = wb.create_sheet('KetLuan')
for c, w in zip('ABCDEFGH', (52, 13, 13, 13, 13, 3, 3, 3)):
    kl.column_dimensions[c].width = w
ghi(kl, 'A1', 'BƯỚC 5 – KẾT QUẢ ĐIỀU TIẾT NĂM VÀ KẾT LUẬN', f_tieu_de, border=False)
ghi(kl, 'A3', 'Trường hợp dùng nước', f_dam)
ghi(kl, 'B3', '=TH_chon', f_dam, fill=nen_kq, can=can_giua)
ghi(kl, 'A5', 'A. QUÁ TRÌNH LẶP', f_muc, border=False)
tieu_de_cot(kl, 6, 1, ['Lần tính', 'Vhi (10⁶ m³)', 'Tổng tổn thất năm', 'Sai số so với lần trước', 'MNDBT cần (m)'], cao=30)
LAP = [('Không kể tổn thất (B1)', 'B1_KhongTonThat', False), ('Có tổn thất – lần 1 (B2)', 'B2_Lan1', True),
       ('Có tổn thất – lần 2 (B3)', 'B3_Lan2', True), ('Có tổn thất – lần 3 (B4)', 'B4_Lan3', True)]
for k, (nhan, ws_ten, co) in enumerate(LAP):
    r = 7 + k
    ghi(kl, 'A%d' % r, nhan, f_dam)
    ghi(kl, 'B%d' % r, "='%s'!H20" % ws_ten, f_thuong, '0.00', can=can_giua)
    ghi(kl, 'C%d' % r, ("='%s'!V18" % ws_ten) if co else 0, f_thuong, '0.00', can=can_giua)
    ghi(kl, 'D%d' % r, ("='%s'!H26" % ws_ten) if k >= 1 else '–', f_thuong, '0.00%', can=can_giua)
    ghi(kl, 'E%d' % r, "='%s'!H22" % ws_ten, f_thuong, '0.00', can=can_giua)
ghi(kl, 'A12', 'B. SO SÁNH VỚI CÔNG TRÌNH HIỆN CÓ (lần tính cuối – B4)', f_muc, border=False)
SS = [
    (13, 'Dung tích hữu ích cần Vhi (10⁶ m³)', "='B4_Lan3'!H20", '0.00'),
    (14, 'Dung tích hữu ích hiện có (bình đồ 2022) (10⁶ m³)', '=Vhi_co', '0.00'),
    (15, 'Dư (+) / thiếu (−) dung tích (10⁶ m³)', '=B14-B13', '+0.00;-0.00;0.00'),
    (16, 'MNDBT cần (m)', "='B4_Lan3'!H22", '0.00'),
    (17, 'MNDBT hiện có (m)', '=Z_MNDBT', '0.00'),
    (18, 'Tổng tổn thất năm: bốc hơi / thấm (10⁶ m³)', "='B4_Lan3'!T18", '0.00'),
    (19, 'KẾT LUẬN', '=IF(B13<=B14,"ĐẠT – MNDBT hiện có đủ cấp nước năm thiết kế","KHÔNG ĐẠT – dung tích hữu ích không đủ")', None),
]
for hang, nhan, ct, fmt in SS:
    ghi(kl, 'A%d' % hang, nhan, f_dam)
    ghi(kl, 'B%d' % hang, ct, f_dam, fmt, nen_kq, can_giua)
ghi(kl, 'C18', "='B4_Lan3'!U18", f_dam, '0.00', nen_kq, can_giua)
kl.merge_cells('B19:E19')
kl['B19'].font = Font(name=FONT, size=12, bold=True, color='A65A10')
kl.conditional_formatting.add('B19', FormulaRule(formula=['LEFT($B$19,5)="KHÔNG"'], fill=nen_do))
for k, t in enumerate([
        'Kết quả tham khảo (chương trình Python): TH1A Vhi = 27,85 (MNDBT cần +74,42) – đạt; TH2 28,81 (+74,57) – đạt theo năm thiết kế; '
        'TH1 36,19 (+75,71) – không đạt.',
        'Lưu ý: phương pháp năm thiết kế chỉ xét một năm đại biểu. Kiểm tra toàn chuỗi 44 năm trong thuyết minh: TH1A 39/44 năm '
        '(86,7%) – đạt; TH2 37/44 năm (82,2%) – không đạt.']):
    ghi(kl, 'A%d' % (21 + k), t, f_ghi_chu, border=False)
# số liệu biểu đồ
ghi(kl, 'A24', 'C. QUÁ TRÌNH DUNG TÍCH HỒ TRONG NĂM THIẾT KẾ (lần tính cuối)', f_muc, border=False)
tieu_de_cot(kl, 25, 1, ['Tháng (năm thủy lợi)', 'V cuối tháng', 'Vc', 'Vc + Vhi cần', 'Vbt hiện có'], cao=30)
for i in range(12):
    r = 26 + i
    ghi(kl, 'A%d' % r, "=\"Tháng \"&'B4_Lan3'!A%d" % (R0 + i), can=can_giua)
    ghi(kl, 'B%d' % r, "='B4_Lan3'!I%d" % (R0 + i), fmt='0.00', can=can_giua)
    ghi(kl, 'C%d' % r, '=V_chet', fmt='0.00', can=can_giua)
    ghi(kl, 'D%d' % r, "='B4_Lan3'!H21", fmt='0.00', can=can_giua)
    ghi(kl, 'E%d' % r, '=V_bt_co', fmt='0.00', can=can_giua)
ch = LineChart()
ch.title = 'Dung tích hồ cuối tháng – năm thiết kế P = 85%'
ch.y_axis.title = 'Dung tích hồ (10⁶ m³)'
ch.x_axis.title = 'Tháng (năm thủy lợi)'
ch.height, ch.width = 9.5, 18
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 40
ch.y_axis.majorUnit = 5
ch.y_axis.number_format = '0'
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.legend.position = 'b'
ch.add_data(Reference(kl, min_col=2, max_col=5, min_row=25, max_row=37), titles_from_data=True)
ch.set_categories(Reference(kl, min_col=1, min_row=26, max_row=37))
for s, (mau, net, rong) in zip(ch.series, (('2A78D6', None, 22225), ('52514E', 'dash', 15875),
                                           ('EB6834', 'dash', 19050), ('52514E', 'sysDot', 15875))):
    s.smooth = False
    s.graphicalProperties.line.solidFill = mau
    s.graphicalProperties.line.width = rong
    if net:
        s.graphicalProperties.line.dashStyle = net
    s.marker.symbol = 'circle' if mau == '2A78D6' else 'none'
    if mau == '2A78D6':
        s.marker.size = 6
        s.marker.graphicalProperties.solidFill = mau
        s.marker.graphicalProperties.line.solidFill = mau
kl.add_chart(ch, 'G3')
in_an(kl)

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'BangTra': '8A8984',
                                    'B1_KhongTonThat': '2A78D6', 'B2_Lan1': '2A78D6', 'B3_Lan2': '2A78D6',
                                    'B4_Lan3': '2A78D6', 'KetLuan': '1BAF7A'}[ws.title]
in_an(hd)
in_an(sl)
in_an(bt, ngang=False)
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
