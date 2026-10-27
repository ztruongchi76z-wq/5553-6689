# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL: ĐIỀU TIẾT LŨ HỒ CÀ GIÂY – PHƯƠNG PHÁP PÔ-TA-PỐP VỚI HAI ĐƯỜNG PHỤ TRỢ q ~ f₁ VÀ q ~ f₂
Thông số (theo yêu cầu SV, 26/10/2026):
  tràn chính 9 khoang có cửa van, Σb = 21,6 m (9 × 2,4 m), ngưỡng +74,70; 8 trụ pin đầu lượn tròn: 6 trụ giữa dày 0,6 m
  (2 trụ mới trong mỗi khoang của tràn 3 khoang ban đầu) và 2 trụ chính ban đầu dày 1,0 m (bổ sung 26/10);
  tràn sự cố 1 khoang có cửa van, b = 5 m, ngưỡng +71,00.
  Vận hành bình thường: khi đón lũ mọi cửa van kéo lên hoàn toàn. Sự cố: 1 cửa tràn chính không mở được (8 khoang).
  Lũ đến: đường quá trình QTVH 2026 (thu phóng trận lũ trạm Đá Bàn, 72 giờ, 8 tần suất – SV sửa 26/10, thay lũ BCNCKT
  36 giờ), tính P = 1% và 0,2%; hồ ở MNDBT khi lũ đến, Δt = 1 giờ. Vận hành trước đỉnh lũ (ô VH_mo): 1 – mở hết cửa từ
  giờ 0; 2 – giữ MNDBT (xả bằng lũ đến) tới khi lũ đến lớn hơn khả năng xả khi mở hết cửa ở MNDBT.
  f₁ = V/Δt − q/2 ;  f₂ = V/Δt + q/2 ;  f₂ (cuối thời đoạn) = f₁ (đầu thời đoạn) + Qtb
Mọi ô kết quả là công thức; ô nền vàng chữ xanh là số liệu nhập.
  python3 xuat_excel_dieu_tiet_lu_potapov.py [đường dẫn .xlsx ra]
Mặc định: ../TAI_VE/DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx
"""
import math
import os
import sys

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import so_lieu as S

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx')
T_MAX = 72                          # số giờ tính = thời gian lũ QTVH
Z_DAU, Z_CUOI, BUOC = 71.0, 81.0, 0.1
N_QH = int(round((Z_CUOI - Z_DAU) / BUOC)) + 1
Z_BD0, Z_BD1 = 74.0, 80.0           # khoảng mực nước vẽ biểu đồ phụ trợ

# Đường quá trình lũ đến (m³/s), giờ 1 ÷ 72 – QTVH 2026, thu phóng trận lũ trạm Đá Bàn (QTVH Bảng 3-22, 3-23); SV nhập
# 26/10/2026, ba tần suất 0,1%, 0,2%, 1% trùng so_lieu.Q_QTVH
P_LU = [0.1, 0.2, 0.5, 1, 1.5, 2, 10, 20]
Q_LU = {
    0.1: S.Q_QTVH['0.1%'], 0.2: S.Q_QTVH['0.2%'], 1: S.Q_QTVH['1%'],
    0.5: [9, 10.4, 11.7, 25.5, 39.2, 46.8, 54.4, 64.2, 74, 83.3, 92.6, 79.2, 58, 74, 94, 116, 187, 369, 590, 844, 965, 1029,
          1044, 1040, 1028, 722, 502, 405, 331, 306, 247, 224, 208, 191, 178, 161, 154, 151, 143, 136, 128, 121, 119, 114, 110,
          108, 103, 100, 97, 93, 91, 90, 86, 86, 70, 55, 53, 51, 50, 50, 49, 48, 47, 47, 46, 46, 45, 44, 44, 43, 43, 43],
    1.5: [8, 9.2, 10.4, 22.6, 34.9, 41.6, 48.4, 57.1, 65.8, 74, 82.3, 70.4, 47, 60, 76, 93, 151, 297, 475, 679, 777, 818, 819,
          794, 688, 581, 404, 326, 266, 246, 199, 180, 168, 154, 143, 130, 124, 121, 115, 110, 103, 98, 95, 92, 89, 87, 83, 81,
          78, 75, 73, 72, 69, 69, 57, 44, 43, 41, 40, 40, 39, 39, 38, 38, 37, 37, 36, 36, 35, 35, 35, 34],
    2: [6.6, 7.6, 8.6, 18.7, 28.8, 34.4, 39.9, 47.1, 54.3, 61.1, 67.9, 58.1, 43, 55, 70, 86, 139, 275, 439, 628, 719, 752, 748,
        716, 627, 538, 374, 302, 246, 228, 184, 167, 155, 142, 132, 120, 115, 112, 106, 101, 96, 90, 88, 85, 82, 80, 77, 75,
        72, 69, 68, 67, 64, 64, 52, 41, 39, 38, 37, 37, 36, 36, 35, 35, 34, 34, 33, 33, 32, 32, 32, 32],
    10: [3.6, 4.2, 4.7, 10.2, 15.7, 18.8, 21.8, 25.8, 29.7, 33.4, 37.1, 31.8, 33, 42, 53, 65, 105, 207, 330, 472, 540, 551, 536,
         485, 430, 374, 281, 227, 185, 171, 138, 125, 117, 107, 99, 90, 86, 84, 80, 76, 72, 68, 66, 64, 62, 60, 58, 56, 54, 52,
         51, 50, 48, 48, 39, 31, 30, 29, 28, 28, 27, 27, 27, 26, 26, 26, 25, 25, 24, 24, 24, 24],
    20: [2.4, 2.8, 3.1, 6.8, 10.4, 12.5, 14.5, 17.1, 19.7, 22.2, 24.6, 21.1, 27, 35, 44, 54, 87, 172, 274, 392, 449, 449, 427,
         367, 321, 277, 233, 188, 154, 142, 115, 104, 97, 89, 83, 75, 72, 70, 66, 63, 60, 56, 55, 53, 51, 50, 48, 47, 45, 43,
         42, 42, 40, 40, 33, 25, 25, 24, 23, 23, 23, 22, 22, 22, 22, 21, 21, 21, 20, 20, 20, 20],
}
W_LU_QTVH = [67.1, 60.2, 52.2, 46.1, 42.0, 38.81, 29.2, 24.31]      # tổng lượng lũ (10⁶ m³) [QTVH Bảng 3-22]
assert all(len(Q_LU[p]) == T_MAX for p in P_LU)
Z_TRUOC = {'1%': (77.61, 77.67), '0.2%': (78.51, 78.60)}            # Zmax bản trước (lũ BCNCKT 36 giờ): bình thường, sự cố

FONT = 'Times New Roman'
f_thuong = Font(name=FONT, size=11)
f_dam = Font(name=FONT, size=11, bold=True)
f_tieu_de = Font(name=FONT, size=14, bold=True, color='12324A')
f_muc = Font(name=FONT, size=12, bold=True, color='A65A10')
f_nhap = Font(name=FONT, size=11, color='0000FF')
f_nhap_dam = Font(name=FONT, size=11, color='0000FF', bold=True)
f_trang = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_ghi_chu = Font(name=FONT, size=10, italic=True, color='52514E')
f_phu = Font(name=FONT, size=10, color='8A8984')
nen_nhap = PatternFill('solid', fgColor='FFF2CC')
nen_tieu_de = PatternFill('solid', fgColor='12324A')
nen_kq = PatternFill('solid', fgColor='FBF3E8')
nen_tc = PatternFill('solid', fgColor='EEF3FA')
nen_sc = PatternFill('solid', fgColor='F4F1FA')
nen_f = PatternFill('solid', fgColor='EAF6EF')
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
    # Excel không phân biệt chữ hoa, chữ thường trong tên (TC_b và TC_B là một tên)
    assert all(n.upper() != name.upper() for n in wb.defined_names), name
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


def bieu_do(ws, tieu_de, truc_x, truc_y, day, vi_tri, x_lim=None, y_lim=None, rong=18, cao=10, tron=False, so=('0', '0')):
    """day: (Reference x, Reference y, tên, màu, nét đứt?[, kiểu điểm])"""
    ch = ScatterChart()
    ch.title = tieu_de
    ch.style = 13
    ch.x_axis.title, ch.y_axis.title = truc_x, truc_y
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.x_axis.number_format, ch.y_axis.number_format = so
    ch.width, ch.height = rong, cao
    ch.legend.position = 'b'
    for truc, lim in ((ch.x_axis, x_lim), (ch.y_axis, y_lim)):
        if lim:
            truc.scaling.min, truc.scaling.max, truc.majorUnit = lim
    for x, y, nhan, mau, net, *diem in day:
        s = Series(y, x, title=nhan)
        s.smooth = tron
        s.marker.symbol = diem[0] if diem else 'none'
        if diem:
            s.marker.size = 6
            s.marker.graphicalProperties.solidFill = s.marker.graphicalProperties.line.solidFill = mau
        s.graphicalProperties.line.solidFill = mau
        s.graphicalProperties.line.width = 22225
        if net:
            s.graphicalProperties.line.dashStyle = 'dash'
        ch.series.append(s)
    ws.add_chart(ch, vi_tri)


# ------------------------------------------------------------------ giá trị gần đúng để đặt giới hạn trục biểu đồ
def _V(z):
    zt, ft = np.asarray(S.Z_TAB, float), np.asarray(S.F_TAB, float)
    wt = np.concatenate([[0.0], np.cumsum(np.diff(zt) * (ft[1:] + ft[:-1]) / 2)])
    k = int(np.searchsorted(zt, z, side='right') - 1)
    f = ft[k] + (z - zt[k]) * (ft[k + 1] - ft[k]) / (zt[k + 1] - zt[k])
    return wt[k] + (z - zt[k]) * (ft[k] + f) / 2


def _q(z, n_tc):
    def q1(Zn, n, b, H=None):
        H = max(z - Zn, 0.0)
        return (1 - 0.2 * (0.7 + (n - 1) * 0.45) * H / (n * b)) * 0.45 * n * b * math.sqrt(2 * 9.81) * H ** 1.5
    return q1(74.70, n_tc, 2.40) + q1(71.00, 1, 5.0)


F_MIN = math.floor((_V(Z_BD0) * 1e6 / 3600 - _q(Z_BD0, 9) / 2) / 500) * 500
F_MAX = math.ceil((_V(Z_BD1) * 1e6 / 3600 + _q(Z_BD1, 9) / 2) / 500) * 500
Q_MAX_BD = math.ceil(_q(Z_BD1, 9) / 100) * 100
Q_LU_MAX = math.ceil(max(max(q) for q in Q_LU.values()) / 200) * 200

# ================================================================== HƯỚNG DẪN
hd = wb.active
hd.title = 'HuongDan'
hd.column_dimensions['A'].width = 3
hd.column_dimensions['B'].width = 125
hd.sheet_view.showGridLines = False
ND = [
    ('T', 'ĐIỀU TIẾT LŨ HỒ CÀ GIÂY – PHƯƠNG PHÁP PÔ-TA-PỐP VỚI HAI ĐƯỜNG PHỤ TRỢ q ~ f₁, q ~ f₂'),
    ('g', 'Mọi ô kết quả là công thức. Sửa thông số tràn, mực nước trước lũ ở sheet SoLieu; đổi tần suất, trường hợp ở ô vàng '
          'đầu mỗi sheet điều tiết.'),
    ('M', '1. Công trình tháo lũ và các trường hợp tính'),
    ('-', 'Tràn chính: 9 khoang có cửa van, mỗi khoang b = 2,4 m (Σb = 21,6 m), ngưỡng +74,70 m; 8 trụ pin đầu lượn tròn gồm '
          '6 trụ giữa dày 0,6 m (2 trụ trong mỗi khoang của tràn 3 khoang ban đầu) và 2 trụ chính ban đầu dày 1,0 m → tổng '
          'bề rộng tràn 21,6 + 6 × 0,6 + 2 × 1,0 = 27,2 m. Tràn sự cố: 1 khoang có cửa van, b = 5 m, ngưỡng +71,00 m.'),
    ('-', 'Vận hành bình thường: khi đón lũ, tất cả cửa van (9 cửa tràn chính và cửa tràn sự cố) kéo lên hoàn toàn – hai tràn '
          'chảy tự do. Sự cố: 1 cửa tràn chính không mở được – tràn chính còn 8 khoang làm việc (Σb = 19,2 m), tràn sự cố mở '
          'hoàn toàn. Mỗi trường hợp tính cho lũ thiết kế P = 1% (MNLTK) và lũ kiểm tra P = 0,2% (MNLKT).'),
    ('-', 'Lũ đến: đường quá trình 72 giờ của QTVH 2026 (thu phóng trận lũ trạm Đá Bàn), 8 tần suất 0,1 ÷ 20%. Hồ ở MNDBT '
          '+74,70 m khi lũ đến (ô SoLieu!B6). 16 giờ đầu lưu lượng nhỏ (P = 1%: 8 ÷ 85 m³/s), đỉnh lũ chính ở giờ 21 ÷ 25.'),
    ('-', 'Ô VH_mo (SoLieu!B10) chọn cách vận hành trước đỉnh lũ. 1 – mở hết cửa từ giờ 0 (như yêu cầu): ngưỡng tràn sự cố thấp '
          'hơn MNDBT 3,7 m nên xả ngay ≈ 63,6 m³/s, lớn hơn lũ đến – mực nước hạ (thấp nhất ≈ +74,54 m ở giờ 8 với P = 1%) rồi '
          'mới dâng. 2 – giữ MNDBT: mở cửa từng phần cho xả đúng bằng lưu lượng đến, tới khi lũ đến lớn hơn khả năng xả khi mở '
          'hết cửa ở MNDBT thì kéo hết cửa – không hạ mực nước trước lũ, Zmax cao hơn so với cách 1 khoảng 4 ÷ 7 cm.'),
    ('M', '2. Lưu lượng qua tràn (TCVN 9147:2012 – công thức đập tràn chảy tự do)'),
    ('c', 'q = σn·ε·m·n·b·√(2g)·H^1,5        ε = 1 − 0,2·[ξmb + Σξtrụ]·H / (n·b)        Σξtrụ = n₁·ξ₁ + n₂·ξ₂'),
    ('-', 'H = Z − Z ngưỡng; n – số khoang làm việc; b – bề rộng một khoang; m = 0,45; ξmb = 0,70 (mố bên lượn tròn). Σξtrụ – '
          'tổng hệ số hình dạng của các trụ nằm giữa các khoang làm việc: 6 trụ giữa ξ₁ = 0,45 và 2 trụ chính ξ₂ = 0,45 (đều '
          'đầu lượn tròn) → Σξ = 3,60; khi các trụ cùng loại thì Σξ = (n − 1)·ξmt như công thức TCVN 9147:2012. Bề dày trụ '
          '(0,6 m, 1,0 m) không có trong công thức – chỉ hình dạng đầu trụ ảnh hưởng tới co hẹp ε. '
          'Sự cố: n = 8 (coi cửa hỏng ở khoang biên – trụ giữa 0,6 m cạnh nó không còn nằm giữa hai khoang làm việc: Σξ = '
          '5 × 0,45 + 2 × 0,45 = 3,15; nếu cửa hỏng ở giữa – hai nhóm 4 khoang – Zmax chỉ cao thêm 0,3 ÷ 0,7 cm). '
          'Hai tràn cùng một mực nước: q(Z) = q tràn chính(Z) + q tràn sự cố(Z).'),
    ('M', '3. Phương pháp Pô-ta-pốp với hai đường phụ trợ'),
    ('c', '(Q₁ + Q₂)/2·Δt − (q₁ + q₂)/2·Δt = V₂ − V₁    ⇒    V₂/Δt + q₂/2 = (V₁/Δt − q₁/2) + Qtb    ⇒    f₂ = f₁ + Qtb'),
    ('-', 'f₁ = V/Δt − q/2 và f₂ = V/Δt + q/2 (m³/s) chỉ phụ thuộc mực nước Z → lập sẵn bảng Z ~ V ~ q ~ f₁ ~ f₂ (sheet QuanHe, '
          'bước 0,1 m) và vẽ hai đường q ~ f₁, q ~ f₂ (sheet BieuDoPhuTro) cho từng trường hợp.'),
    ('-', 'Mỗi giờ: (1) Qtb = (Q₁ + Q₂)/2; (2) từ q₁ (cuối giờ trước) tra đường q ~ f₁ được f₁; (3) f₂ = f₁ + Qtb; (4) từ f₂ tra '
          'đường q ~ f₂ được q₂, đồng thời Z₂, V₂; (5) sang giờ sau với q₁ = q₂.'),
    ('-', 'Tra bảng trong Excel: k = MATCH(giá trị; cột tra; 1) – dòng có giá trị ≤ giá trị cần tra; w = tỷ lệ nội suy giữa dòng k '
          'và k + 1 (cột phụ màu xám). Z₂, q₂, V₂ nội suy với CÙNG k, w nên V₂/Δt + q₂/2 = f₂ đúng tuyệt đối (cột kiểm tra L).'),
    ('-', 'Giờ 0: Z = MNDBT, q = q(MNDBT) khi mở hết cửa (VH_mo = 1) hoặc q = Q (VH_mo = 2). Dòng “Giữ MNDBT” (cột Q, nền xám): '
          'q₂ = Q₂, Z₂ = MNDBT, nước qua tràn sự cố mở một phần. Giờ đầu tiên kéo hết cửa: trạng thái đầu giờ chưa nằm trên đường '
          'q ~ f₁ nên f₁ = V₁/Δt − q₁/2 tính trực tiếp. Kết quả: Zmax = MNLTK (P = 1%) hoặc MNLKT (P = 0,2%), lưu lượng xả lớn '
          'nhất, phân chia cho từng tràn; kiểm tra W đến − W xả = ΔV (≈ 0).'),
    ('M', '4. Trình tự trong file'),
    ('-', 'SoLieu: thông số hai tràn, mực nước trước lũ, cách vận hành trước đỉnh lũ, đường quá trình lũ 72 giờ của 8 tần suất '
          '(QTVH 2026) kèm tổng lượng lũ ghi trong QTVH và tính lại để kiểm tra số liệu nhập, bảng Z ~ F ~ W 2022.'),
    ('-', 'QuanHe: bảng Z ~ V ~ q (tràn chính 9 / 8 khoang, tràn sự cố) ~ f₁ ~ f₂ cho hai trường hợp (101 dòng, +71,00 ÷ +81,00 m).'),
    ('-', 'BieuDoPhuTro: hai đường q ~ f₁, q ~ f₂ (bình thường, sự cố) và ví dụ tra một giờ.'),
    ('-', 'P1_BinhThuong, P1_SuCo, P02_BinhThuong, P02_SuCo: bốn bảng điều tiết giống nhau, chỉ khác ô tần suất (C3 – chọn '
          'được 0,1 ÷ 20%) và trường hợp (F3); mỗi sheet có biểu đồ đường quá trình lũ đến và xả.'),
    ('-', 'BangIn: bảng quan hệ phụ trợ để in vào thuyết minh (chọn Z đầu, cuối, bước). KetQua: tổng hợp 4 trường hợp và biểu đồ.'),
    ('M', '5. Chú giải màu ô'),
    ('-', 'Nền vàng chữ xanh: ô được sửa. Chữ đen: công thức. Nền xanh lá: f₁, f₂. Nền đỏ nhạt: dòng có mực nước lớn nhất. '
          'Nền xám: giờ giữ MNDBT (VH_mo = 2).'),
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
ghi(sl, 'A1', 'SỐ LIỆU ĐẦU VÀO – ĐIỀU TIẾT LŨ (PÔ-TA-PỐP, HAI ĐƯỜNG f₁, f₂)', f_tieu_de, border=False)
ghi(sl, 'A2', 'Ô nền vàng chữ xanh: được sửa. Chữ đen: công thức.', f_ghi_chu, border=False)


def khoi_thong_so(h0, tieu_de, dong):
    ghi(sl, 'A%d' % h0, tieu_de, f_muc, border=False)
    tieu_de_cot(sl, h0 + 1, 1, ['Thông số', 'Giá trị', 'Đơn vị'])
    ghi(sl, 'D%d' % (h0 + 1), 'Ghi chú, nguồn', f_trang, fill=nen_tieu_de, can=giua)
    sl.merge_cells('D%d:H%d' % (h0 + 1, h0 + 1))
    for k, (nhan, gt, dv, ten_o, gc, fmt) in enumerate(dong):
        r = h0 + 2 + k
        nhap = not (isinstance(gt, str) and gt.startswith('='))
        ghi(sl, 'A%d' % r, nhan, f_dam)
        ghi(sl, 'B%d' % r, gt, f_nhap_dam if nhap else f_dam, fmt, nen_nhap if nhap else nen_kq, cg)
        ghi(sl, 'C%d' % r, dv, can=cg)
        ghi(sl, 'D%d' % r, gc, f_ghi_chu)
        sl.merge_cells('D%d:H%d' % (r, r))
        if ten_o:
            ten(ten_o, 'SoLieu!$B$%d' % r)
    return h0 + 2 + len(dong)


h = khoi_thong_so(4, 'A. THÔNG SỐ CHUNG', [
    ('Mực nước hồ khi lũ đến (MNDBT)', 74.70, 'm', 'Z0', 'Giả thiết bất lợi: lũ đến khi hồ đầy ở MNDBT', '0.00'),
    ('Bước thời gian Δt', 1, 'giờ', 'dt_h', 'Bằng bước của đường quá trình lũ (cố định 1 giờ)', '0'),
    ('Δt tính bằng giây', '=dt_h*3600', 's', 'dt_s', '', '0'),
    ('Gia tốc trọng trường g', 9.81, 'm/s²', 'g_', '', '0.00'),
    ('Vận hành trước đỉnh lũ (1 hoặc 2)', 1, '', 'VH_mo', '1 – mở hết cửa từ giờ 0 (như yêu cầu); 2 – giữ MNDBT, xả bằng lũ '
     'đến tới khi lũ đến > khả năng xả khi mở hết cửa', '0'),
])
dv_vh = DataValidation(type='list', formula1='"1,2"', allow_blank=False)
sl.add_data_validation(dv_vh)
dv_vh.add('B%d' % (h - 1))
h = khoi_thong_so(h + 1, 'B. TRÀN CHÍNH (9 khoang có cửa van – mở hoàn toàn khi đón lũ)', [
    ('Cao trình ngưỡng', 74.70, 'm', 'TC_Z', 'Không thay đổi', '0.00'),
    ('Số khoang n', 9, 'khoang', 'TC_n', 'Tràn ban đầu 3 khoang, nay chia mỗi khoang thành 3 (QTVH 2026)', '0'),
    ('Bề rộng một khoang b', 2.40, 'm', 'TC_b', 'Σb = 9 × 2,4 = 21,6 m', '0.00'),
    ('Số trụ giữa n₁', 6, 'trụ', 'TC_n1', '2 trụ mới trong mỗi khoang của tràn ban đầu', '0'),
    ('Bề dày trụ giữa d₁', 0.60, 'm', 'TC_d1', '', '0.00'),
    ('Hệ số hình dạng trụ giữa ξ₁', 0.45, '', 'TC_x1', 'Đầu trụ lượn tròn', '0.00'),
    ('Số trụ chính ban đầu n₂', 2, 'trụ', 'TC_n2', 'Trụ của tràn 3 khoang ban đầu', '0'),
    ('Bề dày trụ chính d₂', 1.00, 'm', 'TC_d2', '', '0.00'),
    ('Hệ số hình dạng trụ chính ξ₂', 0.45, '', 'TC_x2', 'Đầu trụ lượn tròn như trụ giữa; đầu trụ hình khác thì tra lại', '0.00'),
    ('Kiểm tra: n₁ + n₂ = n − 1', '=IF(TC_n1+TC_n2=TC_n-1,"Đúng","Sai – sửa số trụ")', '', None,
     '9 khoang có 8 trụ nằm giữa các khoang', None),
    ('Tổng bề rộng tràn kể trụ pin', '=TC_n*TC_b+TC_n1*TC_d1+TC_n2*TC_d2', 'm', 'TC_Btong',
     '21,6 + 6 × 0,6 + 2 × 1,0 = 27,2 m; bề dày trụ không có trong công thức lưu lượng', '0.00'),
    ('Hệ số lưu lượng m', 0.45, '', 'TC_m', 'Như thuyết minh (mục 2.3.2)', '0.000'),
    ('Hệ số hình dạng mố bên ξmb', 0.70, '', 'TC_xmb', 'Mố bên lượn tròn', '0.00'),
    ('Σξ trụ khi 9 khoang làm việc', '=TC_n1*TC_x1+TC_n2*TC_x2', '', 'TC_sx', 'n₁·ξ₁ + n₂·ξ₂ = 6 × 0,45 + 2 × 0,45', '0.00'),
    ('Hệ số chảy ngập σn', 1.0, '', 'TC_s', 'Chảy không ngập', '0.00'),
    ('Số cửa không mở được khi sự cố', 1, 'cửa', 'TC_hong', 'Trường hợp sự cố – coi cửa hỏng ở khoang biên', '0'),
    ('Số khoang làm việc khi sự cố', '=TC_n-TC_hong', 'khoang', 'TC_n_sc', '', '0'),
    ('Σξ trụ khi sự cố', '=TC_sx-TC_hong*TC_x1', '', 'TC_sx_sc',
     'Trụ giữa 0,6 m cạnh khoang biên bị đóng không còn nằm giữa hai khoang làm việc: 5 × 0,45 + 2 × 0,45', '0.00'),
])
h = khoi_thong_so(h + 1, 'C. TRÀN SỰ CỐ (1 khoang có cửa van – mở hoàn toàn khi đón lũ)', [
    ('Cao trình ngưỡng', 71.00, 'm', 'SC_Z', 'KTTV Bảng 1-2', '0.00'),
    ('Số khoang n', 1, 'khoang', 'SC_n', '', '0'),
    ('Bề rộng khoang b', 5.00, 'm', 'SC_b', '', '0.00'),
    ('Hệ số lưu lượng m', 0.45, '', 'SC_m', 'Như tràn chính', '0.000'),
    ('Hệ số hình dạng mố bên ξmb', 0.70, '', 'SC_xmb', '', '0.00'),
    ('Hệ số hình dạng trụ pin ξmt', 0.45, '', 'SC_xmt', 'Không dùng khi n = 1', '0.00'),
    ('Hệ số chảy ngập σn', 1.0, '', 'SC_s', 'Chảy không ngập', '0.00'),
])
# D. đường quá trình lũ
HL = h + 1
ghi(sl, 'A%d' % HL, 'D. ĐƯỜNG QUÁ TRÌNH LŨ ĐẾN Q (m³/s) – QTVH 2026: thu phóng trận lũ trạm Đá Bàn, 72 giờ (QTVH Bảng 3-22, '
                    '3-23)', f_muc, border=False)
ghi(sl, 'A%d' % (HL + 1), 'Tần suất P (%) →', f_dam, can=Alignment(horizontal='right'))
CL = get_column_letter(1 + len(P_LU))                    # cột tần suất cuối
for k, p in enumerate(P_LU):
    ghi(sl, '%s%d' % (get_column_letter(2 + k), HL + 1), p, f_nhap_dam, 'General', nen_nhap, cg)
tieu_de_cot(sl, HL + 2, 1, ['t (giờ)'] + ['P = %s%%' % ('%g' % p).replace('.', ',') for p in P_LU])
for t in range(0, T_MAX + 1):
    r = HL + 3 + t
    ghi(sl, 'A%d' % r, t, f_nhap, '0', can=cg)
    for k, p in enumerate(P_LU):
        ghi(sl, '%s%d' % (get_column_letter(2 + k), r), Q_LU[p][max(t, 1) - 1], f_nhap, '0.0', nen_nhap, cg)
RL0, RL1 = HL + 3, HL + 3 + T_MAX
ten('Lu_t', 'SoLieu!$A$%d:$A$%d' % (RL0, RL1))
ten('Lu_Q', 'SoLieu!$B$%d:$%s$%d' % (RL0, CL, RL1))
ten('Lu_P', 'SoLieu!$B$%d:$%s$%d' % (HL + 1, CL, HL + 1))
ghi(sl, '%s%d' % (get_column_letter(3 + len(P_LU)), RL0), 'Giờ 0 lấy bằng giờ 1 (bảng QTVH bắt đầu từ giờ 1).', f_ghi_chu,
    border=False)
for k, (nhan, fmt) in enumerate((('W ghi trong QTVH Bảng 3-22 (10⁶ m³)', '0.00'), ('W tính từ đường quá trình (10⁶ m³)', '0.00'),
                                 ('Chênh lệch W tính so với W ghi (%)', '0.0'))):
    ghi(sl, 'A%d' % (RL1 + 1 + k), nhan, f_dam)
for k, p in enumerate(P_LU):
    c = get_column_letter(2 + k)
    ghi(sl, '%s%d' % (c, RL1 + 1), W_LU_QTVH[k], f_nhap, '0.00', nen_nhap, cg)
    ghi(sl, '%s%d' % (c, RL1 + 2), '=SUMPRODUCT(({c}{a}:{c}{b}+{c}{a1}:{c}{b1})/2)*dt_s/1000000'
        .format(c=c, a=RL0, b=RL1 - 1, a1=RL0 + 1, b1=RL1), f_dam, '0.00', nen_kq, cg)
    ghi(sl, '%s%d' % (c, RL1 + 3), '=({c}{w}/{c}{g}-1)*100'.format(c=c, w=RL1 + 2, g=RL1 + 1), fmt='0.0', can=cg)
# E. Z ~ F ~ W
HZ = RL1 + 6
ghi(sl, 'A%d' % HZ, 'E. QUAN HỆ Z ~ F ~ W – BÌNH ĐỒ 2022 (QTVH Bảng 1-2)', f_muc, border=False)
tieu_de_cot(sl, HZ + 1, 1, ['Z (m)', 'F (km²)', 'W (10⁶ m³)'])
for k, (z, f) in enumerate(zip(S.Z_TAB, S.F_TAB)):
    r = HZ + 2 + k
    ghi(sl, 'A%d' % r, float(z), f_nhap, '0.00', can=cg)
    ghi(sl, 'B%d' % r, float(f), f_nhap, '0.00', can=cg)
    ghi(sl, 'C%d' % r, 0 if k == 0 else '=C{p}+(A{h}-A{p})*(B{p}+B{h})/2'.format(p=r - 1, h=r), fmt='0.000', can=cg)
RZ1 = HZ + 1 + len(S.Z_TAB)
ghi(sl, 'D%d' % (HZ + 2), 'W cộng dồn từ F (F biến đổi tuyến tính giữa hai cao trình) – V(+74,70) = 36,41·10⁶ m³ như '
                         'thuyết minh.', f_ghi_chu, border=False)
ten('Bang_Z', 'SoLieu!$A$%d:$A$%d' % (HZ + 2, RZ1))
ten('Bang_F', 'SoLieu!$B$%d:$B$%d' % (HZ + 2, RZ1))
ten('Bang_W', 'SoLieu!$C$%d:$C$%d' % (HZ + 2, RZ1))
sl.freeze_panes = 'A4'
in_an(sl)

# ================================================================== QUAN HỆ Z ~ V ~ q ~ f₁ ~ f₂
qh = wb.create_sheet('QuanHe')
RONG = (8, 8, 9, 10, 7, 8, 8, 9, 9, 7, 7, 9, 9, 9, 10, 10, 10, 10)
for k, w in enumerate(RONG):
    qh.column_dimensions[get_column_letter(k + 1)].width = w
ghi(qh, 'A1', 'BẢNG QUAN HỆ PHỤ TRỢ Z ~ V ~ q ~ f₁ ~ f₂ (bước 0,1 m) – dùng để tra trong các sheet điều tiết', f_tieu_de,
    border=False)
ghi(qh, 'A2', 'f₁ = V/Δt − q/2;  f₂ = V/Δt + q/2 (m³/s). BT – vận hành bình thường (tràn chính 9 khoang); SC – sự cố '
              '(tràn chính 8 khoang). q của hồ = q tràn chính + q tràn sự cố tại CÙNG mực nước Z.', f_ghi_chu, border=False)
tieu_de_cot(qh, 4, 1, ['Z (m)', 'F (km²)', 'V (10⁶ m³)', 'V/Δt (m³/s)', 'H tràn chính (m)', 'ε tràn chính 9 khoang',
                       'ε tràn chính 8 khoang', 'q tràn chính 9 khoang (m³/s)', 'q tràn chính 8 khoang (m³/s)',
                       'H tràn sự cố (m)', 'ε tràn sự cố', 'q tràn sự cố (m³/s)', 'q BT (m³/s)', 'q SC (m³/s)',
                       'f₁ BT (m³/s)', 'f₁ SC (m³/s)', 'f₂ BT (m³/s)', 'f₂ SC (m³/s)'], cao=58)
for i in range(N_QH):
    r = 5 + i
    k = 'MATCH(A{r},Bang_Z,1)'.format(r=r)
    ct = {
        'A': Z_DAU if i == 0 else '=ROUND(A%d+%s,2)' % (r - 1, BUOC),
        'B': '=INDEX(Bang_F,{k})+(A{r}-INDEX(Bang_Z,{k}))*(INDEX(Bang_F,{k}+1)-INDEX(Bang_F,{k}))'.format(k=k, r=r),
        'C': '=INDEX(Bang_W,{k})+(A{r}-INDEX(Bang_Z,{k}))*(INDEX(Bang_F,{k})+B{r})/2'.format(k=k, r=r),
        'D': '=C%d*1000000/dt_s' % r,
        'E': '=MAX(A%d-TC_Z,0)' % r,
        'F': '=1-0.2*(TC_xmb+TC_sx)*E%d/(TC_n*TC_b)' % r,
        'G': '=1-0.2*(TC_xmb+TC_sx_sc)*E%d/(TC_n_sc*TC_b)' % r,
        'H': '=TC_s*F{r}*TC_m*TC_n*TC_b*SQRT(2*g_)*E{r}^1.5'.format(r=r),
        'I': '=TC_s*G{r}*TC_m*TC_n_sc*TC_b*SQRT(2*g_)*E{r}^1.5'.format(r=r),
        'J': '=MAX(A%d-SC_Z,0)' % r,
        'K': '=1-0.2*(SC_xmb+(SC_n-1)*SC_xmt)*J%d/(SC_n*SC_b)' % r,
        'L': '=SC_s*K{r}*SC_m*SC_n*SC_b*SQRT(2*g_)*J{r}^1.5'.format(r=r),
        'M': '=H%d+L%d' % (r, r),
        'N': '=I%d+L%d' % (r, r),
        'O': '=D%d-M%d/2' % (r, r),
        'P': '=D%d-N%d/2' % (r, r),
        'Q': '=D%d+M%d/2' % (r, r),
        'R': '=D%d+N%d/2' % (r, r),
    }
    fmts = dict(A='0.00', B='0.000', C='0.000', D='0.0', E='0.00', F='0.000', G='0.000', H='0.0', I='0.0', J='0.00',
                K='0.000', L='0.0', M='0.0', N='0.0', O='0.0', P='0.0', Q='0.0', R='0.0')
    for col, f in ct.items():
        c = qh['%s%d' % (col, r)]
        c.value = f
        c.font = f_dam if col in 'AMN' else f_thuong
        c.number_format = fmts[col]
        c.alignment = cg
        c.fill = nen_tc if col in 'EFGHI' else (nen_sc if col in 'JKL' else (nen_f if col in 'OPQR' else PatternFill()))
R_QH = 4 + N_QH
ten('QH_Z', 'QuanHe!$A$5:$A$%d' % R_QH)
ten('QH_V', 'QuanHe!$C$5:$C$%d' % R_QH)
ten('QH_qtc', 'QuanHe!$H$5:$I$%d' % R_QH)
ten('QH_qsc', 'QuanHe!$L$5:$L$%d' % R_QH)
ten('QH_q', 'QuanHe!$M$5:$N$%d' % R_QH)
ten('QH_f1', 'QuanHe!$O$5:$P$%d' % R_QH)
ten('QH_f2', 'QuanHe!$Q$5:$R$%d' % R_QH)
qh.freeze_panes = 'B5'
in_an(qh)

# ================================================================== BIỂU ĐỒ PHỤ TRỢ
bp = wb.create_sheet('BieuDoPhuTro')
for k, w in enumerate((30, 13, 13, 13, 13)):
    bp.column_dimensions[get_column_letter(k + 1)].width = w
ghi(bp, 'A1', 'BIỂU ĐỒ PHỤ TRỢ PÔ-TA-PỐP: HAI ĐƯỜNG q ~ f₁ VÀ q ~ f₂', f_tieu_de, border=False)
for k, t in enumerate([
        'Cách dùng mỗi giờ: (1) từ q₁ trên trục đứng gióng ngang sang đường q ~ f₁, đọc f₁ trên trục ngang; (2) f₂ = f₁ + Qtb;',
        '(3) từ f₂ trên trục ngang gióng thẳng lên đường q ~ f₂, đọc q₂ trên trục đứng. Hai đường cách nhau đúng bằng q theo '
        'phương ngang (f₂ − f₁ = q).',
        'Số liệu vẽ lấy từ sheet QuanHe, mực nước +%s ÷ +%s m; muốn đổi khoảng vẽ: bấm chuột phải vào biểu đồ → Select Data.'
        % (('%.2f' % Z_BD0).replace('.', ','), ('%.2f' % Z_BD1).replace('.', ','))]):
    ghi(bp, 'A%d' % (2 + k), t, f_ghi_chu, border=False)
rb0, rb1 = 5 + int(round((Z_BD0 - Z_DAU) / BUOC)), 5 + int(round((Z_BD1 - Z_DAU) / BUOC))


def ref(col):
    c = 'ABCDEFGHIJKLMNOPQR'.index(col) + 1
    return Reference(qh, min_col=c, min_row=rb0, max_row=rb1)


# ví dụ tra một giờ (lấy từ sheet P1_BinhThuong) và đường gióng vẽ trên biểu đồ vận hành bình thường
R0 = 16                          # dòng giờ 0 trong các sheet điều tiết
R1 = R0 + T_MAX
T_VD = 24                        # giờ đỉnh lũ đến P = 1%
ghi(bp, 'A49', 'Ví dụ tra giờ thứ %d – sheet P1_BinhThuong (số liệu tự cập nhật)' % T_VD, f_muc, border=False)
tieu_de_cot(bp, 50, 1, ['Bước tra', 'Giá trị'])
VD = [('Qtb = (Q₁ + Q₂)/2 (m³/s)', 'C'), ('q₁ – cuối giờ trước (m³/s)', 'D'), ('f₁ – tra đường q ~ f₁ với q₁ (m³/s)', 'E'),
      ('f₂ = f₁ + Qtb (m³/s)', 'F'), ('q₂ – tra đường q ~ f₂ với f₂ (m³/s)', 'G'), ('Z₂ (m)', 'H')]
for k, (nhan, col) in enumerate(VD):
    ghi(bp, 'A%d' % (51 + k), nhan, f_dam)
    ghi(bp, 'B%d' % (51 + k), '=P1_BinhThuong!%s%d' % (col, R0 + T_VD), f_dam, '0.000' if col == 'H' else '0.0', nen_kq, cg)
tieu_de_cot(bp, 50, 4, ['Điểm gióng', 'f (m³/s)', 'q (m³/s)'])
for k, (diem, x, y) in enumerate((('q₁ trên trục q', F_MIN, '=B52'), ('gặp đường f₁', '=B53', '=B52'),
                                  ('+ Qtb → f₂', '=B54', '=B52'), ('gặp đường f₂', '=B54', '=B55'),
                                  ('q₂ trên trục q', F_MIN, '=B55'))):
    ghi(bp, 'D%d' % (51 + k), diem, can=cg)
    ghi(bp, 'E%d' % (51 + k), x, fmt='0', can=cg)
    ghi(bp, 'F%d' % (51 + k), y, fmt='0.0', can=cg)
ghi(bp, 'D56', 'Đường nét đứt trên biểu đồ vận hành bình thường.', f_ghi_chu, border=False)
bp.column_dimensions['F'].width = 13
for k, (nhan, cq, c1, c2, vt) in enumerate((('Vận hành bình thường (tràn chính 9 khoang + tràn sự cố)', 'M', 'O', 'Q', 'A6'),
                                              ('Sự cố – 1 cửa tràn chính không mở (8 khoang + tràn sự cố)', 'N', 'P', 'R',
                                               'A28'))):
    day = [(ref(c1), ref(cq), 'q ~ f₁ = V/Δt − q/2', '2A78D6', False),
           (ref(c2), ref(cq), 'q ~ f₂ = V/Δt + q/2', 'EB6834', False)]
    if k == 0:
        day.append((Reference(bp, min_col=5, min_row=51, max_row=55), Reference(bp, min_col=6, min_row=51, max_row=55),
                    'Ví dụ tra giờ %d (P = 1%%)' % T_VD, '52514E', True, 'circle'))
    bieu_do(bp, 'Biểu đồ phụ trợ – ' + nhan, 'f (m³/s)', 'q (m³/s)', day,
            vt, x_lim=(F_MIN, F_MAX, 1000), y_lim=(0, Q_MAX_BD, 100), rong=24, cao=11)
in_an(bp)

# ================================================================== CÁC SHEET ĐIỀU TIẾT
DT_HDR = ['t (giờ)', 'Q đến (m³/s)', 'Qtb = (Q₁+Q₂)/2', 'q₁ (cuối giờ trước)', 'f₁ – tra đường q ~ f₁ với q₁',
          'f₂ = f₁ + Qtb', 'q₂ – tra đường q ~ f₂ với f₂', 'Z₂ (m)', 'q tràn chính (m³/s)', 'q tràn sự cố (m³/s)',
          'V₂ (10⁶ m³)', 'Kiểm tra V₂/Δt + q₂/2 (= f₂)', 'k₁ – dòng tra q₁', 'w₁', 'k – dòng tra f₂', 'w', 'Trạng thái cửa van']
nen_giu = PatternFill('solid', fgColor='ECECEC')
CA = [('P1_BinhThuong', 1, 'Bình thường'), ('P1_SuCo', 1, 'Sự cố'), ('P02_BinhThuong', 0.2, 'Bình thường'),
      ('P02_SuCo', 0.2, 'Sự cố')]


def nd(bang, kc, wc, r, cot=True):
    """Nội suy cột 'bang' giữa dòng k (cột kc) và k+1 với tỷ lệ w (cột wc); cot: bảng 2 cột chọn theo $F$4."""
    c = ',$F$4' if cot else ''
    return 'INDEX({b},{k}{r}{c})+{w}{r}*(INDEX({b},{k}{r}+1{c})-INDEX({b},{k}{r}{c}))'.format(b=bang, k=kc, w=wc, r=r, c=c)


def sheet_dieu_tiet(ten_ws, P, th):
    ws = wb.create_sheet(ten_ws)
    for k, w in enumerate((7, 10, 11, 11, 12, 11, 12, 10, 11, 11, 11, 12, 8, 8, 8, 8, 11)):
        ws.column_dimensions[get_column_letter(k + 1)].width = w
    ghi(ws, 'A1', 'ĐIỀU TIẾT LŨ – PHƯƠNG PHÁP PÔ-TA-PỐP, HAI ĐƯỜNG q ~ f₁, q ~ f₂ (Δt = 1 giờ)', f_tieu_de, border=False)
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
    ghi(ws, 'D4', 'Cột bảng tra', f_ghi_chu)
    ws.merge_cells('D4:E4')
    ghi(ws, 'F4', '=IF(F3="Sự cố",2,1)', fmt='0', can=cg)
    ghi(ws, 'H3', '← sửa ô vàng: P (%) = 0,1 ÷ 20 như bảng lũ ở SoLieu; trường hợp “Bình thường” hoặc “Sự cố” (1 cửa tràn chính '
                  'không mở)', f_ghi_chu, border=False)
    dv1 = DataValidation(type='list', formula1='"%s"' % ','.join('%g' % p for p in P_LU), allow_blank=False)
    dv2 = DataValidation(type='list', formula1='"Bình thường,Sự cố"', allow_blank=False)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add('C3')
    dv2.add('F3')
    rg = lambda c: '%s%d:%s%d' % (c, R0, c, R1)
    vt = 'MATCH(MAX({h}),{h},0)'.format(h=rg('H'))
    TH = [
        (6, 'Mực nước lớn nhất Zmax (m)', '=MAX(%s)' % rg('H'), '0.00'),
        (7, 'Thời điểm xuất hiện (giờ)', '=INDEX(%s,%s)' % (rg('A'), vt), '0'),
        (8, 'Lưu lượng xả lớn nhất qmax (m³/s)', '=MAX(%s)' % rg('G'), '0.0'),
        (9, 'Lúc Zmax: q tràn chính / q tràn sự cố (m³/s)', '=INDEX(%s,%s)' % (rg('I'), vt), '0.0'),
        (10, 'Cột nước trên ngưỡng lúc Zmax: tràn chính / tràn sự cố (m)', '=F6-TC_Z', '0.00'),
        (11, 'Dung tích phòng lũ V(Zmax) − V(MNDBT) (10⁶ m³)', '=INDEX(%s,%s)-K%d' % (rg('K'), vt, R0), '0.000'),
        (12, 'Cân bằng nước: W đến − W xả − ΔV (10⁶ m³, phải ≈ 0)',
         '=SUMPRODUCT((B{a}:B{b}+B{a1}:B{b1})/2)*dt_s/1000000-SUMPRODUCT((G{a}:G{b}+G{a1}:G{b1})/2)*dt_s/1000000-(K{b1}-K{a})'
         .format(a=R0, b=R1 - 1, a1=R0 + 1, b1=R1), '0.0000'),
    ]
    for r, nhan, ct, fmt in TH:
        ghi(ws, 'A%d' % r, nhan, f_dam)
        ws.merge_cells('A%d:E%d' % (r, r))
        ghi(ws, 'F%d' % r, ct, f_dam, fmt, nen_kq, cg)
    ghi(ws, 'G9', '=INDEX(%s,%s)' % (rg('J'), vt), f_dam, '0.0', nen_kq, cg)
    ghi(ws, 'G10', '=F6-SC_Z', f_dam, '0.00', nen_kq, cg)
    ghi(ws, 'A13', 'Khả năng xả khi mở hết cửa ở mực nước trước lũ q(Z0) (m³/s)', f_dam)
    ws.merge_cells('A13:E13')
    ghi(ws, 'F13', '=' + nd('QH_q', 'O', 'P', R0), f_dam, '0.0', nen_kq, cg)
    ghi(ws, 'G13', '=IF(VH_mo=1,"VH_mo = 1: mở hết cửa từ giờ 0","VH_mo = 2: giữ MNDBT tới khi Q > q(Z0)")', f_ghi_chu,
        border=False)
    tieu_de_cot(ws, 15, 1, DT_HDR, cao=58)
    ws['M15'].comment = Comment('k₁ = MATCH(q₁; cột q của QuanHe; 1): dòng có q(k₁) ≤ q₁ < q(k₁+1); w₁ = tỷ lệ nội suy.\n'
                                'f₁ = f₁(k₁) + w₁ × [f₁(k₁+1) − f₁(k₁)] – tức tra đường q ~ f₁ với q₁.\n'
                                'Cột phụ – không cần đưa vào thuyết minh.', 'Hướng dẫn', 110, 300)
    ws['O15'].comment = Comment('k = MATCH(f₂; cột f₂ của QuanHe; 1): dòng có f₂(k) ≤ f₂ < f₂(k+1); w = tỷ lệ nội suy.\n'
                                'q₂, Z₂, V₂ = giá trị dòng k + w × (dòng k+1 − dòng k) – tức tra đường q ~ f₂ với f₂.\n'
                                'Cột phụ – không cần đưa vào thuyết minh.', 'Hướng dẫn', 110, 300)
    for t in range(0, T_MAX + 1):
        r = R0 + t
        p = r - 1
        giu = 'Q{r}="Giữ MNDBT"'.format(r=r)              # giờ giữ MNDBT: xả bằng lũ đến
        ct = {
            'A': t,
            'B': '=IFERROR(INDEX(Lu_Q,MATCH(A{r},Lu_t,0),$C$4),0)'.format(r=r),
            'G': '=IF({g},B{r},{f})'.format(g=giu, r=r, f=nd('QH_q', 'O', 'P', r)),
            'I': '=IF({g},0,{f})'.format(g=giu, f=nd('QH_qtc', 'O', 'P', r)),
            'J': '=IF({g},B{r},{f})'.format(g=giu, r=r, f=nd('QH_qsc', 'O', 'P', r, cot=False)),
        }
        if t == 0:
            ct.update({'C': '', 'D': '', 'E': '', 'F': '', 'M': '', 'N': '',
                       'H': '=' + nd('QH_Z', 'O', 'P', r, cot=False),
                       'K': '=' + nd('QH_V', 'O', 'P', r, cot=False),
                       'L': '=K{r}*1000000/dt_s+G{r}/2'.format(r=r),
                       'O': '=MATCH(Z0,QH_Z,1)',
                       'P': '=(Z0-INDEX(QH_Z,O{r}))/(INDEX(QH_Z,O{r}+1)-INDEX(QH_Z,O{r}))'.format(r=r),
                       'Q': '=IF(VH_mo=1,"Mở hết","Giữ MNDBT")'})
        else:
            ct.update({
                'C': '=(B%d+B%d)/2' % (p, r),
                'D': '=G%d' % p,
                # giờ đầu kéo hết cửa: trạng thái đầu giờ (MNDBT, q₁ = Q₁) không nằm trên đường q ~ f₁ → tính trực tiếp
                'E': '=IF({g},"",IF(Q{p}="Giữ MNDBT",K{p}*1000000/dt_s-D{r}/2,{f}))'
                     .format(g=giu, p=p, r=r, f=nd('QH_f1', 'M', 'N', r)),
                'F': '=IF({g},"",E{r}+C{r})'.format(g=giu, r=r),
                'H': '=IF({g},Z0,{f})'.format(g=giu, f=nd('QH_Z', 'O', 'P', r, cot=False)),
                'K': '=IF({g},$K${r0},{f})'.format(g=giu, r0=R0, f=nd('QH_V', 'O', 'P', r, cot=False)),
                'L': '=IF({g},"",K{r}*1000000/dt_s+G{r}/2)'.format(g=giu, r=r),
                'M': '=IF(OR({g},Q{p}="Giữ MNDBT"),"",MATCH(D{r},INDEX(QH_q,0,$F$4),1))'.format(g=giu, p=p, r=r),
                'N': '=IF(M{r}="","",(D{r}-INDEX(QH_q,M{r},$F$4))/(INDEX(QH_q,M{r}+1,$F$4)-INDEX(QH_q,M{r},$F$4)))'
                     .format(r=r),
                'O': '=IF({g},"",MATCH(F{r},INDEX(QH_f2,0,$F$4),1))'.format(g=giu, r=r),
                'P': '=IF(O{r}="","",(F{r}-INDEX(QH_f2,O{r},$F$4))/(INDEX(QH_f2,O{r}+1,$F$4)-INDEX(QH_f2,O{r},$F$4)))'
                     .format(r=r),
                'Q': '=IF(OR(VH_mo=1,Q{p}="Mở hết",B{r}>$F$13),"Mở hết","Giữ MNDBT")'.format(p=p, r=r),
            })
        fm = dict(A='0', B='0.0', C='0.0', D='0.0', E='0.0', F='0.0', G='0.0', H='0.000', I='0.0', J='0.0', K='0.000',
                  L='0.0', M='0', N='0.000', O='0', P='0.000', Q='General')
        for col, f in ct.items():
            c = ws['%s%d' % (col, r)]
            c.value = f
            c.font = f_dam if col in 'AGH' else (f_phu if col in 'MNOP' else f_thuong)
            c.number_format = fm[col]
            c.alignment = cg
            c.border = vien
            if col in 'EF':
                c.fill = nen_f
            elif col == 'I':
                c.fill = nen_tc
            elif col == 'J':
                c.fill = nen_sc
    ws.conditional_formatting.add('A%d:Q%d' % (R0, R1), FormulaRule(formula=['$H%d=MAX($H$%d:$H$%d)' % (R0, R0, R1)],
                                                                   fill=nen_dinh))
    ws.conditional_formatting.add('A%d:Q%d' % (R0, R1), FormulaRule(formula=['$Q%d="Giữ MNDBT"' % R0], fill=nen_giu))
    ghi(ws, 'H5', 'Giờ 0: Z = MNDBT. Giờ “Mở hết”: q₁ → tra đường q ~ f₁ được f₁ → f₂ = f₁ + Qtb → tra đường q ~ f₂ được q₂, '
                  'Z₂, V₂. Giờ “Giữ MNDBT” (chỉ khi VH_mo = 2): q₂ = Q₂, Z₂ = MNDBT.', f_ghi_chu, border=False)
    ghi(ws, 'H6', 'Cột M, N: vị trí tra q₁; cột O, P: vị trí tra f₂ trong bảng QuanHe. Cột L kiểm tra V₂/Δt + q₂/2 = f₂.',
        f_ghi_chu, border=False)
    x = Reference(ws, min_col=1, min_row=R0, max_row=R1)
    bieu_do(ws, 'Đường quá trình lũ đến và xả', 't (giờ)', 'Lưu lượng (m³/s)',
            [(x, Reference(ws, min_col=2, min_row=R0, max_row=R1), 'Q đến', '52514E', False),
             (x, Reference(ws, min_col=7, min_row=R0, max_row=R1), 'q xả (tổng)', '2A78D6', False),
             (x, Reference(ws, min_col=9, min_row=R0, max_row=R1), 'q tràn chính', '2A78D6', True),
             (x, Reference(ws, min_col=10, min_row=R0, max_row=R1), 'q tràn sự cố', '1BAF7A', True)],
            'S3', x_lim=(0, T_MAX, 12), y_lim=(0, Q_LU_MAX, 200), rong=18, cao=10)
    ws.freeze_panes = 'B16'
    in_an(ws)
    return ws


for ten_ws, P, th in CA:
    sheet_dieu_tiet(ten_ws, P, th)

# ================================================================== BẢNG IN THUYẾT MINH
bi = wb.create_sheet('BangIn')
for k, w in enumerate((5, 8, 10, 10, 10, 10, 10, 10, 10, 11, 11, 11, 11)):
    bi.column_dimensions[get_column_letter(k + 1)].width = w
ghi(bi, 'A1', 'BẢNG QUAN HỆ PHỤ TRỢ PÔ-TA-PỐP (q ~ f₁, q ~ f₂) – DÙNG ĐỂ IN VÀO THUYẾT MINH', f_tieu_de, border=False)
ghi(bi, 'A2', 'Chọn mực nước đầu, cuối, bước ở ô vàng; bảng tự rút gọn. Chọn vùng có số liệu → Copy → Word: dán “Giữ văn bản” '
              '(Keep Text Only) rồi chuyển thành bảng, hoặc Excel khác: Paste Special → Values.', f_ghi_chu, border=False)
for c1, c2, nhan, gt, nm, fmt in (('A', 'C', 'Z đầu (m)', 74.70, 'Z_dau_in', '0.00'),
                                  ('D', 'F', 'Z cuối (m)', 79.00, 'Z_cuoi_in', '0.00'),
                                  ('G', 'I', 'Bước (m)', 0.1, 'Buoc_in', '0.0')):
    ghi(bi, c1 + '4', nhan, f_dam)
    ghi(bi, get_column_letter(bi[c1 + '4'].column + 1) + '4', None)
    bi.merge_cells('%s4:%s4' % (c1, get_column_letter(bi[c1 + '4'].column + 1)))
    ghi(bi, c2 + '4', gt, f_nhap_dam, fmt, nen_nhap, cg)
    ten(nm, 'BangIn!$%s$4' % c2)
dv_b = DataValidation(type='list', formula1='"0.1,0.2,0.5"', allow_blank=False)
bi.add_data_validation(dv_b)
dv_b.add('I4')
tieu_de_cot(bi, 6, 1, ['TT', 'Z (m)', 'q tràn chính 9 khoang', 'q tràn chính 8 khoang', 'q tràn sự cố', 'q BT', 'q SC',
                       'V (10⁶ m³)', 'V/Δt (m³/s)', 'f₁ BT = V/Δt − q/2', 'f₂ BT = V/Δt + q/2', 'f₁ SC', 'f₂ SC'], cao=45)
for i in range(N_QH):
    r = 7 + i
    B = 'B%d' % r
    k = 'MATCH({B},QH_Z,1)'.format(B=B)
    ct = {
        'A': '=IF({B}="","",ROW()-6)'.format(B=B),
        'B': '=Z_dau_in' if i == 0 else
             '=IF(B{p}="","",IF(B{p}+Buoc_in>Z_cuoi_in+0.000001,"",ROUND(B{p}+Buoc_in,2)))'.format(p=r - 1),
    }
    for col, src in (('C', 'H'), ('D', 'I'), ('E', 'L'), ('F', 'M'), ('G', 'N'), ('H', 'C'), ('I', 'D'), ('J', 'O'),
                     ('K', 'Q'), ('L', 'P'), ('M', 'R')):
        ct[col] = '=IF({B}="","",INDEX(QuanHe!${s}$5:${s}${e},{k}))'.format(B=B, s=src, e=R_QH, k=k)
    fm = dict(A='0', B='0.00', C='0.0', D='0.0', E='0.0', F='0.0', G='0.0', H='0.000', I='0.0', J='0.0', K='0.0', L='0.0',
              M='0.0')
    for col, f in ct.items():
        c = bi['%s%d' % (col, r)]
        c.value = f
        c.font = f_dam if col == 'B' else f_thuong
        c.number_format = fm[col]
        c.alignment = cg
bi.conditional_formatting.add('A7:M%d' % (6 + N_QH), FormulaRule(formula=['$B7<>""'], border=vien))
bi.freeze_panes = 'C7'
in_an(bi, ngang=True)
bi.defined_names['Print_Area'] = DefinedName('Print_Area', attr_text='OFFSET(BangIn!$A$1,0,0,6+COUNT(BangIn!$B$7:$B$%d),13)'
                                             % (6 + N_QH))

# ================================================================== KẾT QUẢ
kq = wb.create_sheet('KetQua')
for k, w in enumerate((30, 22, 11, 9, 11, 12, 12, 13, 12, 34, 13, 12)):
    kq.column_dimensions[get_column_letter(k + 1)].width = w
ghi(kq, 'A1', 'KẾT QUẢ ĐIỀU TIẾT LŨ – HỒ CÀ GIÂY (Pô-ta-pốp; lũ đến QTVH 72 giờ; cửa van kéo lên hoàn toàn khi đón lũ)',
    f_tieu_de, border=False)
ghi(kq, 'A2', '="Vận hành trước đỉnh lũ: "&IF(VH_mo=1,"1 – mở hết cửa từ giờ 0","2 – giữ MNDBT tới khi lũ đến lớn hơn khả '
              'năng xả khi mở hết cửa")&" (đổi ở SoLieu!B10)."', f_ghi_chu, border=False)
tieu_de_cot(kq, 3, 1, ['Sheet', 'P (%) – trường hợp', 'Zmax (m)', 't (giờ)', 'qmax (m³/s)', 'q tràn chính lúc Zmax',
                       'q tràn sự cố lúc Zmax', 'Dung tích phòng lũ (10⁶ m³)', 'Đỉnh đập sau SC − Zmax (m)', 'Ý nghĩa',
                       'Zmax bản trước – lũ BCNCKT 36 giờ (m)', 'Chênh lệch so với bản trước (m)'], cao=58)
Y_NGHIA = {'P1_BinhThuong': 'MNLTK – lũ thiết kế, vận hành bình thường',
           'P1_SuCo': 'Lũ thiết kế, 1 cửa tràn chính không mở',
           'P02_BinhThuong': 'MNLKT – lũ kiểm tra, vận hành bình thường',
           'P02_SuCo': 'Lũ kiểm tra, 1 cửa tràn chính không mở'}
for k, (ten_ws, P, th) in enumerate(CA):
    r = 4 + k
    ghi(kq, 'A%d' % r, ten_ws, f_dam)
    ghi(kq, 'B%d' % r, "=\"P = \"&'%s'!C3&\"%% – \"&'%s'!F3" % (ten_ws, ten_ws), can=cg)
    for col, src, fmt in (('C', 'F6', '0.00'), ('D', 'F7', '0'), ('E', 'F8', '0.0'), ('F', 'F9', '0.0'),
                          ('G', 'G9', '0.0'), ('H', 'F11', '0.00')):
        ghi(kq, '%s%d' % (col, r), "='%s'!%s" % (ten_ws, src), f_dam if col == 'C' else f_thuong, fmt, nen_kq, cg)
    ghi(kq, 'I%d' % r, '=Z_dinh_SC-C%d' % r, fmt='0.00', can=cg)
    ghi(kq, 'J%d' % r, Y_NGHIA[ten_ws], f_ghi_chu)
    ghi(kq, 'K%d' % r, Z_TRUOC['%g%%' % P][th == 'Sự cố'], f_nhap, '0.00', nen_nhap, cg)
    ghi(kq, 'L%d' % r, '=C{r}-K{r}'.format(r=r), fmt='+0.00;-0.00;0.00', can=cg)
ghi(kq, 'A9', 'Cao trình đỉnh đập sau sửa chữa (m)', f_dam)
ghi(kq, 'B9', S.Z_DINH_DAP_SC, f_nhap_dam, '0.00', nen_nhap, cg)
ten('Z_dinh_SC', 'KetQua!$B$9')
ghi(kq, 'A10', 'Ảnh hưởng của sự cố (Zmax tăng thêm, m): P = 1% / P = 0,2%', f_dam)
ghi(kq, 'B10', '=C5-C4', f_dam, '0.00', nen_kq, cg)
ghi(kq, 'C10', '=C7-C6', f_dam, '0.00', nen_kq, cg)
ghi(kq, 'A11', 'Kết luận: MNLTK = Zmax (P = 1%, bình thường) =', f_dam, border=False)
ghi(kq, 'D11', '=C4', f_dam, '0.00', nen_kq, cg)
ghi(kq, 'A12', '             MNLKT = Zmax (P = 0,2%, bình thường) =', f_dam, border=False)
ghi(kq, 'D12', '=C6', f_dam, '0.00', nen_kq, cg)
ghi(kq, 'A13', 'Lũ QTVH có tổng lượng lớn hơn lũ BCNCKT 36 giờ nhưng phần lớn nằm ở 16 giờ đầu và đuôi lũ kéo dài với lưu lượng '
               'nhỏ; phần đỉnh lũ (Q > 300 m³/s) có thể tích nhỏ hơn (P = 1%: 14,8 so với 16,5 triệu m³) nên Zmax thấp hơn bản '
               'trước (cột L). Sự cố 1 cửa làm mực nước cao thêm vài cm, vẫn thấp hơn đỉnh đập sau sửa chữa.', f_ghi_chu,
    border=False)
# số liệu biểu đồ
HB = 16
ghi(kq, 'A%d' % (HB - 1), 'SỐ LIỆU VẼ BIỂU ĐỒ (lấy từ các sheet điều tiết)', f_muc, border=False)
tieu_de_cot(kq, HB, 1, ['t (giờ)', 'Q đến P = 1%', 'q xả P = 1% bình thường', 'q xả P = 1% sự cố', 'Q đến P = 0,2%',
                        'q xả P = 0,2% bình thường', 'q xả P = 0,2% sự cố', 'Z P = 1% bình thường', 'Z P = 1% sự cố',
                        'Z P = 0,2% bình thường', 'Z P = 0,2% sự cố'], cao=58)
for t in range(0, T_MAX + 1):
    r = HB + 1 + t
    src = R0 + t
    for col, f in (('A', '=P1_BinhThuong!A%d' % src), ('B', '=P1_BinhThuong!B%d' % src),
                   ('C', '=P1_BinhThuong!G%d' % src), ('D', '=P1_SuCo!G%d' % src), ('E', '=P02_BinhThuong!B%d' % src),
                   ('F', '=P02_BinhThuong!G%d' % src), ('G', '=P02_SuCo!G%d' % src), ('H', '=P1_BinhThuong!H%d' % src),
                   ('I', '=P1_SuCo!H%d' % src), ('J', '=P02_BinhThuong!H%d' % src), ('K', '=P02_SuCo!H%d' % src)):
        ghi(kq, '%s%d' % (col, r), f, fmt='0' if col == 'A' else ('0.0' if col in 'BCDEFG' else '0.00'), can=cg)
HB1 = HB + 1 + T_MAX
xk = Reference(kq, min_col=1, min_row=HB + 1, max_row=HB1)
cot = lambda c: Reference(kq, min_col=c, min_row=HB + 1, max_row=HB1)
bieu_do(kq, 'Đường quá trình lũ đến và xả – P = 1%', 't (giờ)', 'Lưu lượng (m³/s)',
        [(xk, cot(2), 'Q đến', '52514E', False), (xk, cot(3), 'q xả – bình thường', '2A78D6', False),
         (xk, cot(4), 'q xả – sự cố', 'EB6834', True)], 'M3', x_lim=(0, T_MAX, 12), y_lim=(0, Q_LU_MAX, 200))
bieu_do(kq, 'Đường quá trình lũ đến và xả – P = 0,2%', 't (giờ)', 'Lưu lượng (m³/s)',
        [(xk, cot(5), 'Q đến', '52514E', False), (xk, cot(6), 'q xả – bình thường', '2A78D6', False),
         (xk, cot(7), 'q xả – sự cố', 'EB6834', True)], 'M24', x_lim=(0, T_MAX, 12), y_lim=(0, Q_LU_MAX, 200))
bieu_do(kq, 'Mực nước hồ theo thời gian', 't (giờ)', 'Z (m)',
        [(xk, cot(8), 'P = 1% – bình thường', '2A78D6', False), (xk, cot(9), 'P = 1% – sự cố', '2A78D6', True),
         (xk, cot(10), 'P = 0,2% – bình thường', 'EB6834', False), (xk, cot(11), 'P = 0,2% – sự cố', 'EB6834', True)],
        'M45', x_lim=(0, T_MAX, 12), y_lim=(74, 79, 0.5), so=('0', '0.0'))
in_an(kq)

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'QuanHe': '8A8984', 'BieuDoPhuTro': '1BAF7A',
                                    'BangIn': '8A8984', 'KetQua': '1BAF7A'}.get(ws.title, '2A78D6')
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
