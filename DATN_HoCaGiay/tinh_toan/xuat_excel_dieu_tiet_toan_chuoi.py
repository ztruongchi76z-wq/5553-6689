# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL HƯỚNG DẪN TÍNH ĐIỀU TIẾT TOÀN CHUỖI 44 NĂM (mục 2.2 của đồ án)
Mọi kết quả trong file Excel là CÔNG THỨC (đổi số liệu nhập thì kết quả tự tính lại); thuật toán giống hệt
kt_dong_chay_toi_thieu.dieu_tiet_chuoi(): bước tháng, lặp 3 lần dung tích bình quân, tổn thất bốc hơi ΔZ·F(Z̄),
tổn thất thấm k_t·V̄, giới hạn Vc ≤ V ≤ Vbt, mức bảo đảm P = m/(n + 1).

  python3 xuat_excel_dieu_tiet_toan_chuoi.py [đường dẫn .xlsx ra]
Mặc định: ../TAI_VE/HuongDan_DieuTiet_ToanChuoi_44nam.xlsx
"""
import json
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
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'HuongDan_DieuTiet_ToanChuoi_44nam.xlsx')

D = json.load(open(os.path.join(DIR, 'du_lieu', 'can_bang_nuoc.json'), encoding='utf-8'))
NHU_CAU = {}
for v in D['nhu_cau'].values():
    ten = 'TH1' if 'TH1 ' in v['nhan'][0] else ('TH1A' if 'TH1A' in v['nhan'][0] else 'TH2')
    NHU_CAU[ten] = v
MO_TA_TH = {
    'TH1': 'Tưới 8.965 ha theo nhiệm vụ thiết kế, 100% cây trồng cạn tưới truyền thống + sinh hoạt (QTVH Bảng 4-20)',
    'TH1A': 'Tưới 8.965 ha, 50% diện tích cây trồng cạn tưới tiết kiệm + sinh hoạt (QTVH Bảng 4-21)',
    'TH2': 'Tưới tiềm năng 9.540 ha, 50% cây trồng cạn tưới tiết kiệm + sinh hoạt (QTVH Bảng 4-22)',
}
DZ_MM = [68.4, 64.4, 76.9, 66.6, 62.3, 55.1, 49.8, 51.8, 43.9, 40.8, 52.2, 66.7]   # mm [QTVH Bảng 3-24], như kt_dong_chay_toi_thieu.DZ

# ------------------------------------------------------------------ định dạng
FONT = 'Times New Roman'
f_thuong = Font(name=FONT, size=11)
f_dam = Font(name=FONT, size=11, bold=True)
f_tieu_de = Font(name=FONT, size=14, bold=True, color='12324A')
f_muc = Font(name=FONT, size=12, bold=True, color='A65A10')
f_nhap = Font(name=FONT, size=11, color='0000FF')            # số liệu nhập / số liệu gốc
f_nhap_dam = Font(name=FONT, size=11, color='0000FF', bold=True)
f_trang = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_ghi_chu = Font(name=FONT, size=10, italic=True, color='52514E')
nen_nhap = PatternFill('solid', fgColor='FFF2CC')               # ô được sửa
nen_tieu_de = PatternFill('solid', fgColor='12324A')
nen_lap = [PatternFill('solid', fgColor=c) for c in ('EEF3FA', 'F4F1FA', 'EEF7F1')]
nen_kq = PatternFill('solid', fgColor='FBF3E8')
nen_do = PatternFill('solid', fgColor='F8D7D7')
vien = Border(*(Side(style='thin', color='B8BEC6'),) * 4)
giua = Alignment(horizontal='center', vertical='center', wrap_text=True)
trai = Alignment(horizontal='left', vertical='top', wrap_text=True)

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


# Nội suy tuyến tính đúng thứ tự phép tính của numpy.interp: slope*(x − x_j) + y_j
def noi_suy(x, bx, by):
    k = 'MATCH(%s,%s,1)' % (x, bx)
    return ('((INDEX({by},{k}+1)-INDEX({by},{k}))/(INDEX({bx},{k}+1)-INDEX({bx},{k})))*({x}-INDEX({bx},{k}))'
            '+INDEX({by},{k})').format(x=x, bx=bx, by=by, k=k)


def Z_tu_V(v):
    return noi_suy(v, 'BT_V', 'BT_Z')


def F_tu_Z(z):
    return noi_suy(z, 'Bang_Z', 'Bang_F')


# ================================================================== 1. HƯỚNG DẪN
hd = wb.active
hd.title = 'HuongDan'
hd.column_dimensions['A'].width = 4
hd.column_dimensions['B'].width = 125
hd.sheet_view.showGridLines = False

NOI_DUNG = [
    ('T', 'HƯỚNG DẪN TÍNH ĐIỀU TIẾT TOÀN CHUỖI 44 NĂM – HỒ CHỨA NƯỚC CÀ GIÂY'),
    ('g', 'File đi kèm thuyết minh ĐATN mục 2.2 (Đánh giá khả năng cấp nước). Mọi ô kết quả là công thức – đổi số liệu ở sheet SoLieu thì toàn bộ file tự tính lại.'),
    ('M', '1. Bài toán và trường hợp tính'),
    ('-', 'Câu hỏi: với MNC +67,50 m, MNDBT +74,70 m (giữ như hiện có) và quan hệ Z ~ F ~ W đo năm 2022, hồ cấp đủ nước cho nhu cầu dùng nước bao nhiêu năm trong 44 năm, và mức bảo đảm có đạt 85% không?'),
    ('-', 'Trường hợp mặc định: TH1A – tưới 8.965 ha, 50% diện tích cây trồng cạn tưới tiết kiệm + cấp nước sinh hoạt (QTVH Bảng 4-21). Đây là trường hợp sát yêu cầu nhất: 39/44 năm đủ nước, P = 86,7% – vừa đạt (thiếu thêm 1 năm là không đạt).'),
    ('-', 'Có thể chọn TH1 hoặc TH2 ở ô SoLieu!B5 (TH2: 37/44 năm, 82,2% – gần đạt nhưng chưa đạt; TH1: 18/44 năm, 40,0%).'),
    ('-', 'Dung tích lấy theo bình đồ 2022: Vc = V(+67,50) = 6,82; Vbt = V(+74,70) = 36,41; Vhi = 29,60 triệu m³ (thiết kế ban đầu 8,41 / 37,24 / 28,83 chỉ để so sánh – bản đồ cũ không còn nên không lập lại được đường V(Z) cũ).'),
    ('M', '2. Vì sao tính 44 năm?'),
    ('-', '44 năm là TOÀN BỘ chuỗi dòng chảy tháng đến hồ mà QTVH 2026 đã tính (Bảng 3-13: năm 1981-1982 đến 2024-2025) – không phải con số tự chọn. Dùng hết số liệu có thì chuỗi đại diện tốt nhất cho các tổ hợp năm nhiều nước – ít nước.'),
    ('-', 'TCVN 10778:2024 mục 8.2.1.2: khi có chuỗi dòng chảy đến từ 15 năm trở lên thì tính điều tiết theo phương pháp toàn chuỗi – mô phỏng lần lượt từng tháng của cả chuỗi – thay cho phương pháp năm đại biểu.'),
    ('-', 'Toàn chuỗi giữ đúng thứ tự các năm và phân phối dòng chảy thực của từng năm: nếu có năm hồ không tích đầy thì năm sau bắt đầu với dung tích thấp (dung tích chuyển tiếp) – phương pháp một năm đại biểu không thể hiện được. Với hồ Cà Giây, cột “Hồ tích đầy trong năm?” ở sheet KetQua cho thấy năm nào hồ cũng tích đầy MNDBT (dòng chảy đến bình quân 220 triệu m³/năm, gấp khoảng 7 lần dung tích hữu ích 29,6 triệu m³ – hồ điều tiết năm), nên năm thiếu nước là do mùa khô của chính năm đó.'),
    ('-', 'Mức bảo đảm được xác định bằng cách ĐẾM NĂM: P = m/(n + 1) (TCVN 10778:2024 công thức (14)); điều kiện đạt: m ≥ (n + 1)·P_tk (công thức (9)). Với n = 44: m ≥ 45 × 0,85 = 38,25 → cần ít nhất 39 năm đủ nước, tức chỉ được thiếu tối đa 5 năm; mỗi năm thiếu làm P giảm 1/45 = 2,2%.'),
    ('-', 'Chuỗi càng ngắn kết luận càng thô: n = 15 năm thì mỗi năm thiếu làm P giảm 6,25% và chỉ được thiếu 1 năm (bảng minh họa ở sheet KetQua). Vì vậy dùng hết 44 năm cho kết quả tin cậy hơn.'),
    ('-', 'QTVH cũng tính với đúng chuỗi 44 năm này, nên kết quả của đồ án so sánh trực tiếp được với QTVH.'),
    ('-', 'Bước thời gian tháng: hồ điều tiết năm (tích nước mùa lũ tháng 9 ÷ 11, cấp nước mùa khô tháng 12 ÷ 4) nên thời đoạn tháng đủ chi tiết; 44 năm × 12 tháng = 528 bước tính (sheet DieuTiet).'),
    ('M', '3. Phương trình cân bằng nước cho tháng i (thuyết minh công thức (2-1))'),
    ('c', 'V cuối = V đầu + W đến − W dùng − W bốc hơi − W thấm − W xả,    với Vc ≤ V cuối ≤ Vbt'),
    ('-', 'W bốc hơi = ΔZ_i × F(Z̄): ΔZ_i – tổn thất bốc hơi phụ thêm tháng i (m, QTVH Bảng 3-24); F(Z̄) – diện tích mặt hồ (km²) ứng với mực nước bình quân tháng. Đơn vị: km² × m = 10⁶ m³.'),
    ('-', 'W thấm = k_t × V̄: k_t = 1%/tháng dung tích bình quân tháng (TCVN 10778:2024 Bảng 1: 0,5 ÷ 1,0%/tháng với địa chất lòng hồ tốt; đồ án lấy 1% – thiên an toàn).'),
    ('-', 'Vì W bốc hơi, W thấm phụ thuộc dung tích bình quân V̄ = (V đầu + V cuối)/2 mà V cuối chưa biết nên phải TÍNH THỬ DẦN: lần 1 giả thiết V̄ = V đầu tháng; lần 2 và 3 dùng V̄ = (V đầu + V cuối lần trước)/2. Sau 3 lần, sai khác giữa hai lần cuối lớn nhất 0,002 triệu m³ – đủ chính xác.'),
    ('-', 'Nếu V cuối > Vbt: hồ đầy, phần thừa xả qua tràn (W xả = V cuối − Vbt), lấy V cuối = Vbt.'),
    ('-', 'Nếu V cuối < Vc: hồ đã về mực nước chết, không cấp đủ nhu cầu – tháng thiếu nước, W thiếu = Vc − V cuối, lấy V cuối = Vc. Năm có ít nhất 1 tháng thiếu là năm thiếu nước.'),
    ('-', 'Đầu chuỗi (tháng 1 năm 1981-1982) giả thiết hồ đầy ở MNDBT – thời điểm cuối mùa lũ; tháng sau lấy V đầu = V cuối tháng trước, nối liên tục qua 44 năm.'),
    ('M', '4. Trình tự thao tác trong file'),
    ('-', 'Bước 1 – sheet SoLieu: kiểm tra số liệu và nguồn (Z ~ F ~ W 2022, bốc hơi, nhu cầu dùng nước, dòng chảy 44 năm); chọn trường hợp ở ô B5; Vc, Vbt tự tra theo MNC, MNDBT.'),
    ('-', 'Bước 2 – sheet BangTraZV: bảng tra chi tiết 0,01 m (F nội suy tuyến tính, V cộng dồn theo công thức chóp cụt) – dùng để đổi qua lại Z ↔ V khi tính.'),
    ('-', 'Bước 3 – sheet ViDu: tự tính tay tháng đầu tiên theo từng bước rồi so với dòng đầu của sheet DieuTiet; xem thêm ví dụ tháng thiếu nước đầu tiên.'),
    ('-', 'Bước 4 – sheet DieuTiet: 528 dòng, mỗi dòng một tháng; các khối cột Lần lặp 1, 2, 3 và Kết quả tháng. Dòng nền đỏ là tháng thiếu nước.'),
    ('-', 'Bước 5 – sheet KetQua: tổng hợp theo năm, đếm số năm đủ nước m, tính P = m/(n + 1), so với điều kiện m ≥ (n + 1)·P_tk và kết luận; kiểm tra khép kín cân bằng nước toàn chuỗi.'),
    ('-', 'Bước 6 – sheet BieuDo: quá trình mực nước hồ cuối tháng 44 năm, thấy rõ các năm hồ chạm MNC.'),
    ('M', '5. Kết quả kiểm chứng với chương trình của đồ án (kt_dong_chay_toi_thieu.py, k_t = 1%/tháng)'),
    ('-', 'TH1A: 39/44 năm, P = 86,7% – ĐẠT (cần 39).   TH2: 37/44 năm, P = 82,2% – KHÔNG ĐẠT.   TH1: 18/44 năm, P = 40,0% – KHÔNG ĐẠT. File Excel cho đúng các số này (sai khác với chương trình < 10⁻¹³ triệu m³).'),
    ('-', 'TH1A thiếu nước 6 tháng trong 5 năm: tháng 4 các năm 2005-2006, 2006-2007, 2008-2009 (cuối mùa khô, nhu cầu tháng 4 lớn 14,20 triệu m³); tháng 5, 6 năm 1986-1987 và tháng 7 năm 1990-1991 (mùa mưa đến muộn). Xem các dòng nền đỏ ở sheet DieuTiet và ví dụ 2 ở sheet ViDu.'),
    ('M', '6. Dung tích thay đổi (Vc 8,41 → 6,82) có làm thay đổi cao trình, chiều cao đập không?'),
    ('-', 'Không trực tiếp. Cao trình đỉnh đập xác định từ mực nước lũ (MNLTK, MNLKT) cộng sóng và độ vượt cao an toàn. Mực nước lũ tính bằng điều tiết lũ bắt đầu từ MNDBT, chỉ dùng phần dung tích PHÍA TRÊN MNDBT (+74,70 ÷ +78,85 m: 30,3 triệu m³) và khả năng xả của tràn.'),
    ('-', 'Dung tích chết và dung tích hữu ích nằm dưới MNDBT, không tham gia điều tiết lũ. Kiểm tra bằng chương trình: dịch cả đường dung tích để V(MNDBT) = 37,24 như thiết kế → MNLTK, MNLKT vẫn +77,63 / +78,54 m; dùng nguyên cột W của bảng hồ sơ → +77,64 / +78,51 m (lệch ≤ 0,03 m).'),
    ('-', 'Phần trên MNDBT của đường 2022 so với thiết kế chỉ làm MNLTK lệch 0,07 m (lũ 847 m³/s, tràn 25 m: +77,31 so với +77,24 m). Đỉnh đập +78,85 m (TKCS +78,75 m) cao hơn chủ yếu do lũ thiết kế cập nhật (+0,20 m) và tràn chính thực tế hẹp hơn 21,6 m (+0,12 m).'),
    ('-', 'Cấp công trình không đổi: V(MNDBT) = 36,41 triệu m³ vẫn trong khoảng 20 ÷ 200 triệu m³ → cấp II → lũ thiết kế 1%, kiểm tra 0,2% như cũ. Dung tích giảm chỉ ảnh hưởng cấp nước (sheet này) và điều kiện bồi lắng – lấy nước (MNC).'),
    ('M', '7. Chú giải màu ô'),
    ('-', 'Nền vàng, chữ xanh: số liệu nhập hoặc lựa chọn – được sửa để thử (trường hợp, MNC, MNDBT, k_t, P_tk).'),
    ('-', 'Chữ xanh trên nền trắng: số liệu gốc lấy từ hồ sơ (dòng chảy, nhu cầu nước, Z ~ F ~ W, bốc hơi) – chỉ sửa khi có số liệu mới, có ghi nguồn.'),
    ('-', 'Chữ đen: công thức – không gõ đè. Nền đỏ nhạt: tháng / năm thiếu nước.'),
    ('M', '8. Bài tập tự kiểm tra (đổi ở sheet SoLieu, xem kết quả ở KetQua)'),
    ('-', 'Đổi k_t từ 1% sang 0,5% → TH1A còn bao nhiêu năm đủ? (đáp số: 42/44). Chọn TH2 với k_t = 0,5% → 39/44, đạt.'),
    ('-', 'Giảm MNDBT (ô B7) để tìm MNDBT tối thiểu cho TH1A đủ 39 năm (đáp số: khoảng +74,63 m; chương trình: +74,623 m – thuyết minh ghi +74,62 m).'),
    ('-', 'Tính tay tháng 1 năm đầu theo sheet ViDu (3 lần lặp) rồi so với dòng 5 của sheet DieuTiet.'),
]
h = 1
for loai, txt in NOI_DUNG:
    o = hd['B%d' % h]
    o.value = ('•  ' + txt) if loai == '-' else txt
    o.alignment = trai
    if loai == 'T':
        o.font = f_tieu_de
        hd.row_dimensions[h].height = 24
    elif loai == 'M':
        o.font = f_muc
        hd.row_dimensions[h].height = 22
    elif loai == 'c':
        o.font = Font(name=FONT, size=12, bold=True, color='12324A')
        o.alignment = Alignment(horizontal='center', vertical='center')
        hd.row_dimensions[h].height = 22
    elif loai == 'g':
        o.font = f_ghi_chu
        hd.row_dimensions[h].height = 30
    else:
        o.font = f_thuong
        hd.row_dimensions[h].height = 15.5 * max(1, math.ceil(len(txt) / 145))
    h += 1

# ================================================================== 2. SỐ LIỆU
sl = wb.create_sheet('SoLieu')
for c, w in zip('ABCDEFGHIJKLMNO', (44, 12, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 11, 60)):
    sl.column_dimensions[c].width = w
ghi(sl, 'A1', 'SỐ LIỆU ĐẦU VÀO – ĐIỀU TIẾT TOÀN CHUỖI HỒ CÀ GIÂY', f_tieu_de, border=False)
ghi(sl, 'A2', 'Ô nền vàng chữ xanh: được sửa để thử. Chữ xanh: số liệu gốc hồ sơ. Chữ đen: công thức. '
              'Dung tích: 10⁶ m³ (triệu m³).', f_ghi_chu, border=False)

ghi(sl, 'A4', 'A. THÔNG SỐ TÍNH TOÁN', f_muc, border=False)
tieu_de_cot(sl, 4, 2, ['Giá trị', 'Đơn vị'])
ghi(sl, 'D4', 'Ghi chú, nguồn', f_trang, fill=nen_tieu_de, can=giua)
sl.merge_cells('D4:O4')
TS = [
    (5, 'Trường hợp dùng nước (chọn TH1 / TH1A / TH2)', 'TH1A', None, '', 'Mặc định TH1A – trường hợp sát yêu cầu nhất (39/44 năm). Mô tả từng trường hợp ở mục C.'),
    (6, 'Mực nước chết MNC', 67.5, '0.00', 'm', 'Giữ như thiết kế; kiểm tra bồi lắng, lấy nước theo TCVN 10778:2024 mục 7 (thuyết minh 2.2.1)'),
    (7, 'Mực nước dâng bình thường MNDBT', 74.7, '0.00', 'm', 'Bằng cao trình ngưỡng tràn chính'),
    (8, 'Tổn thất thấm lòng hồ k_t', 0.01, '0.0%', '/tháng', 'Phần trăm dung tích bình quân tháng; TCVN 10778:2024 Bảng 1: 0,5 ÷ 1,0%/tháng (địa chất lòng hồ tốt) – đồ án lấy 1%'),
    (9, 'Mức bảo đảm cấp nước thiết kế P_tk', 0.85, '0.0%', '', 'Công trình cấp II (thuyết minh 1.1.3). TH1A tính chung tưới 85% + sinh hoạt 90% theo trọng số: 85,15% (TCVN 10778:2024 mục 8.2.2.2) – vẫn cần 39 năm'),
]
for hang, nhan, gt, fmt, dv, gc in TS:
    ghi(sl, 'A%d' % hang, nhan, f_dam)
    ghi(sl, 'B%d' % hang, gt, f_nhap_dam, fmt, nen_nhap, Alignment(horizontal='center'))
    ghi(sl, 'C%d' % hang, dv, can=Alignment(horizontal='center'))
    ghi(sl, 'D%d' % hang, gc, f_ghi_chu)
    sl.merge_cells('D%d:O%d' % (hang, hang))
dv_th = DataValidation(type='list', formula1='"TH1,TH1A,TH2"', allow_blank=False)
dv_th.error, dv_th.errorTitle = 'Chỉ chọn TH1, TH1A hoặc TH2', 'Trường hợp dùng nước'
sl.add_data_validation(dv_th)
dv_th.add('B5')

V_Z = lambda z: noi_suy(z, 'BT_Z', 'BT_V')
KQ = [
    (10, 'Dung tích chết Vc = V(MNC)', '=' + V_Z('Z_MNC'), 'Tra bảng V ~ Z bình đồ 2022 (sheet BangTraZV). Thiết kế ban đầu: 8,41 (chỉ để so sánh)'),
    (11, 'Dung tích ứng với MNDBT Vbt = V(MNDBT)', '=' + V_Z('Z_MNDBT'), 'Thiết kế ban đầu: 37,24'),
    (12, 'Dung tích hữu ích Vhi = Vbt − Vc', '=B11-B10', 'Thiết kế ban đầu: 28,83'),
    (13, 'Số năm của chuỗi dòng chảy n', '=COUNTA(Nam_lbl)', 'Đếm số năm trong bảng dòng chảy (mục E)'),
]
for hang, nhan, ct, gc in KQ:
    ghi(sl, 'A%d' % hang, nhan, f_dam)
    ghi(sl, 'B%d' % hang, ct, f_dam, '0' if hang == 13 else '0.000', can=Alignment(horizontal='center'))
    ghi(sl, 'C%d' % hang, '' if hang == 13 else '10⁶ m³', can=Alignment(horizontal='center'))
    ghi(sl, 'D%d' % hang, gc, f_ghi_chu)
    sl.merge_cells('D%d:O%d' % (hang, hang))
for nm, o in (('TH_chon', 'B5'), ('Z_MNC', 'B6'), ('Z_MNDBT', 'B7'), ('k_tham', 'B8'), ('P_tk', 'B9'),
              ('V_chet', 'B10'), ('V_bt', 'B11'), ('n_nam', 'B13')):
    ten(nm, 'SoLieu!$%s$%s' % (o[0], o[1:]))

# B. bốc hơi
ghi(sl, 'A15', 'B. TỔN THẤT BỐC HƠI PHỤ THÊM ΔZ (mm) – QTVH Bảng 3-24', f_muc, border=False)
tieu_de_cot(sl, 16, 1, ['Tháng'] + list(range(1, 13)) + ['Cả năm'])
ghi(sl, 'A17', 'ΔZ (mm)', f_dam)
for k, x in enumerate(DZ_MM):
    ghi(sl, '%s17' % get_column_letter(2 + k), x, f_nhap, '0.0', can=Alignment(horizontal='center'))
ghi(sl, 'N17', '=SUM(B17:M17)', f_dam, '0.0', can=Alignment(horizontal='center'))
ten('BocHoi', 'SoLieu!$B$17:$M$17')

# C. nhu cầu
ghi(sl, 'A19', 'C. NHU CẦU DÙNG NƯỚC TỔNG (tưới + sinh hoạt), 10⁶ m³/tháng – QTVH Bảng 4-20 ÷ 4-22', f_muc, border=False)
tieu_de_cot(sl, 20, 1, ['Trường hợp'] + list(range(1, 13)) + ['Cả năm', 'Nội dung'])
for k, th in enumerate(('TH1', 'TH1A', 'TH2')):
    hang = 21 + k
    ghi(sl, 'A%d' % hang, th, f_dam, can=Alignment(horizontal='center'))
    for m in range(12):
        ghi(sl, '%s%d' % (get_column_letter(2 + m), hang), NHU_CAU[th]['tong'][m], f_nhap, '0.00',
            can=Alignment(horizontal='center'))
    ghi(sl, 'N%d' % hang, '=SUM(B%d:M%d)' % (hang, hang), f_dam, '0.00', can=Alignment(horizontal='center'))
    ghi(sl, 'O%d' % hang, MO_TA_TH[th] + '; trong đó sinh hoạt 3,22 triệu m³/năm', f_ghi_chu)
ghi(sl, 'A24', '→ Trường hợp đang chọn', f_dam, fill=nen_kq)
for m in range(12):
    cl = get_column_letter(2 + m)
    ghi(sl, '%s24' % cl, '=INDEX(%s$21:%s$23,MATCH(TH_chon,$A$21:$A$23,0))' % (cl, cl), f_dam, '0.00',
        nen_kq, Alignment(horizontal='center'))
ghi(sl, 'N24', '=SUM(B24:M24)', f_dam, '0.00', nen_kq, Alignment(horizontal='center'))
ghi(sl, 'O24', '=TH_chon', f_dam, fill=nen_kq)
ten('NhuCau_chon', 'SoLieu!$B$24:$M$24')

# D. Z ~ F ~ W
ghi(sl, 'A26', 'D. QUAN HỆ Z ~ F ~ W LÒNG HỒ – BÌNH ĐỒ KHẢO SÁT 2022 (QTVH Bảng 1-2; BCNCKT mục 2.1.2)', f_muc, border=False)
tieu_de_cot(sl, 27, 1, ['Z (m)', 'F (km²)', 'W hồ sơ', 'W tính lại từ F', 'Lệch (hồ sơ − tính lại)'], cao=30)
ghi(sl, 'F27', 'Ghi chú', f_trang, fill=nen_tieu_de, can=giua)
sl.merge_cells('F27:O27')
for k, (z, f, w) in enumerate(zip(S.Z_TAB, S.F_TAB, S.W_TAB_HOSO)):
    hang = 28 + k
    ghi(sl, 'A%d' % hang, float(z), f_nhap, '0.00', can=Alignment(horizontal='center'))
    ghi(sl, 'B%d' % hang, float(f), f_nhap, '0.00', can=Alignment(horizontal='center'))
    ghi(sl, 'C%d' % hang, float(w), f_nhap, '0.00', can=Alignment(horizontal='center'))
    ct = '=0' if k == 0 else '=D{p}+(A{h}-A{p})/3*(B{p}+B{h}+SQRT(B{p}*B{h}))'.format(p=hang - 1, h=hang)
    ghi(sl, 'D%d' % hang, ct, fmt='0.00', can=Alignment(horizontal='center'))
    ghi(sl, 'E%d' % hang, '=C%d-D%d' % (hang, hang), fmt='+0.00;-0.00;0.00', can=Alignment(horizontal='center'))
GHI_CHU_ZFW = [
    'Cột F trong hồ sơ ghi đơn vị “ha” nhưng đúng là km² (TKCS Hình 1-1; hồ 37 triệu m³ không thể có mặt hồ vài ha).',
    'W tính lại theo công thức hình chóp cụt: ΔW = ΔZ/3·(F₁ + F₂ + √(F₁·F₂)) (thuyết minh công thức (1-1)).',
    'Dưới +69 m cột W hồ sơ khớp W tính lại (lệch ≤ 0,01) → V(MNC) = 6,82 chắc chắn.',
    'Từ +70 m trở lên cột W hồ sơ tăng giảm thất thường, không khớp cột F → đồ án dùng W tính lại từ F.',
    'Khi tính, chương trình dùng bảng chi tiết 0,01 m (sheet BangTraZV) để nội suy chính xác hơn bảng 1 m này.',
    'Đáy hồ thấp nhất trên bình đồ 2022 ở khoảng +59 m.',
]
for k, t in enumerate(GHI_CHU_ZFW):
    ghi(sl, 'F%d' % (28 + k), t, f_ghi_chu)
    sl.merge_cells('F%d:O%d' % (28 + k, 28 + k))
ten('Bang_Z', 'SoLieu!$A$28:$A$54')
ten('Bang_F', 'SoLieu!$B$28:$B$54')

# E. dòng chảy đến
H0 = 57
ghi(sl, 'A%d' % (H0 - 1), 'E. LƯỢNG DÒNG CHẢY ĐẾN HỒ THEO THÁNG (10⁶ m³) – QTVH Bảng 3-13: khu giữa 244 km² đến đập dâng Sông Lũy '
    '+ nước bổ sung từ hồ Sông Lũy (kênh tiếp nước Q ≤ 8,6 m³/s) + nội lưu vực hồ Cà Giây 141 km²', f_muc, border=False)
tieu_de_cot(sl, H0, 1, ['Năm (nhãn QTVH)'] + ['Tháng %d' % m for m in range(1, 13)] + ['Cả năm'])
for j, r in enumerate(D['dong_chay']):
    hang = H0 + 1 + j
    ghi(sl, 'A%d' % hang, r[0], f_nhap, can=Alignment(horizontal='center'))
    for m in range(12):
        ghi(sl, '%s%d' % (get_column_letter(2 + m), hang), r[1 + m], f_nhap, '0.00', can=Alignment(horizontal='center'))
    ghi(sl, 'N%d' % hang, '=SUM(B%d:M%d)' % (hang, hang), f_dam, '0.00', can=Alignment(horizontal='center'))
H_CUOI = H0 + len(D['dong_chay'])
ghi(sl, 'A%d' % (H_CUOI + 1), 'Trung bình 44 năm', f_dam, fill=nen_kq)
for m in range(13):
    cl = get_column_letter(2 + m)
    ghi(sl, '%s%d' % (cl, H_CUOI + 1), '=AVERAGE(%s%d:%s%d)' % (cl, H0 + 1, cl, H_CUOI), f_dam, '0.00', nen_kq,
        Alignment(horizontal='center'))
ghi(sl, 'O%d' % (H0 + 1), 'Các cột là tháng 1 ÷ 12 như bảng QTVH: mùa khô tháng 1 ÷ 4 (dòng chảy nhỏ, nhu cầu tưới lớn), '
    'mùa lũ tháng 9 ÷ 11. Chuỗi nối liên tục: sau tháng 12 của một dòng là tháng 1 của dòng kế tiếp.', f_ghi_chu)
sl.merge_cells('O%d:O%d' % (H0 + 1, H0 + 6))
sl['O%d' % (H0 + 1)].alignment = trai
ten('Nam_lbl', 'SoLieu!$A$%d:$A$%d' % (H0 + 1, H_CUOI))
ten('DongChay', 'SoLieu!$B$%d:$M$%d' % (H0 + 1, H_CUOI))
sl.freeze_panes = 'B5'

# ================================================================== 3. BẢNG TRA Z ~ F ~ V 0,01 m
bt = wb.create_sheet('BangTraZV')
for c, w in zip('ABCDE', (10, 12, 14, 4, 90)):
    bt.column_dimensions[c].width = w
ghi(bt, 'A1', 'BẢNG TRA Z ~ F ~ V CHI TIẾT (bước 0,01 m)', f_tieu_de, border=False)
ghi(bt, 'A2', 'F nội suy tuyến tính từ bảng 1 m (SoLieu mục D); V cộng dồn theo chóp cụt ΔV = ΔZ/3·(F₁ + F₂ + √(F₁F₂)) – '
              'giống chương trình so_lieu.py của đồ án.', f_ghi_chu, border=False)
tieu_de_cot(bt, 4, 1, ['Z (m)', 'F (km²)', 'V (10⁶ m³)'])
N_BT = 2601
for k in range(N_BT):
    hang = 5 + k
    ghi(bt, 'A%d' % hang, 59.0 if k == 0 else '=ROUND(A%d+0.01,2)' % (hang - 1), fmt='0.00', border=False)
    ghi(bt, 'B%d' % hang, '=IF(A{h}>=MAX(Bang_Z),INDEX(Bang_F,ROWS(Bang_F)),{f})'.format(h=hang, f=F_tu_Z('A%d' % hang)),
        fmt='0.0000', border=False)
    ghi(bt, 'C%d' % hang, 0 if k == 0 else '=C{p}+(A{h}-A{p})/3*(B{p}+B{h}+SQRT(B{p}*B{h}))'.format(p=hang - 1, h=hang),
        fmt='0.0000', border=False)
H_BT = 4 + N_BT
ten('BT_Z', 'BangTraZV!$A$5:$A$%d' % H_BT)
ten('BT_F', 'BangTraZV!$B$5:$B$%d' % H_BT)
ten('BT_V', 'BangTraZV!$C$5:$C$%d' % H_BT)
ghi(bt, 'E4', 'Cách tra (nội suy tuyến tính)', f_trang, fill=nen_tieu_de, can=giua)
for k, t in enumerate([
        'Biết Z, tìm V: tìm dòng có Z_k ≤ Z < Z_k+1 (hàm MATCH), V = V_k + (Z − Z_k)·(V_k+1 − V_k)/(Z_k+1 − Z_k).',
        'Biết V, tìm Z: tương tự, đổi vai Z và V. Ví dụ: Vc = V(+67,50) = 6,816 triệu m³; Vbt = V(+74,70) = 36,412 triệu m³.',
        'F(Z) lấy từ bảng 1 m của SoLieu mục D (nội suy tuyến tính) – dùng để tính tổn thất bốc hơi.']):
    ghi(bt, 'E%d' % (5 + k), t, f_ghi_chu)
bt.freeze_panes = 'A5'

# ================================================================== 4. ĐIỀU TIẾT
dt = wb.create_sheet('DieuTiet')
N_TH = 12 * len(D['dong_chay'])
R0 = 5
R1 = R0 + N_TH - 1
ghi(dt, 'A1', 'ĐIỀU TIẾT TOÀN CHUỖI THEO THÁNG – %d NĂM × 12 THÁNG = %d THÁNG' % (len(D['dong_chay']), N_TH), f_tieu_de,
    border=False)
for o, t in (('A2', 'Trường hợp'), ('C2', '=TH_chon'), ('D2', 'Vc ='), ('E2', '=V_chet'), ('F2', 'Vbt ='), ('G2', '=V_bt'),
             ('H2', 'k_t ='), ('I2', '=k_tham')):
    ghi(dt, o, t, f_dam, '0.0%' if o == 'I2' else ('0.000' if o in ('E2', 'G2') else None), nen_kq)
ghi(dt, 'J2', '← đổi ở sheet SoLieu. Dòng nền đỏ: tháng thiếu nước.', f_ghi_chu, border=False)
NHOM = [('A3:G3', 'SỐ LIỆU THÁNG', None), ('H3:M3', 'LẦN LẶP 1 – giả thiết V̄ = V đầu tháng', nen_lap[0]),
        ('N3:S3', 'LẦN LẶP 2 – V̄ = (V đầu + V′₁)/2', nen_lap[1]), ('T3:Y3', 'LẦN LẶP 3 – V̄ = (V đầu + V′₂)/2', nen_lap[2]),
        ('Z3:AD3', 'KẾT QUẢ THÁNG', nen_kq)]
for rg, t, nen in NHOM:
    dt.merge_cells(rg)
    ghi(dt, rg.split(':')[0], t, f_dam, fill=nen or PatternFill('solid', fgColor='E6E8EB'), can=giua)
COT = ['TT tháng', 'Năm', 'Tháng', 'W đến (10⁶ m³)', 'W dùng (10⁶ m³)', 'ΔZ bốc hơi (m)', 'V đầu tháng (10⁶ m³)',
       'V̄₁ = V đầu', 'Z̄₁ (m) tra V~Z', 'F₁ (km²)', 'W bh₁ = ΔZ·F₁', 'W thấm₁ = k_t·V̄₁',
       'V′₁ = V đầu + W đến − W dùng − W bh₁ − W thấm₁',
       'V̄₂ = (V đầu + V′₁*)/2', 'Z̄₂ (m)', 'F₂ (km²)', 'W bh₂', 'W thấm₂', 'V′₂',
       'V̄₃ = (V đầu + V′₂*)/2', 'Z̄₃ (m)', 'F₃ (km²)', 'W bh₃', 'W thấm₃', 'V′₃ (kết quả lặp)',
       'W xả thừa = max(V′₃ − Vbt; 0)', 'W thiếu = max(Vc − V′₃; 0)', 'V cuối tháng (giới hạn Vc ÷ Vbt)',
       'Z cuối tháng (m)', 'Tháng thiếu (1 = có)']
tieu_de_cot(dt, 4, 1, COT, cao=58)
RONG = [7, 10, 6.5, 9, 9, 8.5, 10, 9, 8.5, 8, 8.5, 8.5, 14, 10, 8.5, 8, 8.5, 8.5, 9.5, 10, 8.5, 8, 8.5, 8.5, 10.5, 10,
        10, 11, 9, 8]
for k, w in enumerate(RONG):
    dt.column_dimensions[get_column_letter(k + 1)].width = w
KEP = 'MIN(MAX({v},V_chet),V_bt)'
for t in range(1, N_TH + 1):
    r = R0 + t - 1
    o = lambda c: '%s%d' % (c, r)
    lap = lambda col: nen_lap[(col - 8) // 6] if 8 <= col <= 25 else None
    ct = {
        'A': t,
        'B': '=INDEX(Nam_lbl,INT((A{r}-1)/12)+1)',
        'C': '=MOD(A{r}-1,12)+1',
        'D': '=INDEX(DongChay,INT((A{r}-1)/12)+1,C{r})',
        'E': '=INDEX(NhuCau_chon,1,C{r})',
        'F': '=INDEX(BocHoi,1,C{r})/1000',
        'G': '=V_bt' if t == 1 else '=AB%d' % (r - 1),
        'H': '=G{r}',
        'I': '=' + Z_tu_V('H{r}'), 'J': '=' + F_tu_Z('I{r}'), 'K': '=F{r}*J{r}', 'L': '=k_tham*H{r}',
        'M': '=G{r}+D{r}-E{r}-K{r}-L{r}',
        'N': '=(G{r}+' + KEP.format(v='M{r}') + ')/2',
        'O': '=' + Z_tu_V('N{r}'), 'P': '=' + F_tu_Z('O{r}'), 'Q': '=F{r}*P{r}', 'R': '=k_tham*N{r}',
        'S': '=G{r}+D{r}-E{r}-Q{r}-R{r}',
        'T': '=(G{r}+' + KEP.format(v='S{r}') + ')/2',
        'U': '=' + Z_tu_V('T{r}'), 'V': '=' + F_tu_Z('U{r}'), 'W': '=F{r}*V{r}', 'X': '=k_tham*T{r}',
        'Y': '=G{r}+D{r}-E{r}-W{r}-X{r}',
        'Z': '=MAX(Y{r}-V_bt,0)', 'AA': '=MAX(V_chet-Y{r},0)', 'AB': '=' + KEP.format(v='Y{r}'),
        'AC': '=' + Z_tu_V('AB{r}'), 'AD': '=IF(AA{r}>0.000001,1,0)',
    }
    FMT = {'A': '0', 'B': 'General', 'C': '0', 'F': '0.0000', 'I': '0.000', 'O': '0.000', 'U': '0.000', 'AC': '0.00',
           'J': '0.000', 'P': '0.000', 'V': '0.000', 'AD': '0'}
    for col, f in ct.items():
        ci = sum((ord(ch) - 64) * 26 ** (len(col) - 1 - k) for k, ch in enumerate(col))
        c = dt[o(col)]
        c.value = f.format(r=r) if isinstance(f, str) else f
        c.font = f_thuong
        c.number_format = FMT.get(col, '0.000')
        c.alignment = Alignment(horizontal='center')
        nen = lap(ci)
        if nen:
            c.fill = nen
        elif ci >= 26:
            c.fill = nen_kq
dt.freeze_panes = 'D5'
dt.conditional_formatting.add('A%d:AD%d' % (R0, R1),
                              FormulaRule(formula=['$AD%d=1' % R0], fill=nen_do, stopIfTrue=False))

# ================================================================== 5. KẾT QUẢ
kq = wb.create_sheet('KetQua')
for c, w in zip('ABCDEFGHIJKLMN', (9, 12, 11, 11, 11, 11, 11, 11, 10, 12, 12, 12, 10, 11)):
    kq.column_dimensions[c].width = w
ghi(kq, 'A1', 'KẾT QUẢ ĐIỀU TIẾT THEO NĂM VÀ MỨC BẢO ĐẢM CẤP NƯỚC', f_tieu_de, border=False)
TOM = [
    (3, 'Trường hợp dùng nước', '=TH_chon', None),
    (4, 'Số năm tính toán n', '=n_nam', '0'),
    (5, 'Số năm đủ nước m (năm không có tháng thiếu)', '=SUM(M19:M%d)' % (18 + len(D['dong_chay'])), '0'),
    (6, 'Số năm thiếu nước n − m', '=B4-B5', '0'),
    (7, 'Mức bảo đảm P = m/(n + 1) – TCVN 10778:2024 công thức (14)', '=B5/(B4+1)', '0.0%'),
    (8, 'Điều kiện đạt (công thức (9)): m ≥ (n + 1)·P_tk =', '=(B4+1)*P_tk', '0.00'),
    (9, '→ cần ít nhất (năm đủ nước)', '=ROUNDUP(B8,0)', '0'),
    (10, 'KẾT LUẬN', '=IF(B5>=B8,"ĐẠT","KHÔNG ĐẠT")', None),
    (11, 'Tham khảo: P = m/n (cách QTVH tính)', '=B5/B4', '0.0%'),
    (12, 'Số tháng thiếu nước / tổng lượng thiếu (10⁶ m³)', '=SUM(I19:I%d)' % (18 + len(D['dong_chay'])), '0'),
    (13, 'Số năm hồ KHÔNG tích đầy MNDBT (> 0 thì năm sau bắt đầu với dung tích thấp – ảnh hưởng chuyển tiếp)', '=COUNTIF(N19:N%d,0)' % (18 + len(D['dong_chay'])), '0'),
    (14, 'Kiểm tra cân bằng nước toàn chuỗi: ΔV − (ΣW đến − ΣW dùng − ΣW bh − ΣW thấm − ΣW xả + ΣW thiếu) =', None, '0.000000'),
]
for hang, nhan, ct, fmt in TOM:
    kq.merge_cells('C%d:L%d' % (hang, hang))
    ghi(kq, 'C%d' % hang, nhan, f_dam)
    kq['C%d' % hang].alignment = Alignment(horizontal='left', vertical='center')
    if ct is not None:
        ghi(kq, 'B%d' % hang, ct, f_dam, fmt, nen_kq, Alignment(horizontal='center'))
    ghi(kq, 'A%d' % hang, '', border=False)
NAM = len(D['dong_chay'])
HN = 18 + NAM
ghi(kq, 'M12', '=SUM(H19:H%d)' % HN, f_dam, '0.000', nen_kq, Alignment(horizontal='center'))
ghi(kq, 'B14', '=(DieuTiet!AB%d-DieuTiet!G%d)-(SUM(C19:C%d)-SUM(D19:D%d)-SUM(E19:E%d)-SUM(F19:F%d)-SUM(G19:G%d)+SUM(H19:H%d))'
    % (R1, R0, HN, HN, HN, HN, HN, HN), f_dam, '0.000000', nen_kq, Alignment(horizontal='center'))
kq['B10'].font = Font(name=FONT, size=12, bold=True, color='A65A10')
kq.conditional_formatting.add('B10', FormulaRule(formula=['$B$10="KHÔNG ĐẠT"'], fill=nen_do))
ghi(kq, 'A15', 'Kiểm chứng: chương trình Python của đồ án (kt_dong_chay_toi_thieu.py, k_t = 1%) cho TH1A 39/44 năm (86,7%), '
               'TH2 37/44 (82,2%), TH1 18/44 (40,0%).', f_ghi_chu, border=False)
tieu_de_cot(kq, 18, 1, ['Năm thứ', 'Năm', 'W đến', 'W dùng', 'W bốc hơi', 'W thấm', 'W xả thừa', 'W thiếu',
                         'Số tháng thiếu', 'V đầu năm (tháng 1)', 'V cuối năm (tháng 12)', 'Z thấp nhất trong năm (m)',
                         'Đủ nước? (1 = đủ)', 'Hồ tích đầy trong năm? (1 = có)'], cao=45)
ghi(kq, 'A17', 'Đơn vị dung tích: 10⁶ m³. Mỗi dòng cộng 12 tháng tương ứng của sheet DieuTiet.', f_ghi_chu, border=False)
for j in range(1, NAM + 1):
    r = 18 + j
    ofs = lambda col: 'OFFSET(DieuTiet!$%s$%d,(A%d-1)*12,0,12,1)' % (col, R0, r)
    vals = {
        'A': j, 'B': '=INDEX(Nam_lbl,A%d)' % r,
        'C': '=SUM(%s)' % ofs('D'), 'D': '=SUM(%s)' % ofs('E'), 'E': '=SUM(%s)' % ofs('W'),
        'F': '=SUM(%s)' % ofs('X'), 'G': '=SUM(%s)' % ofs('Z'), 'H': '=SUM(%s)' % ofs('AA'),
        'I': '=SUM(%s)' % ofs('AD'),
        'J': '=INDEX(DieuTiet!$G$%d:$G$%d,(A%d-1)*12+1)' % (R0, R1, r),
        'K': '=INDEX(DieuTiet!$AB$%d:$AB$%d,A%d*12)' % (R0, R1, r),
        'L': '=MIN(%s)' % ofs('AC'),
        'M': '=IF(I%d=0,1,0)' % r,
        'N': '=IF(MAX(%s)>=V_bt-0.000001,1,0)' % ofs('AB'),
    }
    for col, v in vals.items():
        fmt = {'A': '0', 'I': '0', 'L': '0.00', 'M': '0', 'N': '0'}.get(col, '0.00')
        ghi(kq, '%s%d' % (col, r), v, fmt=fmt, can=Alignment(horizontal='center'))
ghi(kq, 'A%d' % (HN + 1), 'Cộng', f_dam, fill=nen_kq)
ghi(kq, 'B%d' % (HN + 1), '', fill=nen_kq)
for col in 'CDEFGHIMN':
    ghi(kq, '%s%d' % (col, HN + 1), '=SUM(%s19:%s%d)' % (col, col, HN), f_dam, '0' if col in 'IMN' else '0.00', nen_kq,
        Alignment(horizontal='center'))
kq.conditional_formatting.add('A19:N%d' % HN, FormulaRule(formula=['$M19=0'], fill=nen_do))
kq.freeze_panes = 'C19'
# bảng minh họa: vì sao chuỗi dài
HM = HN + 4
ghi(kq, 'A%d' % (HM - 1), 'MINH HỌA – VÌ SAO DÙNG HẾT 44 NĂM: chuỗi càng dài, mức bảo đảm càng “mịn” và đánh giá càng tin cậy',
    f_muc, border=False)
tieu_de_cot(kq, HM, 1, ['n (năm)', 'Mỗi năm thiếu làm P giảm', 'm cần (P ≥ P_tk)', 'Được thiếu tối đa (năm)'], cao=45)
for k, nn in enumerate((15, 20, 30, 44)):
    r = HM + 1 + k
    ghi(kq, 'A%d' % r, nn, f_nhap, '0', can=Alignment(horizontal='center'))
    ghi(kq, 'B%d' % r, '=1/(A%d+1)' % r, fmt='0.0%', can=Alignment(horizontal='center'))
    ghi(kq, 'C%d' % r, '=ROUNDUP((A%d+1)*P_tk,0)' % r, fmt='0', can=Alignment(horizontal='center'))
    ghi(kq, 'D%d' % r, '=A%d-C%d' % (r, r), fmt='0', can=Alignment(horizontal='center'))
ghi(kq, 'E%d' % (HM + 1), 'n ≥ 15 năm là điều kiện tối thiểu để dùng phương pháp toàn chuỗi (TCVN 10778:2024 mục 8.2.1.2); '
    'hồ Cà Giây có 44 năm nên dùng hết.', f_ghi_chu, border=False)

# ================================================================== 6. VÍ DỤ TÍNH TAY
vd = wb.create_sheet('ViDu')
for c, w in zip('ABCDE', (7, 46, 52, 13, 10)):
    vd.column_dimensions[c].width = w
ghi(vd, 'A1', 'VÍ DỤ TÍNH TAY – THÁNG ĐẦU TIÊN CỦA CHUỖI (dòng 5 sheet DieuTiet)', f_tieu_de, border=False)
ghi(vd, 'A2', 'Tự bấm máy theo cột “Công thức” rồi so với cột “Giá trị” (lấy từ sheet DieuTiet).', f_ghi_chu, border=False)
tieu_de_cot(vd, 4, 1, ['Bước', 'Nội dung', 'Công thức', 'Giá trị', 'Đơn vị'])
r0 = R0
VD1 = [
    ('1', 'Năm, tháng', 'Tháng đầu tiên của chuỗi', '=DieuTiet!B%d&" – tháng "&DieuTiet!C%d' % (r0, r0), '', 'General'),
    ('', 'Dung tích đầu tháng V đầu', 'Đầu chuỗi giả thiết hồ đầy: V đầu = Vbt', '=DieuTiet!G%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Lượng nước đến W đến', 'SoLieu mục E, tháng 1 năm đầu', '=DieuTiet!D%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Lượng nước dùng W dùng', 'SoLieu mục C, tháng 1 của trường hợp chọn', '=DieuTiet!E%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Bốc hơi phụ thêm ΔZ', 'SoLieu mục B: mm ÷ 1000', '=DieuTiet!F%d' % r0, 'm', '0.0000'),
    ('2', 'Lần lặp 1: V̄₁', 'Giả thiết V̄₁ = V đầu', '=DieuTiet!H%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Mực nước bình quân Z̄₁', 'Tra bảng V ~ Z (BangTraZV), nội suy', '=DieuTiet!I%d' % r0, 'm', '0.000'),
    ('', 'Diện tích mặt hồ F₁', 'Tra bảng Z ~ F (SoLieu mục D), nội suy', '=DieuTiet!J%d' % r0, 'km²', '0.000'),
    ('', 'Tổn thất bốc hơi W bh₁', 'ΔZ × F₁', '=DieuTiet!K%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Tổn thất thấm W thấm₁', 'k_t × V̄₁', '=DieuTiet!L%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Dung tích cuối tháng V′₁', 'V đầu + W đến − W dùng − W bh₁ − W thấm₁', '=DieuTiet!M%d' % r0, '10⁶ m³', '0.000'),
    ('3', 'Lần lặp 2: V̄₂', '(V đầu + V′₁)/2  (V′₁ giới hạn trong Vc ÷ Vbt)', '=DieuTiet!N%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Z̄₂', 'Tra bảng V ~ Z', '=DieuTiet!O%d' % r0, 'm', '0.000'),
    ('', 'F₂', 'Tra bảng Z ~ F', '=DieuTiet!P%d' % r0, 'km²', '0.000'),
    ('', 'W bh₂', 'ΔZ × F₂', '=DieuTiet!Q%d' % r0, '10⁶ m³', '0.000'),
    ('', 'W thấm₂', 'k_t × V̄₂', '=DieuTiet!R%d' % r0, '10⁶ m³', '0.000'),
    ('', 'V′₂', 'V đầu + W đến − W dùng − W bh₂ − W thấm₂', '=DieuTiet!S%d' % r0, '10⁶ m³', '0.000'),
    ('4', 'Lần lặp 3: V̄₃', '(V đầu + V′₂)/2', '=DieuTiet!T%d' % r0, '10⁶ m³', '0.000'),
    ('', 'V′₃', 'Như lần 2 với V̄₃ → Z̄₃ → F₃', '=DieuTiet!Y%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Sai khác |V′₃ − V′₂|', 'Nhỏ → dừng lặp', '=ABS(DieuTiet!Y%d-DieuTiet!S%d)' % (r0, r0), '10⁶ m³', '0.0000'),
    ('5', 'So với giới hạn: W xả thừa', 'max(V′₃ − Vbt; 0)', '=DieuTiet!Z%d' % r0, '10⁶ m³', '0.000'),
    ('', 'W thiếu', 'max(Vc − V′₃; 0)', '=DieuTiet!AA%d' % r0, '10⁶ m³', '0.000'),
    ('', 'V cuối tháng', 'Giới hạn V′₃ trong khoảng Vc ÷ Vbt', '=DieuTiet!AB%d' % r0, '10⁶ m³', '0.000'),
    ('', 'Z cuối tháng', 'Tra bảng V ~ Z', '=DieuTiet!AC%d' % r0, 'm', '0.00'),
    ('6', 'Sang tháng sau', 'V đầu tháng 2 = V cuối tháng 1', '=DieuTiet!G%d' % (r0 + 1), '10⁶ m³', '0.000'),
]
for k, (b, nd, cthuc, gt, dv, fm) in enumerate(VD1):
    r = 5 + k
    ghi(vd, 'A%d' % r, b, f_dam, can=Alignment(horizontal='center'))
    ghi(vd, 'B%d' % r, nd, f_dam if b else f_thuong)
    ghi(vd, 'C%d' % r, cthuc)
    ghi(vd, 'D%d' % r, gt, f_dam, fm, nen_kq, Alignment(horizontal='center'))
    ghi(vd, 'E%d' % r, dv, can=Alignment(horizontal='center'))
H2 = 5 + len(VD1) + 2
ghi(vd, 'A%d' % H2, 'VÍ DỤ 2 – THÁNG THIẾU NƯỚC ĐẦU TIÊN TRONG CHUỖI', f_muc, border=False)
tieu_de_cot(vd, H2 + 1, 1, ['Bước', 'Nội dung', 'Công thức', 'Giá trị', 'Đơn vị'])
KHONG = '"không có"'
idx = '$D$%d' % (H2 + 2)
VD2 = [
    ('1', 'Thứ tự tháng trong chuỗi', 'Tháng đầu tiên có cột “Tháng thiếu” = 1', '=IFERROR(MATCH(1,DieuTiet!$AD$%d:$AD$%d,0),%s)' % (R0, R1, KHONG), ''),
    ('', 'Năm, tháng', '', '=IFERROR(INDEX(DieuTiet!$B$%d:$B$%d,%s)&" – tháng "&INDEX(DieuTiet!$C$%d:$C$%d,%s),%s)' % (R0, R1, idx, R0, R1, idx, KHONG), ''),
    ('', 'V đầu tháng', '', '=IFERROR(INDEX(DieuTiet!$G$%d:$G$%d,%s),%s)' % (R0, R1, idx, KHONG), '10⁶ m³'),
    ('', 'W đến', '', '=IFERROR(INDEX(DieuTiet!$D$%d:$D$%d,%s),%s)' % (R0, R1, idx, KHONG), '10⁶ m³'),
    ('', 'W dùng', '', '=IFERROR(INDEX(DieuTiet!$E$%d:$E$%d,%s),%s)' % (R0, R1, idx, KHONG), '10⁶ m³'),
    ('', 'W bh₃ + W thấm₃', '', '=IFERROR(INDEX(DieuTiet!$W$%d:$W$%d,%s)+INDEX(DieuTiet!$X$%d:$X$%d,%s),%s)' % (R0, R1, idx, R0, R1, idx, KHONG), '10⁶ m³'),
    ('2', 'V′₃ (chưa giới hạn)', 'V đầu + W đến − W dùng − tổn thất', '=IFERROR(INDEX(DieuTiet!$Y$%d:$Y$%d,%s),%s)' % (R0, R1, idx, KHONG), '10⁶ m³'),
    ('', 'Dung tích chết Vc', 'V(MNC)', '=V_chet', '10⁶ m³'),
    ('3', 'W thiếu', 'Vc − V′₃ > 0 → tháng thiếu nước, năm này là năm thiếu nước', '=IFERROR(INDEX(DieuTiet!$AA$%d:$AA$%d,%s),%s)' % (R0, R1, idx, KHONG), '10⁶ m³'),
    ('', 'Lượng cấp được trong tháng', 'W dùng − W thiếu (hồ về MNC, cắt giảm cấp nước tưới)', '=IFERROR(INDEX(DieuTiet!$E$%d:$E$%d,%s)-INDEX(DieuTiet!$AA$%d:$AA$%d,%s),%s)' % (R0, R1, idx, R0, R1, idx, KHONG), '10⁶ m³'),
    ('', 'V cuối tháng', 'Lấy bằng Vc (hồ ở MNC)', '=IFERROR(INDEX(DieuTiet!$AB$%d:$AB$%d,%s),%s)' % (R0, R1, idx, KHONG), '10⁶ m³'),
]
for k, (b, nd, cthuc, gt, dv) in enumerate(VD2):
    r = H2 + 2 + k
    ghi(vd, 'A%d' % r, b, f_dam, can=Alignment(horizontal='center'))
    ghi(vd, 'B%d' % r, nd, f_dam if b else f_thuong)
    ghi(vd, 'C%d' % r, cthuc)
    ghi(vd, 'D%d' % r, gt, f_dam, '0' if k == 0 else '0.000', nen_kq, Alignment(horizontal='center'))
    ghi(vd, 'E%d' % r, dv, can=Alignment(horizontal='center'))

# ================================================================== 7. BIỂU ĐỒ
bd = wb.create_sheet('BieuDo')
for c, w in zip('TUVWX', (8, 11, 11, 11, 11)):
    bd.column_dimensions[c].width = w
ghi(bd, 'A1', 'QUÁ TRÌNH MỰC NƯỚC HỒ CUỐI THÁNG – ĐIỀU TIẾT TOÀN CHUỖI', f_tieu_de, border=False)
ghi(bd, 'T1', 'Số liệu vẽ (lấy từ sheet DieuTiet)', f_dam, border=False)
tieu_de_cot(bd, 2, 20, ['TT', 'Năm', 'Z cuối tháng', 'MNC', 'MNDBT'])
for t in range(1, N_TH + 1):
    r = 2 + t
    ghi(bd, 'T%d' % r, t, fmt='0', border=False)
    ghi(bd, 'U%d' % r, '=DieuTiet!B%d' % (R0 + t - 1), border=False)
    ghi(bd, 'V%d' % r, '=DieuTiet!AC%d' % (R0 + t - 1), fmt='0.00', border=False)
    ghi(bd, 'W%d' % r, '=Z_MNC', fmt='0.00', border=False)
    ghi(bd, 'X%d' % r, '=Z_MNDBT', fmt='0.00', border=False)
ch = LineChart()
ch.title = 'Mực nước hồ cuối tháng – 44 năm (trường hợp chọn ở SoLieu!B5)'
ch.y_axis.title = 'Mực nước hồ (m)'
ch.x_axis.title = 'Năm (nhãn QTVH)'
ch.height, ch.width = 11, 30
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 66, 76
ch.y_axis.majorUnit = 1
ch.y_axis.number_format = '0'
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.x_axis.tickLblSkip = 36
ch.x_axis.tickMarkSkip = 12
ch.legend.position = 'b'
data = Reference(bd, min_col=22, max_col=24, min_row=2, max_row=2 + N_TH)
ch.add_data(data, titles_from_data=True)
ch.set_categories(Reference(bd, min_col=21, min_row=3, max_row=2 + N_TH))
for s, (mau, net, rong) in zip(ch.series, (('2A78D6', None, 19050), ('52514E', 'sysDash', 15875),
                                           ('52514E', 'dash', 15875))):
    s.marker.symbol = 'none'
    s.smooth = False
    s.graphicalProperties.line.solidFill = mau
    s.graphicalProperties.line.width = rong
    if net:
        s.graphicalProperties.line.dashStyle = net
bd.add_chart(ch, 'A3')
ghi(bd, 'A26', 'Đường liền: mực nước hồ cuối tháng. Đường gạch: MNC (+67,50) và MNDBT (+74,70). Những đoạn chạm MNC là các tháng '
               'thiếu nước (hồ đã về mực nước chết).', f_ghi_chu, border=False)

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'BangTraZV': '8A8984', 'DieuTiet': '2A78D6',
                                    'KetQua': '1BAF7A', 'ViDu': '6250D6', 'BieuDo': '2A78D6'}[ws.title]
# in ấn: khổ A4, vừa bề ngang trang
for ws in wb.worksheets:
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait' if ws.title in ('ViDu', 'BangTraZV') else 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
dt.print_title_rows = '3:4'
kq.print_title_rows = '18:18'
bd.print_area = 'A1:R28'
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
