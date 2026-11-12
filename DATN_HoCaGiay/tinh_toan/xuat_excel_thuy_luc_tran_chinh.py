# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL: TÍNH TOÁN THỦY LỰC TRÀN CHÍNH THEO TRÌNH TỰ FILE MẪU “THIẾT KẾ ĐƯỜNG TRÀN THÁO LŨ” (SV yêu cầu 05/11/2026)
Cùng nội dung, cùng số với Word TAI_VE/ThuyLuc_TranChinh_TheoMau.docx (thuy_luc_tran_chinh.py). Mọi ô kết quả là công thức;
ô nền vàng chữ xanh là số liệu nhập. Các đại lượng phải thử dần (H₀ theo Q, độ sâu co hẹp h_c, độ sâu dòng đều h₀, độ sâu cuối
mỗi đoạn trên dốc) giải bằng phép lặp điểm bất động viết sẵn trong các cột phụ màu xám (6 ÷ 8 vòng, hội tụ < 10⁻⁶ m).
Số liệu tràn chính, kết quả điều tiết lũ lấy từ file DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx (ghi vào SoLieu dưới dạng số –
khi điều tiết lũ thay đổi thì sửa các ô vàng ở mục A, C của SoLieu).
  python3 xuat_excel_thuy_luc_tran_chinh.py [đường dẫn .xlsx ra]
Mặc định: ../TAI_VE/ThuyLuc_TranChinh_TheoMau.xlsx
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

import so_lieu as S
import thuy_luc_tran_chinh as T

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'ThuyLuc_TranChinh_TheoMau.xlsx')
TC = S.TRAN_CHINH
N_B = T.N_BUOC_LT                       # số bước trên mỗi đoạn dốc (thu hẹp, lăng trụ)
NR = 2 * N_B + 1                        # số mặt cắt trên dốc
CAP = [('0,4Q', 0.4, 'MN_04Q'), ('0,6Q', 0.6, 'MN_06Q'), ('0,8Q', 0.8, 'MN_08Q'), ('Q (P = 1%)', 1.0, 'MN_Q1'),
       ('Q (P = 0,2%)', None, 'MN_Q02')]

FONT = 'Times New Roman'
f_thuong = Font(name=FONT, size=11)
f_dam = Font(name=FONT, size=11, bold=True)
f_tieu_de = Font(name=FONT, size=14, bold=True, color='12324A')
f_muc = Font(name=FONT, size=12, bold=True, color='A65A10')
f_nhap = Font(name=FONT, size=11, color='0000FF', bold=True)
f_trang = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_ghi_chu = Font(name=FONT, size=10, italic=True, color='52514E')
f_phu = Font(name=FONT, size=9, color='8A8984')
nen_nhap = PatternFill('solid', fgColor='FFF2CC')
nen_tieu_de = PatternFill('solid', fgColor='12324A')
nen_kq = PatternFill('solid', fgColor='FBF3E8')
nen_phu = PatternFill('solid', fgColor='F2F2F2')
vien = Border(*(Side(style='thin', color='B8BEC6'),) * 4)
giua = Alignment(horizontal='center', vertical='center', wrap_text=True)
trai = Alignment(horizontal='left', vertical='center', wrap_text=True)
cg = Alignment(horizontal='center', vertical='center')

wb = Workbook()


def ghi(ws, o, gt, font=f_thuong, fmt=None, fill=None, can=cg, border=True):
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
    assert all(n.upper() != name.upper() for n in wb.defined_names), name
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def tieu_de_cot(ws, hang, cot0, nhan, cao=None):
    for k, t in enumerate(nhan):
        ghi(ws, '%s%d' % (get_column_letter(cot0 + k), hang), t, f_trang, fill=nen_tieu_de, can=giua)
    if cao:
        ws.row_dimensions[hang].height = cao


def rong(ws, ds):
    for k, w in enumerate(ds):
        ws.column_dimensions[get_column_letter(k + 1)].width = w


def in_an(ws, ngang=True):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if ngang else 'portrait'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.4


def sh(s):
    return "'%s'" % s


# ================================================================== SỐ LIỆU
ts = T.doc_excel()
hd = wb.active
hd.title = 'HuongDan'
sl = wb.create_sheet('SoLieu')
rong(sl, (46, 13, 10, 70))
ghi(sl, 'A1', 'SỐ LIỆU ĐẦU VÀO – THỦY LỰC TRÀN CHÍNH (theo file mẫu “Thiết kế đường tràn tháo lũ”)', f_tieu_de, can=None,
    border=False)
ghi(sl, 'A2', 'Ô nền vàng chữ xanh: số liệu nhập (được sửa). Chữ đen nền nhạt: công thức. Mục A, C lấy từ file điều tiết lũ '
              'DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx.', f_ghi_chu, can=None, border=False)
r_sl = [4]


def khoi(tieu_de, dong):
    r = r_sl[0]
    ghi(sl, 'A%d' % r, tieu_de, f_muc, can=None, border=False)
    tieu_de_cot(sl, r + 1, 1, ['Thông số', 'Giá trị', 'Đơn vị', 'Ghi chú, nguồn'])
    for k, (nhan, gt, dv, ten_o, gc, fmt) in enumerate(dong):
        rr = r + 2 + k
        nhap = not (isinstance(gt, str) and gt.startswith('='))
        ghi(sl, 'A%d' % rr, nhan, f_dam, can=trai)
        ghi(sl, 'B%d' % rr, gt, f_nhap if nhap else f_dam, fmt, nen_nhap if nhap else nen_kq)
        ghi(sl, 'C%d' % rr, dv)
        ghi(sl, 'D%d' % rr, gc, f_ghi_chu, can=trai)
        if ten_o:
            ten(ten_o, 'SoLieu!$B$%d' % rr)
    r_sl[0] = r + 3 + len(dong)


khoi('A. TRÀN CHÍNH HIỆN TRẠNG (file điều tiết lũ, SoLieu mục B)', [
    ('Cao trình ngưỡng tràn Z_ngưỡng', ts['TC_Z'], 'm', 'SL_Zn', '', '0.00'),
    ('Số khoang n', ts['TC_n'], 'khoang', 'SL_nk', '', '0'),
    ('Bề rộng một khoang b', ts['TC_b'], 'm', 'SL_b', '', '0.00'),
    ('Số trụ giữa n₁', ts['TC_n1'], 'trụ', 'SL_n1', '', '0'),
    ('Bề dày trụ giữa d₁', ts['TC_d1'], 'm', 'SL_d1', '', '0.00'),
    ('Hệ số hình dạng trụ giữa ξ₁', ts['TC_x1'], '', 'SL_x1', 'Số liệu SV (03/11/2026)', '0.00'),
    ('Số trụ chính ban đầu n₂', ts['TC_n2'], 'trụ', 'SL_n2', '', '0'),
    ('Bề dày trụ chính d₂', ts['TC_d2'], 'm', 'SL_d2', '', '0.00'),
    ('Hệ số hình dạng trụ chính ξ₂', ts['TC_x2'], '', 'SL_x2', 'Đầu trụ lượn tròn', '0.00'),
    ('Hệ số hình dạng mố bên ξmb', ts['TC_xmb'], '', 'SL_xmb', 'Mố bên lượn tròn', '0.00'),
    ('Hệ số lưu lượng m', ts['TC_m'], '', 'SL_m', 'Tràn thực dụng thấp; thiên an toàn so với m = 0,467 tra trong BCNCKT',
     '0.000'),
    ('Hệ số chảy ngập σn', ts['TC_s'], '', 'SL_sn', 'Chảy không ngập (sau ngưỡng là bậc hạ và dốc nước)', '0.00'),
    ('Tổng bề rộng các khoang Σb = n·b', '=SL_nk*SL_b', 'm', 'SL_Sb', '', '0.00'),
    ('Tổng hệ số hình dạng trụ Σξ = n₁ξ₁ + n₂ξ₂', '=SL_n1*SL_x1+SL_n2*SL_x2', '', 'SL_Sxi', '', '0.00'),
    ('Tổng bề rộng tràn kể trụ pin B', '=SL_nk*SL_b+SL_n1*SL_d1+SL_n2*SL_d2', 'm', 'SL_Btong', '', '0.00'),
])
khoi('B. NGƯỠNG TRÀN, KÊNH DẪN THƯỢNG LƯU', [
    ('Chiều cao ngưỡng so với sân trước P₁', TC['P1'], 'm', 'SL_P1', 'BCC mục 2.4.4', '0.00'),
    ('Chiều cao ngưỡng so với đáy dốc (bậc hạ) P₂', TC['P2'], 'm', 'SL_P2', 'BCC mục 2.4.4', '0.00'),
    ('Chiều dày đỉnh ngưỡng δ', TC['delta'], 'm', 'SL_delta', 'Mái thượng lưu 1:3,0; mái hạ lưu 1:1,75', '0.00'),
    ('Cao trình đáy kênh dẫn Z_đáy kênh = Z_ngưỡng − P₁', '=SL_Zn-SL_P1', 'm', 'SL_Zday', '', '0.00'),
    ('Cao trình đầu dốc Z_đầu dốc = Z_ngưỡng − P₂', '=SL_Zn-SL_P2', 'm', 'SL_Zdd', '', '0.00'),
    ('Hệ số động năng dòng tới gần α', T.ALPHA_TL, '', 'SL_aTL', 'Mẫu: α = 1 ÷ 1,1', '0.00'),
    ('Góc mở rộng tường cánh thượng lưu', T.GOC_TUONG, 'độ', 'SL_goc', 'Như mẫu', '0'),
    ('Chiều dài tường cánh L_s', T.L_S, 'm', 'SL_Ls', 'Giả thiết: L_s = (3 ÷ 5)H với cả lũ thiết kế và lũ kiểm tra – hồ sơ '
     'chưa có số đo tường cánh hiện trạng', '0.0'),
])
khoi('C. KẾT QUẢ ĐIỀU TIẾT LŨ (file điều tiết lũ, sheet KetQua – vận hành bình thường)', [
    ('MNLTK – Zmax lũ thiết kế P = 1%', ts['P1_BinhThuong']['Z'], 'm', 'SL_ZTK', 'KetQua!C4', '0.000'),
    ('q xả max tràn chính – lũ thiết kế', ts['P1_BinhThuong']['qtc'], 'm³/s', 'SL_qTK', 'KetQua!F4', '0.00'),
    ('q xả max cả hai tràn – lũ thiết kế', ts['P1_BinhThuong']['q'], 'm³/s', 'SL_qTK2', 'KetQua!E4', '0.00'),
    ('MNLKT – Zmax lũ kiểm tra P = 0,2%', ts['P02_BinhThuong']['Z'], 'm', 'SL_ZKT', 'KetQua!C6', '0.000'),
    ('q xả max tràn chính – lũ kiểm tra', ts['P02_BinhThuong']['qtc'], 'm³/s', 'SL_qKT', 'KetQua!F6', '0.00'),
    ('q xả max cả hai tràn – lũ kiểm tra', ts['P02_BinhThuong']['q'], 'm³/s', 'SL_qKT2', 'KetQua!E6', '0.00'),
])
khoi('D. DỐC NƯỚC', [
    ('Chiều dài dốc L_dốc', TC['L_doc'], 'm', 'SL_Ld', 'KTTV Bảng 1-2', '0.0'),
    ('Độ dốc đáy i', TC['i_doc'], '', 'SL_i', 'KTTV Bảng 1-2', '0.000'),
    ('Hệ số nhám n (bê tông)', T.N_NHAM, '', 'SL_nn', 'Như thuyết minh mục 4.2.3', '0.000'),
    ('Bề rộng đầu dốc B_đ', TC['b_doc_dau'], 'm', 'SL_Bd', 'PL Bảng PL1-5', '0.00'),
    ('Bề rộng đoạn lăng trụ B', TC['b_doc_cuoi'], 'm', 'SL_Bc', 'PL Bảng PL1-5', '0.00'),
    ('Chiều dài đoạn thu hẹp L_tt', TC['L_chuyen_tiep'], 'm', 'SL_Ltt', 'Bề rộng thu hẹp đều từ B_đ xuống B', '0.0'),
    ('Hệ số lưu tốc chảy qua ngưỡng xuống chân bậc φ', T.PHI, '', 'SL_phi', '', '0.00'),
    ('Hệ số động năng, động lượng trên dốc α', T.ALPHA, '', 'SL_a', 'Mẫu: α = 1', '0.00'),
    ('Độ vượt cao an toàn tường bên a', T.A_TUONG, 'm', 'SL_at', 'Mẫu: a = 0,5 m', '0.00'),
    ('Hệ số nền trong công thức Đômbrốpxki α', T.ALPHA_NEN, '', 'SL_anen', 'Nền sét 0,8 ÷ 1,0; nền cát, á cát 1,5 ÷ 2,0', '0.00'),
    ('Lưu tốc không xói [v]kx – chiều sâu 0,5 m', T.V_KX[0], 'm/s', 'SL_vkx1', 'Theo mẫu: TCVN 4118:2012, phụ lục I, Bảng I5 – '
     'BTCT M200', '0.0'),
    ('Lưu tốc không xói [v]kx – chiều sâu 1 m', T.V_KX[1], 'm/s', 'SL_vkx2', 'Như trên', '0.0'),
    ('Gia tốc trọng trường g', 9.81, 'm/s²', 'SL_g', '', '0.00'),
])
# E. chiều cao tường bên hiện trạng
r0 = r_sl[0]
ghi(sl, 'A%d' % r0, 'E. CHIỀU CAO TƯỜNG BÊN DỐC NƯỚC HIỆN TRẠNG (PL Bảng PL1-4; sau x = 40 m giả thiết)', f_muc, can=None,
    border=False)
tieu_de_cot(sl, r0 + 1, 1, ['x (m)', 'h tường (m)', '', 'Ghi chú'])
X_T = list(range(0, int(TC['L_doc']) + 1, 5))
for k, x in enumerate(X_T):
    rr = r0 + 2 + k
    ghi(sl, 'A%d' % rr, x, f_nhap, '0', nen_nhap)
    ht = TC['h_tuong'][k] if k < len(TC['h_tuong']) else TC['h_tuong'][-1]
    ghi(sl, 'B%d' % rr, ht, f_nhap, '0.00', nen_nhap)
    ghi(sl, 'D%d' % rr, 'Hồ sơ' if k < len(TC['h_tuong']) else 'Giả thiết bằng %s m' % ('%.2f' % TC['h_tuong'][-1])
        .replace('.', ','), f_ghi_chu, can=trai)
ten('SL_xt', 'SoLieu!$A$%d:$A$%d' % (r0 + 2, r0 + 1 + len(X_T)))
ten('SL_ht', 'SoLieu!$B$%d:$B$%d' % (r0 + 2, r0 + 1 + len(X_T)))
sl.freeze_panes = 'A4'
in_an(sl, ngang=False)

# ================================================================== KHẢ NĂNG THÁO
kt = wb.create_sheet('KhaNangThao')
rong(kt, (52, 30, 13, 13, 16, 52))
ghi(kt, 'A1', 'KIỂM TRA KHẢ NĂNG THÁO CỦA TRÀN CHÍNH (TCVN 9147:2012 – trình tự file mẫu)', f_tieu_de, can=None, border=False)
ghi(kt, 'A2', 'Q = σn·ε·m·Σb·√(2g)·H₀^1,5;  H₀ = H + α·V₀²/2g;  V₀ = q/Ω_TL;  Ω_TL = B_TL·H_TL;  B_TL = B + 2·L_s·tan14°;  '
              'ε = 1 − 0,2·(ξmb + Σξ)·H₀/(n·b);  δ = |Q_tt − q|/Q_tt·100 %', f_ghi_chu, can=None, border=False)
tieu_de_cot(kt, 4, 1, ['Đại lượng', 'Công thức', 'P = 1%', 'P = 0,2%', 'Đơn vị', 'Ghi chú'], cao=24)
DONG_KT = [
    ('Z', 'Mực nước lũ Z_max', 'Kết quả điều tiết lũ', ('=SL_ZTK', '=SL_ZKT'), 'm', '0.000', 'MNLTK, MNLKT'),
    ('q', 'Lưu lượng xả lớn nhất qua tràn chính q_xả max', 'Kết quả điều tiết lũ', ('=SL_qTK', '=SL_qKT'), 'm³/s', '0.00', ''),
    ('H', 'Cột nước trên ngưỡng H', 'Z − Z_ngưỡng', '={c}{Z}-SL_Zn', 'm', '0.000', ''),
    ('H2', '2H', '', '=2*{c}{H}', 'm', '0.00', 'Điều kiện tràn đỉnh rộng: (2 ÷ 3)H < δ < (8 ÷ 10)H'),
    ('H3', '3H', '', '=3*{c}{H}', 'm', '0.00', ''),
    ('de', 'Chiều dày đỉnh ngưỡng δ', '', '=SL_delta', 'm', '0.00', ''),
    ('ht', 'Hình thức ngưỡng', '', '=IF(SL_delta<{c}{H2},"Thực dụng (δ < 2H)","Đỉnh rộng")', '', None, 'Tràn thực dụng → '
     'm = 0,45'),
    ('Zd', 'Cao trình đáy kênh dẫn Z_đáy kênh', 'Z_ngưỡng − P₁', '=SL_Zday', 'm', '0.00', ''),
    ('HTL', 'Chiều sâu nước thượng lưu H_TL', 'Z − Z_đáy kênh', '={c}{Z}-{c}{Zd}', 'm', '0.000', ''),
    ('Ls1', 'L_s nhỏ nhất = 3H', '', '=3*{c}{H}', 'm', '0.00', 'Chiều dài tường cánh L_s = (3 ÷ 5)H'),
    ('Ls2', 'L_s lớn nhất = 5H', '', '=5*{c}{H}', 'm', '0.00', ''),
    ('Ls', 'L_s chọn', 'SoLieu', '=SL_Ls', 'm', '0.00', ''),
    ('BTL', 'Bề rộng kênh thượng lưu B_TL', 'B + 2·L_s·tan14°', '=SL_Btong+2*{c}{Ls}*TAN(RADIANS(SL_goc))', 'm', '0.00', ''),
    ('Om', 'Diện tích mặt cắt ướt thượng lưu Ω_TL', 'B_TL·H_TL', '={c}{BTL}*{c}{HTL}', 'm²', '0.00', ''),
    ('Om4', '4·Σb·H', '', '=4*SL_Sb*{c}{H}', 'm²', '0.00', ''),
    ('xet', 'Xét lưu tốc tới gần', 'Ω_TL so với 4·Σb·H',
     '=IF({c}{Om}<{c}{Om4},"Ω_TL < 4ΣbH → phải xét","Ω_TL > 4ΣbH → bỏ qua")', '', None, 'Điều 6.1.5 TCVN 9147:2012'),
    ('V0', 'Lưu tốc tới gần V₀', 'q/Ω_TL', '=IF({c}{Om}<{c}{Om4},{c}{q}/{c}{Om},0)', 'm/s', '0.000', ''),
    ('hv', 'Cột nước lưu tốc α·V₀²/2g', '', '=SL_aTL*{c}{V0}^2/(2*SL_g)', 'm', '0.0000', ''),
    ('H0', 'Cột nước toàn phần H₀', 'H + α·V₀²/2g', '={c}{H}+{c}{hv}', 'm', '0.000', ''),
    ('ep', 'Hệ số co hẹp bên ε', '1 − 0,2·(ξmb + Σξ)·H₀/(n·b)', '=1-0.2*(SL_xmb+SL_Sxi)*{c}{H0}/(SL_nk*SL_b)', '', '0.0000',
     'Mẫu dùng ε = (ΣB − Σd)/ΣB; ở đây theo TCVN 9147:2012 như khi điều tiết lũ'),
    ('m', 'Hệ số lưu lượng m', 'SoLieu', '=SL_m', '', '0.000', ''),
    ('sn', 'Hệ số ngập σn', 'SoLieu', '=SL_sn', '', '0.00', ''),
    ('Qtt', 'Lưu lượng tháo thực tế Q_tt', 'σn·ε·m·Σb·√(2g)·H₀^1,5',
     '={c}{sn}*{c}{ep}*{c}{m}*SL_Sb*SQRT(2*SL_g)*{c}{H0}^1.5', 'm³/s', '0.00', ''),
    ('ss', 'So sánh với q_xả max', '',
     '=IF({c}{Qtt}>={c}{q},"Q_tt > q_xả max → thỏa mãn","Q_tt < q_xả max → tính lại điều tiết lũ")', '', None, ''),
    ('dl', 'Sai số δ', '|Q_tt − q|/Q_tt·100', '=ABS({c}{Qtt}-{c}{q})/{c}{Qtt}*100', '%', '0.0', ''),
    ('nx', 'Nhận xét', '',
     '=IF({c}{Qtt}<{c}{q},"Tính lại điều tiết lũ",IF({c}{dl}<=5,"δ ≤ 5 %: kết quả điều tiết chính xác",'
     '"Q_tt > q: kết quả điều tiết thiên an toàn"))', '', None, 'Kể V₀ khi điều tiết: MNLTK thấp hơn 2,4 cm, MNLKT 3,3 cm '
     '(thuy_luc_tran_chinh)'),
]
DK = {k: 5 + j for j, (k, *_) in enumerate(DONG_KT)}
for j, (k, nhan, ct, gt, dv, fmt, gc) in enumerate(DONG_KT):
    r = 5 + j
    ghi(kt, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(kt, 'B%d' % r, ct, f_ghi_chu, can=trai)
    for cc, g_ in zip('CD', gt if isinstance(gt, tuple) else (gt, gt)):
        ghi(kt, '%s%d' % (cc, r), g_.format(c=cc, **DK), f_dam if k in ('Qtt', 'H0', 'ep') else f_thuong, fmt, nen_kq,
            giua if fmt is None else cg)
    ghi(kt, 'E%d' % r, dv)
    ghi(kt, 'F%d' % r, gc, f_ghi_chu, can=trai)
    if fmt is None:
        kt.row_dimensions[r].height = 30
r = 5 + len(DONG_KT) + 1
ghi(kt, 'A%d' % r, 'BẢNG TỔNG HỢP', f_muc, can=None, border=False)
tieu_de_cot(kt, r + 1, 1, ['Tần suất', 'H_tr (m)', 'q_xả max tràn chính (m³/s)', 'q_xả max cả hai tràn (m³/s)',
                           'Mực nước lũ (m)'], cao=36)
for j, (nhan, cc, q2, ten_mn) in enumerate((('P = 1%', 'C', 'SL_qTK2', 'MNLTK'), ('P = 0,2%', 'D', 'SL_qKT2', 'MNLKT'))):
    rr = r + 2 + j
    ghi(kt, 'A%d' % rr, nhan, f_dam)
    ghi(kt, 'B%d' % rr, '=%s%d' % (cc, DK['H']), fmt='0.00', fill=nen_kq)
    ghi(kt, 'C%d' % rr, '=%s%d' % (cc, DK['q']), fmt='0.00', fill=nen_kq)
    ghi(kt, 'D%d' % rr, '=%s' % q2, fmt='0.00', fill=nen_kq)
    ghi(kt, 'E%d' % rr, '="%s = +"&FIXED(%s%d,2)' % (ten_mn, cc, DK['Z']), fill=nen_kq)
in_an(kt, ngang=False)

# ================================================================== CẤP LƯU LƯỢNG: ĐỘ SÂU ĐẦU DỐC, i_k, h_0, ĐỊNH TÍNH
cq = wb.create_sheet('CapLuuLuong')
rong(cq, (14, 10, 11, 10, 10, 10, 10, 10, 10, 22, 3) + (9,) * 16)
ghi(cq, 'A1', 'CÁC CẤP LƯU LƯỢNG – ĐỘ SÂU ĐẦU DỐC, ĐỘ DỐC PHÂN GIỚI, ĐỘ SÂU DÒNG ĐỀU, ĐỊNH TÍNH ĐƯỜNG MẶT NƯỚC', f_tieu_de,
    can=None, border=False)
ghi(cq, 'A2', 'Cột phụ màu xám (từ cột L): phép lặp giải H₀ theo Q (8 vòng), h_c (6 vòng), h₀ (6 vòng) – không cần sửa.',
    f_ghi_chu, can=None, border=False)
# --- 1. độ sâu đầu dốc
ghi(cq, 'A4', '1. Các cấp lưu lượng và độ sâu đầu dốc (độ sâu co hẹp ở chân bậc, B = B_đ)', f_muc, can=None, border=False)
tieu_de_cot(cq, 5, 1, ['Cấp', 'Hệ số k', 'Q (m³/s)', 'H₀ (m)', 'E₀ = P₂ + H₀ (m)', 'h_c (m)', 'q = Q/B_đ (m²/s)',
                       'h_k (m)', 'Fr = v²/(g·h_c)', 'Trạng thái'], cao=45)
tieu_de_cot(cq, 5, 12, ['H₀ lặp %d' % k for k in range(8)] + ['h_c lặp %d' % k for k in range(6)], cao=45)
R1 = {}
for j, (nhan, k, _) in enumerate(CAP):
    r = 6 + j
    R1[nhan] = r
    ghi(cq, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(cq, 'B%d' % r, k if k is not None else '=C%d/SL_qTK' % r, f_nhap if k is not None else f_thuong, '0.00',
        nen_nhap if k is not None else nen_kq)
    ghi(cq, 'C%d' % r, ('=B%d*SL_qTK' % r) if k is not None else '=SL_qKT', f_dam, '0.00', nen_kq)
    # H₀ theo Q: H₀ = [Q/(σn·ε(H₀)·m·Σb·√2g)]^(2/3)
    for it in range(8):
        col = get_column_letter(12 + it)
        prev = get_column_letter(11 + it)
        e_ = '1' if it == 0 else '(1-0.2*(SL_xmb+SL_Sxi)*%s%d/(SL_nk*SL_b))' % (prev, r)
        ghi(cq, '%s%d' % (col, r), '=(C{r}/(SL_sn*{e}*SL_m*SL_Sb*SQRT(2*SL_g)))^(2/3)'.format(r=r, e=e_), f_phu, '0.0000',
            nen_phu)
    ghi(cq, 'D%d' % r, '=S%d' % r, fmt='0.000', fill=nen_kq)
    ghi(cq, 'E%d' % r, '=SL_P2+D%d' % r, fmt='0.000', fill=nen_kq)
    # h_c: h = Q/(φ·B_đ·√(2g(E₀ − h)))
    for it in range(6):
        col = get_column_letter(20 + it)
        prev = get_column_letter(19 + it)
        hp = '0' if it == 0 else '%s%d' % (prev, r)
        ghi(cq, '%s%d' % (col, r), '=C{r}/(SL_phi*SL_Bd*SQRT(2*SL_g*(E{r}-{hp})))'.format(r=r, hp=hp), f_phu, '0.0000',
            nen_phu)
    ghi(cq, 'F%d' % r, '=Y%d' % r, f_dam, '0.000', nen_kq)
    ghi(cq, 'G%d' % r, '=C%d/SL_Bd' % r, fmt='0.00', fill=nen_kq)
    ghi(cq, 'H%d' % r, '=(SL_a*G%d^2/SL_g)^(1/3)' % r, fmt='0.000', fill=nen_kq)
    ghi(cq, 'I%d' % r, '=(C{r}/(SL_Bd*F{r}))^2/(SL_g*F{r})'.format(r=r), fmt='0.00', fill=nen_kq)
    ghi(cq, 'J%d' % r, '=IF(F{r}<H{r},"h_c < h_k: dòng xiết","h_c ≥ h_k")'.format(r=r), fill=nen_kq, can=giua)
ten('Q_cap', 'CapLuuLuong!$C$6:$C$10')
ten('hc_cap', 'CapLuuLuong!$F$6:$F$10')
# --- 2. độ dốc phân giới
ghi(cq, 'A12', '2. Độ dốc phân giới i_k = Q²/(ω_k²·C_k²·R_k) – đoạn lăng trụ B = 18 m (cột J: i_k với B_đ = 31 m)', f_muc,
    can=None, border=False)
tieu_de_cot(cq, 13, 1, ['Cấp', 'Q (m³/s)', 'q (m²/s)', 'h_k (m)', 'ω_k (m²)', 'χ_k (m)', 'R_k (m)', 'C_k (m^0,5/s)', 'i_k',
                        'i_k (B = B_đ)'], cao=36)
for j, (nhan, *_) in enumerate(CAP):
    r = 14 + j
    r1 = R1[nhan]
    ghi(cq, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(cq, 'B%d' % r, '=C%d' % r1, fmt='0.00', fill=nen_kq)
    ghi(cq, 'C%d' % r, '=B%d/SL_Bc' % r, fmt='0.00', fill=nen_kq)
    ghi(cq, 'D%d' % r, '=(SL_a*C%d^2/SL_g)^(1/3)' % r, fmt='0.000', fill=nen_kq)
    ghi(cq, 'E%d' % r, '=SL_Bc*D%d' % r, fmt='0.00', fill=nen_kq)
    ghi(cq, 'F%d' % r, '=SL_Bc+2*D%d' % r, fmt='0.00', fill=nen_kq)
    ghi(cq, 'G%d' % r, '=E%d/F%d' % (r, r), fmt='0.000', fill=nen_kq)
    ghi(cq, 'H%d' % r, '=G%d^(1/6)/SL_nn' % r, fmt='0.00', fill=nen_kq)
    ghi(cq, 'I%d' % r, '=B{r}^2/(E{r}^2*H{r}^2*G{r})'.format(r=r), f_dam, '0.00000', nen_kq)
    ghi(cq, 'J%d' % r, '=B{r}^2/((SL_Bd*H{r1})^2*((SL_Bd*H{r1}/(SL_Bd+2*H{r1}))^(1/6)/SL_nn)^2*(SL_Bd*H{r1}/(SL_Bd+2*H{r1})))'
        .format(r=r, r1=r1), fmt='0.00000', fill=nen_kq)
# --- 3. độ sâu dòng đều
ghi(cq, 'A20', '3. Độ sâu dòng đều h₀ (Agơrốtskin: f(R_ln) = 4m₀√i/Q, m = 0 → 4m₀ = 8; R_ln từ Q = 4m₀√i·R_ln^(8/3)/n; '
               'h₀ giải trực tiếp – cột phụ)', f_muc, can=None, border=False)
tieu_de_cot(cq, 21, 1, ['Cấp', 'Q (m³/s)', 'f(R_ln)', 'R_ln (m)', 'B (m)', 'B/R_ln', 'h₀ (m)', 'h₀/R_ln', 'v₀ (m/s)'], cao=36)
tieu_de_cot(cq, 21, 12, ['h₀ lặp %d' % k for k in range(6)], cao=36)
for j, (nhan, *_) in enumerate(CAP):
    r = 22 + j
    ghi(cq, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(cq, 'B%d' % r, '=C%d' % R1[nhan], fmt='0.00', fill=nen_kq)
    ghi(cq, 'C%d' % r, '=8*SQRT(SL_i)/B%d' % r, fmt='0.0000', fill=nen_kq)
    ghi(cq, 'D%d' % r, '=(SL_nn/C%d)^(3/8)' % r, fmt='0.000', fill=nen_kq)
    ghi(cq, 'E%d' % r, '=SL_Bc', fmt='0.00', fill=nen_kq)
    ghi(cq, 'F%d' % r, '=E%d/D%d' % (r, r), fmt='0.00', fill=nen_kq)
    # Q = B·h·(1/n)·R^(2/3)·√i → h = (Q·n/(B·√i))^(3/5)·((B + 2h)/B)^(2/5)
    for it in range(6):
        col = get_column_letter(12 + it)
        prev = get_column_letter(11 + it)
        nhan_ = '1' if it == 0 else '((SL_Bc+2*%s%d)/SL_Bc)^(2/5)' % (prev, r)
        ghi(cq, '%s%d' % (col, r), '=(B{r}*SL_nn/(SL_Bc*SQRT(SL_i)))^(3/5)*{k}'.format(r=r, k=nhan_), f_phu, '0.0000', nen_phu)
    ghi(cq, 'G%d' % r, '=Q%d' % r, f_dam, '0.000', nen_kq)
    ghi(cq, 'H%d' % r, '=G%d/D%d' % (r, r), fmt='0.000', fill=nen_kq)
    ghi(cq, 'I%d' % r, '=B%d/(SL_Bc*G%d)' % (r, r), fmt='0.00', fill=nen_kq)
# --- 4. định tính
ghi(cq, 'A28', '4. Định tính đường mặt nước trên đoạn lăng trụ', f_muc, can=None, border=False)
tieu_de_cot(cq, 29, 1, ['Cấp', 'Q (m³/s)', 'i', 'i_k', 'h_k (m)', 'h₀ (m)', 'h tại x = L_tt (m)', 'So sánh', 'Kết luận'], cao=36)
for j, (nhan, _, mn) in enumerate(CAP):
    r = 30 + j
    ghi(cq, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(cq, 'B%d' % r, '=C%d' % R1[nhan], fmt='0.00', fill=nen_kq)
    ghi(cq, 'C%d' % r, '=SL_i', fmt='0.000', fill=nen_kq)
    ghi(cq, 'D%d' % r, '=I%d' % (14 + j), fmt='0.0000', fill=nen_kq)
    ghi(cq, 'E%d' % r, '=D%d' % (14 + j), fmt='0.000', fill=nen_kq)
    ghi(cq, 'F%d' % r, '=G%d' % (22 + j), fmt='0.000', fill=nen_kq)
    ghi(cq, 'G%d' % r, '=INDEX(%s!$C$6:$C$%d,MATCH(SL_Ltt+0.0001,%s!$O$6:$O$%d,1))' % (sh(mn), 5 + NR, sh(mn), 5 + NR),
        fmt='0.000', fill=nen_kq)
    ghi(cq, 'H%d' % r, '=IF(C{r}>D{r},"i > i_k","i < i_k")&IF(AND(F{r}<G{r},G{r}<E{r}),"; h₀ < h < h_k","")'.format(r=r),
        fill=nen_kq, can=giua)
    ghi(cq, 'I%d' % r, '=IF(AND(C{r}>D{r},F{r}<G{r},G{r}<E{r}),"Đường nước đổ b_II",IF(AND(C{r}>D{r},G{r}<=F{r}),'
                       '"Đường nước dâng c_II","Kiểm tra lại"))'.format(r=r), f_dam, fill=nen_kq, can=giua)
in_an(cq)

# ================================================================== ĐƯỜNG MẶT NƯỚC TRÊN DỐC (mỗi cấp lưu lượng một sheet)
HDR15 = ['MC', 'B (m)', 'h (m)', 'v (m/s)', 'Э (m)', 'ΔЭ (m)', 'ω (m²)', 'χ (m)', 'R (m)', 'C', 'J', 'J_tb', 'i − J_tb',
         'ΔL = ΔЭ/(i − J_tb) (m)', 'L (m)']
for j, (nhan, _, mn) in enumerate(CAP):
    ws = wb.create_sheet(mn)
    rong(ws, (5, 8, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 9, 11, 8, 3, 9, 10) + (10,) * 6 + (14,))
    r1 = R1[nhan]
    ghi(ws, 'A1', 'ĐƯỜNG MẶT NƯỚC TRÊN DỐC NƯỚC – %s (phương pháp cộng trực tiếp, chọn trước ΔL, thử dần h)' % nhan,
        f_tieu_de, can=None, border=False)
    for c_, nh, ct, fmt in (('A', 'Q (m³/s) =', '=CapLuuLuong!C%d' % r1, '0.00'), ('E', 'h đầu dốc h_c (m) =',
                                                                               '=CapLuuLuong!F%d' % r1, '0.000')):
        ghi(ws, '%s3' % c_, nh, f_dam, can=None, border=False)
    ghi(ws, 'C3', '=CapLuuLuong!C%d' % r1, f_dam, '0.00', nen_kq)
    ghi(ws, 'G3', '=CapLuuLuong!F%d' % r1, f_dam, '0.000', nen_kq)
    ghi(ws, 'I3', '="i = "&FIXED(SL_i*100,0)&" %;  n = "&FIXED(SL_nn,3)&";  B = "&FIXED(SL_Bd,0)&" → "&FIXED(SL_Bc,0)&" m '
                  '(x = 0 ÷ "&FIXED(SL_Ltt,0)&" m), "&FIXED(SL_Bc,0)&" m sau đó"', f_ghi_chu, can=None, border=False)
    tieu_de_cot(ws, 5, 1, HDR15, cao=36)
    tieu_de_cot(ws, 5, 17, ['ΔL chọn (m)', 'K = Э₁ + i·ΔL − J₁·ΔL/2'] + ['h lặp %d' % k for k in range(6)] +
                ['Kiểm tra Э₁ + i·ΔL − Э₂ − J_tb·ΔL'], cao=36)
    for k in range(NR):
        r = 6 + k                     # hàng 6 ÷ 6 + NR − 1
        p = r - 1
        ct = {'A': k + 1,
              'B': '=IF(O{r}<=SL_Ltt,SL_Bd+(SL_Bc-SL_Bd)*O{r}/SL_Ltt,SL_Bc)'.format(r=r),
              'D': '=$C$3/(B{r}*C{r})'.format(r=r),
              'E': '=C{r}+SL_a*D{r}^2/(2*SL_g)'.format(r=r),
              'G': '=B{r}*C{r}'.format(r=r),
              'H': '=B{r}+2*C{r}'.format(r=r),
              'I': '=G{r}/H{r}'.format(r=r),
              'J': '=I{r}^(1/6)/SL_nn'.format(r=r),
              'K': '=D{r}^2/(J{r}^2*I{r})'.format(r=r)}
        if k == 0:
            ct.update({'C': '=$G$3', 'O': 0})
        else:
            ct.update({
                'Q': '=IF(O{p}<SL_Ltt-0.000001,SL_Ltt/{n},(SL_Ld-SL_Ltt)/{n})'.format(p=p, n=N_B),
                'R': '=E{p}+SL_i*Q{r}-0.5*K{p}*Q{r}'.format(p=p, r=r),
                'O': '=O{p}+Q{r}'.format(p=p, r=r),
                'C': '=X%d' % r,
                'F': '=E{r}-E{p}'.format(r=r, p=p),
                'L': '=(K{p}+K{r})/2'.format(p=p, r=r),
                'M': '=SL_i-L%d' % r,
                'N': '=F{r}/M{r}'.format(r=r),
                'Y': '=E{p}+SL_i*Q{r}-E{r}-L{r}*Q{r}'.format(p=p, r=r)})
            # lặp: h = Q/(B·√(2g/α·(K − h − J(h)·ΔL/2)))
            for it in range(6):
                col = get_column_letter(19 + it)
                hp = ('C%d' % p) if it == 0 else '%s%d' % (get_column_letter(18 + it), r)
                ct[col] = ('=$C$3/(B{r}*SQRT(2*SL_g/SL_a*(R{r}-{h}-0.5*Q{r}*(($C$3/(B{r}*{h}))^2*SL_nn^2/'
                           '((B{r}*{h}/(B{r}+2*{h}))^(4/3))))))'.format(r=r, h=hp))
        fm = dict(A='0', B='0.00', C='0.000', D='0.00', E='0.000', F='0.000', G='0.00', H='0.00', I='0.000', J='0.00',
                  K='0.0000', L='0.0000', M='0.0000', N='0.00', O='0.00', Q='0.00', R='0.0000', Y='0.0E+00')
        for col, f_ in ct.items():
            phu = col >= 'Q'
            ghi(ws, '%s%d' % (col, r), f_, f_phu if phu else (f_dam if col in 'CO' else f_thuong), fm.get(col, '0.0000'),
                nen_phu if phu else None)
        for col in 'ABCDEFGHIJKLMNO':
            ws['%s%d' % (col, r)].border = vien
    ghi(ws, 'A%d' % (7 + NR), 'Cột (14) tính theo công thức của mẫu ΔL = ΔЭ/(i − J_tb), đúng bằng ΔL chọn (cột Q) khi h đã thử '
                              'đúng; cột Y kiểm tra phương trình năng lượng (≈ 0).', f_ghi_chu, can=None, border=False)
    ws.freeze_panes = 'B6'
    in_an(ws)

# ================================================================== KẾT QUẢ DỐC NƯỚC: CUỐI DỐC, XÓI, TƯỜNG BÊN, BẢN ĐÁY
kd = wb.create_sheet('KetQuaDoc')
rong(kd, (14, 11, 10, 11, 11, 11, 13, 13, 15, 11, 3, 9, 11, 11, 11, 13, 13))
ghi(kd, 'A1', 'KẾT QUẢ THỦY LỰC DỐC NƯỚC – CUỐI DỐC, KIỂM TRA XÓI, TƯỜNG BÊN, BẢN ĐÁY', f_tieu_de, can=None, border=False)
RC = 6 + NR - 1                        # hàng mặt cắt cuối dốc trong sheet MN
ghi(kd, 'A3', '1. Chiều sâu, lưu tốc cuối dốc và kiểm tra không xói v < [v]kx', f_muc, can=None, border=False)
tieu_de_cot(kd, 4, 1, ['Cấp', 'Q (m³/s)', 'L (m)', 'B cuối dốc (m)', 'h cuối dốc (m)', 'v cuối dốc (m/s)', 'Fr = v/√(g·h)',
                       '[v]kx (m/s)', 'Kiểm tra'], cao=36)
for j, (nhan, _, mn) in enumerate(CAP):
    r = 5 + j
    ghi(kd, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(kd, 'B%d' % r, '=%s!$C$3' % sh(mn), fmt='0.00', fill=nen_kq)
    ghi(kd, 'C%d' % r, '=%s!O%d' % (sh(mn), RC), fmt='0.00', fill=nen_kq)
    ghi(kd, 'D%d' % r, '=%s!B%d' % (sh(mn), RC), fmt='0.00', fill=nen_kq)
    ghi(kd, 'E%d' % r, '=%s!C%d' % (sh(mn), RC), f_dam, '0.000', nen_kq)
    ghi(kd, 'F%d' % r, '=%s!D%d' % (sh(mn), RC), f_dam, '0.00', nen_kq)
    ghi(kd, 'G%d' % r, '=F%d/SQRT(SL_g*E%d)' % (r, r), fmt='0.00', fill=nen_kq)
    ghi(kd, 'H%d' % r, '=FIXED(SL_vkx1,1)&" ÷ "&FIXED(SL_vkx2,1)', fill=nen_kq)
    ghi(kd, 'I%d' % r, '=IF(F%d<SL_vkx1,"Không xói","Kiểm tra lại")' % r, f_dam, fill=nen_kq)
RT0 = 12
ghi(kd, 'A%d' % (RT0 - 1), '2. Chiều cao tường bên – lũ kiểm tra: h_hk = h·(1 + v/100); h_t = h_hk + a', f_muc, can=None,
    border=False)
tieu_de_cot(kd, RT0, 1, ['x (m)', 'B (m)', 'h (m)', 'v (m/s)', 'h_hk (m)', 'h_t yêu cầu (m)', 'h_t hiện trạng (m)',
                         'Độ vượt cao h_t − h_hk (m)', 'Đánh giá', 'Cao trình đáy (m)'], cao=45)
MN_KT, MN_TK = CAP[-1][2], CAP[-2][2]
for k, x in enumerate(X_T):
    r = RT0 + 1 + k
    tim = 'MATCH(A{r}+0.0001,{w}!$O$6:$O${e},1)'.format(r=r, w=sh(MN_KT), e=RC)
    ghi(kd, 'A%d' % r, '=INDEX(SL_xt,%d)' % (k + 1), f_dam, '0')
    ghi(kd, 'B%d' % r, '=INDEX(%s!$B$6:$B$%d,%s)' % (sh(MN_KT), RC, tim), fmt='0.00', fill=nen_kq)
    ghi(kd, 'C%d' % r, '=INDEX(%s!$C$6:$C$%d,%s)' % (sh(MN_KT), RC, tim), fmt='0.000', fill=nen_kq)
    ghi(kd, 'D%d' % r, '=INDEX(%s!$D$6:$D$%d,%s)' % (sh(MN_KT), RC, tim), fmt='0.00', fill=nen_kq)
    ghi(kd, 'E%d' % r, '=C{r}*(1+D{r}/100)'.format(r=r), fmt='0.000', fill=nen_kq)
    ghi(kd, 'F%d' % r, '=E%d+SL_at' % r, f_dam, '0.00', nen_kq)
    ghi(kd, 'G%d' % r, '=INDEX(SL_ht,%d)' % (k + 1), fmt='0.00', fill=nen_kq)
    ghi(kd, 'H%d' % r, '=G%d-E%d' % (r, r), f_dam, '0.00', nen_kq)
    ghi(kd, 'I%d' % r, '=IF(H%d>=SL_at,"Đạt","Không đạt – nâng tường")' % r, f_dam, fill=nen_kq)
    ghi(kd, 'J%d' % r, '=SL_Zdd-SL_i*A%d' % r, fmt='0.00', fill=nen_kq)
RT1 = RT0 + len(X_T)
ghi(kd, 'A%d' % (RT1 + 1), '=IF(COUNTIF(I%d:I%d,"Đạt")=%d,"Tường bên đủ chiều cao ở mọi mặt cắt; độ vượt cao nhỏ nhất "&'
                           'FIXED(MIN(H%d:H%d),2)&" m","Có mặt cắt tường bên không đủ chiều cao")'
    % (RT0 + 1, RT1, len(X_T), RT0 + 1, RT1), f_dam, can=None, border=False)
RB = RT1 + 4
ghi(kd, 'A%d' % (RB - 1), '3. Chiều dày bản đáy dốc nước (Đômbrốpxki): t = (0,03 ÷ 0,035)·α·v_tb·√h_tb – lũ kiểm tra', f_muc,
    can=None, border=False)
tieu_de_cot(kd, RB, 1, ['Q (m³/s)', 'v đầu (m/s)', 'v cuối (m/s)', 'v_tb (m/s)', 'h đầu (m)', 'h cuối (m)', 'h_tb (m)',
                        't (0,03) (m)', 't (0,035) (m)'], cao=36)
r = RB + 1
for col, ct, fmt in (('A', '=%s!$C$3' % sh(MN_KT), '0.00'), ('B', '=%s!D6' % sh(MN_KT), '0.00'),
                     ('C', '=%s!D%d' % (sh(MN_KT), RC), '0.00'), ('D', '=(B{r}+C{r})/2'.format(r=r), '0.00'),
                     ('E', '=%s!C6' % sh(MN_KT), '0.000'), ('F', '=%s!C%d' % (sh(MN_KT), RC), '0.000'),
                     ('G', '=(E{r}+F{r})/2'.format(r=r), '0.000'), ('H', '=0.03*SL_anen*D{r}*SQRT(G{r})'.format(r=r), '0.00'),
                     ('I', '=0.035*SL_anen*D{r}*SQRT(G{r})'.format(r=r), '0.00')):
    ghi(kd, '%s%d' % (col, r), ct, f_dam if col in 'HI' else f_thuong, fmt, nen_kq)
ghi(kd, 'A%d' % (r + 1), 'Đối chiếu với chiều dày bản đáy dốc nước hiện trạng theo hồ sơ hoàn công (chưa có trong hồ sơ hiện có).',
    f_ghi_chu, can=None, border=False)
# số liệu vẽ hình: cao trình đáy, mặt nước lũ thiết kế, lũ kiểm tra (có, không kể hàm khí) dọc dốc
RV = r + 4
ghi(kd, 'L%d' % (RV - 1), 'Số liệu vẽ biểu đồ (cao trình, m)', f_muc, can=None, border=False)
tieu_de_cot(kd, RV, 12, ['x (m)', 'Đáy dốc', 'Mặt nước lũ thiết kế', 'Mặt nước lũ kiểm tra', 'Lũ kiểm tra có kể hàm khí',
                         'Đỉnh tường hiện trạng'], cao=36)
for k in range(NR):
    rr = RV + 1 + k
    ghi(kd, 'L%d' % rr, '=%s!O%d' % (sh(MN_KT), 6 + k), fmt='0.0')
    ghi(kd, 'M%d' % rr, '=SL_Zdd-SL_i*L%d' % rr, fmt='0.00')
    ghi(kd, 'N%d' % rr, '=M%d+%s!C%d' % (rr, sh(MN_TK), 6 + k), fmt='0.00')
    ghi(kd, 'O%d' % rr, '=M%d+%s!C%d' % (rr, sh(MN_KT), 6 + k), fmt='0.00')
    ghi(kd, 'P%d' % rr, '=M{rr}+{w}!C{k}*(1+{w}!D{k}/100)'.format(rr=rr, w=sh(MN_KT), k=6 + k), fmt='0.00')
for k in range(len(X_T)):
    rr = RV + 1 + k
    ghi(kd, 'Q%d' % rr, '=J{r}+G{r}'.format(r=RT0 + 1 + k), fmt='0.00')
ch = ScatterChart()
ch.title = 'Đường mặt nước trên dốc tràn chính và đỉnh tường bên hiện trạng'
ch.style = 13
ch.x_axis.title, ch.y_axis.title = 'x (m)', 'Cao trình (m)'
ch.x_axis.delete = ch.y_axis.delete = False
ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0, TC['L_doc'], 10
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 68, 79
ch.width, ch.height = 22, 11
ch.legend.position = 'b'
xr = Reference(kd, min_col=12, min_row=RV + 1, max_row=RV + NR)
for col, nhan, mau, net in ((13, 'Đáy dốc', '52514E', None), (14, 'Mặt nước lũ thiết kế', '2A78D6', None),
                            (15, 'Mặt nước lũ kiểm tra', 'EB6834', None), (16, 'Lũ kiểm tra có kể hàm khí', 'EB6834', 'dash')):
    s = Series(Reference(kd, min_col=col, min_row=RV + 1, max_row=RV + NR), xr, title=nhan)
    s.smooth = False
    s.marker.symbol = 'none'
    s.graphicalProperties.line.solidFill = mau
    s.graphicalProperties.line.width = 22225
    if net:
        s.graphicalProperties.line.dashStyle = net
    ch.series.append(s)
s = Series(Reference(kd, min_col=17, min_row=RV + 1, max_row=RV + len(X_T)),
           Reference(kd, min_col=1, min_row=RT0 + 1, max_row=RT1), title='Đỉnh tường bên hiện trạng')
s.smooth = False
s.marker.symbol = 'circle'
s.marker.size = 5
s.graphicalProperties.line.solidFill = '0B0B0B'
s.graphicalProperties.line.dashStyle = 'sysDot'
s.marker.graphicalProperties.solidFill = s.marker.graphicalProperties.line.solidFill = '0B0B0B'
ch.series.append(s)
kd.add_chart(ch, 'L3')
in_an(kd)

# ================================================================== HƯỚNG DẪN
rong(hd, (3, 125))
hd.sheet_view.showGridLines = False
ND = [
    ('T', 'THỦY LỰC TRÀN CHÍNH HỒ CÀ GIÂY – TÍNH THEO TRÌNH TỰ FILE MẪU “THIẾT KẾ ĐƯỜNG TRÀN THÁO LŨ”'),
    ('g', 'Cùng số với Word ThuyLuc_TranChinh_TheoMau.docx. Ô nền vàng chữ xanh: số liệu nhập; chữ đen: công thức; cột xám: '
          'phép lặp giải các bước thử dần (không cần sửa).'),
    ('M', 'Các sheet'),
    ('-', 'SoLieu: thông số tràn chính (mục A – từ file điều tiết lũ), ngưỡng và kênh dẫn (B), kết quả điều tiết lũ (C), dốc '
          'nước (D), chiều cao tường bên hiện trạng (E). Khi điều tiết lũ thay đổi, sửa các ô vàng mục A, C theo file '
          'DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx (sheet SoLieu, KetQua).'),
    ('-', 'KhaNangThao: kiểm tra khả năng tháo (TCVN 9147:2012) ở MNLTK, MNLKT – hình thức ngưỡng, lưu tốc tới gần (Ω_TL so với '
          '4·Σb·H), H₀, ε, Q_tt so với q xả max, sai số δ; bảng tổng hợp.'),
    ('-', 'CapLuuLuong: 5 cấp lưu lượng (0,4Q; 0,6Q; 0,8Q; Q lũ thiết kế; Q lũ kiểm tra), độ sâu đầu dốc h_c (bậc hạ P₂ sau '
          'ngưỡng), h_k, i_k, độ sâu dòng đều h₀ theo Agơrốtskin, định tính đường mặt nước.'),
    ('-', 'MN_04Q … MN_Q02: đường mặt nước trên dốc – 15 cột như mẫu; chọn trước ΔL (2,5 m trên đoạn thu hẹp, 5 m trên đoạn '
          'lăng trụ), độ sâu cuối đoạn thử dần (cột S ÷ X), cột (14) ΔL = ΔЭ/(i − J_tb) tính lại bằng ΔL chọn, cột Y kiểm tra '
          'phương trình năng lượng.'),
    ('-', 'KetQuaDoc: chiều sâu, lưu tốc cuối dốc và kiểm tra không xói; chiều cao tường bên với lũ kiểm tra (hàm khí '
          'h_hk = h(1 + v/100), h_t = h_hk + a) so với tường hiện trạng; chiều dày bản đáy (Đômbrốpxki); biểu đồ.'),
    ('M', 'Khác mẫu (do tràn Cà Giây khác tràn trong mẫu)'),
    ('-', 'Ngưỡng thực dụng (δ = 0,5 m < 2H), sau ngưỡng có bậc hạ P₂ = 1,40 m: độ sâu đầu dốc là độ sâu co hẹp h_c '
          '(Q = φ·B_đ·h_c·√(2g(E₀ − h_c)), E₀ = P₂ + H₀) thay cho h_đd = h_k của tràn đỉnh rộng.'),
    ('-', 'ε theo TCVN 9147:2012 có hệ số hình dạng trụ (như file điều tiết lũ) thay cho ε = (ΣB − Σd)/ΣB; cấp lưu lượng '
          'lớn nhất là lũ kiểm tra (1,47Q) thay cho 1,2Q; dốc thu hẹp 31 → 18 m trên 20 m đầu nên chọn trước ΔL và thử dần h.'),
    ('M', 'Giả thiết cần xác nhận'),
    ('-', 'L_s = 12 m (tường cánh thượng lưu); [v]kx = 19,2 ÷ 21,1 m/s theo mẫu (TCVN 4118:2012 Bảng I5); tường bên sau '
          'x = 40 m cao 2,5 m; α nền = 1,0; chiều dày bản đáy hiện trạng chưa có. Phần tiêu năng mũi phun của mẫu chưa tính '
          '(thiếu số liệu mũi phun, kênh xả hạ lưu).'),
]
for h, (loai, txt) in enumerate(ND, 1):
    o = hd['B%d' % h]
    o.value = ('•  ' + txt) if loai == '-' else txt
    o.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    o.font = {'T': f_tieu_de, 'M': f_muc, 'g': f_ghi_chu}.get(loai, f_thuong)
    hd.row_dimensions[h].height = 22 if loai in 'TM' else 15.5 * max(1, -(-len(txt) // 140))
in_an(hd, ngang=False)

MAU = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'KhaNangThao': '1BAF7A', 'CapLuuLuong': '2A78D6', 'KetQuaDoc': '1BAF7A'}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = MAU.get(ws.title, '8A8984')
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
