# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL: KIỂM TRA TRÀN CHÍNH VÀ TRÀN SỰ CỐ HIỆN TRẠNG TRONG ĐỒ ÁN SỬA CHỮA (theo hồ sơ kiểm định an toàn đập – QTVH
chương 6). Cùng nội dung, cùng số với Word TAI_VE/KiemTra_TranChinh_SuaChua.docx (kiem_tra_tran_chinh.py, kiem_tra_tran_su_co.py
– sheet TSC_… do SV yêu cầu 08/11/2026). Mọi ô kết quả là công thức;
ô nền vàng chữ xanh là số liệu nhập. Các bước thử dần (H theo Q, độ sâu co hẹp h_c, độ sâu cuối mỗi đoạn dốc) giải bằng phép
lặp trong các cột phụ màu xám.
  python3 xuat_excel_kiem_tra_tran_chinh.py [đường dẫn .xlsx ra]
Mặc định: ../TAI_VE/KiemTra_TranChinh_SuaChua.xlsx
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

import kiem_tra_tran_chinh as K
import kiem_tra_tran_su_co as SC
import thuy_luc_tran_chinh as T

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'KiemTra_TranChinh_SuaChua.xlsx')
TC = T.TC
N_B = T.N_BUOC_LT
NR = 2 * N_B + 1                       # số mặt cắt trên dốc (hàng 6 ÷ 5 + NR)
RC = 5 + NR                            # hàng mặt cắt cuối dốc
# cấp lưu lượng: (nhãn, giá trị cố định hoặc tên ô, sheet đường mặt nước)
CAP = [('Q = 50 m³/s', 50.0, 'MN_50'), ('Q = 100 m³/s', 100.0, 'MN_100'), ('Q = 150 m³/s', 150.0, 'MN_150'),
       ('Lũ thiết kế P = 1%', 'SL_qTK', 'MN_P1'), ('Q = 200 m³/s', 200.0, 'MN_200'), ('Lũ kiểm tra P = 0,2%', 'SL_qKT', 'MN_P02')]

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
    for k, t_ in enumerate(nhan):
        ghi(ws, '%s%d' % (get_column_letter(cot0 + k), hang), t_, f_trang, fill=nen_tieu_de, can=giua)
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


def tieu_de_sheet(ws, chu_, ghi_chu=None):
    ghi(ws, 'A1', chu_, f_tieu_de, can=None, border=False)
    if ghi_chu:
        ghi(ws, 'A2', ghi_chu, f_ghi_chu, can=None, border=False)


# ================================================================== SỐ LIỆU
ts = T.doc_excel()
hd = wb.active
hd.title = 'HuongDan'
sl = wb.create_sheet('SoLieu')
rong(sl, (50, 13, 10, 72))
tieu_de_sheet(sl, 'SỐ LIỆU ĐẦU VÀO – KIỂM TRA TRÀN CHÍNH, TRÀN SỰ CỐ HIỆN TRẠNG (đồ án sửa chữa)',
              'Ô nền vàng chữ xanh: số liệu nhập. Mục A, C lấy từ file điều tiết lũ DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx; mục F, '
              'G từ hồ sơ kiểm định an toàn đập (QTVH chương 6), QTVH Bảng 13; cao độ ngưỡng, đầu và cuối dốc đối chiếu bình đồ hiện '
              'trạng (mặt bằng nâng cấp cụm công trình đầu mối).')
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


khoi('A. TRÀN CHÍNH HIỆN TRẠNG VÀ TRÀN SỰ CỐ (file điều tiết lũ, SoLieu mục B, C)', [
    ('Cao trình ngưỡng tràn chính', ts['TC_Z'], 'm', 'SL_Zn', '', '0.00'),
    ('Số khoang n', ts['TC_n'], 'khoang', 'SL_nk', '', '0'),
    ('Bề rộng một khoang b', ts['TC_b'], 'm', 'SL_b', '', '0.00'),
    ('Số trụ giữa n₁ / hệ số ξ₁', ts['TC_n1'], 'trụ', 'SL_n1', '', '0'),
    ('Hệ số hình dạng trụ giữa ξ₁', ts['TC_x1'], '', 'SL_x1', 'Số liệu SV (03/11/2026)', '0.00'),
    ('Số trụ chính n₂', ts['TC_n2'], 'trụ', 'SL_n2', '', '0'),
    ('Hệ số hình dạng trụ chính ξ₂', ts['TC_x2'], '', 'SL_x2', '', '0.00'),
    ('Hệ số hình dạng mố bên ξmb', ts['TC_xmb'], '', 'SL_xmb', 'Mố bên lượn tròn (tường cánh thượng lưu cong – bình đồ)',
     '0.00'),
    ('Hệ số lưu lượng m', ts['TC_m'], '', 'SL_m', 'Hồ sơ kiểm định: 0,445', '0.000'),
    ('Hệ số chảy ngập σn', ts['TC_s'], '', 'SL_sn', 'Chảy không ngập (i > i_k)', '0.00'),
    ('Số cửa không mở được khi sự cố', 1, 'cửa', 'SL_hong', 'Cửa khoang biên', '0'),
    ('Σb = n·b', '=SL_nk*SL_b', 'm', 'SL_Sb', 'Hồ sơ kiểm định: 22 m', '0.00'),
    ('Σξ = n₁ξ₁ + n₂ξ₂', '=SL_n1*SL_x1+SL_n2*SL_x2', '', 'SL_Sxi', '', '0.00'),
    ('Ngưỡng tràn sự cố', ts['SC_Z'], 'm', 'SC_Z', '1 khoang, cửa van phẳng; sau ngưỡng không có bậc hạ (P = 0)', '0.00'),
    ('Bề rộng tràn sự cố', ts['SC_b'], 'm', 'SC_b', '', '0.00'),
    ('Hệ số lưu lượng tràn sự cố', ts['SC_m'], '', 'SC_m', '', '0.000'),
    ('Hệ số mố bên tràn sự cố ξmb', ts['SC_xmb'], '', 'SC_xmb', 'Mố bên vuông góc (bình đồ cửa vào)', '0.00'),
])
khoi('B. NGƯỠNG TRÀN, DỐC NƯỚC (QTVH Bảng 13, hồ sơ dự án, bình đồ hiện trạng)', [
    ('Chiều cao bậc hạ sau ngưỡng P₂', TC['P2'], 'm', 'SL_P2',
     'Bình đồ: ngưỡng TP5 +74,70, đáy sau ngưỡng TP7 ≈ +73,30 (hồ sơ kiểm định ghi P = 3,3 m)', '0.00'),
    ('Cao trình đầu dốc = Z_ngưỡng − P₂', '=SL_Zn-SL_P2', 'm', 'SL_Zdd', '', '0.00'),
    ('Chiều dài dốc', TC['L_doc'], 'm', 'SL_Ld', 'QTVH Bảng 13; bình đồ: chân ngưỡng → cuối mũi phun 61,87 m', '0.0'),
    ('Độ dốc đáy i', TC['i_doc'], '', 'SL_i', 'QTVH Bảng 13 (hồ sơ kiểm định ghi 10 %)', '0.000'),
    ('Hệ số nhám n', T.N_NHAM, '', 'SL_nn', 'Bê tông', '0.000'),
    ('Bề rộng đầu dốc', TC['b_doc_dau'], 'm', 'SL_Bd', 'Hồ sơ dự án', '0.00'),
    ('Bề rộng đoạn lăng trụ', TC['b_doc_cuoi'], 'm', 'SL_Bc', 'Hồ sơ dự án; hồ sơ kiểm định Bảng 6-7: 18 m', '0.00'),
    ('Chiều dài đoạn thu hẹp', TC['L_chuyen_tiep'], 'm', 'SL_Ltt', '', '0.0'),
    ('Hệ số lưu tốc φ (độ sâu co hẹp)', T.PHI, '', 'SL_phi', 'Như hồ sơ kiểm định', '0.00'),
    ('Hệ số động năng α', T.ALPHA, '', 'SL_a', '', '0.00'),
    ('Độ vượt cao yêu cầu của tường bên', T.A_TUONG, 'm', 'SL_at', '', '0.00'),
    ('Lưu tốc không xói của dốc BT M200 [v]kx', K.V_KX, 'm/s', 'SL_vkx', 'Nhỏ nhất trong Bảng 6-7 hồ sơ kiểm định (17,6 ÷ 19,4)',
     '0.0'),
    ('Gia tốc trọng trường g', 9.81, 'm/s²', 'SL_g', '', '0.00'),
])
khoi('C. KẾT QUẢ ĐIỀU TIẾT LŨ (file điều tiết lũ, sheet KetQua)', [
    ('MNLTK – P = 1%, bình thường', ts['P1_BinhThuong']['Z'], 'm', 'SL_ZTK', '', '0.000'),
    ('q tràn chính – P = 1%, bình thường', ts['P1_BinhThuong']['qtc'], 'm³/s', 'SL_qTK', '', '0.00'),
    ('q tràn sự cố – P = 1%, bình thường', ts['P1_BinhThuong']['qsc'], 'm³/s', 'SL_qscTK', '', '0.00'),
    ('Zmax – P = 1%, 1 cửa tràn chính không mở', ts['P1_SuCo']['Z'], 'm', 'SL_ZTKs', '', '0.000'),
    ('q tràn chính – P = 1%, sự cố', ts['P1_SuCo']['qtc'], 'm³/s', 'SL_qTKs', '', '0.00'),
    ('q tràn sự cố – P = 1%, sự cố', ts['P1_SuCo']['qsc'], 'm³/s', 'SL_qscTKs', '', '0.00'),
    ('MNLKT – P = 0,2%, bình thường', ts['P02_BinhThuong']['Z'], 'm', 'SL_ZKT', '', '0.000'),
    ('q tràn chính – P = 0,2%, bình thường', ts['P02_BinhThuong']['qtc'], 'm³/s', 'SL_qKT', '', '0.00'),
    ('q tràn sự cố – P = 0,2%, bình thường', ts['P02_BinhThuong']['qsc'], 'm³/s', 'SL_qscKT', '', '0.00'),
    ('Zmax – P = 0,2%, 1 cửa tràn chính không mở', ts['P02_SuCo']['Z'], 'm', 'SL_ZKTs', '', '0.000'),
    ('q tràn chính – P = 0,2%, sự cố', ts['P02_SuCo']['qtc'], 'm³/s', 'SL_qKTs', '', '0.00'),
    ('q tràn sự cố – P = 0,2%, sự cố', ts['P02_SuCo']['qsc'], 'm³/s', 'SL_qscKTs', '', '0.00'),
    ('Cao trình đỉnh đập sau sửa chữa', ts['Z_dinh_SC'], 'm', 'SL_Zdinh', 'Mục 2.3', '0.00'),
])
khoi('D. TIÊU NĂNG MŨI PHUN – HỐ XÓI (TCVN 8420:2010; số liệu suy từ Bảng 6-8 hồ sơ kiểm định)', [
    ('Cao trình mũi phun', K.Z_MUI, 'm', 'SL_Zmui', 'Z_mũi = 74,70 + H_tr − Z_H = +69,0 ÷ +69,07 trong Bảng 6-8', '0.00'),
    ('Góc hắt mũi phun α', K.ALPHA_MP, 'độ', 'SL_amp', 'Tính ngược Bảng 6-8; như thuyết minh 4.2.4', '0'),
    ('Hệ số k (hàm khí, tách dòng)', K.K_MP, '', 'SL_kmp', 'Fr_H < 30 ÷ 35', '0.00'),
    ('Kích thước khối đá tương đương d = c', K.D_DA, 'm', 'SL_d', 'Tính ngược Bảng 6-8 (cho lại đúng h_x)', '0.00'),
    ('Hệ số mái thượng lưu hố xói', K.M_HO_TL, '', 'SL_mTL', 'Hồ sơ kiểm định', '0.0'),
    ('Hệ số mái hạ lưu hố xói', K.M_HO_HL, '', 'SL_mHL', 'Hồ sơ kiểm định', '0.0'),
])
# E. tường bên; F. quan hệ hạ lưu
r0 = r_sl[0]
ghi(sl, 'A%d' % r0, 'E. CHIỀU CAO TƯỜNG BÊN DỐC NƯỚC HIỆN TRẠNG (Bảng 6-7 hồ sơ kiểm định – lập đến x = 40 m; sau đó giả thiết)',
    f_muc, can=None, border=False)
tieu_de_cot(sl, r0 + 1, 1, ['x (m)', 'h tường (m)', '', 'Ghi chú'])
X_T = list(range(0, int(TC['L_doc']) + 1, 5))
for k, x in enumerate(X_T):
    rr = r0 + 2 + k
    ghi(sl, 'A%d' % rr, x, f_nhap, '0', nen_nhap)
    ghi(sl, 'B%d' % rr, TC['h_tuong'][k] if k < len(TC['h_tuong']) else TC['h_tuong'][-1], f_nhap, '0.00', nen_nhap)
    ghi(sl, 'D%d' % rr, 'Hồ sơ' if k < len(TC['h_tuong']) else 'Giả thiết (bằng đoạn x = 10 ÷ 40 m)', f_ghi_chu, can=trai)
ten('SL_xt', 'SoLieu!$A$%d:$A$%d' % (r0 + 2, r0 + 1 + len(X_T)))
ten('SL_htgc', 'SoLieu!$D$%d:$D$%d' % (r0 + 2, r0 + 1 + len(X_T)))
ten('SL_ht', 'SoLieu!$B$%d:$B$%d' % (r0 + 2, r0 + 1 + len(X_T)))
r0 = r0 + 3 + len(X_T)
ghi(sl, 'A%d' % r0, 'F. QUAN HỆ HẠ LƯU THEO BẢNG 6-8 HỒ SƠ KIỂM ĐỊNH (nội suy)', f_muc, can=None, border=False)
tieu_de_cot(sl, r0 + 1, 1, ['Q (m³/s)', 'h hạ lưu (m)', 'MNHL (m)', 'Ghi chú'])
for k, (q, h, mn) in enumerate(zip(K.KD_Q, K.KD_h, K.KD_MNHL)):
    rr = r0 + 2 + k
    ghi(sl, 'A%d' % rr, q, f_nhap, '0', nen_nhap)
    ghi(sl, 'B%d' % rr, h, f_nhap, '0.00', nen_nhap)
    ghi(sl, 'C%d' % rr, mn, f_nhap, '0.00', nen_nhap)
    ghi(sl, 'D%d' % rr, 'MNHL = 74,70 + H_tr − Z₀ của Bảng 6-8', f_ghi_chu, can=trai)
ten('KD_Q', 'SoLieu!$A$%d:$A$%d' % (r0 + 2, r0 + 1 + len(K.KD_Q)))
ten('KD_h', 'SoLieu!$B$%d:$B$%d' % (r0 + 2, r0 + 1 + len(K.KD_Q)))
ten('KD_MNHL', 'SoLieu!$C$%d:$C$%d' % (r0 + 2, r0 + 1 + len(K.KD_Q)))
r_sl[0] = r0 + 3 + len(K.KD_Q)
khoi('G. TRÀN SỰ CỐ – DỐC NƯỚC (QTVH Bảng 13; tính như dốc tràn chính)', [
    ('Chiều dài dốc nước tràn sự cố', SC.TSC['L_doc'], 'm', 'TSC_L', 'QTVH Bảng 13', '0.0'),
    ('Bề rộng dốc (lăng trụ) = bề rộng ngưỡng', '=SC_b', 'm', 'TSC_B', 'QTVH Bảng 13: 5 m', '0.00'),
    ('Độ dốc đáy i', SC.TSC['i_doc'], '', 'TSC_i', 'QTVH Bảng 13', '0.000'),
    ('Hệ số nhám n', T.N_NHAM, '', 'TSC_nn', 'Bê tông', '0.000'),
    ('Chiều dài đoạn chọn trước Δl', SC.DL, 'm', 'TSC_dL', 'Chia dốc thành 16 đoạn', '0.0'),
    ('Độ sâu bắt đầu = k·h_k, k', SC.K0, '', 'TSC_k0', 'Đầu dốc chảy phân giới – như thuyết minh mục 4.3', '0.00'),
])
sl.freeze_panes = 'A4'
in_an(sl, ngang=False)

# ================================================================== KHẢ NĂNG XẢ
kx = wb.create_sheet('KhaNangXa')
rong(kx, (12, 11, 11, 11, 15, 15, 12, 14, 15, 3, 40))
tieu_de_sheet(kx, 'KHẢ NĂNG XẢ LŨ CỦA NGƯỠNG TRÀN – QUAN HỆ Q ~ Z (TCVN 9147:2012, H₀ = H: bỏ qua lưu tốc tới gần)',
              'Q = σn·ε·m·Σb·√(2g)·H^1,5;  ε = 1 − 0,2·(ξmb + Σξ)·(H/b)/n, lấy H/b ≤ 1 – như khi điều tiết lũ (tràn sự cố 1 '
              'khoang: ε = 1 − 0,2·ξmb·H/b). Hồ sơ kiểm định: Ω_T ≫ 4·H·b nên không xét V₀; ε = 0,895 (ξmt = 0,25, H₀/b ≤ 1), '
              'm = 0,445, Σb = 22 m.')
tieu_de_cot(kx, 4, 1, ['Z (m)', 'H (m)', 'ε (9 khoang)', 'ε (8 khoang)', 'q tràn chính 9 khoang (m³/s)',
                       'q tràn chính 8 khoang (m³/s)', 'ε tràn sự cố', 'q tràn sự cố (m³/s)', 'q tổng – bình thường (m³/s)'], cao=45)
Z_DS = [round(74.70 + 0.25 * k, 2) for k in range(20)]
for k, z in enumerate(Z_DS + ['=SL_ZTK', '=SL_ZKT']):
    r = 5 + k
    nhap = not isinstance(z, str)
    ghi(kx, 'A%d' % r, z, f_nhap if nhap else f_dam, '0.00', nen_nhap if nhap else nen_kq)
    ghi(kx, 'B%d' % r, '=MAX(A%d-SL_Zn,0)' % r, fmt='0.000', fill=nen_kq)
    ghi(kx, 'C%d' % r, '=1-0.2*(SL_xmb+SL_Sxi)*MIN(B%d/SL_b,1)/SL_nk' % r, fmt='0.000', fill=nen_kq)
    ghi(kx, 'D%d' % r, '=1-0.2*(SL_xmb+SL_Sxi-SL_hong*SL_x1)*MIN(B%d/SL_b,1)/(SL_nk-SL_hong)' % r, fmt='0.000', fill=nen_kq)
    ghi(kx, 'E%d' % r, '=SL_sn*C{r}*SL_m*SL_Sb*SQRT(2*SL_g)*B{r}^1.5'.format(r=r), f_dam, '0.00', nen_kq)
    ghi(kx, 'F%d' % r, '=SL_sn*D{r}*SL_m*(SL_nk-SL_hong)*SL_b*SQRT(2*SL_g)*B{r}^1.5'.format(r=r), fmt='0.00', fill=nen_kq)
    ghi(kx, 'G%d' % r, '=1-0.2*SC_xmb*MIN(MAX(A{r}-SC_Z,0)/SC_b,1)'.format(r=r), fmt='0.000', fill=nen_kq)
    ghi(kx, 'H%d' % r, '=G{r}*SC_m*SC_b*SQRT(2*SL_g)*MAX(A{r}-SC_Z,0)^1.5'.format(r=r), fmt='0.00', fill=nen_kq)
    ghi(kx, 'I%d' % r, '=E%d+H%d' % (r, r), f_dam, '0.00', nen_kq)
    if not nhap:
        ghi(kx, 'K%d' % r, 'MNLTK (kết quả điều tiết)' if k == len(Z_DS) else 'MNLKT (kết quả điều tiết)', f_ghi_chu, can=trai)
RK = 5 + len(Z_DS) + 4
ghi(kx, 'A%d' % (RK - 1), 'Kết quả điều tiết lũ – đồ án (lũ QTVH 2026) và hồ sơ kiểm định (lũ kiểm định, Σb = 22 m)', f_muc,
    can=None, border=False)
tieu_de_cot(kx, RK, 1, ['Trường hợp', 'Zmax đồ án (m)', 'q tràn chính đồ án', 'q tràn sự cố đồ án', 'MNL kiểm định (m)',
                        'q tràn chính kiểm định', 'q tràn sự cố kiểm định', 'Đỉnh đập SC − Zmax đồ án (m)'], cao=45)
for k, (nhan, z, qt, qs, kd) in enumerate((
        ('P = 1%, bình thường', 'SL_ZTK', 'SL_qTK', 'SL_qscTK', (77.87, 242.00, 159.54)),
        ('P = 1%, 1 cửa không mở', 'SL_ZTKs', 'SL_qTKs', 'SL_qscTKs', None),
        ('P = 0,2%, bình thường', 'SL_ZKT', 'SL_qKT', 'SL_qscKT', (79.13, 399.79, 205.39)),
        ('P = 0,2%, 1 cửa không mở', 'SL_ZKTs', 'SL_qKTs', 'SL_qscKTs', None))):
    r = RK + 1 + k
    ghi(kx, 'A%d' % r, nhan, f_dam, can=trai)
    for col, ct, fmt in (('B', '=' + z, '0.00'), ('C', '=' + qt, '0.00'), ('D', '=' + qs, '0.00'), ('H', '=SL_Zdinh-' + z, '0.00')):
        ghi(kx, '%s%d' % (col, r), ct, f_dam if col in 'BH' else f_thuong, fmt, nen_kq)
    for col, v in zip('EFG', kd or ('–',) * 3):
        ghi(kx, '%s%d' % (col, r), v, f_nhap if kd else f_thuong, '0.00', nen_nhap if kd else None)
kx.row_dimensions[RK + 1].height = kx.row_dimensions[RK + 2].height = 30
in_an(kx)

# ================================================================== ĐẦU DỐC: cấp lưu lượng, H, MNL, h_c
dd = wb.create_sheet('DauDoc')
rong(dd, (22, 11, 10, 10, 10, 10, 10, 18, 3) + (9,) * 14)
tieu_de_sheet(dd, 'CÁC CẤP LƯU LƯỢNG – CỘT NƯỚC, MỰC NƯỚC HỒ, ĐỘ SÂU CO HẸP ĐẦU DỐC',
              'H giải ngược từ công thức khả năng xả, ε có H/b ≤ 1 (8 vòng lặp, cột J ÷ Q); h_c từ Q = φ·B_đ·h_c·√(2g(E₀ − h_c)), E₀ = P₂ + H '
              '(6 vòng lặp, cột R ÷ W). Cấp 50 ÷ 200 m³/s để vẽ đường bao hố xói.')
tieu_de_cot(dd, 4, 1, ['Cấp', 'Q (m³/s)', 'H (m)', 'MNL (m)', 'E₀ (m)', 'h_c (m)', 'h_k (B_đ) (m)', 'Trạng thái'], cao=36)
tieu_de_cot(dd, 4, 10, ['H lặp %d' % k for k in range(8)] + ['h_c lặp %d' % k for k in range(6)], cao=36)
R_DD = {}
for j, (nhan, q, mn) in enumerate(CAP):
    r = 5 + j
    R_DD[mn] = r
    ghi(dd, 'A%d' % r, nhan, f_dam, can=trai)
    ghi(dd, 'B%d' % r, q if not isinstance(q, str) else '=' + q, f_nhap if not isinstance(q, str) else f_dam, '0.00',
        nen_nhap if not isinstance(q, str) else nen_kq)
    for it in range(8):
        col, prev = get_column_letter(10 + it), get_column_letter(9 + it)
        e_ = '1' if it == 0 else '(1-0.2*(SL_xmb+SL_Sxi)*MIN(%s%d/SL_b,1)/SL_nk)' % (prev, r)
        ghi(dd, '%s%d' % (col, r), '=(B{r}/(SL_sn*{e}*SL_m*SL_Sb*SQRT(2*SL_g)))^(2/3)'.format(r=r, e=e_), f_phu, '0.0000', nen_phu)
    ghi(dd, 'C%d' % r, '=Q%d' % r, fmt='0.000', fill=nen_kq)
    ghi(dd, 'D%d' % r, '=SL_Zn+C%d' % r, f_dam, '0.000', nen_kq)
    ghi(dd, 'E%d' % r, '=SL_P2+C%d' % r, fmt='0.000', fill=nen_kq)
    for it in range(6):
        col, prev = get_column_letter(18 + it), get_column_letter(17 + it)
        hp = '0' if it == 0 else '%s%d' % (prev, r)
        ghi(dd, '%s%d' % (col, r), '=B{r}/(SL_phi*SL_Bd*SQRT(2*SL_g*(E{r}-{hp})))'.format(r=r, hp=hp), f_phu, '0.0000', nen_phu)
    ghi(dd, 'F%d' % r, '=W%d' % r, f_dam, '0.000', nen_kq)
    ghi(dd, 'G%d' % r, '=(SL_a*(B%d/SL_Bd)^2/SL_g)^(1/3)' % r, fmt='0.000', fill=nen_kq)
    ghi(dd, 'H%d' % r, '=IF(F%d<G%d,"h_c < h_k: dòng xiết","Kiểm tra lại")' % (r, r), fill=nen_kq)
in_an(dd)

# ================================================================== ĐƯỜNG MẶT NƯỚC (mỗi cấp lưu lượng một sheet)
HDR15 = ['MC', 'B (m)', 'h (m)', 'v (m/s)', 'E (m)', 'ΔE (m)', 'ω (m²)', 'χ (m)', 'R (m)', 'C', 'J', 'J_tb', 'i − J_tb',
         'Δl = ΔE/(i − J_tb) (m)', 'L (m)']
for nhan, _, mn in CAP:
    ws = wb.create_sheet(mn)
    rong(ws, (5, 8, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 9, 11, 8, 3, 9, 10) + (10,) * 6 + (14,))
    r1 = R_DD[mn]
    tieu_de_sheet(ws, 'ĐƯỜNG MẶT NƯỚC TRÊN DỐC NƯỚC – %s (sai phân, chọn trước Δl, thử dần h – như hồ sơ kiểm định)' % nhan)
    ghi(ws, 'A3', 'Q (m³/s) =', f_dam, can=None, border=False)
    ghi(ws, 'C3', '=DauDoc!B%d' % r1, f_dam, '0.00', nen_kq)
    ghi(ws, 'E3', 'h đầu dốc h_c (m) =', f_dam, can=None, border=False)
    ghi(ws, 'G3', '=DauDoc!F%d' % r1, f_dam, '0.000', nen_kq)
    tieu_de_cot(ws, 5, 1, HDR15, cao=36)
    tieu_de_cot(ws, 5, 17, ['Δl chọn (m)', 'K = E₁ + i·Δl − J₁·Δl/2'] + ['h lặp %d' % k for k in range(6)] +
                ['Kiểm tra E₁ + i·Δl − E₂ − J_tb·Δl'], cao=36)
    for k in range(NR):
        r, p = 6 + k, 5 + k
        ct = {'A': k + 1,
              'B': '=IF(O{r}<=SL_Ltt,SL_Bd+(SL_Bc-SL_Bd)*O{r}/SL_Ltt,SL_Bc)'.format(r=r),
              'D': '=$C$3/(B{r}*C{r})'.format(r=r), 'E': '=C{r}+SL_a*D{r}^2/(2*SL_g)'.format(r=r),
              'G': '=B{r}*C{r}'.format(r=r), 'H': '=B{r}+2*C{r}'.format(r=r), 'I': '=G{r}/H{r}'.format(r=r),
              'J': '=I{r}^(1/6)/SL_nn'.format(r=r), 'K': '=D{r}^2/(J{r}^2*I{r})'.format(r=r)}
        if k == 0:
            ct.update({'C': '=$G$3', 'O': 0})
        else:
            ct.update({'Q': '=IF(O{p}<SL_Ltt-0.000001,SL_Ltt/{n},(SL_Ld-SL_Ltt)/{n})'.format(p=p, n=N_B),
                       'R': '=E{p}+SL_i*Q{r}-0.5*K{p}*Q{r}'.format(p=p, r=r), 'O': '=O{p}+Q{r}'.format(p=p, r=r),
                       'C': '=X%d' % r, 'F': '=E{r}-E{p}'.format(r=r, p=p), 'L': '=(K{p}+K{r})/2'.format(p=p, r=r),
                       'M': '=SL_i-L%d' % r, 'N': '=F{r}/M{r}'.format(r=r),
                       'Y': '=E{p}+SL_i*Q{r}-E{r}-L{r}*Q{r}'.format(p=p, r=r)})
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
    ghi(ws, 'A%d' % (RC + 2), 'Cột (14) tính theo Δl = ΔE/(i − J_tb), bằng Δl chọn (cột Q) khi h thử đúng; cột Y kiểm tra phương '
                              'trình năng lượng (≈ 0).', f_ghi_chu, can=None, border=False)
    ws.freeze_panes = 'B6'
    in_an(ws)

# ================================================================== DỐC NƯỚC: TƯỜNG BÊN, XÓI
dn = wb.create_sheet('DocNuoc')
rong(dn, (8, 9, 9, 10, 9, 9, 10, 9, 11, 12, 12, 12, 11, 13, 18, 3) + (10,) * 6)
tieu_de_sheet(dn, 'KẾT QUẢ KIỂM TRA THỦY LỰC DỐC NƯỚC – TƯỜNG BÊN, XÓI (như Bảng 6-7 hồ sơ kiểm định)',
              'h_hk = h·(1 + v/100); độ vượt cao = chiều cao tường − h_hk; không xói khi v < [v]kx; tường cần tối thiểu = h_hk lớn '
              'nhất + a (dùng để đánh giá đoạn tường giả thiết).')
tieu_de_cot(dn, 4, 1, ['x (m)', 'B (m)', 'h TK (m)', 'h_hk TK (m)', 'v TK (m/s)', 'h KT (m)', 'h_hk KT (m)', 'v KT (m/s)',
                       'Tường bên (m)', 'Độ vượt cao TK (m)', 'Độ vượt cao KT (m)', 'Không tràn tường', 'Không xói',
                       'Tường cần tối thiểu h_hk + a (m)', 'Chiều cao tường bên'], cao=45)
for k, x in enumerate(X_T):
    r = 5 + k
    tim = lambda w: 'MATCH(A{r}+0.0001,{w}!$O$6:$O${e},1)'.format(r=r, w=sh(w), e=RC)
    ghi(dn, 'A%d' % r, '=INDEX(SL_xt,%d)' % (k + 1), f_dam, '0')
    ghi(dn, 'B%d' % r, '=INDEX(%s!$B$6:$B$%d,%s)' % (sh('MN_P02'), RC, tim('MN_P02')), fmt='0.00', fill=nen_kq)
    for col, w, src, fmt in (('C', 'MN_P1', 'C', '0.000'), ('E', 'MN_P1', 'D', '0.00'), ('F', 'MN_P02', 'C', '0.000'),
                             ('H', 'MN_P02', 'D', '0.00')):
        ghi(dn, '%s%d' % (col, r), '=INDEX(%s!$%s$6:$%s$%d,%s)' % (sh(w), src, src, RC, tim(w)), fmt=fmt, fill=nen_kq)
    ghi(dn, 'D%d' % r, '=C{r}*(1+E{r}/100)'.format(r=r), fmt='0.000', fill=nen_kq)
    ghi(dn, 'G%d' % r, '=F{r}*(1+H{r}/100)'.format(r=r), fmt='0.000', fill=nen_kq)
    ghi(dn, 'I%d' % r, '=INDEX(SL_ht,%d)' % (k + 1), fmt='0.00', fill=nen_kq)
    ghi(dn, 'J%d' % r, '=I%d-D%d' % (r, r), fmt='0.00', fill=nen_kq)
    ghi(dn, 'K%d' % r, '=I%d-G%d' % (r, r), f_dam, '0.00', nen_kq)
    ghi(dn, 'L%d' % r, '=IF(MIN(J{r}:K{r})>=SL_at,"Thỏa mãn","Kiểm tra")'.format(r=r), fill=nen_kq)
    ghi(dn, 'M%d' % r, '=IF(MAX(E{r},H{r})<SL_vkx,"Thỏa mãn","Xói")'.format(r=r), fill=nen_kq)
    ghi(dn, 'N%d' % r, '=MAX(D{r},G{r})+SL_at'.format(r=r), fmt='0.00', fill=nen_kq)
    ghi(dn, 'O%d' % r, '=INDEX(SL_htgc,%d)' % (k + 1), f_ghi_chu, can=trai)
RT1 = 4 + len(X_T)
ghi(dn, 'A%d' % (RT1 + 2), '="Độ vượt cao nhỏ nhất "&FIXED(MIN(K5:K%d),2)&" m; lưu tốc lớn nhất "&FIXED(MAX(H5:H%d),2)&'
                           '" m/s so với [v]kx = "&FIXED(SL_vkx,1)&" m/s"' % (RT1, RT1), f_dam, can=None, border=False)
RGT = 5 + len(TC['h_tuong'])                  # hàng đầu tiên của đoạn tường giả thiết (x > 40 m)
ghi(dn, 'A%d' % (RT1 + 3), '="Đoạn x = "&A{a}&" ÷ "&A{b}&" m (hồ sơ không ghi chiều cao tường, giả thiết "&FIXED(MIN(I{a}:I{b}),2)&'
                           '" m): tường cần tối thiểu "&FIXED(MAX(N{a}:N{b}),2)&" m – giả thiết không quyết định kết luận; đo kiểm '
                           'tra thực tế để bỏ giả thiết."'.format(a=RGT, b=RT1), f_dam, can=None, border=False)
# số liệu vẽ hình
tieu_de_cot(dn, 4, 17, ['x (m)', 'Đáy dốc', 'MN lũ thiết kế (hàm khí)', 'MN lũ kiểm tra (hàm khí)', 'Đỉnh tường'], cao=45)
for k in range(len(X_T)):
    r = 5 + k
    ghi(dn, 'Q%d' % r, '=A%d' % r, fmt='0')
    ghi(dn, 'R%d' % r, '=SL_Zdd-SL_i*Q%d' % r, fmt='0.00')
    ghi(dn, 'S%d' % r, '=R%d+D%d' % (r, r), fmt='0.00')
    ghi(dn, 'T%d' % r, '=R%d+G%d' % (r, r), fmt='0.00')
    ghi(dn, 'U%d' % r, '=R%d+I%d' % (r, r), fmt='0.00')
ch = ScatterChart()
ch.title = 'Đường mặt nước trên dốc tràn chính và đỉnh tường bên'
ch.style = 13
ch.x_axis.title, ch.y_axis.title = 'x (m)', 'Cao trình (m)'
ch.x_axis.delete = ch.y_axis.delete = False
ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0, TC['L_doc'], 10
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 68, 79
ch.width, ch.height = 20, 10
ch.legend.position = 'b'
xr = Reference(dn, min_col=17, min_row=5, max_row=RT1)
for col, nhan, mau, net, diem in ((18, 'Đáy dốc', '52514E', None, None), (19, 'Lũ thiết kế (có hàm khí)', '2A78D6', None, None),
                                  (20, 'Lũ kiểm tra (có hàm khí)', 'EB6834', None, None),
                                  (21, 'Đỉnh tường bên', '0B0B0B', 'sysDot', 'circle')):
    s = Series(Reference(dn, min_col=col, min_row=5, max_row=RT1), xr, title=nhan)
    s.smooth = False
    s.marker.symbol = diem or 'none'
    s.graphicalProperties.line.solidFill = mau
    s.graphicalProperties.line.width = 22225
    if net:
        s.graphicalProperties.line.dashStyle = net
    if diem:
        s.marker.size = 5
        s.marker.graphicalProperties.solidFill = s.marker.graphicalProperties.line.solidFill = mau
    ch.series.append(s)
dn.add_chart(ch, 'A%d' % (RT1 + 5))
in_an(dn)

# ================================================================== MŨI PHUN – HỐ XÓI
mp = wb.create_sheet('MuiPhun')
rong(mp, (20,) + (9,) * 20)
tieu_de_sheet(mp, 'TIÊU NĂNG MŨI PHUN – HỐ XÓI (TCVN 8420:2010, công thức (7) ÷ (10) hồ sơ kiểm định)',
              'L = k·φ²·Z_H·sin2α·[1 + √(1 + (Z₀ − Z_H)/(φ²·Z_H·sin²α))], φ = v_H/√(2g·Z_H); L₁ = L + t/tgα₁; '
              'tgα₁ = √(tg²α + 2g(0,5h_H·cosα + Z_mũi − MNHL)/(v_H²·cos²α)); t = (0,1 + 0,45·Fr^0,25)·√Fr·(h/d)^0,25·h_H; '
              'h_x = t − h + c; b_x = 2,5h + 3,7h_x; x_TL = L₁ − 1,25h − 2,2h_x')
HDR_MP = ['Cấp', 'Q (m³/s)', 'MNL (m)', 'h_H (m)', 'v_H (m/s)', 'Fr_H', 'h hạ lưu (m)', 'MNHL (m)', 'Z_H (m)', 'Z₀ (m)', 'φ',
          'L (m)', 'tgα₁', 'α₁ (°)', 't (m)', 'h_x (m)', 'L₁ (m)', 'b_x (m)', 'x_TL (m)', 'x_HL (m)', 'Kết luận']
tieu_de_cot(mp, 4, 1, HDR_MP, cao=36)
mp.column_dimensions['U'].width = 24
for j, (nhan, _, mn) in enumerate(CAP):
    r = 5 + j
    rd = R_DD[mn]
    k_ = 'MIN(MATCH(B{r},KD_Q,1),COUNT(KD_Q)-1)'.format(r=r)
    noi = lambda bang_: ('INDEX({b},{k})+(B{r}-INDEX(KD_Q,{k}))*(INDEX({b},{k}+1)-INDEX({b},{k}))/(INDEX(KD_Q,{k}+1)-'
                         'INDEX(KD_Q,{k}))'.format(b=bang_, k=k_, r=r))
    ct = [('A', nhan, None), ('B', '=DauDoc!B%d' % rd, '0.00'), ('C', '=DauDoc!D%d' % rd, '0.000'),
          ('D', '=%s!C%d' % (sh(mn), RC), '0.000'), ('E', '=%s!D%d' % (sh(mn), RC), '0.00'),
          ('F', '=E{r}^2/(SL_g*D{r})'.format(r=r), '0.00'), ('G', '=' + noi('KD_h'), '0.00'), ('H', '=' + noi('KD_MNHL'), '0.00'),
          ('I', '=C%d-SL_Zmui' % r, '0.00'), ('J', '=C%d-H%d' % (r, r), '0.00'), ('K', '=E{r}/SQRT(2*SL_g*I{r})'.format(r=r), '0.000'),
          ('L', '=SL_kmp*K{r}^2*I{r}*SIN(2*RADIANS(SL_amp))*(1+SQRT(1+(J{r}/I{r}-1)/(K{r}^2*SIN(RADIANS(SL_amp))^2)))'
           .format(r=r), '0.00'),
          ('M', '=SQRT(TAN(RADIANS(SL_amp))^2+2*SL_g*(0.5*D{r}*COS(RADIANS(SL_amp))+SL_Zmui-H{r})/(E{r}^2*COS(RADIANS(SL_amp))^2))'
           .format(r=r), '0.000'),
          ('N', '=DEGREES(ATAN(M%d))' % r, '0.0'),
          ('O', '=(0.1+0.45*F{r}^0.25)*SQRT(F{r})*(G{r}/SL_d)^0.25*D{r}'.format(r=r), '0.00'),
          ('P', '=O{r}-G{r}+SL_d'.format(r=r), '0.00'), ('Q', '=L{r}+O{r}/M{r}'.format(r=r), '0.00'),
          ('R', '=2.5*G{r}+3.7*P{r}'.format(r=r), '0.00'), ('S', '=Q{r}-1.25*G{r}-SL_mTL*P{r}'.format(r=r), '0.00'),
          ('T', '=Q{r}+1.25*G{r}+SL_mHL*P{r}'.format(r=r), '0.00'),
          ('U', '=IF(S{r}>0,"Ngoài chân mũi phun","Lấn tới mũi phun")'.format(r=r), None)]
    for col, f_, fmt in ct:
        ghi(mp, '%s%d' % (col, r), f_, f_dam if col in 'APQSU' else f_thuong, fmt, None if col == 'A' else nen_kq,
            trai if col == 'A' else cg)
RM1 = 4 + len(CAP)
ghi(mp, 'A%d' % (RM1 + 2), '=IF(COUNTIF(U5:U%d,"Ngoài chân mũi phun")=%d,"Chân mũi phun nằm ngoài đường bao hố xói (x_TL nhỏ nhất "&'
                           'FIXED(MIN(S5:S%d),2)&" m) – tiêu năng mũi phun an toàn","Hố xói lấn tới chân mũi phun – cần gia cố")'
    % (RM1, len(CAP), RM1), f_dam, can=None, border=False)
# đường bao hố xói: từng hố (4 điểm) và đường bao lấy mẫu 0,5 m
RH = RM1 + 5
ghi(mp, 'A%d' % (RH - 1), 'Số liệu vẽ hố xói (độ sâu so với đáy lòng dẫn hạ lưu, m)', f_muc, can=None, border=False)
for j in range(len(CAP)):
    c0 = 2 + 2 * j
    tieu_de_cot(mp, RH, c0, ['x (%d)' % (j + 1), 'y (%d)' % (j + 1)])
    r = 5 + j
    for k, (fx, fy) in enumerate((('=S{r}', '0'), ('=S{r}+SL_mTL*P{r}', '=-P{r}'), ('=T{r}-SL_mHL*P{r}', '=-P{r}'), ('=T{r}', '0'))):
        ghi(mp, '%s%d' % (get_column_letter(c0), RH + 1 + k), fx.format(r=r), fmt='0.00')
        ghi(mp, '%s%d' % (get_column_letter(c0 + 1), RH + 1 + k), fy.format(r=r) if fy.startswith('=') else 0, fmt='0.00')
RB = RH + 7
tieu_de_cot(mp, RB, 1, ['x (m)'] + ['sâu hố %d' % (j + 1) for j in range(len(CAP))] + ['Đường bao (m)'])
N_MAU = 101
for k in range(N_MAU):
    r = RB + 1 + k
    ghi(mp, 'A%d' % r, '=MAX(T5:T%d)*%d/%d' % (RM1, k, N_MAU - 1), fmt='0.00')
    for j in range(len(CAP)):
        rr = 5 + j
        ghi(mp, '%s%d' % (get_column_letter(2 + j), r),
            '=MAX(0,MIN(P{rr},($A{r}-S{rr})/SL_mTL,(T{rr}-$A{r})/SL_mHL))'.format(rr=rr, r=r), f_phu, '0.00', nen_phu)
    ghi(mp, '%s%d' % (get_column_letter(2 + len(CAP)), r), '=-MAX(B{r}:{c}{r})'.format(r=r, c=get_column_letter(1 + len(CAP))),
        fmt='0.00')
ch = ScatterChart()
ch.title = 'Đường bao hố xói sau mũi phun'
ch.style = 13
ch.x_axis.title, ch.y_axis.title = 'Khoảng cách từ mũi phun (m)', 'So với đáy lòng dẫn (m)'
ch.x_axis.crosses = 'min'
ch.x_axis.delete = ch.y_axis.delete = False
ch.width, ch.height = 20, 9
ch.legend.position = 'b'
MAU = ['B7D3F6', '8BB7EE', '5B97E2', '2A78D6', '1F5FB0', '13427E']
for j, (nhan, _, _) in enumerate(CAP):
    c0 = 2 + 2 * j
    s = Series(Reference(mp, min_col=c0 + 1, min_row=RH + 1, max_row=RH + 4), Reference(mp, min_col=c0, min_row=RH + 1, max_row=RH + 4),
               title=nhan)
    s.smooth = False
    s.marker.symbol = 'none'
    s.graphicalProperties.line.solidFill = MAU[j]
    s.graphicalProperties.line.width = 15875
    ch.series.append(s)
s = Series(Reference(mp, min_col=2 + len(CAP), min_row=RB + 1, max_row=RB + N_MAU),
           Reference(mp, min_col=1, min_row=RB + 1, max_row=RB + N_MAU), title='Đường bao hố xói')
s.smooth = False
s.marker.symbol = 'none'
s.graphicalProperties.line.solidFill = '0B0B0B'
s.graphicalProperties.line.width = 28575
ch.series.append(s)
mp.add_chart(ch, 'N%d' % (RH - 1))
in_an(mp)

# ================================================================== TRÀN SỰ CỐ (SV yêu cầu 08/11/2026)
CA_SC = [('Lũ thiết kế P = 1%', 'SL_ZTK', 'SL_qscTK', 'SL_qTK', 'TSC_MN_P1'),
         ('P = 1%, 1 cửa tràn chính không mở', 'SL_ZTKs', 'SL_qscTKs', 'SL_qTKs', 'TSC_MN_P1SC'),
         ('Lũ kiểm tra P = 0,2%', 'SL_ZKT', 'SL_qscKT', 'SL_qKT', 'TSC_MN_P02'),
         ('P = 0,2%, 1 cửa tràn chính không mở', 'SL_ZKTs', 'SL_qscKTs', 'SL_qKTs', 'TSC_MN_P02SC')]
NS = int(round(SC.TSC['L_doc'] / SC.DL)) + 1            # số mặt cắt dốc tràn sự cố (hàng 6 ÷ 5 + NS)
RS = 5 + NS
N_IT = 30                                               # vòng lặp h (đoạn đầu sát phân giới hội tụ chậm)
N_H0 = 8
sd = wb.create_sheet('TSC_DauDoc')
rong(sd, (30, 9, 9, 7, 8, 8, 11, 12, 11, 10, 9, 9, 9, 9, 8, 9, 15, 9, 11, 3) + (9,) * N_H0)
tieu_de_sheet(sd, 'TRÀN SỰ CỐ – KHẢ NĂNG XẢ, ĐỘ SÂU PHÂN GIỚI, ĐỘ DỐC PHÂN GIỚI, ĐỘ SÂU DÒNG ĐỀU, ĐỘ SÂU ĐẦU DỐC',
              'ε = 1 − 0,2·ξmb·(H/b), lấy H/b ≤ 1 (1 khoang); sau ngưỡng không có bậc hạ nên đầu dốc là mặt cắt khống chế, chảy phân giới: '
              'h_k = (α·q²/g)^(1/3); i_k = Q²/(ω_k²·C_k²·R_k); h₀ từ Q = ω·C·√(R·i) (lặp, cột U ÷ AB). Độ sâu bắt đầu tính đường '
              'mặt nước = k·h_k.')
tieu_de_cot(sd, 4, 1, ['Trường hợp', 'Z_max (m)', 'H = Z − Z_ng (m)', 'H/b', 'ε', 'ε·m', 'Q tràn sự cố (m³/s)',
                       'Q theo công thức ở Z_max (m³/s)', 'Tỷ lệ trong lưu lượng xả (%)', 'q = Q/b (m³/s/m)', 'h_k (m)', 'ω_k (m²)',
                       'χ_k (m)', 'R_k (m)', 'C_k', 'i_k', 'So sánh i', 'h₀ (m)', 'h đầu dốc = k·h_k (m)'], cao=48)
tieu_de_cot(sd, 4, 21, ['h₀ lặp %d' % k for k in range(N_H0)], cao=48)
R_SC = {}
for j, (nhan, zn, qn, qtn, mn) in enumerate(CA_SC):
    r = 5 + j
    R_SC[mn] = r
    ct = [('A', nhan, None), ('B', '=' + zn, '0.000'), ('C', '=B%d-SC_Z' % r, '0.000'), ('D', '=C%d/SC_b' % r, '0.000'),
          ('E', '=1-0.2*SC_xmb*MIN(D%d,1)' % r, '0.000'), ('F', '=E%d*SC_m' % r, '0.000'), ('G', '=' + qn, '0.00'),
          ('H', '=F{r}*SC_b*SQRT(2*SL_g)*C{r}^1.5'.format(r=r), '0.00'),
          ('I', '=G{r}/(G{r}+{q})*100'.format(r=r, q=qtn), '0.0'), ('J', '=G%d/TSC_B' % r, '0.00'),
          ('K', '=(SL_a*J%d^2/SL_g)^(1/3)' % r, '0.000'), ('L', '=TSC_B*K%d' % r, '0.00'), ('M', '=TSC_B+2*K%d' % r, '0.00'),
          ('N', '=L{r}/M{r}'.format(r=r), '0.000'), ('O', '=N%d^(1/6)/TSC_nn' % r, '0.00'),
          ('P', '=G{r}^2/(L{r}^2*O{r}^2*N{r})'.format(r=r), '0.00000'),
          ('Q', '=IF(TSC_i>P%d,"i > i_k: dốc nước","Kiểm tra")' % r, None),
          ('R', '=%s%d' % (get_column_letter(20 + N_H0), r), '0.000'), ('S', '=TSC_k0*K%d' % r, '0.000')]
    for col, f_, fmt in ct:
        ghi(sd, '%s%d' % (col, r), f_, f_dam if col in 'AFGKS' else f_thuong, fmt, None if col == 'A' else nen_kq,
            trai if col == 'A' else cg)
    for it in range(N_H0):
        col = get_column_letter(21 + it)
        hp = 'K%d' % r if it == 0 else '%s%d' % (get_column_letter(20 + it), r)
        ghi(sd, '%s%d' % (col, r), '=(G{r}*TSC_nn/(TSC_B*SQRT(TSC_i)))^(3/5)*((TSC_B+2*{h})/TSC_B)^(2/5)'.format(r=r, h=hp), f_phu,
            '0.0000', nen_phu)
    sd.row_dimensions[r].height = 30
RSC1 = 4 + len(CA_SC)
ghi(sd, 'A%d' % (RSC1 + 2), 'Giới hạn ε·m của ngưỡng không có bậc hạ (chảy phân giới ở đầu dốc) m_max = φ·(2/3)·√(1/3) =', f_dam,
    can=None, border=False)
ghi(sd, 'H%d' % (RSC1 + 2), '=SL_phi*(2/3)*SQRT(1/3)', f_dam, '0.000', nen_kq)
ten('TSC_mmax', 'TSC_DauDoc!$H$%d' % (RSC1 + 2))
ghi(sd, 'A%d' % (RSC1 + 3), '=IF(MAX(F5:F{e})<=TSC_mmax*1.01,"ε·m của tràn sự cố ("&IF(ROUND(MIN(F5:F{e}),3)=ROUND(MAX(F5:F{e}),3),'
                           'FIXED(MIN(F5:F{e}),3),FIXED(MIN(F5:F{e}),3)&" ÷ "&FIXED(MAX(F5:F{e}),3))&") không vượt giới hạn "&'
                           'FIXED(TSC_mmax,3)&" – phù hợp chảy phân giới ở đầu dốc","ε·m vượt giới hạn – xem lại hệ số")'.format(e=RSC1),
    f_dam, can=None, border=False)
in_an(sd)

HDR_SC = ['MC', 'x (m)', 'h (m)', 'v (m/s)', 'E (m)', 'ΔE (m)', 'ω (m²)', 'χ (m)', 'R (m)', 'C', 'J', 'J_tb', 'i − J_tb',
          'Δl = ΔE/(i − J_tb) (m)', 'h_hk (m)', 'Tường cần h_hk + a (m)', 'Fr = v²/(g·h)']
C_CUOI = get_column_letter(19 + N_IT)                    # cột lặp cuối cùng
C_KT = get_column_letter(20 + N_IT)
for nhan, _, _, _, mn in CA_SC:
    ws = wb.create_sheet(mn)
    rong(ws, (5, 7, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 9, 11, 8, 10, 8, 3, 11) + (8,) * N_IT + (14,))
    r1 = R_SC[mn]
    tieu_de_sheet(ws, 'TRÀN SỰ CỐ – ĐƯỜNG MẶT NƯỚC TRÊN DỐC – %s (sai phân, chọn trước Δl, thử dần h – như dốc tràn chính)' % nhan)
    ghi(ws, 'A3', 'Q (m³/s) =', f_dam, can=None, border=False)
    ghi(ws, 'C3', '=TSC_DauDoc!G%d' % r1, f_dam, '0.00', nen_kq)
    ghi(ws, 'E3', 'h đầu dốc (m) =', f_dam, can=None, border=False)
    ghi(ws, 'G3', '=TSC_DauDoc!S%d' % r1, f_dam, '0.000', nen_kq)
    ghi(ws, 'I3', 'h_k (m) =', f_dam, can=None, border=False)
    ghi(ws, 'K3', '=TSC_DauDoc!K%d' % r1, f_dam, '0.000', nen_kq)
    tieu_de_cot(ws, 5, 1, HDR_SC, cao=36)
    tieu_de_cot(ws, 5, 19, ['K = E₁ + i·Δl − J₁·Δl/2'] + ['h lặp %d' % k for k in range(N_IT)] +
                ['Kiểm tra E₁ + i·Δl − E₂ − J_tb·Δl'], cao=36)
    for k in range(NS):
        r, p = 6 + k, 5 + k
        ct = {'A': k + 1, 'D': '=$C$3/(TSC_B*C{r})'.format(r=r), 'E': '=C{r}+SL_a*D{r}^2/(2*SL_g)'.format(r=r),
              'G': '=TSC_B*C{r}'.format(r=r), 'H': '=TSC_B+2*C{r}'.format(r=r), 'I': '=G{r}/H{r}'.format(r=r),
              'J': '=I{r}^(1/6)/TSC_nn'.format(r=r), 'K': '=D{r}^2/(J{r}^2*I{r})'.format(r=r),
              'O': '=C{r}*(1+D{r}/100)'.format(r=r), 'P': '=O{r}+SL_at'.format(r=r), 'Q': '=D{r}^2/(SL_g*C{r})'.format(r=r)}
        if k == 0:
            ct.update({'B': 0, 'C': '=$G$3'})
        else:
            ct.update({'B': '=B{p}+TSC_dL'.format(p=p), 'C': '=%s%d' % (C_CUOI, r), 'F': '=E{r}-E{p}'.format(r=r, p=p),
                       'L': '=(K{p}+K{r})/2'.format(p=p, r=r), 'M': '=TSC_i-L%d' % r, 'N': '=F{r}/M{r}'.format(r=r),
                       'S': '=E{p}+TSC_i*TSC_dL-0.5*K{p}*TSC_dL'.format(p=p),
                       C_KT: '=E{p}+TSC_i*TSC_dL-E{r}-L{r}*TSC_dL'.format(p=p, r=r)})
            for it in range(N_IT):
                col = get_column_letter(20 + it)
                hp = ('C%d' % p) if it == 0 else '%s%d' % (get_column_letter(19 + it), r)
                ct[col] = ('=$C$3/(TSC_B*SQRT(2*SL_g/SL_a*(S{r}-{h}-0.5*TSC_dL*(($C$3/(TSC_B*{h}))^2*TSC_nn^2/'
                           '((TSC_B*{h}/(TSC_B+2*{h}))^(4/3))))))'.format(r=r, h=hp))
        fm = dict(A='0', B='0.0', C='0.000', D='0.00', E='0.000', F='0.000', G='0.00', H='0.00', I='0.000', J='0.00', K='0.00000',
                  L='0.00000', M='0.00000', N='0.00', O='0.000', P='0.00', Q='0.00', S='0.0000')
        fm[C_KT] = '0.0E+00'
        for col, f_ in ct.items():
            phu = len(col) > 1 or col >= 'S'
            ghi(ws, '%s%d' % (col, r), f_, f_phu if phu else (f_dam if col in 'BCP' else f_thuong), fm.get(col, '0.0000'),
                nen_phu if phu else None)
        for col in 'ABCDEFGHIJKLMNOPQ':
            ws['%s%d' % (col, r)].border = vien
    ghi(ws, 'A%d' % (RS + 2), 'Cột (14) tính theo Δl = ΔE/(i − J_tb), bằng Δl chọn (TSC_dL) khi h thử đúng; cột %s kiểm tra phương '
                              'trình năng lượng (≈ 0). Đoạn đầu sát trạng thái phân giới nên phép lặp cần nhiều vòng (%d vòng).'
        % (C_KT, N_IT), f_ghi_chu, can=None, border=False)
    ws.freeze_panes = 'C6'
    in_an(ws)

st = wb.create_sheet('TSC_DocNuoc')
rong(st, (7, 10, 9, 9, 9, 11, 9, 9, 9, 11, 10, 12, 11, 3, 8, 10, 11, 11, 11))
tieu_de_sheet(st, 'TRÀN SỰ CỐ – KẾT QUẢ THỦY LỰC DỐC NƯỚC: CHIỀU SÂU CÓ KỂ HÀM KHÍ, CHIỀU CAO TƯỜNG CẦN, LƯU TỐC',
              'h_hk = h·(1 + v/100); tường cần = h_hk lớn nhất + a; không xói khi v < [v]kx. Hồ sơ hiện có chưa ghi chiều cao tường bên dốc '
              'nước tràn sự cố – cần đo đạc đối chiếu với chiều cao tường cần.')
tieu_de_cot(st, 4, 1, ['x (m)', 'Cao trình đáy (m)', 'h TK (m)', 'h_hk TK (m)', 'v TK (m/s)', 'h_hk TK kẹt 1 cửa TC (m)', 'h KT (m)',
                       'h_hk KT (m)', 'v KT (m/s)', 'h_hk KT kẹt 1 cửa TC (m)', 'v lớn nhất (m/s)', 'Tường cần h_hk + a (m)',
                       'Không xói'], cao=48)
m1, m1s, m2, m2s = (sh(c[4]) for c in CA_SC)
for k in range(NS):
    r, rr = 5 + k, 6 + k
    ct = [('A', '=%s!B%d' % (m2, rr), '0', f_dam), ('B', '=SC_Z-TSC_i*A%d' % r, '0.00', f_thuong),
          ('C', '=%s!C%d' % (m1, rr), '0.000', f_thuong), ('D', '=%s!O%d' % (m1, rr), '0.000', f_thuong),
          ('E', '=%s!D%d' % (m1, rr), '0.00', f_thuong), ('F', '=%s!O%d' % (m1s, rr), '0.000', f_thuong),
          ('G', '=%s!C%d' % (m2, rr), '0.000', f_thuong), ('H', '=%s!O%d' % (m2, rr), '0.000', f_thuong),
          ('I', '=%s!D%d' % (m2, rr), '0.00', f_thuong), ('J', '=%s!O%d' % (m2s, rr), '0.000', f_thuong),
          ('K', '=MAX(E{r},I{r},{a}!D{rr},{b}!D{rr})'.format(r=r, rr=rr, a=m1s, b=m2s), '0.00', f_thuong),
          ('L', '=MAX(D{r},F{r},H{r},J{r})+SL_at'.format(r=r), '0.00', f_dam),
          ('M', '=IF(K%d<SL_vkx,"Thỏa mãn","Xói")' % r, None, f_thuong)]
    for col, f_, fmt, fo in ct:
        ghi(st, '%s%d' % (col, r), f_, fo, fmt, nen_kq)
RT2 = 4 + NS
ghi(st, 'A%d' % (RT2 + 2), '="Tường bên dốc tràn sự cố cần cao tối thiểu "&FIXED(L5,2)&" m ở đầu dốc, "&FIXED(L%d,2)&" m ở cuối dốc; '
                           'lưu tốc lớn nhất "&FIXED(MAX(K5:K%d),2)&" m/s < [v]kx = "&FIXED(SL_vkx,1)&" m/s – không xói theo lý thuyết."'
    % (RT2, RT2), f_dam, can=None, border=False)
# thông số cuối dốc – căn cứ kiểm tra tiêu năng
RE = RT2 + 4
ghi(st, 'A%d' % (RE - 1), 'Thông số dòng chảy cuối dốc tràn sự cố (căn cứ kiểm tra tiêu năng khi có hồ sơ công trình tiêu năng, mực nước hạ lưu)',
    f_muc, can=None, border=False)
tieu_de_cot(st, RE, 1, ['Trường hợp', '', '', 'Q (m³/s)', 'q (m³/s/m)', 'h cuối dốc (m)', 'v (m/s)', 'Fr = v²/(g·h)', 'E (m)',
                        'h_hk (m)', 'Tường cần (m)'], cao=36)
st.merge_cells('A%d:C%d' % (RE, RE))
for j, (nhan, _, _, _, mn) in enumerate(CA_SC):
    r = RE + 1 + j
    st.merge_cells('A%d:C%d' % (r, r))
    ghi(st, 'A%d' % r, nhan, f_dam, can=trai)
    for col, f_, fmt in (('D', '=%s!C3' % sh(mn), '0.00'), ('E', '=D%d/TSC_B' % r, '0.00'), ('F', '=%s!C%d' % (sh(mn), RS), '0.000'),
                         ('G', '=%s!D%d' % (sh(mn), RS), '0.00'), ('H', '=%s!Q%d' % (sh(mn), RS), '0.00'),
                         ('I', '=%s!E%d' % (sh(mn), RS), '0.00'), ('J', '=%s!O%d' % (sh(mn), RS), '0.000'),
                         ('K', '=%s!P%d' % (sh(mn), RS), '0.00')):
        ghi(st, '%s%d' % (col, r), f_, fmt=fmt, fill=nen_kq)
# số liệu vẽ hình
tieu_de_cot(st, 4, 15, ['x (m)', 'Đáy dốc', 'MN lũ thiết kế (hàm khí)', 'MN lũ kiểm tra (hàm khí)', 'Đỉnh tường cần'], cao=48)
for k in range(NS):
    r = 5 + k
    ghi(st, 'O%d' % r, '=A%d' % r, fmt='0')
    ghi(st, 'P%d' % r, '=B%d' % r, fmt='0.00')
    ghi(st, 'Q%d' % r, '=B%d+D%d' % (r, r), fmt='0.00')
    ghi(st, 'R%d' % r, '=B%d+H%d' % (r, r), fmt='0.00')
    ghi(st, 'S%d' % r, '=B%d+L%d' % (r, r), fmt='0.00')
ch = ScatterChart()
ch.title = 'Đường mặt nước trên dốc tràn sự cố và đỉnh tường cần'
ch.style = 13
ch.x_axis.title, ch.y_axis.title = 'x (m)', 'Cao trình (m)'
ch.x_axis.delete = ch.y_axis.delete = False
ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0, SC.TSC['L_doc'], 4
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 68, 77
ch.width, ch.height = 20, 10
ch.legend.position = 'b'
xr = Reference(st, min_col=15, min_row=5, max_row=RT2)
for col, nhan, mau, net in ((16, 'Đáy dốc', '52514E', None), (17, 'Lũ thiết kế (có hàm khí)', '2A78D6', None),
                            (18, 'Lũ kiểm tra (có hàm khí)', 'EB6834', None), (19, 'Đỉnh tường cần (h_hk + a)', '0B0B0B', 'sysDot')):
    s = Series(Reference(st, min_col=col, min_row=5, max_row=RT2), xr, title=nhan)
    s.smooth = False
    s.marker.symbol = 'none'
    s.graphicalProperties.line.solidFill = mau
    s.graphicalProperties.line.width = 22225
    if net:
        s.graphicalProperties.line.dashStyle = net
    ch.series.append(s)
st.add_chart(ch, 'A%d' % (RE + 7))
in_an(st)

# ================================================================== KẾT LUẬN (công thức)
kl = wb.create_sheet('KetLuan')
rong(kl, (3, 130))
tieu_de_sheet(kl, 'KẾT LUẬN KIỂM TRA TRÀN CHÍNH, TRÀN SỰ CỐ (tự cập nhật)')
DONG_KL = [
    '="Khả năng xả: MNLTK = +"&FIXED(SL_ZTK,2)&" m, MNLKT = +"&FIXED(SL_ZKT,2)&" m (1 cửa tràn chính không mở: +"&FIXED(SL_ZTKs,2)&'
    '" / +"&FIXED(SL_ZKTs,2)&" m); đỉnh đập sau sửa chữa +"&FIXED(SL_Zdinh,2)&" m cao hơn MNLKT "&FIXED(SL_Zdinh-SL_ZKT,2)&" m."',
    '="Dốc nước: độ vượt cao tường bên nhỏ nhất "&FIXED(MIN(DocNuoc!K5:K%d),2)&" m (yêu cầu "&FIXED(SL_at,2)&" m) – "&'
    'IF(MIN(DocNuoc!K5:K%d)>=SL_at,"tường bên đủ chiều cao","cần nâng tường")&"; lưu tốc lớn nhất "&FIXED(MAX(DocNuoc!H5:H%d),2)&'
    '" m/s < [v]kx = "&FIXED(SL_vkx,1)&" m/s – không xói theo lý thuyết; bề mặt nửa cuối dốc đã mài mòn: sửa chữa bề mặt."'
    % (RT1, RT1, RT1),
    '="Dốc nước – đoạn tường giả thiết: "&DocNuoc!A%d' % (RT1 + 3),
    '="Tiêu năng mũi phun: "&MuiPhun!A%d' % (RM1 + 2),
    'Số liệu đã đối chiếu bình đồ hiện trạng (mặt bằng nâng cấp cụm đầu mối): sân trước ngưỡng ≈ +74,0 m, ngưỡng TP5 +74,70 m, đáy '
    'sau ngưỡng TP7 ≈ +73,30 m → P₁ ≈ 0,7 m, P₂ ≈ 1,40 m (khớp số liệu đang dùng; hồ sơ kiểm định ghi 2,3 / 3,3 m); cuối dốc TP11 '
    '+68,90 m ≈ cao trình mũi phun +69,0 m; chân ngưỡng → cuối mũi phun dài 61,87 m (dốc 60 m, i = 8 % + mũi phun ≈ 1,9 m). Nếu '
    'coi cả đoạn là dốc đều (i = 7,1 %) hoặc theo hồ sơ kiểm định (i = 10 %): lưu tốc cuối dốc 11,4 ÷ 12,7 m/s, độ vượt cao tường '
    'nhỏ nhất 0,74 ÷ 0,85 m – kết luận không đổi. Còn cần đo: chiều cao tường sau x = 40 m (nhất là cuối dốc), góc hắt mũi phun.',
    '="Tràn sự cố – khả năng xả: ε = "&FIXED(TSC_DauDoc!E5,3)&" (mố bên vuông góc, H/b ≤ 1), ε·m = "&FIXED(TSC_DauDoc!F5,3)&'
    '" ≤ giới hạn "&FIXED(TSC_mmax,3)&" của ngưỡng không có bậc hạ; xả "&FIXED(TSC_DauDoc!G5,1)&" m³/s (P = 1%) và "&'
    'FIXED(TSC_DauDoc!G7,1)&" m³/s (P = 0,2%), chiếm "&FIXED(MIN(TSC_DauDoc!I5:I8),0)&" ÷ "&FIXED(MAX(TSC_DauDoc!I5:I8),0)&'
    '" % lưu lượng xả của hồ."',
    '="Tràn sự cố – dốc nước: "&TSC_DocNuoc!A%d&" Hồ sơ chưa ghi chiều cao tường – cần đo đối chiếu."' % (RT2 + 2),
    'Tràn sự cố – tiêu năng cuối dốc: hồ sơ chưa ghi hình thức, kích thước công trình tiêu năng và mực nước hạ lưu nên chưa kiểm tra '
    'được như mũi phun tràn chính; thông số dòng chảy cuối dốc (sheet TSC_DocNuoc) dùng để kiểm tra khi có số liệu khảo sát kênh xả, '
    'hố xói hạ lưu.',
]
for k, txt in enumerate(DONG_KL):
    o = kl['B%d' % (3 + k)]
    o.value = txt
    o.font = f_thuong
    o.alignment = Alignment(wrap_text=True, vertical='top')
    kl.row_dimensions[3 + k].height = 34
in_an(kl)

# ================================================================== HƯỚNG DẪN
rong(hd, (3, 125))
hd.sheet_view.showGridLines = False
ND = [
    ('T', 'KIỂM TRA TRÀN CHÍNH VÀ TRÀN SỰ CỐ HIỆN TRẠNG – ĐỒ ÁN SỬA CHỮA HỒ CÀ GIÂY'),
    ('g', 'Cùng số với Word KiemTra_TranChinh_SuaChua.docx. Ô vàng chữ xanh: số liệu nhập; chữ đen: công thức; cột xám: phép lặp.'),
    ('M', 'Vì sao không tính theo file mẫu thiết kế tràn mới'),
    ('-', 'Tràn chính đang vận hành: không chọn lại hình thức, bề rộng, kênh dẫn, chiều dày bản đáy, kích thước công trình tiêu '
          'năng. Đồ án kiểm tra công trình hiện có với lũ cập nhật như hồ sơ kiểm định an toàn đập (QTVH chương 6): (1) khả năng '
          'xả lũ – điều tiết; (2) thủy lực dốc nước – tường bên, xói; (3) tiêu năng mũi phun – hố xói; (4) không tràn đỉnh đập '
          '(mục 2.3).'),
    ('-', 'Lưu tốc tới gần: hồ sơ kiểm định xác định Ω_T ≫ 4·H·b nên H₀ = H – khả năng xả của ngưỡng chính là công thức đã dùng '
          'khi điều tiết lũ, không cần kiểm tra riêng.'),
    ('M', 'Các sheet'),
    ('-', 'SoLieu: thông số tràn, dốc nước, kết quả điều tiết lũ (file f1f2), số liệu mũi phun, tường bên, quan hệ hạ lưu (Bảng 6-8 '
          'hồ sơ kiểm định).'),
    ('-', 'KhaNangXa: quan hệ Q ~ Z (tràn chính 9 / 8 khoang, tràn sự cố) và kết quả điều tiết – đồ án so với hồ sơ kiểm định.'),
    ('-', 'DauDoc: 6 cấp lưu lượng (50, 100, 150, 200 m³/s, lũ thiết kế, lũ kiểm tra) – cột nước, mực nước hồ, độ sâu co hẹp đầu dốc.'),
    ('-', 'MN_…: đường mặt nước trên dốc (sai phân, chọn trước Δl, thử dần h – như hồ sơ kiểm định).'),
    ('-', 'DocNuoc: chiều sâu có kể hàm khí, độ vượt cao tường bên, lưu tốc so với [v]kx tại x = 0, 5, …, 60 m; biểu đồ.'),
    ('-', 'MuiPhun: chiều dài phóng xa, góc đổ, chiều sâu xói, kích thước hố xói, mép hố xói so với mũi phun; đường bao hố xói.'),
    ('-', 'TSC_DauDoc: tràn sự cố – ε (1 khoang, H/b ≤ 1), ε·m so với giới hạn chảy phân giới, tỷ lệ trong lưu lượng xả, h_k, i_k, '
          'h₀, độ sâu đầu dốc (ngưỡng không có bậc hạ – đầu dốc chảy phân giới).'),
    ('-', 'TSC_MN_…: đường mặt nước dốc tràn sự cố, 4 trường hợp (chọn trước Δl = 2 m, thử dần h – 30 vòng lặp vì đoạn đầu sát '
          'phân giới hội tụ chậm).'),
    ('-', 'TSC_DocNuoc: chiều sâu có kể hàm khí, chiều cao tường cần, lưu tốc tại x = 0, 2, …, 32 m; thông số cuối dốc; biểu đồ.'),
    ('-', 'KetLuan: các câu kết luận tự cập nhật theo số liệu.'),
    ('M', 'Số liệu suy ra, giả thiết'),
    ('-', 'Cao trình mũi phun +69,0 m, góc hắt 25°, kích thước khối đá d = c = 0,8 m, quan hệ Q ~ h, Q ~ MNHL: suy từ Bảng 6-8 hồ sơ '
          'kiểm định (cùng công thức cho lại đúng h_x, b_x của bảng). L tính theo lưu tốc thực ở mũi phun (φ²·Z_H = v_H²/2g) nên ngắn '
          'hơn bảng của hồ sơ – thiên an toàn khi xét hố xói lấn về mũi phun.'),
    ('-', 'Hệ số co hẹp bên: mố bên tràn chính lượn tròn ξmb = 0,70 (tường cánh thượng lưu cong), mố bên tràn sự cố vuông góc '
          'ξmb = 1,00 (bình đồ cửa vào); ε lấy H/b ≤ 1 (khi H/b > 1 lấy H/b = 1) – như file điều tiết lũ, hồ sơ kiểm định.'),
    ('-', 'Chiều cao tường bên dốc tràn chính sau x = 40 m chưa có trong hồ sơ (Bảng 6-7 hồ sơ kiểm định lập đến x = 40 m): giả '
          'thiết 2,5 m; tường cần tối thiểu ở đoạn này nhỏ hơn nhiều (sheet DocNuoc) nên giả thiết không quyết định kết luận. Tràn '
          'sự cố: hồ sơ chưa có chiều cao tường bên và kích thước công trình tiêu năng – chỉ tính chiều cao tường cần và thông số '
          'dòng chảy cuối dốc.'),
]
for h_, (loai, txt) in enumerate(ND, 1):
    o = hd['B%d' % h_]
    o.value = ('•  ' + txt) if loai == '-' else txt
    o.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    o.font = {'T': f_tieu_de, 'M': f_muc, 'g': f_ghi_chu}.get(loai, f_thuong)
    hd.row_dimensions[h_].height = 22 if loai in 'TM' else 15.5 * max(1, -(-len(txt) // 140))
in_an(hd, ngang=False)

MAU_TAB = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'KhaNangXa': '1BAF7A', 'DauDoc': '2A78D6', 'DocNuoc': '1BAF7A',
           'MuiPhun': '1BAF7A', 'TSC_DauDoc': '2A78D6', 'TSC_DocNuoc': '1BAF7A', 'KetLuan': 'D03B3B'}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = MAU_TAB.get(ws.title, '8A8984')
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
