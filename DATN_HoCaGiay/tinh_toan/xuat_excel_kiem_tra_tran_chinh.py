# -*- coding: utf-8 -*-
"""
XUẤT FILE EXCEL: KIỂM TRA TRÀN CHÍNH HIỆN TRẠNG TRONG ĐỒ ÁN SỬA CHỮA (theo hồ sơ kiểm định an toàn đập – QTVH chương 6)
Cùng nội dung, cùng số với Word TAI_VE/KiemTra_TranChinh_SuaChua.docx (kiem_tra_tran_chinh.py). Mọi ô kết quả là công thức;
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
tieu_de_sheet(sl, 'SỐ LIỆU ĐẦU VÀO – KIỂM TRA TRÀN CHÍNH HIỆN TRẠNG (đồ án sửa chữa)',
              'Ô nền vàng chữ xanh: số liệu nhập. Mục A, C lấy từ file điều tiết lũ DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx; mục F, '
              'G từ hồ sơ kiểm định an toàn đập (QTVH chương 6).')
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
    ('Hệ số hình dạng mố bên ξmb', ts['TC_xmb'], '', 'SL_xmb', '', '0.00'),
    ('Hệ số lưu lượng m', ts['TC_m'], '', 'SL_m', 'Hồ sơ kiểm định: 0,445', '0.000'),
    ('Hệ số chảy ngập σn', ts['TC_s'], '', 'SL_sn', 'Chảy không ngập (i > i_k)', '0.00'),
    ('Số cửa không mở được khi sự cố', 1, 'cửa', 'SL_hong', 'Cửa khoang biên', '0'),
    ('Σb = n·b', '=SL_nk*SL_b', 'm', 'SL_Sb', 'Hồ sơ kiểm định: 22 m', '0.00'),
    ('Σξ = n₁ξ₁ + n₂ξ₂', '=SL_n1*SL_x1+SL_n2*SL_x2', '', 'SL_Sxi', '', '0.00'),
    ('Ngưỡng tràn sự cố', ts['SC_Z'], 'm', 'SC_Z', '1 khoang, có cửa van', '0.00'),
    ('Bề rộng tràn sự cố', 5.0, 'm', 'SC_b', '', '0.00'),
    ('Hệ số lưu lượng tràn sự cố', 0.45, '', 'SC_m', '', '0.000'),
    ('Hệ số mố bên tràn sự cố', 0.70, '', 'SC_xmb', '', '0.00'),
])
khoi('B. NGƯỠNG TRÀN, DỐC NƯỚC (KTTV Bảng 1-2, BCC, PL Bảng PL1-5)', [
    ('Chiều cao bậc hạ sau ngưỡng P₂', TC['P2'], 'm', 'SL_P2', 'BCC; hồ sơ kiểm định ghi P = 3,3 m – đối chiếu hoàn công', '0.00'),
    ('Cao trình đầu dốc = Z_ngưỡng − P₂', '=SL_Zn-SL_P2', 'm', 'SL_Zdd', '', '0.00'),
    ('Chiều dài dốc', TC['L_doc'], 'm', 'SL_Ld', 'KTTV Bảng 1-2', '0.0'),
    ('Độ dốc đáy i', TC['i_doc'], '', 'SL_i', 'KTTV Bảng 1-2 (hồ sơ kiểm định ghi 10 %)', '0.000'),
    ('Hệ số nhám n', T.N_NHAM, '', 'SL_nn', 'Bê tông', '0.000'),
    ('Bề rộng đầu dốc', TC['b_doc_dau'], 'm', 'SL_Bd', 'PL Bảng PL1-5', '0.00'),
    ('Bề rộng đoạn lăng trụ', TC['b_doc_cuoi'], 'm', 'SL_Bc', 'PL Bảng PL1-5', '0.00'),
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
ghi(sl, 'A%d' % r0, 'E. CHIỀU CAO TƯỜNG BÊN DỐC NƯỚC HIỆN TRẠNG (PL Bảng PL1-4 = Bảng 6-7 hồ sơ kiểm định; sau x = 40 m giả thiết)',
    f_muc, can=None, border=False)
tieu_de_cot(sl, r0 + 1, 1, ['x (m)', 'h tường (m)', '', 'Ghi chú'])
X_T = list(range(0, int(TC['L_doc']) + 1, 5))
for k, x in enumerate(X_T):
    rr = r0 + 2 + k
    ghi(sl, 'A%d' % rr, x, f_nhap, '0', nen_nhap)
    ghi(sl, 'B%d' % rr, TC['h_tuong'][k] if k < len(TC['h_tuong']) else TC['h_tuong'][-1], f_nhap, '0.00', nen_nhap)
    ghi(sl, 'D%d' % rr, 'Hồ sơ' if k < len(TC['h_tuong']) else 'Giả thiết', f_ghi_chu, can=trai)
ten('SL_xt', 'SoLieu!$A$%d:$A$%d' % (r0 + 2, r0 + 1 + len(X_T)))
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
sl.freeze_panes = 'A4'
in_an(sl, ngang=False)

# ================================================================== KHẢ NĂNG XẢ
kx = wb.create_sheet('KhaNangXa')
rong(kx, (12, 11, 11, 11, 15, 15, 14, 15, 3, 40))
tieu_de_sheet(kx, 'KHẢ NĂNG XẢ LŨ CỦA NGƯỠNG TRÀN – QUAN HỆ Q ~ Z (TCVN 9147:2012, H₀ = H: bỏ qua lưu tốc tới gần)',
              'Q = σn·ε·m·Σb·√(2g)·H^1,5;  ε = 1 − 0,2·(ξmb + Σξ)·H/(n·b) – như khi điều tiết lũ. Hồ sơ kiểm định: Ω_T ≫ 4·H·b nên '
              'không xét V₀; ε = 0,895 (ξmt = 0,25, H₀/b ≤ 1), m = 0,445, Σb = 22 m.')
tieu_de_cot(kx, 4, 1, ['Z (m)', 'H (m)', 'ε (9 khoang)', 'ε (8 khoang)', 'q tràn chính 9 khoang (m³/s)',
                       'q tràn chính 8 khoang (m³/s)', 'q tràn sự cố (m³/s)', 'q tổng – bình thường (m³/s)'], cao=45)
Z_DS = [round(74.70 + 0.25 * k, 2) for k in range(20)]
for k, z in enumerate(Z_DS + ['=SL_ZTK', '=SL_ZKT']):
    r = 5 + k
    nhap = not isinstance(z, str)
    ghi(kx, 'A%d' % r, z, f_nhap if nhap else f_dam, '0.00', nen_nhap if nhap else nen_kq)
    ghi(kx, 'B%d' % r, '=MAX(A%d-SL_Zn,0)' % r, fmt='0.000', fill=nen_kq)
    ghi(kx, 'C%d' % r, '=1-0.2*(SL_xmb+SL_Sxi)*B%d/(SL_nk*SL_b)' % r, fmt='0.000', fill=nen_kq)
    ghi(kx, 'D%d' % r, '=1-0.2*(SL_xmb+SL_Sxi-SL_hong*SL_x1)*B%d/((SL_nk-SL_hong)*SL_b)' % r, fmt='0.000', fill=nen_kq)
    ghi(kx, 'E%d' % r, '=SL_sn*C{r}*SL_m*SL_Sb*SQRT(2*SL_g)*B{r}^1.5'.format(r=r), f_dam, '0.00', nen_kq)
    ghi(kx, 'F%d' % r, '=SL_sn*D{r}*SL_m*(SL_nk-SL_hong)*SL_b*SQRT(2*SL_g)*B{r}^1.5'.format(r=r), fmt='0.00', fill=nen_kq)
    ghi(kx, 'G%d' % r, '=(1-0.2*SC_xmb*MAX(A{r}-SC_Z,0)/SC_b)*SC_m*SC_b*SQRT(2*SL_g)*MAX(A{r}-SC_Z,0)^1.5'.format(r=r), fmt='0.00',
        fill=nen_kq)
    ghi(kx, 'H%d' % r, '=E%d+G%d' % (r, r), f_dam, '0.00', nen_kq)
    if not nhap:
        ghi(kx, 'J%d' % r, 'MNLTK (kết quả điều tiết)' if k == len(Z_DS) else 'MNLKT (kết quả điều tiết)', f_ghi_chu, can=trai)
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
              'H giải ngược từ công thức khả năng xả (8 vòng lặp, cột J ÷ Q); h_c từ Q = φ·B_đ·h_c·√(2g(E₀ − h_c)), E₀ = P₂ + H '
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
        e_ = '1' if it == 0 else '(1-0.2*(SL_xmb+SL_Sxi)*%s%d/(SL_nk*SL_b))' % (prev, r)
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
rong(dn, (8, 9, 9, 10, 9, 9, 10, 9, 11, 12, 12, 12, 11, 3) + (10,) * 6)
tieu_de_sheet(dn, 'KẾT QUẢ KIỂM TRA THỦY LỰC DỐC NƯỚC – TƯỜNG BÊN, XÓI (như Bảng 6-7 hồ sơ kiểm định)',
              'h_hk = h·(1 + v/100); độ vượt cao = chiều cao tường − h_hk; không xói khi v < [v]kx.')
tieu_de_cot(dn, 4, 1, ['x (m)', 'B (m)', 'h TK (m)', 'h_hk TK (m)', 'v TK (m/s)', 'h KT (m)', 'h_hk KT (m)', 'v KT (m/s)',
                       'Tường bên (m)', 'Độ vượt cao TK (m)', 'Độ vượt cao KT (m)', 'Không tràn tường', 'Không xói'], cao=45)
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
RT1 = 4 + len(X_T)
ghi(dn, 'A%d' % (RT1 + 2), '="Độ vượt cao nhỏ nhất "&FIXED(MIN(K5:K%d),2)&" m; lưu tốc lớn nhất "&FIXED(MAX(H5:H%d),2)&'
                           '" m/s so với [v]kx = "&FIXED(SL_vkx,1)&" m/s"' % (RT1, RT1), f_dam, can=None, border=False)
# số liệu vẽ hình
tieu_de_cot(dn, 4, 15, ['x (m)', 'Đáy dốc', 'MN lũ thiết kế (hàm khí)', 'MN lũ kiểm tra (hàm khí)', 'Đỉnh tường'], cao=45)
for k in range(len(X_T)):
    r = 5 + k
    ghi(dn, 'O%d' % r, '=A%d' % r, fmt='0')
    ghi(dn, 'P%d' % r, '=SL_Zdd-SL_i*O%d' % r, fmt='0.00')
    ghi(dn, 'Q%d' % r, '=P%d+D%d' % (r, r), fmt='0.00')
    ghi(dn, 'R%d' % r, '=P%d+G%d' % (r, r), fmt='0.00')
    ghi(dn, 'S%d' % r, '=P%d+I%d' % (r, r), fmt='0.00')
ch = ScatterChart()
ch.title = 'Đường mặt nước trên dốc tràn chính và đỉnh tường bên'
ch.style = 13
ch.x_axis.title, ch.y_axis.title = 'x (m)', 'Cao trình (m)'
ch.x_axis.delete = ch.y_axis.delete = False
ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0, TC['L_doc'], 10
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 68, 79
ch.width, ch.height = 20, 10
ch.legend.position = 'b'
xr = Reference(dn, min_col=15, min_row=5, max_row=RT1)
for col, nhan, mau, net, diem in ((16, 'Đáy dốc', '52514E', None, None), (17, 'Lũ thiết kế (có hàm khí)', '2A78D6', None, None),
                                  (18, 'Lũ kiểm tra (có hàm khí)', 'EB6834', None, None),
                                  (19, 'Đỉnh tường bên', '0B0B0B', 'sysDot', 'circle')):
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
dn.add_chart(ch, 'A%d' % (RT1 + 4))
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

# ================================================================== KẾT LUẬN (công thức)
kl = wb.create_sheet('KetLuan')
rong(kl, (3, 130))
tieu_de_sheet(kl, 'KẾT LUẬN KIỂM TRA TRÀN CHÍNH (tự cập nhật)')
DONG_KL = [
    '="Khả năng xả: MNLTK = +"&FIXED(SL_ZTK,2)&" m, MNLKT = +"&FIXED(SL_ZKT,2)&" m (1 cửa tràn chính không mở: +"&FIXED(SL_ZTKs,2)&'
    '" / +"&FIXED(SL_ZKTs,2)&" m); đỉnh đập sau sửa chữa +"&FIXED(SL_Zdinh,2)&" m cao hơn MNLKT "&FIXED(SL_Zdinh-SL_ZKT,2)&" m."',
    '="Dốc nước: độ vượt cao tường bên nhỏ nhất "&FIXED(MIN(DocNuoc!K5:K%d),2)&" m (yêu cầu "&FIXED(SL_at,2)&" m) – "&'
    'IF(MIN(DocNuoc!K5:K%d)>=SL_at,"tường bên đủ chiều cao","cần nâng tường")&"; lưu tốc lớn nhất "&FIXED(MAX(DocNuoc!H5:H%d),2)&'
    '" m/s < [v]kx = "&FIXED(SL_vkx,1)&" m/s – không xói theo lý thuyết; bề mặt nửa cuối dốc đã mài mòn: sửa chữa bề mặt."'
    % (RT1, RT1, RT1),
    '="Tiêu năng mũi phun: "&MuiPhun!A%d' % (RM1 + 2),
    'Số liệu cần đối chiếu hoàn công: P₂ (BCC 1,40 m – kiểm định 3,3 m), độ dốc dốc nước (8 % – 10 %), chiều cao tường sau x = 40 m, '
    'cao trình mũi phun, góc hắt. Tính lại với P = 3,3 m, i = 10 %: lưu tốc cuối dốc ≤ 13,5 m/s, độ vượt cao tường ≥ 1,0 m – kết '
    'luận không đổi.',
]
for k, txt in enumerate(DONG_KL):
    o = kl['B%d' % (3 + k)]
    o.value = txt
    o.font = f_thuong
    o.alignment = Alignment(wrap_text=True, vertical='top')
    kl.row_dimensions[3 + k].height = 32
in_an(kl)

# ================================================================== HƯỚNG DẪN
rong(hd, (3, 125))
hd.sheet_view.showGridLines = False
ND = [
    ('T', 'KIỂM TRA TRÀN CHÍNH HIỆN TRẠNG – ĐỒ ÁN SỬA CHỮA HỒ CÀ GIÂY'),
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
    ('-', 'KetLuan: các câu kết luận tự cập nhật theo số liệu.'),
    ('M', 'Số liệu suy ra, giả thiết'),
    ('-', 'Cao trình mũi phun +69,0 m, góc hắt 25°, kích thước khối đá d = c = 0,8 m, quan hệ Q ~ h, Q ~ MNHL: suy từ Bảng 6-8 hồ sơ '
          'kiểm định (cùng công thức cho lại đúng h_x, b_x của bảng). L tính theo lưu tốc thực ở mũi phun (φ²·Z_H = v_H²/2g) nên ngắn '
          'hơn bảng của hồ sơ – thiên an toàn khi xét hố xói lấn về mũi phun.'),
]
for h_, (loai, txt) in enumerate(ND, 1):
    o = hd['B%d' % h_]
    o.value = ('•  ' + txt) if loai == '-' else txt
    o.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    o.font = {'T': f_tieu_de, 'M': f_muc, 'g': f_ghi_chu}.get(loai, f_thuong)
    hd.row_dimensions[h_].height = 22 if loai in 'TM' else 15.5 * max(1, -(-len(txt) // 140))
in_an(hd, ngang=False)

MAU_TAB = {'HuongDan': 'A65A10', 'SoLieu': 'EDA100', 'KhaNangXa': '1BAF7A', 'DauDoc': '2A78D6', 'DocNuoc': '1BAF7A',
           'MuiPhun': '1BAF7A', 'KetLuan': 'D03B3B'}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = MAU_TAB.get(ws.title, '8A8984')
wb.calculation.fullCalcOnLoad = True
wb.save(RA)
print('Đã ghi', os.path.abspath(RA))
