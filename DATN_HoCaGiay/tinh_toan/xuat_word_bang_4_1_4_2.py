# -*- coding: utf-8 -*-
"""
XUẤT WORD: BẢNG 4.1 (CÁC TRỊ SỐ THIẾT KẾ CỦA CÔNG TRÌNH THÁO LŨ) VÀ BẢNG 4.2 (PHÂN LOẠI CÁC TÌNH HUỐNG LŨ VÀ YÊU CẦU ĐỐI VỚI
ĐỈNH ĐẬP, TƯỜNG CHẮN SÓNG) – SV yêu cầu thống kê lại 03/11/2026 (bảng cũ: tràn chính độc lập +78,28 / +79,15, đỉnh đập
+78,75).
Số liệu đọc từ file Excel điều tiết lũ của SV (TAI_VE/DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx: 6 sheet điều tiết tràn hiện
trạng – bình thường, 1 cửa tràn chính không mở, kẹt cửa tràn sự cố; đỉnh đập, đỉnh tường ở KetQua). File được LibreOffice
tính lại trước khi đọc nên số trong Word trùng số trong Excel (sheet Bang4_1_4_2); chữ soạn bằng Python (dấu phẩy thập
phân).
  python3 xuat_word_bang_4_1_4_2.py [đường dẫn .docx ra]
Mặc định: ../TAI_VE/Bang4.1_4.2_CongTrinhThaoLu.docx
"""
import os
import subprocess
import sys
import tempfile
import warnings

import docx
from docx.shared import Cm, Pt
from openpyxl import load_workbook

from xuat_bang_tham_docx import bang, chu, doan, so

warnings.filterwarnings('ignore')
DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(DIR, '..', 'TAI_VE', 'DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx')
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'Bang4.1_4.2_CongTrinhThaoLu.docx')
R0, R1 = 16, 88                     # giờ 0 ÷ 72 trong các sheet điều tiết
# (sheet, nhãn Bảng 4.1, nhãn Bảng 4.2, phân loại, câu khi đỉnh đập cao hơn Zmax ≥ 0,3 m) – như sheet Bang4_1_4_2
TH = [
    ('P1_BinhThuong', 'Lũ thiết kế P = 1% – vận hành bình thường (mở hết cửa hai tràn)',
     'Lũ thiết kế P = 1%, vận hành bình thường', 'Tổ hợp cơ bản (lũ thiết kế)',
     'Đạt yêu cầu an toàn chống tràn; sóng do gió được tường chắn sóng chắn (Z_{2}, mục 2.3)'),
    ('P1_SuCo', 'Lũ thiết kế P = 1% – 1 cửa tràn chính không mở', 'Lũ thiết kế P = 1%, 1 cửa tràn chính không mở',
     'Tổ hợp đặc biệt (sự cố cửa van)', 'Đạt – không tràn đỉnh đập'),
    ('P1_KetTranSuCo', 'Lũ thiết kế P = 1% – kẹt cửa tràn sự cố (tràn chính độc lập)',
     'Lũ thiết kế P = 1%, kẹt cửa tràn sự cố', 'Tổ hợp đặc biệt (sự cố cửa van)', 'Đạt – không tràn đỉnh đập'),
    ('P02_BinhThuong', 'Lũ kiểm tra P = 0,2% – vận hành bình thường (mở hết cửa hai tràn)',
     'Lũ kiểm tra P = 0,2%, vận hành bình thường', 'Tổ hợp đặc biệt (lũ kiểm tra)',
     'Đạt – đỉnh đập ≥ MNLKT + 0,3 m (TCVN 8216:2018, 10.1.5)'),
    ('P02_SuCo', 'Lũ kiểm tra P = 0,2% – 1 cửa tràn chính không mở', 'Lũ kiểm tra P = 0,2%, 1 cửa tràn chính không mở',
     'Lũ kiểm tra + sự cố cửa van', 'Đạt – không tràn đỉnh đập'),
    ('P02_KetTranSuCo', 'Lũ kiểm tra P = 0,2% – kẹt cửa tràn sự cố (tràn chính độc lập)',
     'Lũ kiểm tra P = 0,2%, kẹt cửa tràn sự cố', 'Sự cố chồng sự cố (lũ kiểm tra + kẹt cửa tràn sự cố)',
     'Đạt – không tràn đỉnh đập'),
]
T_VUOT = ('Vượt đỉnh đập, thấp hơn đỉnh tường: tường làm việc như tường chắn nước nhưng chưa được kiểm tra chịu lực, chống '
          'thấm – không coi là điều kiện bảo đảm an toàn. Biện pháp: kiểm tra, bảo dưỡng cửa van tràn sự cố trước mùa lũ; có '
          'nguồn điện và cơ cấu kéo cửa dự phòng; xả trước khi có dự báo lũ lớn')


def doc_excel():
    """Tính lại file Excel bằng LibreOffice (file do openpyxl ghi chưa có giá trị) rồi đọc kết quả."""
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', tmp, XLSX], check=True,
                       capture_output=True, timeout=900)
        wb = load_workbook(os.path.join(tmp, os.path.basename(XLSX)), data_only=True)
        kq = {}
        for ten in ('Z_dinh_SC', 'Z_tuong'):
            ws_ten, dc = next(wb.defined_names[ten].destinations)
            kq[ten] = wb[ws_ten][dc.replace('$', '')].value
        for s_, *_ in TH:
            ws = wb[s_]
            tsc = isinstance(ws['G10'].value, float)
            kq[s_] = dict(Q=max(ws['B%d' % r].value for r in range(R0, R1 + 1)), Z=ws['F6'].value, Htc=ws['F10'].value,
                          qtc=ws['F9'].value, Hsc=ws['G10'].value if tsc else None, qsc=ws['G9'].value if tsc else None,
                          q=ws['F9'].value + ws['G9'].value, P=ws['C3'].value, th=ws['F3'].value)
    return kq


def so_nghin(x):
    return '{:,.0f}'.format(x).replace(',', '.')


def so_voi(chenh):
    """chenh = cao trình so sánh − Zmax."""
    return ('thấp hơn %s m' if chenh >= 0 else 'cao hơn %s m') % so(abs(chenh))


def yeu_cau(z, dinh, tuong, dat):
    if z + 0.3 <= dinh + 1e-9:
        return dat
    if z <= dinh:
        return 'Không tràn đỉnh đập, độ vượt cao %s m < 0,3 m' % so(dinh - z)
    if z <= tuong:
        return T_VUOT
    return 'Tràn qua đỉnh tường chắn sóng – không đạt'


def main():
    kq = doc_excel()
    dinh, tuong = kq['Z_dinh_SC'], kq['Z_tuong']
    for s_, *_ in TH:                                      # kiểm tra đúng sheet, đúng trường hợp
        assert kq[s_]['P'] == (1 if s_.startswith('P1_') else 0.2), s_
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3.0), Cm(2.0), Cm(2.5), Cm(2.5)
    st = d.styles['Normal']
    st.font.name, st.font.size = 'Times New Roman', Pt(13)

    doan(d, 'Kết quả điều tiết lũ (phương pháp Pô-ta-pốp, lũ đến theo QTVH 2026, hồ ở MNDBT +74,70 m khi lũ đến, cửa van kéo '
            'lên hoàn toàn khi đón lũ) cho ba trạng thái làm việc của công trình tháo lũ – vận hành bình thường, 1 cửa tràn chính '
            'không mở được và cửa van tràn sự cố bị kẹt – được thống kê trong Bảng 4.1.', thut=1.0)
    hang = []
    for s_, nhan, *_ in TH:
        r = kq[s_]
        hang.append([nhan, so_nghin(r['Q']), '+' + so(r['Z']), so(r['Htc']), so(r['qtc']),
                     so(r['Hsc']) if r['Hsc'] is not None else '–', so(r['qsc']) if r['qsc'] is not None else '–',
                     so(r['q'])])
    bang(d, 'Bảng 4.1 Các trị số thiết kế của công trình tháo lũ',
         ['Trường hợp', 'Q_{max} đến (m^{3}/s)', 'Z_{max} (m)', 'H tràn chính (m)', 'q tràn chính (m^{3}/s)',
          'H tràn sự cố (m)', 'q tràn sự cố (m^{3}/s)', 'q tổng (m^{3}/s)'], hang,
         cot=[3070, 780, 860, 800, 900, 800, 900, 960])

    bt, sc, ket1, ket = kq['P02_BinhThuong'], kq['P02_SuCo'], kq['P1_KetTranSuCo'], kq['P02_KetTranSuCo']
    assert max(kq[s_]['Z'] for s_, *_ in TH if s_ != 'P02_KetTranSuCo') < dinh < ket['Z'] < tuong   # đoạn văn viết cho thứ tự này
    p = doan(d, 'Các tình huống lũ được phân loại trong Bảng 4.2. Khi vận hành bình thường hoặc khi 1 cửa tràn chính không mở '
                'được, mực nước hồ lớn nhất đều thấp hơn đỉnh đập sau sửa chữa +%s m; với lũ kiểm tra, đỉnh đập cao hơn MNLKT '
                '%s m (≥ 0,3 m) và cao hơn mực nước khi 1 cửa tràn chính không mở %s m. Cửa van tràn sự cố bị kẹt trong lũ thiết '
                'kế làm mực nước lên +%s m, vẫn thấp hơn đỉnh đập %s m. Khi cửa van tràn sự cố bị kẹt trong lũ kiểm tra (sự cố '
                'chồng sự cố), tràn chính phải xả một mình, mực nước hồ lên +%s m – vượt đỉnh đập sau sửa chữa %s m nhưng còn '
                'thấp hơn đỉnh tường chắn sóng (+%s m) %s m; tường chắn sóng khi đó làm việc như tường chắn nước nhưng chưa được '
                'kiểm tra chịu lực, chống thấm cho trạng thái này nên '
             % (so(dinh), so(dinh - bt['Z']), so(dinh - sc['Z']), so(ket1['Z']), so(dinh - ket1['Z']), so(ket['Z']),
                so(ket['Z'] - dinh), so(tuong), so(tuong - ket['Z'])), thut=1.0, truoc=8)
    chu(p, 'không được coi là điều kiện bảo đảm an toàn', dam=True)
    chu(p, '.')
    hang = []
    for s_, _, nhan, loai, dat in TH:
        z = kq[s_]['Z']
        hang.append([nhan, '+' + so(z), so_voi(dinh - z), so_voi(tuong - z), loai, yeu_cau(z, dinh, tuong, dat)])
    bang(d, 'Bảng 4.2 Phân loại các tình huống lũ và yêu cầu đối với đỉnh đập, tường chắn sóng',
         ['Tình huống', 'Z_{max} (m)', 'So với đỉnh đập +%s m' % so(dinh), 'So với đỉnh tường +%s m' % so(tuong),
          'Phân loại', 'Yêu cầu, biện pháp'], hang, cot=[2150, 830, 1100, 1100, 1530, 2360], cot_trai=(0, 4, 5))
    d.save(RA)
    print('Đã ghi', os.path.abspath(RA))
    for s_, *_ in TH:
        r = kq[s_]
        print('  %-16s Zmax %.3f  đỉnh đập − Z %+.3f  tường − Z %+.3f' % (s_, r['Z'], dinh - r['Z'], tuong - r['Z']))


if __name__ == '__main__':
    main()
