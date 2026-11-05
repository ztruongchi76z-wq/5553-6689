# -*- coding: utf-8 -*-
"""
ĐỐI CHIẾU VỚI TCVN 10778:2024 VÀ TCVN 4116:2023 (đêm 11/10/2026, SV tải lên bản đầy đủ hai tiêu chuẩn).

  python3 tiep_thu_gop_y_5.py <vào.docx> <ra.docx> <thư mục ket_qua của tinh_toan>
  python3 chen_muc_luc.py <ra.docx> <bản cuối.docx> <bản cuối.pdf>

TCVN 10778:2024 (Hồ chứa nước – Xác định các mực nước đặc trưng): mực nước chết kiểm tra thêm điều kiện bồi lắng
(mục 7.1 ÷ 7.3); phương trình cân bằng nước (6) có dòng chảy tối thiểu Q_tt – tính lại (kt_dong_chay_toi_thieu.py);
mức bảo đảm theo công thức (14), điều kiện số năm đủ nước (9); tổn thất thấm theo Bảng 1; điều tiết lũ mục 10.
TCVN 4116:2023 (Kết cấu BT và BTCT thủy công): bê tông ghi theo cấp độ bền B (mục 5.1.2), vật liệu vùng chịu mài
mòn, xâm thực tối thiểu B25, W8 (mục 5.1.9), lớp bê tông bảo vệ ≥ 60 mm ở mặt tiếp xúc trực tiếp với nước (mục 6.3.3)
– tính lại Bảng 6.4 (bo_sung_giai_tich.py, a = 7 cm). Không đổi số thứ tự tài liệu tham khảo.
Nhật ký: nhat_ky_tiep_thu_gop_y_5.txt.
"""
import json
import os
import re
import sys

import docx
from docx.oxml.ns import qn

import cong_cu_docx as C
from cong_cu_docx import ghi, tim_doan, tim_bang, thay, van_ban, hang, them_hang, dat_o, doan_moi, chen_sau


def so(x, n=1):
    return ('%.*f' % (n, x)).replace('.', ',')


def thay_trong_bang(tbl, cu, moi):
    """Thay `cu` bằng `moi` trong mọi đoạn của bảng; trả về số chỗ đã thay."""
    n = 0
    for p in tbl.iter(qn('w:p')):
        k = van_ban(p).count(cu)
        if k:
            thay(p, cu, moi, lan=k)
            n += k
    return n


# =========================================================================== DANH MỤC VIẾT TẮT, BẢNG 2.1
def dau_sach(body):
    for tbl in body.iter(qn('w:tbl')):
        for tr in hang(tbl):
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) == 2 and van_ban(tcs[0]).strip() in ('MNLKT', 'MNLTK'):
                ten = van_ban(tcs[0]).strip()
                moi = {'MNLKT': 'Mực nước lũ kiểm tra (mực nước lớn nhất kiểm tra – TCVN 10778:2024)',
                       'MNLTK': 'Mực nước lũ thiết kế (mực nước lớn nhất thiết kế – TCVN 10778:2024)'}[ten]
                dat_o(tcs[1], moi)
    ghi('Danh mục viết tắt: MNLTK, MNLKT ghi thêm tên gọi theo TCVN 10778:2024 (mực nước lớn nhất thiết kế, kiểm tra).')

    cap, tbl = tim_bang(body, 'Bảng 2.1')
    rows = hang(tbl)
    assert van_ban(rows[-1].findall(qn('w:tc'))[0]).strip() == '10', van_ban(rows[-1])
    r11 = them_hang(tbl, len(rows) - 1, [
        '11', 'Bê tông kết cấu thủy công: chỉ tiêu thiết kế; lớp bê tông bảo vệ cốt thép ở mặt tiếp xúc trực tiếp '
              'với nước',
        'Cấp độ bền chịu nén B, mác chống thấm W (mác M ghi kèm để đối chiếu); ≥ 60 mm và ≥ 2d',
        'TCVN 4116:2023, mục 5.1.2, 6.3.3'])
    them_hang(tbl, len(rows), [
        '12', 'Mức bảo đảm cấp nước tính toán; tổn thất thấm lòng hồ; dòng chảy tối thiểu',
        'P = m/(n + 1); 1%/tháng; Q_{tt} = 0,027 m^{3}/s',
        'TCVN 10778:2024, công thức (6), (14), Bảng 1; QTVH [1]'])
    ghi('Bảng 2.1: thêm dòng 11 (TCVN 4116:2023 – cấp độ bền B, mác chống thấm W, lớp bảo vệ ≥ 60 mm) và dòng 12 '
        '(TCVN 10778:2024 – P = m/(n + 1), tổn thất thấm, dòng chảy tối thiểu Q_tt = 0,027 m³/s).')


# =========================================================================== CHƯƠNG 2
def chuong2(body, kq):
    p = tim_doan(body, 'Mực nước chết của hồ MNC = +67,50 m được giữ như thiết kế.')
    thay(p, 'Mực nước chết của hồ MNC = +67,50 m được giữ như thiết kế. Kiểm tra theo điều kiện lấy nước:',
         'Mực nước chết của hồ MNC = +67,50 m được giữ như thiết kế và kiểm tra lại theo các yêu cầu của '
         'TCVN 10778:2024 (mục 7.1 ÷ 7.3). Theo điều kiện bồi lắng: tổng lượng bùn cát đến hồ trong 70 năm khoảng '
         '0,48 triệu m^{3} [1], nhỏ hơn nhiều dung tích dưới ngưỡng cống +65,60 m (3,52 triệu m^{3} theo bình đồ 2022) '
         'nên sau thời gian khai thác mặt bùn cát vẫn thấp hơn ngưỡng cửa lấy nước và MNC. Theo điều kiện lấy nước:')
    ghi('Mục 2.2.1: kiểm tra MNC theo TCVN 10778:2024 mục 7.1 ÷ 7.3 – thêm điều kiện bồi lắng (bùn cát 70 năm '
        '≈ 0,48 triệu m³ theo QTVH mục 3.5.4, nhỏ hơn dung tích dưới ngưỡng cống 3,52 triệu m³).')

    p = tim_doan(body, 'Khả năng cấp nước được đánh giá bằng điều tiết toàn chuỗi')
    thay(p, 'theo thời đoạn tháng [14], [15] với chuỗi dòng chảy đến 44 năm thủy văn',
         'theo thời đoạn tháng [14], [15] (TCVN 10778:2024 mục 8.2.1.2: áp dụng khi chuỗi dòng chảy đến từ 15 năm '
         'trở lên) với chuỗi dòng chảy đến 44 năm thủy văn')

    p = tim_doan(body, 'Trong đó: Vi – dung tích hồ cuối tháng')
    thay(p, 'Wđến, Wdùng – lượng nước đến và lượng nước dùng trong tháng;',
         'W_{đến}, W_{dùng} – lượng nước đến và lượng nước dùng trong tháng (W_{dùng} gồm nhu cầu tưới, sinh hoạt theo QTVH '
         '[1]; khi kiểm tra theo công thức (6) của TCVN 10778:2024 cộng thêm lượng xả duy trì dòng chảy tối thiểu '
         'Q_{tt} = 0,027 m^{3}/s, khoảng 0,07 triệu m^{3}/tháng [1]);')
    thay(p, 'kt = 1%/tháng – tổn thất thấm (nền tương đối tốt);',
         'k_{t} = 1%/tháng – tổn thất thấm (nền tương đối tốt; TCVN 10778:2024 Bảng 1: 0,5 ÷ 1,0%/tháng với địa chất '
         'lòng hồ tốt, 1,0 ÷ 1,5%/tháng với địa chất trung bình);')

    p = tim_doan(body, 'Tháng nào V chạm dung tích chết thì phải cắt giảm cấp nước')
    thay(p, 'Mức bảo đảm tính theo công thức tần suất kinh nghiệm P = m/(n + 1)·100% [14], với m – số năm đủ nước, '
            'n = 44 (QTVH [1] dùng P = m/n·100%; kết quả theo cả hai cách được so sánh ở mục 2.2.3).',
         'Mức bảo đảm tính theo công thức tần suất kinh nghiệm P = m/(n + 1)·100% [14] – công thức (14) của '
         'TCVN 10778:2024 [15], với m – số năm đủ nước, n = 44; điều kiện đạt mức bảo đảm thiết kế P_{CN} = 85% theo '
         'công thức (9) của tiêu chuẩn là m ≥ (n + 1)·P_{CN}/100 = 38,25, tức ít nhất 39 năm đủ nước (QTVH [1] dùng '
         'P = m/n·100% – khác tiêu chuẩn; kết quả theo cả hai cách được so sánh ở mục 2.2.3).')
    ghi('Mục 2.2.2: phương pháp toàn chuỗi theo TCVN 10778:2024 mục 8.2.1.2; Wdùng và Q_tt theo công thức (6); '
        'tổn thất thấm đối chiếu Bảng 1; P = m/(n + 1) là công thức (14), điều kiện m ≥ 39 năm theo công thức (9).')

    a, b = kq['TH1A|khong_Qtt'], kq['TH1A|co_Qtt']
    assert round(a['P'], 1) == round(b['P'], 1) == 86.7 and round(a['MNDBT_can'], 2) == 74.62
    p = tim_doan(body, 'Kết quả điều tiết (Bảng 2.2, Hình 2.1) cho thấy:')
    thay(p, 'mức bảo đảm đạt P = 86,7% ≥ 85%, mực nước dâng bình thường',
         'mức bảo đảm đạt P = 86,7% ≥ 85% (39/44 năm đủ nước – vừa đạt điều kiện m ≥ 39 năm của công thức (9) '
         'TCVN 10778:2024), mực nước dâng bình thường')
    thay(p, 'TH2 đạt 84,1% – kết luận không đổi.',
         'TH2 đạt 84,1%% – kết luận không đổi. Kiểm tra theo công thức (6) của TCVN 10778:2024 có kể dòng chảy tối '
         'thiểu Q_{tt} = 0,027 m^{3}/s (%s triệu m^{3}/năm [1], chưa có trong tổng nhu cầu của QTVH): số năm đủ nước '
         'và mức bảo đảm của cả ba phương án không đổi (%s%%; %s%%; %s%%), tổng lượng nước thiếu tăng lên %s; %s và '
         '%s triệu m^{3}; mực nước dâng bình thường cần thiết để TH1A đạt 85%% tăng từ +%s m lên +%s m – vẫn thấp hơn '
         'MNDBT +74,70 m (chương trình kt_dong_chay_toi_thieu.py).' % (
             so(kq['W_tt_nam'], 2), so(kq['TH1|co_Qtt']['P']), so(b['P']), so(kq['TH2|co_Qtt']['P']),
             so(kq['TH1|co_Qtt']['W_thieu']), so(b['W_thieu']), so(kq['TH2|co_Qtt']['W_thieu']),
             so(a['MNDBT_can'], 2), so(b['MNDBT_can'], 2)))
    ghi('Mục 2.2.3: kết quả kể dòng chảy tối thiểu Q_tt = 0,027 m³/s (TCVN 10778:2024 công thức 6): P không đổi '
        '(40,0 / 86,7 / 82,2%%), ΣW thiếu %s / %s / %s triệu m³, MNDBT cần cho TH1A +74,62 → +%s m (< +74,70 m).' % (
            so(kq['TH1|co_Qtt']['W_thieu']), so(b['W_thieu']), so(kq['TH2|co_Qtt']['W_thieu']),
            so(b['MNDBT_can'], 2)))

    p = tim_doan(body, 'Điều tiết lũ dựa trên phương trình cân bằng nước viết cho thời đoạn Δt [14]:')
    thay(p, 'viết cho thời đoạn Δt [14]:',
         'viết cho thời đoạn Δt [14], [15] (TCVN 10778:2024 mục 10.1.2; bỏ qua tổn thất bốc hơi, thấm và lưu lượng '
         'lấy nước):')
    p = tim_doan(body, 'Phương trình (2-2) kết hợp với quan hệ q = q(Z) của hai tràn')
    thay(p, 'Sai số cân bằng nước toàn trận lũ nhỏ hơn 0,02 triệu m3.',
         'Sai số mực nước khi lặp nhỏ hơn nhiều giới hạn 5 cm của phương pháp tính lặp trực tiếp (TCVN 10778:2024 '
         'mục 10.2); sai số cân bằng nước toàn trận lũ nhỏ hơn 0,02 triệu m^{3}.')
    ghi('Mục 2.3.1: điều tiết lũ dẫn TCVN 10778:2024 mục 10.1.2 (phương trình cân bằng nước), mục 10.2 (sai số lặp ≤ 5 cm).')


# =========================================================================== BÊ TÔNG – TCVN 4116:2023
def be_tong(body):
    p = tim_doan(body, 'Hồ sơ TKCS [2] đề xuất giữ nguyên tường chắn sóng')
    thay(p, 'bằng bê tông M250 dày 20 cm', 'bằng bê tông M250 (cấp độ bền B20) dày 20 cm')
    p = tim_doan(body, 'Đồ án chọn phương án Đ3')
    thay(p, 'bê tông M250 đá 1×2 dày 20 cm.', 'bê tông B20 (M250) đá 1×2 dày 20 cm.')
    thay(p, 'gờ chắn bánh xe BT M250', 'gờ chắn bánh xe BT B20 (M250)')
    p = tim_doan(body, 'Đường quản lý: nâng cấp, bê tông hóa toàn tuyến')
    thay(p, 'mặt đường BT M250', 'mặt đường BT B20 (M250)')
    p = tim_doan(body, 'Rải lớp cấp phối đá dăm đệm dày 5 cm, lu lèn chặt')
    thay(p, 'đổ bê tông M250 đá 1×2', 'đổ bê tông B20 (M250) đá 1×2')
    p = tim_doan(body, 'Gờ chắn bánh xe BT M250 cao 0,25 m phía hạ lưu')
    thay(p, 'Gờ chắn bánh xe BT M250', 'Gờ chắn bánh xe BT B20 (M250)')
    p = tim_doan(body, 'Kết quả kiểm tra (Bảng 5.6) cho thấy các rãnh')
    thay(p, 'bằng bê tông M200 liền khối', 'bằng bê tông B15 (M200) liền khối')
    p = tim_doan(body, 'Kết quả tính (Bảng 6.2, Hình 6.1) cho mức đầy lớn nhất')
    thay(p, 'của bê tông M200 trở lên.', 'của bê tông B15 (M200) trở lên.')
    n = 0
    for dau in ('Bảng 3.1', 'Bảng 3.2', 'Bảng 7.1'):
        cap, tbl = tim_bang(body, dau)
        n += thay_trong_bang(tbl, 'Bê tông M250', 'Bê tông B20 (M250)')
        n += thay_trong_bang(tbl, 'BT M250', 'BT B20 (M250)')
        n += thay_trong_bang(tbl, 'BT M200', 'BT B15 (M200)')
    # không còn mác M đứng riêng (trừ đoạn mô tả TKCS và vữa xây M100)
    for el in body.iter(qn('w:p')):
        t = van_ban(el)
        for m in re.finditer(r'M(200|250)', t):
            truoc = t[max(0, m.start() - 1):m.start()]
            assert truoc == '(' or 'M250 (cấp độ bền B20)' in t, t[:120]
    ghi('Bê tông ghi theo cấp độ bền chịu nén B (TCVN 4116:2023 mục 5.1.2), mác M kèm trong ngoặc: B20 (M250) mặt '
        'đập, gờ chắn bánh, đường quản lý; B15 (M200) rãnh – mục 3.2.1, 3.3, 5.1.3, 5.5.2, 6.3, Bảng 3.1, 3.2, 7.1 '
        '(%d ô bảng).' % n)

    p = tim_doan(body, 'Đáy dốc nước: đục bỏ lớp bê tông bề mặt bị mài mòn')
    thay(p, 'trát vữa xi măng – polime cường độ ≥ 40 MPa,',
         'trát vữa xi măng – polime cường độ ≥ 40 MPa, mác chống thấm ≥ W8 (TCVN 4116:2023 mục 5.1.9: vùng có yêu '
         'cầu chống mài mòn, xâm thực dùng vật liệu tối thiểu B25, W8),')
    ghi('Mục 4.x (sửa chữa đáy dốc tràn chính): vữa xi măng – polime ≥ 40 MPa, thêm yêu cầu W8 theo TCVN 4116:2023 '
        'mục 5.1.9.')


# =========================================================================== CHƯƠNG 6 – KHUNG CỐNG
def chuong6(body, kq):
    p = tim_doan(body, 'Việc nâng đỉnh đập từ +78,35 m lên +78,85 m làm áp lực đất đứng trên cống tăng')
    thay(p, 'lớp bảo vệ a = 5 cm (nguyên tắc tính kết cấu BTCT thủy công theo TCVN 4116:2023 [34]).',
         'lớp bê tông bảo vệ ≥ 60 mm ở mặt tiếp xúc trực tiếp với nước theo TCVN 4116:2023 mục 6.3.3 [34], lấy khoảng '
         'cách từ mép đến trọng tâm cốt thép a = 7 cm; bỏ qua hệ số điều kiện làm việc của cốt thép γ_{s2} = 1,1 '
         '(Bảng 13 của tiêu chuẩn) – thiên an toàn.')
    k = kq['khung_cong']['kq']
    cap, tbl = tim_bang(body, 'Bảng 6.4')
    rows = hang(tbl)
    ma = {'Sau SC: Kđ = 1,0; K0 = 0,674': 'sau_SC_Kd1.0_K0x1.0', 'Sau SC: Kđ = 1,0; 0,5·K0': 'sau_SC_Kd1.0_K0x0.5',
          'Sau SC: Kđ = 1,3; 0,5·K0': 'sau_SC_Kd1.3_K0x0.5', 'Trước SC: Kđ = 1,3; 0,5·K0 (so sánh)': 'truoc_SC_Kd1.3_K0x0.5'}
    n = 0
    for tr in rows[1:]:
        tcs = tr.findall(qn('w:tc'))
        ten = van_ban(tcs[0]).strip()
        r = k[ma[ten]]
        assert van_ban(tcs[1]).strip() == so(r['MA']), (ten, van_ban(tcs[1]))
        for c, key in ((5, 'As_nap_giua'), (6, 'As_goc'), (7, 'As_day_giua')):
            dat_o(tcs[c], so(r[key]))
            n += 1
    assert n == 12
    lo = min(min(k[x][y] for y in ('As_nap_giua', 'As_goc', 'As_day_giua')) for x in k if x.startswith('sau'))
    hi = max(max(k[x][y] for y in ('As_nap_giua', 'As_goc', 'As_day_giua')) for x in k if x.startswith('sau'))
    assert int(lo) == 8 and round(hi) == 17, (lo, hi)
    p = tim_doan(body, 'Cốt thép yêu cầu sau sửa chữa khoảng 8 ÷ 16 cm')
    thay(p, '8 ÷ 16', '8 ÷ 17')
    thay(p, '5Ø16 ÷ 5Ø20', '5Ø16 ÷ 6Ø20')
    ghi('Mục 6.5, Bảng 6.4: lớp bê tông bảo vệ ≥ 60 mm (TCVN 4116:2023 mục 6.3.3) → a = 7 cm thay 5 cm; tính lại cốt '
        'thép (bo_sung_giai_tich.py): sau sửa chữa %s ÷ %s cm²/m (5Ø16 ÷ 6Ø20 mỗi mét), vẫn lớn hơn trước sửa chữa '
        'khoảng 4%%.' % (so(lo), so(hi)))


def main(vao, ra, dir_kq):
    kq_qtt = json.load(open(os.path.join(dir_kq, 'kt_dong_chay_toi_thieu.json'), encoding='utf-8'))
    kq_bs = json.load(open(os.path.join(dir_kq, 'bo_sung_giai_tich.json'), encoding='utf-8'))
    d = docx.Document(vao)
    body = d.element.body
    n_dong, n_bm = C.go_muc_luc(body)
    ghi('Gỡ mục lục cũ (%d dòng, %d bookmark) để điền lại sau khi sửa.' % (n_dong, n_bm))
    dau_sach(body)
    chuong2(body, kq_qtt)
    be_tong(body)
    chuong6(body, kq_bs)
    d.save(ra)
    with open(os.path.splitext(ra)[0] + '_nhat_ky.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(C.NHAT_KY) + '\n')
    print('Đã lưu', ra, '–', len(C.NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(*sys.argv[1:4])
