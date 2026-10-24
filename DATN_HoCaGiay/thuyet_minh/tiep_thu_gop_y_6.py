# -*- coding: utf-8 -*-
"""
TIẾP THU GÓP Ý ĐỢT 6 (12/10/2026)

  python3 tiep_thu_gop_y_6.py <vào.docx> <ra.docx> <thư mục tinh_toan>
  python3 chen_muc_luc.py <ra.docx> <bản cuối.docx> <bản cuối.pdf>

Nguồn góp ý và căn cứ:
 (1) Yêu cầu của giảng viên (12/10): điều tiết lũ, thấm hiện trạng và sau nâng cấp – nhận xét hiệu quả (đập, tràn);
     đánh giá biến động lượng mưa từ khi vận hành (2000), điều tiết lũ khi lũ tăng, so sánh với MNLTK thiết kế ban
     đầu → mục 4.6 (Bảng 4.7 ÷ 4.9, Hình 4.2, 4.3), mục 5.6 (Bảng 5.7), Phụ lục 7 (Bảng PL12).
 (2) Báo cáo rà soát 11/10 (A1 ÷ A12, B1 ÷ B10) – đã kiểm chứng từng ý (RA_SOAT_THUYET_MINH.md, đợt 6).
 (3) Toàn văn bốn tiêu chuẩn SV tải lên 12/10: QCVN 04-05:2022 (γn, γlc, γc – Bảng B.1, B.2; điều 2.1.2, 2.3.2.2,
     2.6.1.2.3, 2.6.2.2.1, 2.6.3.3, 2.6.3.5, 2.6.3.10, Bảng 4, Bảng 9), TCVN 8421:2010 (A.3.3, Bảng 9),
     TCVN 9151:2012 (3.1, 5.15, 7.2.3, 7.2.6, 7.2.8, Bảng 3), TCVN 4116:2023 (5.1.9, 6.3.3, 7.4, 7.15, 8.1.4,
     công thức (35), Bảng 11, 13, 21, 9.2).
Tài liệu tham khảo: thêm [31] (Lê Thị Hải Yến và cs., 2017), đánh số lại [31] ÷ [41] → [32] ÷ [42].
Nhật ký: <ra>_nhat_ky.txt.
"""
import json
import os
import sys

import docx
from docx.oxml.ns import qn

import cong_cu_docx as C
from cong_cu_docx import (ghi, tim_doan, tim_bang, thay, thay_tu, van_ban, hang, o, them_hang, dat_o, doan_moi,
                          chen_sau, chen_truoc, dat_noi_dung, them_cuoi, bang_moi, doan_anh, doan_cong_thuc,
                          phan_so, kieu)

DU_PHONG_LUN = 0.13          # m – ≥ 0,5 %H (H = 24,75 ÷ 24,85 m), TCVN 8216:2018 điều 11.4.4
HINH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hinh_bo_sung')


def so(x, n=2):
    s = ('%.*f' % (n, x)).replace('.', ',')
    if s.startswith('-'):
        s = '−' + s[1:]
    return s


def so_nghin(x):
    """12345 -> '12.345'."""
    return '{:,.0f}'.format(x).replace(',', '.')


def thay_trong_bang(tbl, cu, moi):
    n = 0
    for p in tbl.iter(qn('w:p')):
        k = van_ban(p).count(cu)
        if k:
            thay(p, cu, moi, lan=k)
            n += k
    return n


def dong_bang(tbl, r):
    return [van_ban(tc) for tc in hang(tbl)[r].findall(qn('w:tc'))]


def dat_hang(tbl, r, gia_tri):
    for tc, gt in zip(hang(tbl)[r].findall(qn('w:tc')), gia_tri):
        if gt is not None:
            dat_o(tc, gt)


def tim_hang(tbl, dau, cot=0, dung=False):
    """Chỉ số hàng có ô cột `cot` bắt đầu bằng `dau` (dung=True: bằng đúng `dau`)."""
    kq = [i for i, tr in enumerate(hang(tbl))
          if (van_ban(tr.findall(qn('w:tc'))[cot]).strip() == dau if dung
              else van_ban(tr.findall(qn('w:tc'))[cot]).strip().startswith(dau))]
    if len(kq) != 1:
        raise ValueError('Hàng bắt đầu bằng %r: %d' % (dau, len(kq)))
    return kq[0]


def dat_rong_anh(d, ten_goc, rong_cm):
    """Đặt lại bề rộng ảnh (giữ tỷ lệ) cho ảnh có tên gốc ten_goc."""
    import re
    for pic in d.element.body.iter('{%s}pic' % C.PIC):
        ten = pic.find('{%s}nvPicPr/{%s}cNvPr' % (C.PIC, C.PIC)).get('name')
        if re.fullmatch(re.escape(ten_goc) + r'_[0-9a-f]{10}\.jpg', ten):
            inline = pic
            while inline.tag != '{%s}inline' % C.WP:
                inline = inline.getparent()
            ext = inline.find('{%s}extent' % C.WP)
            cx0, cy0 = int(ext.get('cx')), int(ext.get('cy'))
            cx = int(rong_cm * 360000)
            cy = int(round(cy0 * cx / cx0))
            ext.set('cx', str(cx))
            ext.set('cy', str(cy))
            x = pic.find('{%s}spPr/{%s}xfrm/{%s}ext' % (C.PIC, C.A, C.A))
            x.set('cx', str(cx))
            x.set('cy', str(cy))
            return cx, cy
    raise ValueError(ten_goc)


# =========================================================================== CHƯƠNG 1, 2
def chuong1_2(body, KQ):
    ct = KQ['ct']
    # ---- 1.1.3 cấp công trình: thêm tiêu chí chiều cao đập (QCVN 04-05:2022 Bảng 1, điều 2.1.2.3)
    p = tim_doan(body, 'Theo QCVN 04-05:2022/BNNPTNT [6], cấp của công trình đầu mối')
    thay(p, 'tiêu chí diện tích tưới (2.000 ÷ 10.000 ha) cho cấp III. ',
         'tiêu chí diện tích tưới (2.000 ÷ 10.000 ha) cho cấp III; tiêu chí đặc tính kỹ thuật: đập đất cao lớn nhất '
         'khoảng 24,3 m (24,75 m sau nâng cấp [1]) trên nền đất sét pha nửa cứng (nền loại B, chiều cao > 15 ÷ 35 m) '
         'cũng cho cấp II, nên không thuộc trường hợp được xem xét giảm cấp theo điều 2.1.2.3a của quy chuẩn. ')
    them_cuoi(p, ' Cấp công trình thay đổi theo các thế hệ tiêu chuẩn: khi thiết kế (1996) và lập QTVH năm 2005 theo '
                 'TCVN 5060-90, hồ sơ ghi công trình cấp III [1] hoặc cấp IV [3] và đều tính lũ thiết '
                 'kế P = 1%; nếu theo TCXDVN 285:2002 (thang cấp I ÷ V), công trình thuộc cấp III (dung tích 20 ÷ 200 '
                 'triệu m^{3}; đập nền loại B cao 15 ÷ 35 m), ứng với lũ thiết kế 1,5% và lũ kiểm tra 0,5%; từ QCVN '
                 '04-05:2012 (thang cấp mới: đặc biệt, I ÷ IV) đến QCVN 04-05:2022, công trình là cấp II. Các hồ sơ sửa '
                 'chữa (BCNCKT, TKCS, báo cáo KTTV, QTVH 2026) đều xác định công trình cấp II với tần suất lũ 1% / 0,2%, '
                 'phù hợp quy chuẩn hiện hành; đồ án sử dụng thống nhất cấp II.')
    ghi('Mục 1.1.3: bổ sung tiêu chí chiều cao đập (≈ 24,3 ÷ 24,75 m, nền loại B → cấp II, QCVN 04-05:2022 Bảng 1) – '
        'không thuộc trường hợp giảm cấp theo điều 2.1.2.3a; thêm lịch sử phân cấp (TCVN 5060-90: III/IV; TCXDVN '
        '285:2002: III – 1,5%/0,5%; QCVN 04-05:2012/2022: II – 1%/0,2%); các hồ sơ sửa chữa đều xác định cấp II.')

    # ---- 1.6.2 tràn chính: chảy tự do, không có cửa van (GPT A3)
    p = tim_doan(body, 'Tràn chính có 9 khoang × 2,4 m, ngưỡng +74,70 m')
    thay(p, 'Tràn chính có 9 khoang × 2,4 m, ngưỡng +74,70 m,',
         'Tràn chính có 9 khoang × 2,4 m (mố trụ dày 0,6 m), ngưỡng +74,70 m: theo thiết kế ban đầu tràn rộng 25 m, '
         '3 khoang chảy tự do [3], [7]; đơn vị quản lý đã chia thành 9 khoang và lắp cửa van để tích thêm nước cuối mùa '
         'lũ phục vụ tưới, đầu mùa lũ toàn bộ cửa van được mở hoàn toàn để xả lũ [1], nên khi tháo lũ tràn chính làm '
         'việc như đập tràn chảy tự do với điều kiện 9 cửa mở hết, không bị cản trở;')
    ghi('Mục 1.6.2: tràn chính có 9 cửa van do đơn vị quản lý lắp để tích nước cuối mùa lũ, mở hoàn toàn đầu mùa lũ '
        '(QTVH mục 6.x); khi tháo lũ chảy tự do với điều kiện 9 cửa mở hết (tiếp thu góp ý A3).')

    # ---- Bảng 2.1: ký hiệu hệ số theo QCVN 04-05:2022, tần suất vượt kiểm tra, γc, TCVN 4116:2023
    cap, tbl = tim_bang(body, 'Bảng 2.1')
    r = tim_hang(tbl, '1', 0, dung=True)
    dat_hang(tbl, r, [None, 'Tần suất lũ thiết kế / lũ kiểm tra / lũ vượt kiểm tra',
                      'P = 1% / P = 0,2% / P = 0,1%',
                      'QCVN 04-05:2022 [6], Bảng 4; lũ vượt kiểm tra chỉ xét khi có luận cứ (điều 2.6.3.5)'])
    r = tim_hang(tbl, '6', 0, dung=True)
    dat_hang(tbl, r, [None, 'Hệ số bảo đảm γ_{n} (trước đây ký hiệu K_{n}) – cấp II: kết cấu bê tông, BTCT / mái đất đắp',
                      '1,15 / 1,30', 'QCVN 04-05:2022, Bảng B.1; TCVN 4116:2023, mục 8.1.4'])
    r = tim_hang(tbl, '7', 0, dung=True)
    dat_hang(tbl, r, [None, 'Hệ số tổ hợp tải trọng γ_{lc} (trước đây n_{c}): cơ bản / đặc biệt / thi công – sửa chữa',
                      '1,00 / 0,90 / 0,95', 'QCVN 04-05:2022, điều 2.5.3'])
    r7 = them_hang(tbl, r, ['7a', 'Hệ số điều kiện làm việc γ_{c}: kết cấu BT trên nền đá, mặt trượt qua tiếp xúc '
                                  'bê tông – đá / mái dốc', '0,95 / 1,00', 'QCVN 04-05:2022, Bảng B.2'])
    r = tim_hang(tbl, '11', 0, dung=True)
    tcs = hang(tbl)[r].findall(qn('w:tc'))
    dat_o(tcs[1], 'Bê tông kết cấu thủy công: chỉ tiêu thiết kế; vùng chịu mài mòn, xâm thực; lớp bê tông bảo vệ cốt '
                  'thép ở mặt tiếp xúc trực tiếp với nước; cường độ tính toán cốt thép R_{s}')
    dat_o(tcs[2], 'Cấp độ bền chịu nén B, mác chống thấm W (mác M ghi kèm để đối chiếu); ≥ B25, W8; ≥ 60 mm và ≥ 2d; '
                  'CB300-V 260 MPa, CB400-V 350 MPa')
    dat_o(tcs[3], 'TCVN 4116:2023, mục 5.1.2, 5.1.9, 6.3.3; Bảng 11')
    ghi('Bảng 2.1: dòng 1 thêm lũ vượt kiểm tra P = 0,1% (QCVN 04-05:2022 Bảng 4, điều 2.6.3.5); dòng 6, 7 đổi ký hiệu '
        'theo QCVN 04-05:2022 (γn, γlc; thêm γlc = 0,95 thi công – sửa chữa, γn = 1,30 cho mái đất); thêm dòng 7a γc '
        '(Bảng B.2: 0,95 khi mặt trượt qua tiếp xúc bê tông – đá); dòng 11 thêm mục 5.1.9 (≥ B25, W8) và Bảng 11 (Rs).')

    # ---- 2.2.1 MNC: thời gian bồi lắng 75 năm (QCVN 04-05:2022 Bảng 9)
    p = tim_doan(body, 'Mực nước chết của hồ MNC = +67,50 m được giữ như thiết kế')
    thay(p, 'tổng lượng bùn cát đến hồ trong 70 năm khoảng 0,48 triệu m3 [1],',
         'QCVN 04-05:2022 (Bảng 9) quy định ngưỡng cửa lấy nước của hồ cấp II không bị bồi lấp trong ít nhất 75 năm; '
         'tổng lượng bùn cát đến hồ trong 75 năm khoảng 0,48·75/70 ≈ 0,51 triệu m^{3} (QTVH tính cho 70 năm: '
         '0,48 triệu m^{3} [1]),')
    ghi('Mục 2.2.1: thời gian tính bồi lắng 75 năm cho hồ cấp II (QCVN 04-05:2022 Bảng 9; QTVH dùng 70 năm) – '
        '0,51 triệu m³ ≪ 3,52 triệu m³ dưới ngưỡng cống, kết luận không đổi.')

    # ---- 2.2.3 mức bảo đảm theo từng mục tiêu (GPT A9, TCVN 10778:2024 mục 8.2.2.2)
    cn = KQ['cap_nuoc']['TH1A|khong_Qtt']
    p = tim_doan(body, 'Lượng nước đến bình quân năm của chuỗi: 220,3 triệu m3.')
    p2 = doan_moi(
        'Hồ có hai mục tiêu cấp nước với mức bảo đảm yêu cầu khác nhau (tưới 85%, sinh hoạt 90%). Theo TCVN '
        '10778:2024 mục 8.2.2.2 [15], khi tính chung có thể dùng mức bảo đảm bình quân theo trọng số lượng nước của '
        'từng mục tiêu: với TH1A (tưới 103,42 và sinh hoạt 3,22 triệu m^{3}/năm) mức bảo đảm chung cần đạt là '
        + so(cn['P_chung'], 2) + '%, tức ít nhất 39 năm đủ nước – TH1A đạt 39/44 năm (86,7%). Kết quả này không tự '
        'chứng minh mức 90% cho sinh hoạt, nên đồ án kiểm tra riêng theo quy tắc ưu tiên của QTVH: khi thiếu nước chỉ '
        'cắt giảm cấp nước tưới, giữ nguyên cấp nước sinh hoạt. Với quy tắc này sinh hoạt được cấp đủ trong cả 44 năm '
        '(P = 44/45 = ' + so(cn['P_sinh_hoat'], 1) + '% ≥ 90%); tháng thiếu nước nặng nhất vẫn cấp được '
        + so(cn['cap_nho_nhat_thang_thieu'], 1) + ' triệu m^{3}, lớn hơn nhiều nhu cầu sinh hoạt một tháng (≈ 0,27 '
        'triệu m^{3}); mức cắt giảm cấp nước tưới lớn nhất trong một tháng khoảng '
        + so(100 * cn['cat_giam_nn_lon_nhat'], 0) + '% (chương trình kt_cap_nuoc_bo_sung.py). Quy tắc ưu tiên này '
        'cần được ghi rõ trong quy trình vận hành.', p)
    chen_sau(p, p2)
    ghi('Mục 2.2.3 (sau Bảng 2.2): mức bảo đảm theo từng mục tiêu – TCVN 10778:2024 mục 8.2.2.2: P chung yêu cầu '
        '%s%% (TH1A đạt 86,7%%); sinh hoạt ưu tiên đạt %s%% ≥ 90%%; cắt giảm tưới lớn nhất %s%% (góp ý A9).'
        % (so(cn['P_chung'], 2), so(cn['P_sinh_hoat'], 1), so(100 * cn['cat_giam_nn_lon_nhat'], 0)))

    # ---- 2.3.1 kẹt van: phân loại theo QCVN 04-05:2022 điều 2.6.3.10 (GPT B1)
    p = tim_doan(body, 'Quy tắc vận hành theo QTVH: đầu trận lũ mực nước hồ ở MNDBT')
    thay(p, '– đây là tình huống kiểm tra sự cố thiết bị, không thuộc tổ hợp thiết kế; kết quả dùng để đánh giá mức dự '
            'trữ an toàn và yêu cầu quản lý vận hành (phân loại ở Bảng 4.2).',
         '– theo CHÚ THÍCH điều 2.6.3.10 của QCVN 04-05:2022 [6], lũ thiết kế khi một cửa xả bị kẹt được xếp vào tổ '
         'hợp đặc biệt; khi tính với lũ kiểm tra, quy chuẩn không xét khả năng cửa xả bị kẹt (khoản b) và xếp lũ vượt '
         'kiểm tra vào trường hợp xét thêm (khoản d, điều 2.6.3.5). Các trường hợp lũ kiểm tra hoặc lũ vượt kiểm tra '
         'kèm kẹt van vì vậy được giữ làm thông tin đánh giá rủi ro và yêu cầu quản lý vận hành (phân loại ở Bảng 4.2).')
    ghi('Mục 2.3.1: phân loại trường hợp kẹt van theo QCVN 04-05:2022 điều 2.6.3.10 (a, b, d, CHÚ THÍCH) và 2.6.3.5 '
        '(góp ý B1).')

    # ---- 2.3.2 độ nhạy m (GPT A4)
    mm = KQ['m']
    p = tim_doan(body, 'Đối chiếu với hồ sơ: BCNCKT 2022 tính được MNLTK = +77,43 m')
    them_cuoi(p, ' QTVH lấy m = 0,40 (TCVN 9147:2012 mục 7.6) và ε = Σb/(Σb + Σd) = 21,6/(21,6 + 8 × 0,6) = 0,82 – '
                 'tỷ lệ bề rộng thông thủy trên tổng bề rộng, không phải hệ số co hẹp bên theo hình dạng mố; lưu lượng '
                 'trong bảng quan hệ Q ~ Z của QTVH ứng đúng 0,40 × bề rộng thông thủy 21,6 m (Q/(B·√(2g)·H^{1,5}) = '
                 '0,400 ở mọi cột nước; tràn sự cố 0,370), tức chưa kể co hẹp bên do 8 mố trụ. Đồ án dùng m = 0,45 và '
                 'hệ số co hẹp bên theo hình dạng mố (ε = 0,85 ÷ 0,88 ở cột nước lũ, ε·m = 0,38 ÷ 0,40): dùng hệ số 0,40 '
                 '/ 0,37 như QTVH cho MNLKT +78,52 m, gần như trùng kết quả của đồ án (+78,54 m). Tuy vậy m chưa được '
                 'kiểm định bằng thí nghiệm mô hình hoặc đo lũ nên đồ án kiểm tra cả phía bất lợi: nếu lấy m = 0,40 '
                 'của QTVH và nhân thêm hệ số co hẹp bên do mố trụ thì MNLKT = +' + so(mm['0.40'][1]) + ' m và điều kiện MNLKT + 0,3 m đòi hỏi đỉnh đập +'
                 + so(mm['0.40'][1] + 0.3) + ' m; đỉnh đập +78,85 m còn thỏa khi m ≥ ' + so(mm['m_gh'], 3)
                 + ' (Bảng PL7). Vì vậy cần nạo vét kênh dẫn để giữ điều kiện chảy tới ngưỡng như giả thiết và xác định '
                 'm bằng đo đạc quan hệ Q ~ Z khi có lũ lớn (kiến nghị).')
    ghi('Mục 2.3.2: giải thích εm = 0,40/0,37 của QTVH là tích ε·m; thử m = 0,40: MNLKT +%s m, đỉnh yêu cầu +%s m; '
        'ngưỡng m ≥ %s cho đỉnh +78,85 m (góp ý A4).' % (so(mm['0.40'][1]), so(mm['0.40'][1] + 0.3), so(mm['m_gh'], 3)))

    # ---- 2.4 sóng: A.3.3 (Vw = kfl·kl·Vl), Bảng 9 (kα) của TCVN 8421:2010
    p = tim_doan(body, 'Trong đó: V – vận tốc gió (m/s); D – đà gió (m); H – chiều sâu nước trước đập')
    thay(p, 'Trong đó: V – vận tốc gió (m/s);',
         'Trong đó: V – vận tốc gió tính toán ở độ cao 10 m trên mặt nước, V = V_{w} = k_{fl}·k_{l}·V_{l} (TCVN '
         '8421:2010, Phụ lục A, A.3.3), với V_{l} – vận tốc gió quan trắc ở trạm Phan Thiết (Bảng 1.4), k_{l} = 1,10 '
         '÷ 1,09 – hệ số quy đổi về mặt nước cho địa hình dạng A (vùng ven biển trống trải, Bảng A3), k_{fl} – hệ số '
         'quy đổi số liệu phong tiêu, lấy bằng 1 (thiên an toàn vì chưa rõ dụng cụ đo);')
    thay(p, 'kα – hệ số kể đến góc giữa hướng sóng và pháp tuyến mái (gió hướng Bắc lệch pháp tuyến mái 6,2° nên kα ≈ 1).',
         'k_{α} – hệ số giảm sóng leo theo góc giữa front sóng và pháp tuyến mái, lấy theo Bảng 9 của tiêu chuẩn '
         '(0° – 1,00; 20° – 0,96; 40° – 0,87; 60° – 0,76); gió hướng Bắc lệch pháp tuyến mái 6,2° nên lấy k_{α} = 1.')
    ghi('Mục 2.4.1: vận tốc gió tính toán Vw = kfl·kl·Vl (TCVN 8421:2010 A.3.3, kl = 1,10 địa hình A, kfl = 1) và kα '
        'theo Bảng 9 – trước đây chưa quy đổi gió (thiên nhỏ ≈ 10%) và nội suy kα sai bảng (thiên lớn 5 ÷ 11%).')

    # Bảng 2.4 – tính lại
    r1, r2 = ct['r1'], ct['r2']
    cap, tbl = tim_bang(body, 'Bảng 2.4')
    gt = {
        'Vận tốc gió V': ['Vận tốc gió V_{l} / V_{w} (m/s) – hướng Bắc', so(r1['V_tram'], 1) + ' / ' + so(r1['V'], 2)
                          + ' (P = 4%)', so(r2['V_tram'], 1) + ' / ' + so(r2['V'], 2) + ' (P = 50%)', '–'],
        'gD/V': ['gD/V_{w}^{2}', so(r1['gD'], 1), so(r2['gD'], 1), '–'],
        'Chiều cao sóng trung bình': [None, so(r1['h'], 3), so(r2['h'], 3), '–'],
        'Chu kỳ': [None, so(r1['tau'], 2) + ' / ' + so(r1['lam'], 2), so(r2['tau'], 2) + ' / ' + so(r2['lam'], 2), '–'],
        'K1%': [None, so(r1['K1'], 2) + ' ; ' + so(r1['h1'], 3), so(r2['K1'], 2) + ' ; ' + so(r2['h1'], 3), '–'],
        'kr ;': [None, ' ; '.join(so(r1[k], 2) for k in ('kr', 'kp', 'ksp', 'krun')),
                 ' ; '.join(so(r2[k], 2) for k in ('kr', 'kp', 'ksp', 'krun')), '–'],
        'Sóng leo': [None, so(r1['hsl'], 3), so(r2['hsl'], 3), '–'],
        'Độ dềnh': [None, so(r1['dh'], 3), so(r2['dh'], 3), '–'],
        'Cao trình đỉnh yêu cầu': [None, '+' + so(ct['Z1']), '+' + so(ct['Z2']), '+' + so(ct['Z3'])],
    }
    for k, v in gt.items():
        dat_hang(tbl, tim_hang(tbl, k), v)
    ghi('Bảng 2.4 tính lại (gió hướng Bắc): Vw = %s / %s m/s; hsl = %s / %s m; Z1 = +%s, Z2 = +%s, Z3 = +%s m.'
        % (so(r1['V'], 2), so(r2['V'], 2), so(r1['hsl'], 3), so(r2['hsl'], 3), so(ct['Z1']), so(ct['Z2']),
           so(ct['Z3'])))

    # Bảng 2.5
    cap, tbl = tim_bang(body, 'Bảng 2.5')
    for dau, k in (('Hồ sơ BCNCKT', 'bcnckt'), ('Đồ án – lũ BCNCKT (+77,63', 'da'), ('Đồ án – lũ BCNCKT, gió Đông Bắc', 'ne'),
                   ('Đồ án – lũ QTVH', 'qtvh')):
        z = KQ['bang25'][k]
        dat_hang(tbl, tim_hang(tbl, dau), [None, so(z[0]), so(z[1]), so(z[2]), so(max(z)), None])
    ne, nw = KQ['huong']['NE'], KQ['huong']['NW']
    p = tim_doan(body, 'Kết quả tính chi tiết cho trong Bảng 2.4 và tổng hợp trong Bảng 2.5.')
    thay(p, 'hướng Đông Bắc (góc tới 38,8°, V4% = 18,6 m/s, V50% = 11,9 m/s) cho Z1 = +77,15 m, Z2 = +78,97 m; '
            'hướng Tây Bắc (51,2°) cho Z1 = +76,66 m, Z2 = +78,62 m;',
         'hướng Đông Bắc (góc tới 38,8°, k_{α} = 0,88; V_{4%} = 18,6 m/s, V_{50%} = 11,9 m/s, tương ứng V_{w} = '
         + so(ne['Vw4'], 1) + ' và ' + so(ne['Vw50'], 1) + ' m/s) cho Z_{1} = +' + so(ne['Z1']) + ' m, Z_{2} = +'
         + so(ne['Z2']) + ' m; hướng Tây Bắc (51,2°, k_{α} = 0,81) cho Z_{1} = +' + so(nw['Z1']) + ' m, Z_{2} = +'
         + so(nw['Z2']) + ' m;')
    thay(p, 'Như vậy trị số lớn nhất là Z2 = +78,97 m (gió Đông Bắc).',
         'Như vậy trị số lớn nhất là Z_{2} = +' + so(ne['Z2']) + ' m (gió Đông Bắc).')
    thay(p, 'max(Z1, Z2, Z3) = +78,97 m < +79,40 m – tường chắn sóng hiện trạng đạt yêu cầu (dự trữ 0,43 m);',
         'max(Z_{1}, Z_{2}, Z_{3}) = +' + so(ne['Z2']) + ' m < +79,40 m – tường chắn sóng hiện trạng đạt yêu cầu (dự '
         'trữ ' + so(79.40 - ne['Z2']) + ' m);')
    ghi('Bảng 2.5, mục 2.4.2: hướng Đông Bắc Z1 = +%s, Z2 = +%s m (trước +77,15; +78,97); Tây Bắc Z2 = +%s m; dự trữ '
        'đến đỉnh tường còn %s m. Cao trình đỉnh đập vẫn do MNLKT + 0,3 = +78,84 m quyết định.'
        % (so(ne['Z1']), so(ne['Z2']), so(nw['Z2']), so(79.40 - ne['Z2'])))

    # ---- 2.5 thấm hiện trạng: giá trị lớn nhất, vị trí (GPT A6), bất đẳng hướng (B2)
    gr = KQ['gradien']
    p = tim_doan(body, 'Kết quả tính thấm hiện trạng cho trong Bảng 2.6;')
    thay(p, 'gradien tại vùng tiếp xúc với ống khói (trung bình 0,80 ÷ 0,88, phân vị 95% ≤ 1,19) nhỏ hơn giá trị cho phép '
            '2,0;',
         'gradien tại vùng tiếp xúc với ống khói (trung bình 0,80 ÷ 0,88, phân vị 95% ≤ 1,19; giá trị lớn nhất '
         + so(gr['HT_MNLKT']['Jtiep_max']) + ' tại chân ống khói, x ≈ ' + so(gr['HT_MNLKT']['x_tiep'], 1) + ' m, '
         'z ≈ +' + so(gr['HT_MNLKT']['z_tiep'], 1) + ' m) nhỏ hơn giá trị cho phép 2,0 lấy theo hồ sơ thiết kế tầng '
         'lọc [3];')
    thay(p, 'gradien cửa ra rất nhỏ (≈ 0,02 ≪ 0,85).',
         'gradien cửa ra rất nhỏ (trung bình ≈ 0,02, lớn nhất ' + so(gr['HT_MNLKT']['Jra_max'], 3) + ' tại chân lăng '
         'trụ, ≪ 0,85 – TCVN 8216:2018 Bảng 6). Giá trị lớn nhất được lấy trên toàn bộ phần tử của vùng kiểm tra '
         '(không loại điểm kỳ dị) để bổ sung cho giá trị trung bình và phân vị 95%. Khi xét bất đẳng hướng thấm '
         'k_{h}/k_{v} = 5 (TCVN 8216:2018 mục 11.2.2) lưu lượng thấm tăng 46 ÷ 61%, đường bão hòa tại tim dâng '
         '0,6 ÷ 2,5 m nhưng gradien vẫn nhỏ hơn cho phép (Bảng PL10).')
    ghi('Mục 2.5.2: báo cáo thêm gradien lớn nhất và vị trí (tiếp xúc ống khói %s; cửa ra %s), nguồn của giới hạn 0,85 '
        '(TCVN 8216:2018 Bảng 6) và 2,0 (hồ sơ tầng lọc); ảnh hưởng bất đẳng hướng kh/kv = 5 (Bảng PL10) – góp ý A6, B2.'
        % (so(gr['HT_MNLKT']['Jtiep_max']), so(gr['HT_MNLKT']['Jra_max'], 3)))

    # ---- 2.5.3 tìm cung trượt (GPT B5)
    p = tim_doan(body, 'Tâm trượt nguy hiểm được tìm trên lưới tâm 28 × 26 điểm')
    thay(p, 'và bằng lưới tâm mịn gấp đôi (mục 5.3, Phụ lục 6).',
         'và bằng cách tìm lại trên miền mở rộng (đáy cung tới cao trình +36 m), lưới lồng nhau chứa toàn bộ tâm và '
         'bán kính cũ cộng với tối ưu hóa Nelder – Mead từ 10 điểm xuất phát (mục 5.3, Phụ lục 6, Bảng PL11).')
    ghi('Mục 2.5.3: mô tả kiểm tra tìm cung trượt bằng miền mở rộng, lưới lồng nhau và Nelder – Mead (góp ý B5).')

    # ---- Bảng 2.8
    cap, tbl = tim_bang(body, 'Bảng 2.8')
    r = tim_hang(tbl, 'Đỉnh tường chắn sóng so với sóng leo')
    dat_hang(tbl, r, [None, 'Z_{max} = +' + so(ne['Z2']) + ' m', None, None])
    r = tim_hang(tbl, 'Cấp nước tưới, TH1A')
    them_hang(tbl, r, ['Cấp nước sinh hoạt, TH1A (ưu tiên, cắt giảm tưới trước)', 'P = ' + so(cn['P_sinh_hoat'], 1) + '%',
                       '≥ 90%', 'Đạt'])
    r = tim_hang(tbl, 'Mực nước lũ kiểm tra so với đỉnh tường')
    them_hang(tbl, r - 1, ['MNLTK hiện nay so với MNLTK thiết kế ban đầu (TK 1996)', '+77,63 m', '+77,24 m',
                           'Cao hơn 0,39 m (mục 4.6)'])
    ghi('Bảng 2.8: Zmax sóng leo +%s m; thêm dòng cấp nước sinh hoạt (%s%%) và dòng MNLTK hiện nay +77,63 so với '
        'MNLTK thiết kế ban đầu +77,24 m.' % (so(ne['Z2']), so(cn['P_sinh_hoat'], 1)))


# =========================================================================== CHƯƠNG 3
def chuong3(d, body, KQ):
    ne = KQ['huong']['NE']
    p = tim_doan(body, 'Hồ sơ TKCS [2] đề xuất giữ nguyên tường chắn sóng')
    thay(p, '(Zmax = +78,97 m)', '(Z_{max} = +' + so(ne['Z2']) + ' m, gió Đông Bắc)')
    p = tim_doan(body, 'Phương án Đ3 – như Đ1 nhưng đắp bù đất đến +78,60 m')
    thay(p, 'nên +78,85 m được quy định là cao trình tối thiểu sau lún tại mép thượng lưu, khi thi công lấy dự phòng lún '
            '2 cm (mục 5.1.1);',
         'nên +78,85 m được quy định là cao trình thiết kế (sau lún) tại mép thượng lưu, khi thi công cộng thêm dự phòng '
         'lún 0,13 m (≥ 0,5%H theo TCVN 8216:2018 điều 11.4.4 – mục 5.1.1);')
    thay(p, 'khối lượng đất đắp tăng khoảng 986 × 5 × 0,1 ≈ 490 m3, chi phí tăng không đáng kể;',
         'khối lượng đất đắp tăng khoảng 986 × 5 × 0,1 ≈ 490 m^{3} so với TKCS (chưa kể dự phòng lún), chi phí tăng '
         'không đáng kể;')
    p = tim_doan(body, 'Đồ án chọn phương án Đ3 vì bảo đảm an toàn với đường quá trình lũ bất lợi hơn')
    thay(p, 'Kết cấu mặt đập (Hình 3.1): lớp đất đắp bù đầm chặt K ≥ 0,97; lớp cấp phối đá dăm đệm dày 5 cm; lớp nilon '
            'tái sinh; bê tông B20 (M250) đá 1×2 dày 20 cm.',
         'Kết cấu mặt đập (Hình 3.1): lớp đất đắp bù đầm chặt K ≥ 0,97; lớp cấp phối đá dăm đệm dày 5 cm; lớp nilon '
         'tái sinh; bê tông B20 (M250) đá 1×2 dày 20 cm. Bộ phận chống thấm ở đỉnh đập là khối đất đắp bù (sét pha, '
         'k ≈ 3·10^{-6} cm/s) liền với khối chống thấm thân đập (lớp ĐĐ-I) và đầm sát lưng tường chắn sóng; đỉnh khối '
         'đất +78,60 m (sau lún) cao hơn MNLTK 0,97 m (≥ 0,40 m theo TCVN 8216:2018 Bảng 5) và cao hơn MNDBT có kể '
         'sóng leo, nước dềnh (+' + so(KQ['Z_sl_DBT']) + ' m). Khe giữa tường và tấm bê tông được chèn ma tít, các khe '
         'giữa những đoạn tường được kiểm tra và trám lại – đây là điều kiện "tường chắn sóng liên kết với bộ phận '
         'chống thấm" để áp dụng điều 10.1.5. Tấm bê tông là lớp bảo vệ đỉnh đập, không tính là bộ phận chống thấm. '
         'Nước hồ chỉ tràn qua đỉnh đập khi vượt cao trình mép thượng lưu – điểm cao nhất trên đường chảy (mặt đập '
         'dốc 2% về hạ lưu), nên +78,85 m tại mép thượng lưu là cao trình khống chế; mép hạ lưu tấm (+78,76 m) thấp '
         'hơn 0,09 m do dốc ngang.')
    ghi('Mục 3.2.1: dự phòng lún 0,13 m (≥ 0,5%%H, TCVN 8216:2018 điều 11.4.4); làm rõ bộ phận chống thấm ở đỉnh đập '
        '(khối đất +78,60 ≥ MNLTK + 0,40 – Bảng 5; ≥ MNDBT + sóng leo +%s), mối nối tường – tấm, cao trình khống chế tại '
        'mép thượng lưu (góp ý A1, A2).' % so(KQ['Z_sl_DBT']))
    # Hình 3.1, 3.2 – thay ảnh
    C.thay_anh(d, 'image5_crop', os.path.join(HINH, 'hinh_3_1_mat_dap.jpg'))
    dat_rong_anh(d, 'image5_crop', 16.0)
    p = tim_doan(body, 'Hình 3.1 Cấu tạo sửa chữa mặt đập', style='Chuthichhinh')
    thay(p, 'Cấu tạo sửa chữa mặt đập (nguồn: TKCS [2])',
         'Cấu tạo đỉnh đập sau sửa chữa, bộ phận chống thấm và cao trình khống chế (lập trên cơ sở TKCS [2]; '
         'tỷ lệ đứng phóng to 4 lần)')
    C.thay_anh(d, 'image8_crop', os.path.join(HINH, 'hinh_3_2_ranh.jpg'))
    p = tim_doan(body, 'Hình 3.2 Mặt bằng bố trí hệ thống rãnh thoát nước mái hạ lưu', style='Chuthichhinh')
    thay(p, '(nguồn: TKCS [2])', '(nguồn: TKCS [2]; ô 4 × 4 m đặt xoay 45°, đường chảy dài nhất theo đường chéo '
                                  '≈ 5,66 m)')
    cap, tbl = tim_bang(body, 'Bảng 3.1')
    assert thay_trong_bang(tbl, 'xuống ≤ 4 m,', 'xuống ≤ 5,7 m (đường chéo ô 4 × 4 m),') == 1
    ghi('Hình 3.1 vẽ lại (mặt cắt đỉnh đập: bộ phận chống thấm, mối nối tường – tấm, cao trình thiết kế sau lún và '
        'nghiệm thu); Hình 3.2 trích lại rõ hơn, ghi chú ô xoay 45° (góp ý A1, B8).')


# =========================================================================== CHƯƠNG 4
def chuong4_on_dinh(d, body, KQ):
    ne = KQ['huong']['NE']
    # ---- Bảng 4.2 – phân loại theo QCVN 04-05:2022 (B1)
    cap, tbl = tim_bang(body, 'Bảng 4.2')
    r = tim_hang(tbl, 'Lũ thiết kế P = 1%')
    dat_hang(tbl, r, [None, None, None, None, 'Tổ hợp cơ bản – thiết kế (QCVN 04-05:2022, điều 2.6.3.10a)',
                      'Z_{2} = +' + so(ne['Z2']) + ' ≤ đỉnh tường: đạt'])
    r = tim_hang(tbl, 'Lũ kiểm tra P = 0,2%')
    dat_hang(tbl, r, [None, None, None, None, 'Tổ hợp đặc biệt – kiểm tra (điều 2.6.3.10b)', None])
    r = tim_hang(tbl, 'P = 1% + kẹt van')
    dat_hang(tbl, r, [None, None, None, None, 'Tổ hợp đặc biệt (CHÚ THÍCH điều 2.6.3.10)',
                      'Đỉnh đập ≥ Z_{max} + 0,3 m: đạt (dư 0,34 m)'])
    them_hang(tbl, r, ['Lũ vượt kiểm tra P = 0,1% (QTVH), vận hành đúng QTVH', '+78,83', 'thấp hơn 0,02 m',
                       'thấp hơn 0,57 m', 'Xét thêm (Bảng 4; điều 2.6.3.5, 2.6.3.10d) – không bắt buộc với hồ cấp II',
                       'Không tràn đỉnh nhưng gần như hết dự trữ – xem mục 4.6'])
    r = tim_hang(tbl, 'P = 0,2% + kẹt van')
    dat_hang(tbl, r, [None, None, None, None, 'Ngoài các trường hợp bắt buộc (điều 2.6.3.10b không xét kẹt van khi tính '
                      'lũ kiểm tra) – đánh giá rủi ro', None])
    r = tim_hang(tbl, 'P = 0,1% + kẹt van')
    dat_hang(tbl, r, [None, None, None, None, 'Lũ vượt kiểm tra kèm sự cố thiết bị – kịch bản khẩn cấp', None])
    p = tim_doan(body, 'Các tình huống lũ được phân loại trong Bảng 4.2.')
    thay(p, 'Các tình huống lũ được phân loại trong Bảng 4.2.',
         'Các tình huống lũ được phân loại trong Bảng 4.2 theo điều 2.6.3.10 và 2.6.3.5 của QCVN 04-05:2022 [6]: lũ '
         'thiết kế với các cửa xả mở hoàn toàn là tổ hợp cơ bản; lũ kiểm tra (mọi công trình xả kể cả công trình xả dự '
         'phòng mở hoàn toàn, không xét kẹt van) và lũ thiết kế khi một cửa xả bị kẹt là tổ hợp đặc biệt; lũ vượt kiểm '
         'tra (P = 0,1% với công trình cấp II, Bảng 4 của quy chuẩn) chỉ bắt buộc xét với hồ từ cấp I trở lên, hồ cấp '
         'II xét khi có luận cứ và được cấp có thẩm quyền chấp thuận.')
    ghi('Mục 4.1, Bảng 4.2: phân loại tình huống lũ theo QCVN 04-05:2022 điều 2.6.3.10 (a, b, d, CHÚ THÍCH), 2.6.3.5; '
        'thêm dòng lũ vượt kiểm tra P = 0,1%% vận hành đúng (+78,83 m); Z2 = +%s (góp ý B1).' % so(ne['Z2']))

    # ---- 4.5.1 công thức (4-5): [K] = γn·γlc/γc (QCVN 04-05:2022 điều 2.5.3, Bảng B.1, B.2)
    p = tim_doan(body, 'Điều kiện ổn định trượt phẳng theo QCVN 04-05:2022 [6] và TCVN 4253:2012 [28]:')
    thay(p, 'theo QCVN 04-05:2022 [6] và TCVN 4253:2012 [28]:',
         'theo QCVN 04-05:2022 [6] (điều 2.5.3, công thức (1): γ_{lc}·F ≤ γ_{c}·R/γ_{n}) và TCVN 4253:2012 [28]:')
    ct = p.getnext()
    assert ct.find('.//{%s}oMath' % C.M) is not None
    ct_moi = doan_cong_thuc(ct, ['K = ', phan_so('f·ΣV + c·B', 'ΣH'), ' ≥ [K] = ',
                                 phan_so('γ_{n}·γ_{lc}', 'γ_{c}')], '4-5')
    ct.addprevious(ct_moi)
    body.remove(ct)
    p = tim_doan(body, 'Trong đó: ΣV, ΣH – tổng lực đứng (đã trừ đẩy ngược)')
    thay(p, 'Kn = 1,15; nc = 1,0 (cơ bản) hoặc 0,9 (đặc biệt); m = 1,0 → [K] = 1,15 (MNLTK) và 1,035 (MNLKT).',
         'γ_{n} = 1,15 – hệ số bảo đảm của công trình bê tông cấp II (Bảng B.1); γ_{lc} = 1,00 (tổ hợp cơ bản) hoặc '
         '0,90 (đặc biệt); γ_{c} = 0,95 – hệ số điều kiện làm việc khi mặt trượt đi qua mặt tiếp xúc bê tông – đá '
         '(Bảng B.2; lớp II có cường độ kháng nén bão hòa 10,7 ÷ 29,3 MPa > 5 MPa nên được xếp là nền đá [3]) → '
         '[K] = 1,21 (MNDBT, MNLTK) và 1,09 (MNLKT).')
    cap, tbl = tim_bang(body, 'Bảng 4.5')
    n1 = thay_trong_bang(tbl, '1,150', '1,211')
    n2 = thay_trong_bang(tbl, '1,035', '1,089')
    assert n1 == 4 and n2 == 2, (n1, n2)
    neo = KQ['neo']
    c_tk, c_kt = neo['MNLTK']['c_can'], neo['MNLKT']['c_can']
    p = tim_doan(body, 'Nhận xét (Bảng 4.5): khi bỏ qua hoàn toàn lực dính')
    thay(p, 'chỉ cần lực dính trên mặt tiếp xúc c ≥ 6,4 kPa (MNLTK) và c ≥ 9,1 kPa (MNLKT) là thỏa mãn',
         'chỉ cần lực dính trên mặt tiếp xúc c ≥ ' + so(c_tk, 1) + ' kPa (MNLTK) và c ≥ ' + so(c_kt, 1)
         + ' kPa (MNLKT) là thỏa mãn')
    thay(p, 'chỉ bảo đảm ổn định trượt khi có lực dính c ≥ 9,1 kPa', 'chỉ bảo đảm ổn định trượt khi có lực dính c ≥ '
         + so(c_kt, 1) + ' kPa')
    ghi('Mục 4.5: [K] = γn·γlc/γc với γc = 0,95 (QCVN 04-05:2022 Bảng B.2 – mặt trượt qua tiếp xúc bê tông – đá) → '
        '[K] = 1,21 / 1,09 thay cho 1,15 / 1,035 (trước đây lấy m = 1,0); công thức (4-5) viết lại; Bảng 4.5; lực dính '
        'cần c ≥ %s / %s kPa (trước 6,4 / 9,1).' % (so(c_tk, 1), so(c_kt, 1)))

    # ---- neo (Bảng 4.6, A10, TCVN 4116:2023)
    p = tim_doan(body, 'với As = 4,91 cm2 (D25), Rs = 350 MPa')
    thay(p, 'Rs = 350 MPa – cường độ chịu kéo tính toán của thép CB400-V',
         'R_{s} = 350 MPa – cường độ chịu kéo tính toán của thép CB400-V (TCVN 4116:2023, Bảng 11)')
    cap, tbl = tim_bang(body, 'Bảng 4.6')
    for dau, k in (('MNLTK', 'MNLTK'), ('MNLKT', 'MNLKT')):
        v = neo[k]
        dat_hang(tbl, tim_hang(tbl, dau), [None, None, None, None, so(v['R_can'], 1), so(v['N_can'], 1), None,
                                           so(v['K_cp'], 3), None])
    kb = KQ['kc']['neo']
    p = tim_doan(body, 'Lực kéo cần trong mỗi thanh để đạt [K] chỉ 40 ÷ 57 kN')
    n_lo, n_hi = neo['MNLTK']['N_can'], neo['MNLKT']['N_can']
    As = 490.9
    thay(p, 'Lực kéo cần trong mỗi thanh để đạt [K] chỉ 40 ÷ 57 kN, bằng 23 ÷ 33% Ntt, tương ứng ứng suất trong thép '
            '82 ÷ 117 MPa và cường độ bám dính vữa – đá cần huy động 0,08 ÷ 0,12 MPa',
         'Lực kéo cần trong mỗi thanh để đạt [K] chỉ ' + so(n_lo, 0) + ' ÷ ' + so(n_hi, 0) + ' kN, bằng '
         + so(100 * n_lo / 171.8, 0) + ' ÷ ' + so(100 * n_hi / 171.8, 0) + '% N_{tt}, tương ứng ứng suất trong thép '
         + so(n_lo * 1e3 / As, 0) + ' ÷ ' + so(n_hi * 1e3 / As, 0) + ' MPa và cường độ bám dính vữa – đá cần huy '
         'động ' + so(n_lo / (3.14159 * 0.076 * 2.0) / 1e3, 2) + ' ÷ ' + so(n_hi / (3.14159 * 0.076 * 2.0) / 1e3, 2)
         + ' MPa')
    thay(p, 'thí nghiệm kéo nhổ tối thiểu 3 neo đến 100 kN (≈ 1,7 lần lực kéo cần lớn nhất)',
         'thí nghiệm kéo nhổ tối thiểu 3 neo đến 100 kN (≈ ' + so(100 / n_hi, 1) + ' lần lực kéo cần lớn nhất)')
    p2 = doan_moi(
        'Thí nghiệm kéo nhổ đến 100 kN là thí nghiệm kiểm chứng theo lực yêu cầu (tiêu chí đạt: không tuột, chuyển vị '
        'dư sau dỡ tải trong giới hạn của chỉ dẫn kỹ thuật); nó không xác nhận khả năng 171,8 kN dùng để đánh giá dự trữ '
        '(K = 2,23 / 1,81), nên trong kiểm tra ổn định thiết kế chỉ dùng lực neo không vượt lực đã thử (lực cần '
        + so(n_lo, 0) + ' ÷ ' + so(n_hi, 0) + ' kN < 100 kN). Kiểm tra các dạng phá hoại khác: (i) '
        'phía bê tông – đoạn neo trong khối ngưỡng dài khoảng 1,0 m cho khả năng bám dính thép – vữa '
        + so(kb['phia_be_tong']['N_thep_vua'], 0) + ' kN và vữa – bê tông ' + so(kb['phia_be_tong']['N_vua_bt'], 0)
        + ' kN, đều lớn hơn lực thử; (ii) nhổ khối đá – mỗi neo giữ một lăng trụ đá 1,5 × 2,5 × 2,0 m, trọng lượng '
        'đẩy nổi ' + so(kb['khoi_da']['W'], 0) + ' kN (bỏ qua sức kháng cắt của đá), hệ số ' + so(kb['khoi_da'][
            'K_nho_N_can_max'], 2) + ' với lực kéo cần lớn nhất và ' + so(kb['khoi_da']['K_nho_N_thu'], 2)
        + ' với lực thử; (iii) bảo vệ chống ăn mòn – tham khảo yêu cầu mục 7.4 của TCVN 4116:2023 (quy định cho công '
        'trình neo vào nền cấp đặc biệt, cấp I): thanh neo có định vị tâm để lớp vữa bao không nhỏ hơn 25 mm trên toàn '
        'chiều dài, bơm vữa kín lỗ từ đáy lên; đầu thanh cắt thấp hơn mặt ngưỡng 60 mm và trám bằng vữa xi măng – '
        'polime mác chống thấm ≥ W8 (lớp bảo vệ ≥ 60 mm ở mặt tiếp xúc trực tiếp với dòng chảy – mục 6.3.3, 5.1.9). '
        'Hình học ngưỡng, hệ số ma sát và lực dính hiện là giả thiết, nên neo là phương án gia cường dự kiến, được chốt '
        'sau khi có hồ sơ hoàn công, khoan lõi mặt tiếp xúc và thí nghiệm kéo nhổ.', p)
    chen_sau(p, p2)
    ghi('Mục 4.5.3: Bảng 4.6 tính lại với [K] mới (R cần %s / %s kN/m; lực kéo 1 thanh %s / %s kN); mục đích thử kéo '
        'nhổ (100 kN ≈ %s lần lực cần); kiểm tra phía bê tông, nhổ khối đá; bảo vệ chống ăn mòn (TCVN 4116:2023 mục '
        '7.4, 6.3.3, 5.1.9) – góp ý A10.' % (so(neo['MNLTK']['R_can'], 1), so(neo['MNLKT']['R_can'], 1), so(n_lo, 1),
                                           so(n_hi, 1), so(100 / n_hi, 1)))


def chuong4_muc46(d, body, KQ):
    """Mục 4.6 mới – yêu cầu của giảng viên (12/10/2026)."""
    M = KQ['mua']
    h1 = tim_doan(body, 'THIẾT KẾ SỬA CHỮA, NÂNG CẤP ĐẬP ĐẤT', style='Heading1')
    mau_h2 = tim_doan(body, 'Kiểm tra ổn định ngưỡng tràn chính', style='Heading2')
    mau_h3 = tim_doan(body, 'Sơ đồ và tải trọng', style='Heading3')
    mau_nd = tim_doan(body, 'Nhận xét (Bảng 4.5): khi bỏ qua hoàn toàn lực dính')
    mau_cap = tim_doan(body, 'Bảng 4.5', style='Chuthichbang')
    cap41 = tim_doan(body, 'Hình 4.1', style='Chuthichhinh')
    pic41 = cap41.getprevious()
    cx = int(pic41.find('.//{%s}extent' % C.WP).get('cx'))
    ds = []

    def nd(markup):
        ds.append(doan_moi(markup, mau_nd))

    ds.append(doan_moi('Đánh giá khả năng tháo lũ khi lượng mưa, lũ biến đổi và hiệu quả của giải pháp sửa chữa '
                       'công trình tháo lũ', mau_h2))
    nd('QCVN 04-05:2022 [6] yêu cầu khi xác định các chỉ tiêu dòng chảy phải dự báo khả năng diễn biến của chúng trong '
       'tương lai (điều 2.3.2.2) và lưu lượng xả tính toán phải xét sự thay đổi điều kiện hình thành dòng chảy trên lưu '
       'vực (điều 2.6.3.3). Mục này đánh giá: (1) biến động lượng mưa lưu vực từ khi công trình đưa vào vận hành (năm '
       '2000); (2) mực nước lớn nhất trong hồ và khả năng tháo của tràn hiện trạng khi lũ đến tăng, so sánh với MNLTK '
       'thiết kế ban đầu; (3) hiệu quả của các giải pháp sửa chữa công trình tháo lũ.')
    # ---------------------------------------------------------------- 4.6.1
    ds.append(doan_moi('Biến động lượng mưa lưu vực từ khi công trình vận hành', mau_h3))
    nd('Số liệu: chuỗi lượng mưa tháng, năm của lưu vực 1978 ÷ 2024 (47 năm, QTVH [1] Bảng 2-9 – quy từ các trạm Đại '
       'Ninh, Sông Lũy) và lượng mưa 1 ngày lớn nhất năm X_{1max} 1978 ÷ 2020 (43 năm, báo cáo KTTV [7] Bảng 3-13), số '
       'liệu chi tiết ở Bảng PL12 (Phụ lục 7). Chuỗi được chia thành hai giai đoạn: trước (1978 ÷ 1999) và sau khi '
       'công trình vận hành (2000 ÷ 2024). Ngoài so sánh trị bình quân, hệ số biến đổi C_{v}, đồ án kiểm định xu thế '
       'bằng phương pháp phi tham số Mann – Kendall (có hiệu chỉnh số liệu trùng) kèm độ dốc Sen và kiểm định Welch '
       'cho sai khác trị bình quân của hai giai đoạn (mức ý nghĩa 5%; chương trình kt_mua_lu_bien_doi.py).')
    hang47 = []
    for k, ten in (('mua_nam', 'Lượng mưa năm'), ('mua_mua', 'Mưa mùa mưa (tháng V ÷ X)'),
                   ('mua_thang_max', 'Mưa tháng lớn nhất trong năm'), ('x1max', 'Mưa 1 ngày lớn nhất X_{1max} (đến 2020)')):
        v = M[k]
        mk = v['mk_toan_chuoi']
        hang47.append([ten, so(v['tb_truoc'], 1) + ' ; ' + so(v['cv_truoc'], 2),
                       so(v['tb_sau'], 1) + ' ; ' + so(v['cv_sau'], 2), so(v['thay_doi_pt'], 1) + '%',
                       so(mk['Z'], 2) + ' ; ' + so(mk['p'], 2) + ' ; ' + so(mk['sen'], 2), so(v['p_welch'], 2)])
    cap, tbl, kc = bang_moi('Bảng 4.7 Đặc trưng lượng mưa lưu vực hồ Cà Giây trước và sau khi công trình vận hành',
                            ['Đặc trưng', '1978 ÷ 1999: TB (mm) ; C_{v}', '2000 ÷ 2024: TB (mm) ; C_{v}', 'Thay đổi TB',
                             'Mann – Kendall: Z ; p ; Sen (mm/năm)', 'Welch: p'],
                            hang47, [2300, 1450, 1450, 1000, 1850, 950], mau_cap)
    ds += [cap, tbl, kc]
    ds.append(doan_anh(d, pic41, os.path.join(HINH, 'hinh_4_2_mua.jpg'), cx, 'hinh_4_2_mua.jpg'))
    ds.append(doan_moi('Hình 4.2 Lượng mưa năm và lượng mưa 1 ngày lớn nhất năm của lưu vực trước và sau khi công trình '
                       'vận hành', cap41))
    dem = M['x1max_dem']
    mn, x1 = M['mua_nam'], M['x1max']
    dau = lambda v: ('+' if v > 0 else '') + so(v, 1)
    tb_n = (mn['tb_truoc'] * mn['n_truoc'] + mn['tb_sau'] * mn['n_sau']) / (mn['n_truoc'] + mn['n_sau'])
    tb_x = (x1['tb_truoc'] * x1['n_truoc'] + x1['tb_sau'] * x1['n_sau']) / (x1['n_truoc'] + x1['n_sau'])
    xu_the = (' Tính theo độ dốc Sen, trên toàn chuỗi lượng mưa năm thay đổi '
              + dau(1000 * mn['mk_toan_chuoi']['sen'] / tb_n) + '%/10 năm và X_{1max} '
              + dau(1000 * x1['mk_toan_chuoi']['sen'] / tb_x) + '%/10 năm; riêng giai đoạn công trình vận hành, lượng '
              'mưa năm (2000 ÷ 2024) có xu thế ' + dau(1000 * mn['mk_sau']['sen'] / mn['tb_sau']) + '%/10 năm (p = '
              + so(mn['mk_sau']['p'], 2) + ') và X_{1max} (2000 ÷ 2020) ' + dau(1000 * x1['mk_sau']['sen'] / x1['tb_sau'])
              + '%/10 năm (p = ' + so(x1['mk_sau']['p'], 2) + ') – đều không có ý nghĩa thống kê.')
    nd('Kết quả (Bảng 4.7, Hình 4.2): lượng mưa năm bình quân hai giai đoạn gần như không đổi (' + so(M['mua_nam'][
        'tb_truoc'], 1) + ' và ' + so(M['mua_nam']['tb_sau'], 1) + ' mm, ' + so(M['mua_nam']['thay_doi_pt'], 1)
       + '%); lượng mưa mùa mưa giảm ' + so(-M['mua_mua']['thay_doi_pt'], 1) + '%, mưa tháng lớn nhất giảm '
       + so(-M['mua_thang_max']['thay_doi_pt'], 1) + '% và X_{1max} bình quân giảm ' + so(-M['x1max']['thay_doi_pt'], 1)
       + '%. Không có chuỗi nào có xu thế hoặc sai khác trị bình quân có ý nghĩa thống kê ở mức 5% (p = '
       + so(min(M[k]['mk_toan_chuoi']['p'] for k in ('mua_nam', 'mua_mua', 'mua_thang_max', 'x1max')), 2) + ' ÷ '
       + so(max(M[k]['mk_toan_chuoi']['p'] for k in ('mua_nam', 'mua_mua', 'mua_thang_max', 'x1max')), 2)
       + ').' + xu_the + ' Tuy vậy mức dao động tăng lên: C_{v} của lượng mưa năm tăng từ ' + so(M['mua_nam']['cv_truoc'], 2)
       + ' lên ' + so(M['mua_nam']['cv_sau'], 2) + ', của X_{1max} từ ' + so(M['x1max']['cv_truoc'], 2) + ' lên '
       + so(M['x1max']['cv_sau'], 2) + '; số năm có X_{1max} > 200 mm tăng từ ' + str(dem['truoc_200']) + ' lên '
       + str(dem['sau_200']) + ' và trận mưa 1 ngày lớn nhất của cả chuỗi (280,4 mm, năm 2003) xảy ra sau khi công '
       'trình vận hành, trong khi số năm có X_{1max} > 150 mm giảm từ ' + str(dem['truoc_150']) + ' xuống '
       + str(dem['sau_150']) + '. Như vậy số liệu đo chưa cho thấy lượng mưa tăng; thay đổi chủ yếu là mưa cực trị '
       'phân tán hơn. Các nghiên cứu dự tính cho vùng Nam Trung Bộ – Tây Nguyên theo 11 mô hình khí hậu toàn cầu cho '
       'thấy lượng mưa lớn nhất 1 ÷ 7 ngày có xu hướng tăng theo kịch bản RCP4.5 và RCP8.5, nhưng mức tăng phân tán '
       'rất rộng giữa các mô hình (từ trên 0% đến 70 ÷ 80%) [@31].')
    nd('Bản thân các lần cập nhật thủy văn của công trình đã làm lũ thiết kế tăng: lưu lượng đỉnh lũ P = 1% từ 847 '
       'm^{3}/s (thiết kế năm 1996) lên 914 m^{3}/s (QTVH 2005, BCNCKT 2022) và 921 m^{3}/s (QTVH 2026), tức tăng '
       '7,9 ÷ 8,7%; lũ P = 0,2% là 1.239 m^{3}/s và QTVH 2026 bổ sung lũ P = 0,1% là 1.398 m^{3}/s [1], [7].')
    # ---------------------------------------------------------------- 4.6.2
    ds.append(doan_moi('Mực nước lũ khi lũ đến gia tăng – so sánh với MNLTK thiết kế ban đầu', mau_h3))
    nd('So với MNLTK thiết kế ban đầu +77,24 m (TK 1996, Q_{1%} = 847 m^{3}/s, tràn chính rộng 25 m [7]), MNLTK tính '
       'với hiện trạng là +77,63 m, cao hơn 0,39 m. Tách riêng từng nguyên nhân bằng cùng chương trình điều tiết: với '
       'tràn chính rộng 25 m, lũ 847 m^{3}/s cho +77,31 m (sát giá trị thiết kế cũ, sai khác 0,07 m do đường đặc tính '
       'lòng hồ đo lại năm 2022); thay bằng lũ cập nhật 914 m^{3}/s cho +77,51 m (+0,20 m); thay tiếp bằng bề rộng tràn '
       'hiện trạng 9 × 2,4 = 21,6 m cho +77,63 m (+0,12 m). Để mực nước lũ thiết kế không vượt +77,24 m với tràn hiện '
       'trạng, lũ thiết kế phải nhỏ hơn hiện nay khoảng 14%. Như vậy MNLTK đã vượt MNLTK cũ ngay với bộ số liệu lũ '
       'hiện hành, do lũ thiết kế được cập nhật lớn hơn và tràn chính thực tế hẹp hơn thiết kế ban đầu.')
    nd('Do hồ không có số liệu đo lưu lượng lũ đến (lũ thiết kế được tính từ mưa), đồ án giả thiết lũ đến tăng k = 10, '
       '20 và 30% so với lũ thiết kế hiện hành – nhân toàn bộ tung độ đường quá trình (đỉnh và tổng lượng cùng tăng). '
       'Dải giả thiết bao được mức tăng do cập nhật thủy văn (≈ 8 ÷ 9%) và phần dưới của dải dự tính biến đổi khí hậu '
       '[@31]. Điều tiết với tràn hiện trạng, quy tắc vận hành như mục 2.3, đầu trận lũ ở MNDBT; kết quả trong Bảng 4.8 '
       'và Hình 4.3.')
    kq = {(r['nguon'], r['P'], round(100 * r['k'])): r for r in M['dieu_tiet']}
    hang48 = []
    for nguon, ten in (('BCNCKT', 'BCNCKT'), ('QTVH', 'QTVH')):
        for P, tp in (('1%', 'P = 1%'), ('0.2%', 'P = 0,2%')):
            hang48.append([ten + ', ' + tp] + ['+' + so(kq[(nguon, P, k)]['Zmax']) + ' (' + so(kq[(nguon, P, k)]['qxa'], 0)
                                               + ')' for k in (0, 10, 20, 30)])
    cap, tbl, kc = bang_moi('Bảng 4.8 Mực nước lớn nhất trong hồ Z_{max} (m) và lưu lượng xả lớn nhất (m^{3}/s) khi lũ '
                            'đến tăng k (tràn hiện trạng, đầu trận lũ ở MNDBT)',
                            ['Đường quá trình lũ', 'k = 0', 'k = 10%', 'k = 20%', 'k = 30%'],
                            hang48, [2600, 1600, 1600, 1600, 1600], mau_cap)
    ds += [cap, tbl, kc]
    ds.append(doan_anh(d, pic41, os.path.join(HINH, 'hinh_4_3_lu_tang.jpg'), cx, 'hinh_4_3_lu_tang.jpg'))
    ds.append(doan_moi('Hình 4.3 Mực nước lớn nhất trong hồ theo mức tăng lưu lượng lũ đến (tràn hiện trạng)', cap41))
    gh = M['k_gioi_han']
    z0 = M['Z0_don_lu']
    nd('Nhận xét: (1) với lũ thiết kế P = 1%, ngay cả khi lũ tăng 30% mực nước lớn nhất chỉ +' + so(kq[('BCNCKT', '1%', 30)][
        'Zmax']) + ' m, thấp hơn đỉnh đập sau sửa chữa ' + so(78.85 - kq[('BCNCKT', '1%', 30)]['Zmax']) + ' m – nước lũ '
       'thiết kế chưa tràn qua đỉnh đập; nhưng mọi trường hợp đều cao hơn MNLTK cũ 0,39 ÷ ' + so(kq[('BCNCKT', '1%', 30)][
        'Zmax'] - 77.24) + ' m, cần được xét khi rà soát ranh giới ngập lòng hồ và các công trình ven hồ; (2) với lũ kiểm '
       'tra P = 0,2%, mực nước hồ chạm đỉnh đập +78,85 m khi lũ tăng khoảng ' + so(100 * gh['k_P02_dinh'], 0) + '% '
       '(đường quá trình BCNCKT) hoặc ' + so(100 * KQ['k_qtvh_dinh'], 0) + '% (QTVH), chạm đỉnh tường chắn sóng khi '
       'tăng khoảng ' + so(100 * gh['k_P02_tuong'], 0) + '%; điều kiện đỉnh đập ≥ MNLKT + 0,3 m chỉ còn dự trữ 0,01 m '
       'nên không còn thỏa khi lũ tăng quá khoảng ' + so(100 * gh['k_P02_dinh_tru_03'], 1) + '%; với lũ thiết kế, sóng leo do gió Đông Bắc (Z_{2}) chạm đỉnh tường khi lũ '
       'tăng khoảng ' + so(100 * gh['k_P1_song_NE'], 0) + '%; (3) lũ vượt kiểm tra P = 0,1% (QTVH) cho +78,83 m, thấp '
       'hơn đỉnh đập 0,02 m. Kết luận: theo các điều kiện cao trình đỉnh đập và tường chắn sóng, dự trữ khả năng tháo '
       'của tràn hiện trạng chỉ tương ứng mức tăng lũ khoảng 9 ÷ 12% (lũ kiểm tra chạm đỉnh đập; sóng do gió Đông Bắc '
       'khi có lũ thiết kế chạm đỉnh tường), còn điều kiện MNLKT + 0,3 m hầu như không còn dự trữ.')
    nd('Biện pháp: nếu muốn giữ MNLKT + 0,3 m ≤ +78,85 m khi lũ tăng 10, 20, 30% thì mực nước đầu trận lũ phải hạ '
       'trước xuống +' + so(z0['10']) + ', +' + so(z0['20']) + ' và +' + so(z0['30']) + ' m, tương ứng xả bỏ '
       + so(KQ['dV_don_lu']['10'], 1) + ', ' + so(KQ['dV_don_lu']['20'], 1) + ' và ' + so(KQ['dV_don_lu']['30'], 1)
       + ' triệu m^{3} (' + so(100 * KQ['dV_don_lu']['10'] / 29.60, 0) + ' ÷ ' + so(100 * KQ['dV_don_lu']['30'] / 29.60, 0)
       + '% dung tích hữu ích). Hạ cố định mực nước suốt mùa lũ là không phù hợp vì nhiệm vụ cấp nước TH1A chỉ vừa '
       'đạt mức bảo đảm khi hồ tích đến MNDBT (mục 2.2). Phương án khả thi là xả trước theo dự báo: mở hết cửa van tràn '
       'sự cố (lưu lượng 64 → 44 m^{3}/s) hạ hồ từ +74,70 xuống +73,84 m mất khoảng 27 ÷ 44 giờ (lưu lượng đến nền 0 '
       '÷ 20 m^{3}/s), phù hợp thời gian dự báo mưa lớn 2 ÷ 3 ngày; hạ sâu hơn (+72,87 m) cần 3 ÷ 7 ngày nên không thực '
       'hiện được bằng tràn sự cố. Đồ án kiến nghị: bổ sung vào quy trình vận hành quy định xả trước khi có dự báo mưa '
       'lớn diện rộng trên lưu vực trong nửa đầu mùa lũ; cập nhật tính toán thủy văn định kỳ (khoảng 5 năm hoặc sau '
       'trận lũ lớn), có xét lũ kép theo điều 2.6.1.2.3 của QCVN 04-05:2022; nếu lũ kiểm tra cập nhật tăng quá khoảng '
       '9% thì phải nghiên cứu mở rộng tràn hoặc nâng đỉnh đập (điều 2.6.3.5d).')
    # ---------------------------------------------------------------- 4.6.3
    ds.append(doan_moi('Hiệu quả của giải pháp sửa chữa đối với công trình tháo lũ', mau_h3))
    nd('Quy mô tràn chính và tràn sự cố được giữ nguyên nên mực nước lũ thiết kế, lũ kiểm tra sau sửa chữa không đổi '
       '(Bảng 2.3, Bảng 4.1). Hiệu quả của các giải pháp được so sánh theo các chỉ tiêu an toàn trong Bảng 4.9.')
    neo = KQ['neo']
    hang49 = [
        ['Khoảng cách từ MNLKT (+78,54) đến đỉnh đập', 'MNLKT cao hơn đỉnh đập +78,35 m 0,19 m',
         'thấp hơn đỉnh +78,85 m 0,31 m', '≥ 0,30 m (TCVN 8216:2018, 10.1.5)', 'Đạt (nhờ nâng đỉnh đập, Chương 5)'],
        ['Ổn định trượt ngưỡng tràn chính khi c = 0 (MNLTK / MNLKT)', 'K = 0,82 / 0,65',
         'K = 2,23 / 1,81 (có neo); lực neo cần ' + so(neo['MNLTK']['N_can'], 0) + ' / ' + so(neo['MNLKT']['N_can'], 0)
         + ' kN/thanh', '[K] = 1,21 / 1,09', 'Đạt, không phụ thuộc lực dính'],
        ['Đáy dốc tràn chính', 'Mài mòn lớp mặt 2 ÷ 3 cm', 'Vữa xi măng – polime ≥ 40 MPa, ≥ W8',
         '≥ B25, W8 (TCVN 4116:2023, 5.1.9)', 'Đạt'],
        ['Kênh dẫn thượng lưu tràn chính', 'Bồi lấp một phần', 'Nạo vét đến cao trình thiết kế',
         'Điều kiện chảy tới ngưỡng như tính toán (m = 0,45)', 'Nếu m chỉ 0,40: MNLKT tăng 0,08 m'],
        ['Thấm mang trái tràn sự cố (MNLTK, 50 m)', '4,2 m^{3}/ngày đêm', '1,9 m^{3}/ngày đêm (màng 3 hàng)', '–',
         'Giảm 54%'],
        ['Gradien cửa ra lớn nhất mang tràn (MNLTK / MNLKT)', '0,29 / 0,34', '0,08 / 0,17', '< J_{th} = 1,02',
         'Giảm 2 ÷ 4 lần'],
        ['Điểm thấm ra mái hạ lưu mang tràn (MNLTK, MNLKT)', '+74,95 ÷ +75,14 m (trên mái)',
         'Lùi về chân mái +74,75 m', 'Không thấm ra mái', 'Đạt'],
        ['Cửa van tràn sự cố (nguy cơ kẹt van)', 'Tời điện, chưa có quy định kiểm tra',
         'Chạy thử trước mùa lũ, nguồn điện dự phòng, mở tay', '–', 'Giảm rủi ro các tình huống Bảng 4.2'],
    ]
    cap, tbl, kc = bang_moi('Bảng 4.9 So sánh các chỉ tiêu an toàn của công trình tháo lũ trước và sau sửa chữa',
                            ['Chỉ tiêu', 'Hiện trạng', 'Sau sửa chữa', 'Yêu cầu', 'Nhận xét'], hang49,
                            [2300, 1700, 2000, 1600, 1400], mau_cap)
    ds += [cap, tbl, kc]
    nd('Như vậy các giải pháp sửa chữa không làm tăng khả năng tháo mà nâng mức an toàn của cụm công trình tháo lũ: '
       'mực nước lũ kiểm tra từ chỗ cao hơn đỉnh đập trở thành thấp hơn đỉnh đập 0,31 m; ngưỡng tràn chính ổn định '
       'không phụ thuộc lực dính chưa được kiểm chứng; dòng thấm qua mang tràn sự cố được khống chế (Chương 8). Với lũ '
       'tăng trên khoảng 9% thì các giải pháp này chưa đủ – cần biện pháp vận hành (xả trước theo dự báo) hoặc công '
       'trình (mở rộng tràn) như đã nêu ở mục 4.6.2.')
    chen_truoc(h1, *ds)
    ghi('Mục 4.6 mới (yêu cầu của giảng viên): 4.6.1 biến động mưa từ 2000 (Bảng 4.7, Hình 4.2: mưa năm %s%%, X1max %s%%, '
        'không có xu thế ý nghĩa; Cv tăng); 4.6.2 MNLTK +77,63 so với +77,24 (+0,20 m do lũ cập nhật, +0,12 m do tràn '
        '21,6 m), điều tiết khi lũ tăng 10 ÷ 30%% (Bảng 4.8, Hình 4.3), ngưỡng k chạm đỉnh %s%%, tường %s%%; đón lũ, xả '
        'trước theo dự báo; 4.6.3 hiệu quả giải pháp sửa chữa công trình tháo lũ (Bảng 4.9).'
        % (so(M['mua_nam']['thay_doi_pt'], 1), so(M['x1max']['thay_doi_pt'], 1), so(100 * gh['k_P02_dinh'], 0),
           so(100 * gh['k_P02_tuong'], 0)))


# =========================================================================== CHƯƠNG 5
def chuong5(d, body, KQ):
    ne = KQ['huong']['NE']
    dp = DU_PHONG_LUN
    p = tim_doan(body, 'Từ kết quả mục 2.4 và phương án chọn ở mục 3.2.1')
    thay(p, 'max(Z1, Z2, Z3) = max(+77,15; +78,97; +78,74) = +78,97 m (gió Đông Bắc)',
         'max(Z_{1}, Z_{2}, Z_{3}) = max(+' + so(ne['Z1']) + '; +' + so(ne['Z2']) + '; +78,74) = +' + so(ne['Z2'])
         + ' m (gió Đông Bắc)')
    thay(p, 'mép hạ lưu tấm +78,76 m vẫn cao hơn MNLKT 0,22 m.',
         'mép hạ lưu tấm +78,76 m vẫn cao hơn MNLKT 0,22 m (nước chỉ tràn qua đỉnh khi vượt mép thượng lưu – mục '
         '3.2.1).')
    a = van_ban(p).index('Điều 11.4.4 TCVN 8216:2018')
    b = van_ban(p).index('dung sai chỉ cho phép dương;')
    cu = van_ban(p)[a:b]
    thay(p, cu,
         'Theo TCVN 8216:2018 điều 11.4.4, với đập cao trên 20 m chiều cao dự phòng lún phải kể vào cao trình đỉnh '
         'đập khi nghiệm thu hoàn thành và không nhỏ hơn 0,5%H; QCVN 04-05:2022 điều 2.6.2.2.1a cũng yêu cầu chiều cao '
         'an toàn của đập kể cả chiều cao phòng lún. Lớp đắp bù chỉ dày 0,4 ÷ 0,5 m trên thân đập đã làm việc hơn 25 '
         'năm nên độ lún bổ sung ước tính chỉ khoảng 0,5 ÷ 1 cm (tải trọng tăng thêm ≈ 11 kPa trên dải rộng 5 m, '
         '∫Δσ_{z}·dz ≈ 100 kPa·m trên chiều cao đập 24 m; mô đun biến dạng của đất đắp đã cố kết giả thiết E_{oed} ≈ '
         '10 ÷ 20 MPa vì hồ sơ chưa có thí nghiệm nén cố kết: s ≈ ∫Δσ_{z}·dz/E_{oed} ≈ 0,5 ÷ 1 cm); tuy nhiên chưa có '
         'số liệu quan trắc chứng minh thân đập đã lún ổn định, nên đồ án áp dụng thiên an toàn dự phòng lún theo điều '
         '11.4.4: 0,5% × 24,75 m ≈ 0,124 m, lấy ' + so(dp) + ' m (khối lượng đất tăng không đáng kể). Cao trình ghi '
         'trên bản vẽ là cao trình thiết kế sau lún (mép thượng lưu +78,85 m, mép hạ lưu tấm +78,76 m); cao độ hoàn '
         'thiện khi nghiệm thu (trước lún) không thấp hơn +' + so(78.85 + dp) + ' m tại mép thượng lưu và +'
         + so(78.76 + dp) + ' m tại mép hạ lưu tấm (đất đắp bù đến +' + so(78.60 + dp) + ' m), ')
    thay(p, 'Đỉnh tường chắn sóng hiện trạng +79,40 m được giữ nguyên; sau sửa chữa tường còn nhô cao 0,55 m so với mặt '
            'đập',
         'Đỉnh tường chắn sóng hiện trạng +79,40 m được giữ nguyên; sau sửa chữa tường còn nhô cao ' + so(79.40 - 78.85 - dp)
         + ' m (ngay sau thi công) ÷ 0,55 m (sau lún) so với mặt đập')
    p = tim_doan(body, 'Cào bóc lớp mặt đá dăm láng nhựa hư hỏng')
    thay(p, '(khi thi công đắp đến +78,62 m – kể dự phòng lún 2 cm)',
         '(khi thi công đắp đến +' + so(78.60 + dp) + ' m – kể dự phòng lún ' + so(dp) + ' m)')
    p = tim_doan(body, 'Rải lớp cấp phối đá dăm đệm dày 5 cm')
    thay(p, 'cao độ nghiệm thu trước lún ≥ +78,87 m tại mép thượng lưu (dự phòng lún 2 cm);',
         'cao độ nghiệm thu trước lún ≥ +' + so(78.85 + dp) + ' m tại mép thượng lưu, ≥ +' + so(78.76 + dp)
         + ' m tại mép hạ lưu tấm (dự phòng lún ' + so(dp) + ' m);')
    ghi('Mục 5.1.1, 5.1.3: dự phòng lún %s m (≥ 0,5%%H – TCVN 8216:2018 điều 11.4.4; QCVN 04-05:2022 điều 2.6.2.2.1a); '
        'nghiệm thu ≥ +%s / +%s m, đất đắp đến +%s m; bỏ câu "điều 11.4.4 viết cho đập đắp mới"; ghi rõ Eoed là giả '
        'thiết; Z2 = +%s m (góp ý A2).' % (so(dp), so(78.85 + dp), so(78.76 + dp), so(78.60 + dp), so(ne['Z2'])))

    # ---- 5.2 thấm sau sửa chữa: giá trị lớn nhất (A6), hội tụ (B6)
    gr = KQ['gradien']
    p = tim_doan(body, 'So với hiện trạng, lưu lượng thấm giảm khoảng 2 ÷ 3%')
    thay(p, '(tiếp xúc ống khói ≤ 0,94 < 2,0; cửa ra ≤ 0,16 < 0,85)',
         '(phân vị 95%: tiếp xúc ống khói ≤ 0,94, cửa ra ≤ 0,16; giá trị lớn nhất ' + so(gr['SC_MNLKT']['Jtiep_max'])
         + ' tại chân ống khói và ' + so(gr['SC_MNDBT_tac']['Jra_max']) + ' tại chân lăng trụ khi thoát nước bị tắc – '
         'nhỏ hơn 2,0 và 0,85)')
    p = tim_doan(body, 'Kiểm tra độ tin cậy của lời giải (Bảng PL8)')
    thay(p, 'làm lưu lượng thấm thay đổi không quá 0,2%,', 'làm lưu lượng thấm thay đổi khoảng 0,24% (dưới 0,3%),')
    thay(p, 'gradien cục bộ tại vùng tiếp xúc ống khói hội tụ về khoảng 1,09 (< 2,0);',
         'gradien cục bộ tại vùng tiếp xúc ống khói hội tụ về khoảng 1,09 (< 2,0) – hội tụ của lưu lượng không đồng '
         'nghĩa với hội tụ của gradien cục bộ (phân vị 95% thay đổi 0,94 ÷ 1,09 giữa các lưới), nên khi so sánh với '
         'giới hạn đồ án dùng giá trị lớn nhất trên các lưới;')
    ghi('Mục 5.2: báo cáo gradien lớn nhất (%s; %s); hội tụ q ≈ 0,24%% (không phải ≤ 0,2%%), tách hội tụ gradien '
        '(góp ý A6, B6).' % (so(gr['SC_MNLKT']['Jtiep_max']), so(gr['SC_MNDBT_tac']['Jra_max'])))

    # ---- 5.3 kiểm tra chéo ổn định (B5)
    L = KQ['luoi']
    giam_max = max(v['giam_pt'] for v in L.values())
    p = tim_doan(body, 'Kiểm tra chéo (Bảng PL6): với cung trượt nguy hiểm của các trường hợp TH1, TH4, TH5, TH7, TH10')
    a = van_ban(p)
    cu = a[:a.index('Trong giai đoạn thi công')]
    thay(p, cu,
         'Kiểm tra chéo (Bảng PL6, Bảng PL11): tìm lại cung trượt trên miền mở rộng (đáy cung tới cao trình +36 m), '
         'trên lưới lồng nhau chứa toàn bộ tâm, bán kính của lưới cơ sở và bằng tối ưu hóa Nelder – Mead từ 10 điểm '
         'xuất phát cho hệ số ổn định nhỏ nhất thấp hơn Bảng 5.2 không quá ' + so(giam_max, 2) + '% (TH3); phương pháp '
         'Spencer trên các cung tìm lại cho hệ số nhỏ hơn Bishop không quá 0,5%, riêng TH8 (hạ mực nước, mái thượng '
         'lưu) nhỏ hơn ' + so(100 * (1 - L['TH8']['spencer'] / L['TH8']['K_toi_uu']), 1) + '% (' + so(L['TH8'][
             'spencer'], 3) + ' so với ' + so(L['TH8']['K_toi_uu'], 3) + '); mọi trường hợp vẫn lớn hơn [K]. ')
    ghi('Mục 5.3: kiểm tra chéo tìm cung trượt bằng miền mở rộng, lưới lồng nhau, Nelder – Mead (K giảm ≤ %s%%), '
        'Spencer cho cả TH1 ÷ TH10 (TH8: %s so với %s) – góp ý B5.' % (so(giam_max, 2), so(L['TH8']['spencer'], 3),
                                                                       so(L['TH8']['K_toi_uu'], 3)))

    # ---- 5.5.1 ô rãnh (B8)
    xm = KQ['kc']['xoi_mai']
    cap, tbl = tim_bang(body, 'Bảng 5.5')
    r = tim_hang(tbl, '4', dung=True)
    v = xm['30%_L5.66']
    them_hang(tbl, r, ['5,66 (đường chéo ô)', so(v['q'] * 1e4, 2) + '·10^{-4}', so(v['h_mm'], 2), so(v['v'], 2),
                       so(v['tau'], 2)])
    p = tim_doan(body, 'Đặc trưng dòng chảy tràn theo chiều dài đường chảy cho trong Bảng 5.5.')
    thay(p, 'Khi chia mái thành ô 4 × 4 m bằng rãnh xiên, đường chảy dài nhất ≈ 4 m, ứng suất tiếp còn 1,8 Pa – giảm '
            'khoảng 3,6 lần;',
         'Khi chia mái thành ô 4 × 4 m bằng rãnh xiên, các ô đặt xoay 45° so với đường dốc nhất (Hình 3.2) nên đường '
         'chảy dài nhất là đường chéo của ô ≈ 5,66 m: ứng suất tiếp ' + so(v['tau'], 2) + ' Pa (1,82 Pa nếu tính theo '
         'cạnh ô 4 m) – giảm khoảng ' + so(xm['30%_L34.00']['tau'] / v['tau'], 1) + ' lần;')
    thay(p, 'ứng suất tiếp trên ô 4 × 4 m là 1,4 ÷ 2,2 Pa,',
         'ứng suất tiếp trên ô 4 × 4 m (theo đường chéo) là ' + so(xm['20%_L5.66']['tau'], 2) + ' ÷ '
         + so(xm['40%_L5.66']['tau'], 2) + ' Pa,')
    thay(p, 'ứng suất tiếp khi mưa lớn (≈ 2,2 Pa)', 'ứng suất tiếp khi mưa lớn (≈ 2,2 ÷ 2,7 Pa)')
    thay(p, 'mức giảm ứng suất tiếp (khoảng 3,6 lần)', 'mức giảm ứng suất tiếp (khoảng 2,9 lần theo đường chéo ô, 3,6 '
                                                        'lần theo cạnh ô)')
    ghi('Mục 5.5.1, Bảng 5.5: đường chảy dài nhất là đường chéo ô 5,66 m (ô xoay 45°): τ = %s Pa, giảm %s lần (thay '
        '1,8 Pa, 3,6 lần) – góp ý B8.' % (so(v['tau'], 2), so(xm['30%_L34.00']['tau'] / v['tau'], 1)))


def chuong5_muc56(d, body, KQ):
    """Mục 5.6 mới – so sánh hiện trạng và sau sửa chữa (yêu cầu của giảng viên)."""
    h1 = tim_doan(body, 'THIẾT KẾ SỬA CHỮA, NÂNG CẤP CỐNG LẤY NƯỚC', style='Heading1')
    mau_h2 = tim_doan(body, 'Thiết kế hệ thống thoát nước mái hạ lưu', style='Heading2')
    mau_nd = tim_doan(body, 'Kết quả tính cho trong Bảng 5.3 và Bảng 5.4.')
    mau_cap = tim_doan(body, 'Bảng 5.5', style='Chuthichbang')
    an, gr, L = KQ['anis'], KQ['gradien'], KQ['luoi']
    xm = KQ['kc']['xoi_mai']
    ds = [doan_moi('Đánh giá hiệu quả của giải pháp sửa chữa đối với an toàn đập', mau_h2)]
    ds.append(doan_moi(
        'Các kết quả tính thấm, ổn định hiện trạng (mục 2.5) và sau sửa chữa (mục 5.2 ÷ 5.5) dùng cùng chương trình, '
        'cùng mặt cắt và chỉ tiêu vật liệu nên so sánh trực tiếp được. Các chỉ tiêu chính được tổng hợp trong Bảng 5.7; '
        'hàng cuối là kiểm tra độ nhạy khi xét bất đẳng hướng thấm k_{h}/k_{v} = 5 (Bảng PL10).', mau_nd))

    def q6(k):
        return so(an[k]['q'] * 1e6, 2)
    dq = [100 * (an['SC_%s_r1' % k]['q'] / an['HT_%s_r1' % k]['q'] - 1) for k in ('MNDBT', 'MNLTK', 'MNLKT')]
    dk = [100 * (an['SC_TH%d_r5_K' % i] / an['SC_TH%d_r1_K' % i] - 1) for i in range(1, 11)]
    hang57 = [
        ['Cao trình đỉnh đập (không kể tường) / so với MNLKT +78,54 m', '+78,35 m / thấp hơn 0,19 m',
         '+78,85 m / cao hơn 0,31 m', '≥ MNLKT + 0,3 m', 'Không đạt → Đạt'],
        ['Lưu lượng thấm q (10^{-6} m^{3}/s/m): MNDBT / MNLTK / MNLKT',
         ' / '.join(q6('HT_%s_r1' % k) for k in ('MNDBT', 'MNLTK', 'MNLKT')),
         ' / '.join(q6('SC_%s_r1' % k) for k in ('MNDBT', 'MNLTK', 'MNLKT')), '–',
         'Giảm ' + so(-max(dq), 1) + ' ÷ ' + so(-min(dq), 1) + '%'],
        ['Đường bão hòa tại tim đập: MNDBT / MNLKT (m)',
         '+' + so(an['HT_MNDBT_r1']['z_tim']) + ' / +' + so(an['HT_MNLKT_r1']['z_tim']),
         '+' + so(an['SC_MNDBT_r1']['z_tim']) + ' / +' + so(an['SC_MNLKT_r1']['z_tim']), '–',
         'Hạ ' + so(an['HT_MNDBT_r1']['z_tim'] - an['SC_MNDBT_r1']['z_tim']) + ' ÷ '
         + so(an['HT_MNLKT_r1']['z_tim'] - an['SC_MNLKT_r1']['z_tim']) + ' m'],
        ['Gradien lớn nhất: tiếp xúc ống khói / cửa ra (MNLKT)',
         so(gr['HT_MNLKT']['Jtiep_max']) + ' / ' + so(gr['HT_MNLKT']['Jra_max'], 3),
         so(gr['SC_MNLKT']['Jtiep_max']) + ' / ' + so(gr['SC_MNLKT']['Jra_max'], 3), '2,0 / 0,85', 'Đạt, giảm nhẹ'],
        ['Thiết bị thoát nước bị tắc (MNDBT): đường bão hòa tại tim, K mái hạ lưu',
         '+' + so(an['HT_MNDBT_tac_r1']['z_tim']) + ' m; K = 1,519',
         'Khơi thông lăng trụ: +' + so(an['SC_MNDBT_r1']['z_tim']) + ' m; K = 1,647 (nếu vẫn tắc: +'
         + so(an['SC_MNDBT_tac_r1']['z_tim']) + ' m; K = 1,497)', 'K ≥ 1,17', 'Thấp hơn 2,8 ÷ 3,1 m; K tăng 8 ÷ 10%'],
        ['K_{min} ổn định tổng thể: mái hạ lưu / thượng lưu', '1,364 / 1,510', '1,349 / 1,496', '1,17 / 1,30',
         'Giảm ≈ 1% do tải trọng đắp đỉnh – vẫn đạt'],
        ['Ổn định lớp mặt mái hạ lưu dày 1 m, bão hòa hoàn toàn', 'K = 0,52 (đất tơi xốp, c′ = 0)',
         'K = 4,41 (đắp trả đầm K ≥ 0,95)', '≥ 1,17', 'Không đạt → Đạt'],
        ['Ứng suất tiếp dòng chảy tràn trên mái (mưa 30 phút = 30% mưa ngày)',
         so(xm['30%_L34.00']['tau'], 2) + ' Pa (đường chảy cả mái 34 m)',
         so(xm['30%_L4.00']['tau'], 2) + ' ÷ ' + so(xm['30%_L5.66']['tau'], 2) + ' Pa (ô 4 × 4 m)',
         'Đất có cỏ phủ ≫ 3,6 Pa', 'Giảm 2,9 ÷ 3,6 lần'],
        ['Độ nhạy k_{h}/k_{v} = 5 (sau sửa chữa)', '–',
         'q tăng ' + so(100 * min(an['SC_%s_r5' % k]['q'] / an['SC_%s_r1' % k]['q'] - 1 for k in
                                  ('MNDBT', 'MNLTK', 'MNLKT', 'MNDBT_tac')), 0) + ' ÷ '
         + so(100 * max(an['SC_%s_r5' % k]['q'] / an['SC_%s_r1' % k]['q'] - 1 for k in
                        ('MNDBT', 'MNLTK', 'MNLKT', 'MNDBT_tac')), 0) + '%; đường bão hòa tại tim cao thêm '
         + so(min(an['SC_%s_r5' % k]['z_tim'] - an['SC_%s_r1' % k]['z_tim'] for k in ('MNDBT', 'MNLTK', 'MNLKT')), 1)
         + ' ÷ ' + so(max(an['SC_%s_r5' % k]['z_tim'] - an['SC_%s_r1' % k]['z_tim'] for k in ('MNDBT', 'MNLTK', 'MNLKT')), 1)
         + ' m; K thay đổi ' + so(min(dk), 1) + ' ÷ +' + so(max(dk), 1) + '%', 'Như trên', 'Vẫn đạt'],
    ]
    cap, tbl, kc = bang_moi('Bảng 5.7 So sánh các chỉ tiêu an toàn của đập đất trước và sau sửa chữa (mặt cắt lòng sông)',
                            ['Chỉ tiêu', 'Hiện trạng', 'Sau sửa chữa', 'Yêu cầu', 'Nhận xét'], hang57,
                            [2500, 1750, 2150, 1300, 1300], mau_cap)
    ds += [cap, tbl, kc]
    ds.append(doan_moi(
        'Nhận xét: (1) giải pháp sửa chữa không thay đổi hệ thống chống thấm, thoát nước của thân đập (ống khói, đệm '
        'tiêu nước, lăng trụ) nên lưu lượng thấm và gradien chỉ giảm nhẹ (2 ÷ 3%, đường bão hòa hạ 0,4 ÷ 0,9 m do mái '
        'hạ lưu được đắp trả); an toàn thấm của đập vốn đã bảo đảm ở hiện trạng; (2) hiệu quả rõ nhất là ở các hư hỏng '
        'thực tế của đập: đỉnh đập từ chỗ thấp hơn MNLKT trở thành cao hơn MNLKT 0,31 m; lớp mặt mái hạ lưu từ trạng '
        'thái mất ổn định khi bão hòa (K = 0,52) trở thành ổn định (K = 4,41); ứng suất tiếp do dòng chảy tràn giảm 2,9 '
        '÷ 3,6 lần nhờ hệ thống rãnh và lớp cỏ; (3) việc khơi thông lăng trụ thoát nước và duy trì quan trắc đường bão '
        'hòa có ý nghĩa lớn – nếu thoát nước bị tắc, đường bão hòa tại tim dâng thêm khoảng 2,8 m và hệ số ổn định mái '
        'hạ lưu giảm khoảng 9%; (4) ổn định tổng thể giảm khoảng 1% do tải trọng đắp đỉnh nhưng vẫn lớn hơn yêu cầu, '
        'kể cả khi xét bất đẳng hướng thấm. Các nhận xét trên giới hạn ở mặt cắt lòng sông đại diện, các tổ hợp và chỉ '
        'tiêu cơ lý đã nêu (một số chỉ tiêu cường độ lấy theo trị tiêu chuẩn do hồ sơ thiếu trị tính toán).', mau_nd))
    chen_truoc(h1, *ds)
    ghi('Mục 5.6 mới (yêu cầu của giảng viên): Bảng 5.7 so sánh hiện trạng – sau sửa chữa (đỉnh đập, q, đường bão hòa, '
        'gradien, thoát nước tắc, K tổng thể, lớp mặt, ứng suất tiếp, bất đẳng hướng kh/kv = 5) và nhận xét hiệu quả.')


# =========================================================================== CHƯƠNG 6
def chuong6(d, body, KQ):
    cn = KQ['cap_nuoc']['TH1A|khong_Qtt']
    p = tim_doan(body, 'Khi cửa van mở hoàn toàn và cửa vào chưa bị ngập')
    thay(p, 'Theo [24], [31], cống chảy không áp ở cửa vào khi cột nước trước cống H ≤ (1,2 ÷ 1,4)·hc, với hc = 2,0 m – '
            'chiều cao cống (trị số lớn ứng với cửa vào có mép lượn), tức H ≤ 2,4 ÷ 2,8 m:',
         'Theo TCVN 9151:2012 [31] (điều 7.2.3), khi cột nước trên ngưỡng cửa vào H < 1,15h (h = 2,0 m – chiều cao '
         'cống, tức H < 2,30 m) cống làm việc theo sơ đồ chảy qua đập tràn; khi H ≥ 1,5h = 3,0 m cống làm việc theo sơ '
         'đồ chảy qua lỗ, vòi hoặc ống có áp; khoảng giữa là vùng chuyển tiếp mà tiêu chuẩn lưu ý còn phụ thuộc hình '
         'dạng cửa vào (chú thích 3 điều 3.1); giáo trình [24] cho giới hạn (1,2 ÷ 1,4)h tùy dạng cửa vào. Khả năng tháo '
         'theo sơ đồ đập tràn đỉnh rộng chảy không ngập:')
    p = tim_doan(body, 'Ở MNC: H0 = 1,90 m → Q = 7,31 m3/s > 6,54 m3/s')
    thay(p, '(= 1,17hc, vẫn trong giới hạn chảy không áp ở cửa vào nên công thức (6-1) áp dụng được), tức mực nước hồ '
            '≥ +67,94 m;',
         '(= 1,17h – vượt nhẹ giới hạn 1,15h của sơ đồ đập tràn, thuộc vùng chuyển tiếp), tức mực nước hồ ≥ +67,94 m '
         'theo (6-1);')
    thay(p, 'cho +67,97 m, sai khác 0,03 m nên kết luận không đổi.',
         'cho +67,97 m; đồ án lấy yêu cầu +67,94 ÷ +67,97 m (hai sơ đồ chỉ sai khác 0,03 m nên kết luận không đổi).')
    thay(p, 'điều kiện này cần được thể hiện trong quy trình vận hành (không cần cải tạo cống vì thời gian hồ ở sát MNC '
            'rất ngắn – xem Hình 2.1).',
         'điều kiện này cần được thể hiện trong quy trình vận hành. Theo điều tiết toàn chuỗi (TH1A), mực nước cuối '
         'tháng thấp hơn +67,94 m trong ' + str(cn['thang_Z_duoi_6794']) + '/528 tháng (' + str(
             cn['thang_Z_duoi_6794_mua_kho']) + ' tháng mùa khô), thuộc ' + str(cn['nam_co_thang_Z_duoi_6794'])
         + '/44 năm; nhiệm vụ hiện hữu 6,54 m^{3}/s luôn được bảo đảm, còn nhiệm vụ 10 m^{3}/s của QTVH chỉ được đáp '
         'ứng đầy đủ khi mực nước hồ từ +67,94 ÷ +67,97 m trở lên. Vì mực nước cuối tháng (Hình 2.1) không thể hiện '
         'mực nước thấp nhất trong tháng và thời điểm phát sinh nhu cầu lớn nhất, kết luận "không cần cải tạo cống" '
         'chỉ áp dụng cho nhiệm vụ hiện hữu; với nhiệm vụ mới cần lập lịch cấp nước theo ngày khi lập quy trình vận '
         'hành.')
    ghi('Mục 6.3.1: giới hạn chảy qua đập tràn H < 1,15h theo TCVN 9151:2012 điều 7.2.3 (thay (1,2 ÷ 1,4)h); H0 = 2,34 m '
        'thuộc vùng chuyển tiếp → yêu cầu +67,94 ÷ +67,97 m; số tháng mực nước < +67,94 m (TH1A: %d tháng, %d năm); tách '
        'nhiệm vụ hiện hữu và nhiệm vụ 10 m³/s (góp ý B3).' % (cn['thang_Z_duoi_6794'], cn['nam_co_thang_Z_duoi_6794']))
    p = tim_doan(body, 'Kết quả tính (Bảng 6.2, Hình 6.1) cho mức đầy lớn nhất 55 ÷ 73%')
    thay(p, 'thỏa mãn điều kiện chảy không áp ổn định (khoảng lưu không ≥ 0,4 m và diện tích ướt không vượt quá 75% '
            'diện tích mặt cắt).',
         'Số Froude cuối cống Fr = v^{2}/(g·h) = 5,7 (Q = 6,54 m^{3}/s) và 4,9 (Q = 10 m^{3}/s) < 10, nên theo TCVN '
         '9151:2012 (Bảng 3) độ đầy giới hạn của ống không áp là 0,90 và không bắt buộc cấp khí bổ sung (chú thích 1 '
         'điều 7.2.6 chỉ áp dụng khi L/h > 10 và Fr > 10). Đồ án vẫn dùng giới hạn thiên an toàn: khoảng lưu không ≥ '
         '0,4 m và độ đầy ≤ 75%, vì ở đầu cống chiều sâu xấp xỉ chiều sâu phân giới và độ đầy 0,73 > 0,7 với Fr < 5 – '
         'trường hợp tiêu chuẩn lưu ý có thể phát sinh sóng trong cống (chú thích 3 điều 7.2.8); cần giữ lỗ thông khí '
         'sau cửa van thông thoáng và quan sát dòng chảy khi vận hành thử với lưu lượng lớn.')
    p = tim_doan(body, 'Trường hợp độ dốc đáy chỉ là 3‰')
    thay(p, 'Lưu lượng lớn nhất còn bảo đảm chảy không áp ổn định (mức đầy 75%) khi i = 3‰ là 7,2 m3/s.',
         'Khi i = 3‰, năng lực của cống theo kiểm tra dòng đều ở độ đầy giới hạn 75% là 7,2 m^{3}/s (dòng êm, Fr < 5 – '
         'giới hạn 0,75 theo chú thích 3 điều 7.2.8 TCVN 9151:2012); đây không phải khả năng tháo lớn nhất tuyệt đối '
         'vì chế độ chảy thực tế còn phụ thuộc chiều dài cống, điều kiện cửa vào, cửa ra và mực nước hạ lưu – cần tính '
         'lại khi có trắc dọc đo đạc.')
    ghi('Mục 6.3.2: số Froude Fr = v²/(gh) = 4,9 ÷ 5,7 < 10 → độ đầy giới hạn 0,90 (TCVN 9151:2012 Bảng 3), cấp khí '
        'không bắt buộc (7.2.6); giữ 0,75 thiên an toàn (7.2.8 chú thích 3); 7,2 m³/s là năng lực theo dòng đều ở độ đầy '
        'giới hạn (góp ý B4).')

    # ---- 6.4 trục vít (A11)
    kod = KQ['kc']['truc_vit']['Kod']
    cap, tbl = tim_bang(body, 'Bảng 6.3')
    dat_o(o(tbl, 0, 9), 'K_{ođ} có gối trung gian (μ = 1,0): 1 gối, L = 3 m / 2 gối, L = 2 m')
    for r in range(1, len(hang(tbl))):
        row = dong_bang(tbl, r)
        mn = 'MNDBT' if 'MNDBT' in row[0] else 'MNLTK'
        k = '%s f=%s' % (mn, ('%.1f' % float(row[1].strip().replace(',', '.'))).replace('.', ','))
        dat_o(o(tbl, r, 9), so(kod[k]['goi_giua_mu10_L3'], 2) + ' / ' + so(kod[k]['hai_goi_mu10_L2'], 2))
    p = tim_doan(body, 'Trong đó: E = 2,1·105 MPa; d1 = 72 mm')
    them_cuoi(p, ' Khi bổ sung giá đỡ dẫn hướng trung gian, mỗi đoạn trục giữa hai gối chỉ bị giữ chuyển vị ngang, '
                 'không bị giữ xoay, nên lấy μ = 1,0 (hai đầu khớp) – thiên an toàn; giữ μ = 0,7 cho đoạn ngắn hơn là '
                 'không có cơ sở. Đường kính chân ren, đặc tính máy và lực đóng cưỡng bức cần được xác nhận theo hồ sơ '
                 'thiết bị khi lắp đặt.')
    p = tim_doan(body, 'Nhận xét (Bảng 6.3): (i) ở MNDBT, máy VĐ20 đủ sức nâng cửa')
    k_lo = min(kod[k]['hai_goi_mu10_L2'] for k in kod if 'f=0,5' not in k)
    k_hi = max(kod[k]['hai_goi_mu10_L2'] for k in kod if 'f=0,5' not in k)
    thay(p, 'bổ sung giá đỡ dẫn hướng trung gian chia trục thành hai nhịp L ≤ 3 m (Kođ tăng lên 3,4 ÷ 6,5 khi f ≤ 0,4; '
            'trường hợp bất lợi nhất f = 0,5 ở MNLTK vẫn đạt 2,7 > 2);',
         'bổ sung hai giá đỡ dẫn hướng trung gian chia trục thành ba đoạn L ≤ 2 m (với μ = 1,0: K_{ođ} = '
         + so(k_lo, 1) + ' ÷ ' + so(k_hi, 1) + ' khi f ≤ 0,4; trường hợp bất lợi nhất f = 0,5 ở MNLTK đạt '
         + so(kod['MNLTK f=0,5']['hai_goi_mu10_L2'], 2) + ' > 2; nếu chỉ dùng một gối giữa, L = 3 m thì K_{ođ} = '
         + so(kod['MNLTK f=0,5']['goi_giua_mu10_L3'], 2) + ' ÷ ' + so(kod['MNDBT f=0,3']['goi_giua_mu10_L3'], 2)
         + ', không đạt khi f = 0,5 ở MNDBT và khi f ≥ 0,4 ở MNLTK);')
    ghi('Mục 6.4, Bảng 6.3: đoạn trục giữa các gối dẫn hướng lấy μ = 1,0; đề xuất 2 gối (L = 2 m): Kođ ≥ %s; 1 gối '
        '(L = 3 m) không đạt khi f = 0,5 ở MNDBT và f ≥ 0,4 ở MNLTK (góp ý A11).' % so(kod['MNLTK f=0,5']['hai_goi_mu10_L2'], 2))

    # ---- 6.5 khung cống – TCVN 4116:2023
    p = tim_doan(body, 'Việc nâng đỉnh đập từ +78,35 m lên +78,85 m làm áp lực đất đứng trên cống tăng')
    thay(p, 'cốt thép yêu cầu ước tính theo As ≈ Kn·nc·M/(Rs·0,9·h0) với Kn = 1,15, nc = 1,0, thép CB300-V Rs = 260 MPa,',
         'cốt thép yêu cầu ước tính từ điều kiện bền (35) của TCVN 4116:2023 [34]: A_{s} ≈ γ_{lc}·γ_{n}·M/(γ_{c}·R_{s}·z), '
         'z ≈ 0,9h_{0}, với γ_{n} = 1,15 (cấp II), γ_{lc} = 1,0 (tổ hợp cơ bản, thời kỳ khai thác), γ_{c} = 1,0; thép '
         'CB300-V R_{s} = 260 MPa (Bảng 11),')
    thay(p, '– thiên an toàn. Kết quả cho trong Bảng 6.4.',
         '– thiên an toàn. Kiểm tra với giả thiết thận trọng bê tông thân cống cấp B15 (R_{b} = 8,9 MPa, Bảng 2): chiều '
         'cao vùng nén tương đối ξ = x/h_{0} = 0,11 ÷ 0,14 ≪ ξ_{R} = 0,70 (Bảng 21) – cấu kiện phá hoại dẻo, cánh tay '
         'đòn thực z = 0,93 ÷ 0,94h_{0} nên lấy z = 0,9h_{0} là thiên an toàn. Kết quả cho trong Bảng 6.4.')
    p = tim_doan(body, 'Cốt thép yêu cầu sau sửa chữa khoảng 8 ÷ 17 cm2/m')
    them_cuoi(p, ' Theo TCVN 4116:2023 (mục 7.2, 9.2), kết cấu BTCT của cống còn phải kiểm tra bề rộng khe nứt ở trạng '
                 'thái giới hạn thứ hai (γ_{n} = γ_{lc} = 1,0): bề rộng cho phép lấy theo Bảng 23, 24 của tiêu chuẩn '
                 '(phụ thuộc thành phần hóa học của nước và gradien cột nước), nhân 1,6 với công trình cấp II, nhân 0,5 '
                 'với kết cấu thành mỏng (chiều dày < 1,5 m) và không quá 0,5 mm – việc này cần cốt thép và kết quả phân '
                 'tích nước thực tế. Bề rộng khe nứt tính toán gần như tỷ lệ với ứng suất cốt thép nên sau khi nâng đỉnh '
                 'đập chỉ tăng khoảng 4 ÷ 5% so với hiện trạng; kết quả quan sát vết nứt hiện trạng vì vậy là chỉ dấu '
                 'trực tiếp. Nếu phải gia cường, tính toán dùng cường độ bê tông thực tế (khoan lõi) và trạng thái ứng '
                 'suất hiện có (mục 7.15). Cống có mặt cắt 1,8 × 2,0 m, lớn hơn kích thước tối thiểu 1,2 × 1,6 m để quan '
                 'sát, sửa chữa (TCVN 9151:2012 mục 5.15) nên có thể kiểm tra trực tiếp khi tháo cạn giữa hai vụ.')
    ghi('Mục 6.5: ký hiệu theo TCVN 4116:2023 (công thức 35: γn, γlc, γc; Rs Bảng 11); kiểm tra ξ = 0,11 ÷ 0,14 ≪ ξR = '
        '0,70 (Bảng 21); bổ sung yêu cầu kiểm tra khe nứt TTGH II (mục 9.2, Bảng 23, 24), cường độ thực tế khi gia cường '
        '(7.15), khả năng vào kiểm tra cống (TCVN 9151:2012 mục 5.15).')


# =========================================================================== CHƯƠNG 7
def chuong7(d, body, KQ):
    dp = DU_PHONG_LUN
    neo = KQ['neo']
    cap, tbl = tim_bang(body, 'Bảng 7.1')
    rows = hang(tbl)
    r = tim_hang(tbl, 'Đất đắp bù mặt đập K ≥ 0,97', cot=1)
    dat_hang(tbl, r, [None, 'Đất đắp bù mặt đập K ≥ 0,97 (đến +' + so(78.60 + dp) + ' m, kể dự phòng lún ' + so(dp)
                      + ' m; dốc 2%)', None, so_nghin(KQ['v_dat_dinh'])])
    r = tim_hang(tbl, 'Đào, đắp xử lý hố xói mái hạ lưu', cot=1)
    dat_hang(tbl, r, [None, 'Đào, đắp xử lý hố xói, đắp áp trúc mái hạ lưu (K ≥ 0,95): tổng đất đắp theo TKCS ≈ '
                            '25.000 m^{3} trừ phần nâng đỉnh lên +78,50 m của TKCS (≈ 986 × 4,7 × 0,30 ≈ 1.390 m^{3}) '
                            'đã tính ở dòng 2', None, '≈ 23.600'])
    r = tim_hang(tbl, 'Thay gioăng, tấm trượt, giá đỡ trục trung gian', cot=1)
    dat_hang(tbl, r, [None, 'Thay gioăng, tấm trượt; 02 giá đỡ dẫn hướng trục trung gian cho mỗi máy; sơn cửa van',
                      None, None])
    r = tim_hang(tbl, 'Mốc quan trắc lún mặt / giếng quan trắc thấm tự động', cot=1)
    dat_hang(tbl, r, [None, 'Mốc quan trắc lún mặt / giếng quan trắc thấm tự động (14 giếng thân đập theo TKCS + 2 '
                            'giếng đo áp hạ lưu màng mang tràn sự cố, mục 8.6)', None, '9 ; 14 + 2'])
    ghi('Bảng 7.1: đất đắp bù mặt đập %s m³ (đến +%s m, dự phòng lún %s m); dòng 9 tách phần nâng đỉnh của TKCS – '
        '≈ 23.600 m³ (góp ý B9); 02 giá đỡ trục/máy (A11); giếng quan trắc 14 + 2 (góp ý B10).'
        % (so_nghin(KQ['v_dat_dinh']), so(78.60 + dp), so(dp)))
    p = tim_doan(body, 'Khối lượng khoan phụt tổng hợp từ Bảng 2-2 TKCS')
    thay(p, 'đến +78,62 m – kể dự phòng lún 2 cm, dốc 2%).', 'đến +' + so(78.60 + dp) + ' m – kể dự phòng lún '
         + so(dp) + ' m, dốc 2%).')
    thay(p, 'tổng chiều dài khoan, phụt giảm khoảng 28%.',
         'chiều dài khoan giảm 27,85% và chiều dài phụt giảm 29,01% – đây là tỷ lệ khối lượng, chưa phải tỷ lệ tiết '
         'kiệm chi phí (chi phí còn gồm phụt thí nghiệm, kiểm tra, vật liệu, huy động và xử lý bổ sung).')
    p = tim_doan(body, 'Theo BCNCKT [3], vốn đầu tư sửa chữa, nâng cấp hồ Cà Giây khoảng 48,69 tỷ đồng.')
    thay(p, '(tăng khoảng 490 m3 đất đắp mặt đập – khoảng 590 m3 nếu kể dự phòng lún 2 cm và 36 neo ngưỡng tràn);',
         '(tăng khoảng ' + so_nghin(KQ['v_dat_dinh'] - KQ['v_dat_tkcs']) + ' m^{3} đất đắp mặt đập so với TKCS – kể cả '
         'dự phòng lún ' + so(dp) + ' m – và 36 neo ngưỡng tràn);')
    thay(p, 'khối lượng khoan, phụt giảm khoảng 28%.', 'chiều dài khoan, phụt giảm 27,85 ÷ 29,01%.')
    p = tim_doan(body, 'Bê tông mặt đập, rãnh, gờ chắn bánh và mặt đường quản lý dùng bê tông thương phẩm')
    thay(p, 'đặt thanh neo D25 có định vị tâm,', 'đặt thanh neo D25 có định vị tâm (lớp vữa bao ≥ 25 mm),')
    thay(p, 'bơm vữa xi măng không co ngót từ đáy lỗ lên;',
         'bơm vữa xi măng không co ngót từ đáy lỗ lên; cắt đầu thanh thấp hơn mặt ngưỡng 60 mm và trám bằng vữa xi '
         'măng – polime ≥ W8;')
    thay(p, '(≈ 1,7 lần lực kéo yêu cầu lớn nhất 57 kN)',
         '(≈ ' + so(100 / neo['MNLKT']['N_can'], 1) + ' lần lực kéo yêu cầu lớn nhất ' + so(neo['MNLKT']['N_can'], 0)
         + ' kN)')
    p = tim_doan(body, 'Công tác quan trắc thực hiện theo TCVN 8414:2010 [38]')
    thay(p, 'vượt dự phòng lún 2 cm', 'vượt dự phòng lún ' + so(dp) + ' m')
    cap, tbl = tim_bang(body, 'Bảng 7.3')
    r = tim_hang(tbl, 'Đường bão hòa thân đập (14 giếng)')
    dat_hang(tbl, r, ['Đường bão hòa thân đập (14 giếng) và hạ lưu màng mang tràn sự cố (2 giếng)', None, None])
    r = tim_hang(tbl, 'Cửa van, máy đóng mở, nguồn điện dự phòng')
    them_hang(tbl, r, ['Cập nhật tính toán thủy văn, lũ thiết kế', 'Tính lại với chuỗi mưa, lũ mới; kiểm tra lũ kép',
                       '5 năm/lần hoặc sau trận lũ lớn (mục 4.6)'])
    p = tim_doan(body, 'Cập nhật quy trình vận hành: quy định rõ thao tác mở cửa van tràn sự cố')
    thay(p, 'Cập nhật quy trình vận hành: quy định rõ thao tác mở cửa van tràn sự cố khi mực nước hồ vượt MNDBT, kiểm '
            'tra cửa van và nguồn điện trước mùa lũ;',
         'Cập nhật quy trình vận hành: quy định rõ thao tác mở cửa van tràn sự cố khi mực nước hồ vượt MNDBT, kiểm '
         'tra cửa van và nguồn điện trước mùa lũ; xả trước qua tràn sự cố khi có dự báo mưa lớn diện rộng trên lưu vực '
         '(mục 4.6.2); kiểm tra, mở hết 9 cửa van tràn chính, dọn vật cản ở các khoang tràn và kênh dẫn trước mùa lũ, '
         'không giữ cửa đóng để tích nước khi có dự báo lũ; ưu tiên cấp nước sinh hoạt '
         'khi thiếu nước (mục 2.2.3);')
    ghi('Mục 7.2 ÷ 7.5: neo – định tâm, lớp vữa ≥ 25 mm, đầu neo trám vữa W8, thử 100 kN ≈ %s lần; dự phòng lún %s m '
        'trong ngưỡng cảnh báo lún; Bảng 7.3 thêm 2 giếng mang tràn và cập nhật thủy văn; quy trình vận hành thêm xả '
        'trước theo dự báo, kiểm tra 9 khoang tràn chính, ưu tiên sinh hoạt.' % (so(100 / neo['MNLKT']['N_can'], 1), so(dp)))


# =========================================================================== CHƯƠNG 8
def chuong8(d, body, KQ):
    p = tim_doan(body, '(3) Chiều sâu màng: màng phải cắt qua hết các lớp thấm mạnh')
    thay(p, 'nên đáy màng hợp lý nằm khoảng 3 m dưới mặt lớp II nếu thí nghiệm ép nước xác nhận q ≤ 5 Lu.',
         'nên một phương án thử là đáy màng khoảng 3 m dưới mặt lớp II nếu thí nghiệm ép nước xác nhận q ≤ 5 Lu – trị '
         'số "cộng thêm 3 m" là giới hạn phạm vi khoan của điều 9.2.2.9 đối với đập từ cấp II trở xuống, không phải '
         'công thức chọn đáy màng; chiều sâu chọn phải được đánh giá bằng địa chất, thí nghiệm và tính toán thấm. Quy '
         'đổi 1 Lu ≈ 1,3·10^{-5} cm/s chỉ là gần đúng (phụ thuộc sơ đồ thí nghiệm, chiều dài đoạn ép và mức nứt nẻ), '
         'nên đồ án trình bày riêng lượng mất nước đo được (Lu) và hệ số thấm dùng trong mô hình.')
    p = tim_doan(body, 'Khi mực nước hồ thấp (+74,0 ÷ +74,7 m), lưu lượng thấm rất nhỏ')
    thay(p, 'phù hợp với hiện tượng thấm quan sát được ở mái gia cố hạ lưu khi mực nước hồ cao;',
         'tức mô hình thể hiện được cơ chế thấm ra trên mái khi mực nước lũ cao nhưng chưa tái hiện được ngưỡng xuất '
         'hiện thấm quan sát được từ khoảng +74,0 m (ở +74,0 ÷ +74,7 m mô hình chỉ cho thấm ra ở vùng trũng xa, x > 60 '
         'm) – sai khác cho thấy vai trò của dòng thấm tập trung theo mặt tiếp xúc đất – bê tông và hiệu ứng ba chiều '
         'mà mô hình phẳng không mô tả;')
    p = tim_doan(body, 'Gradien cửa ra khi chưa có màng (0,29 ÷ 0,34) nhỏ hơn gradien tới hạn của lớp 5')
    thay(p, 'như vậy theo mô hình phẳng, đất nền mang tràn chưa bị xói ngầm.',
         'như vậy theo mô hình phẳng, đất nền mang tràn chưa có nguy cơ mất ổn định thấm dạng đẩy trồi do lực thấm; '
         'phép kiểm tra này không bao quát xói tiếp xúc và sự di chuyển hạt mịn (cần cấp phối hạt và kiểm tra theo '
         'TCVN 8216:2018 mục 11.2.3).')
    thay(p, 'chứng tỏ dòng thấm tập trung', 'gợi ý có dòng thấm tập trung')
    thay(p, 'Điều này khẳng định sự cần thiết của màng', 'Điều này cho thấy sự cần thiết của màng')
    p = tim_doan(body, 'Tính thấm của lớp II quyết định chiều sâu hợp lý.')
    thay(p, 'Nếu KII ≤ 1·10-5 cm/s, chỉ cần cắm màng vào lớp II khoảng 1,0 ÷ 1,5 m.',
         'Nếu K_{II} ≤ 1·10^{-5} cm/s, trong mô hình chỉ cần cắm màng vào lớp II khoảng 1,0 ÷ 1,5 m là đạt hiệu quả gần '
         'tối đa – đây là kết quả độ nhạy của mô hình phẳng, không phải chiều sâu thiết kế (phương án đề xuất ở mục 8.6 '
         'cắm 3 m, chờ thí nghiệm ép nước).')
    p = tim_doan(body, 'Đoạn còn lại (khoảng 35 m) có thể giảm còn 2 hàng kề nhau A, B')
    p2 = doan_moi('Điều kiện để chuyển phương án hợp lý hóa thành phương án chọn – cả ba phải thỏa: (1) phụt thí '
                  'nghiệm cho màng 2 hàng liên tục, K ≤ 1·10^{-5} cm/s qua khoan kiểm tra, ép nước; (2) ép nước lớp II '
                  'cho q ≤ 1 Lu và tính lại với K đo được cho thấy gradien cửa ra, đường thấm vòng qua đầu và chân màng '
                  'đạt yêu cầu; (3) nếu điều kiện (1) hoặc (2) không đạt thì giữ 3 hàng trên toàn tuyến như TKCS (phương '
                  'án dự phòng). Mức giảm lưu lượng 47% so với 54% không đồng nghĩa với tương đương về an toàn chống xói '
                  '– vì vậy đoạn sát tràn (dòng thấm tập trung theo mặt tiếp xúc) luôn giữ 3 hàng và phụt tiếp xúc;', p)
    chen_sau(p, p2)
    p = tim_doan(body, 'Cao trình đáy màng xác định theo TCVN 8216:2018 mục 9.2.2.9')
    thay(p, 'Cao trình đáy màng xác định theo TCVN 8216:2018 mục 9.2.2.9 trên cơ sở thí nghiệm ép nước tại các hố phụt '
            'thí nghiệm: nếu lớp II có lượng mất nước q ≤ 5 Lu ngay từ mặt lớp thì đáy màng thấp hơn mặt lớp II 3,0 m '
            '(+69,4 m tại mặt cắt đại diện, so với +68,3 m của TKCS), nhưng không sâu hơn đáy màng TKCS;',
         'Cao trình đáy màng chọn trên cơ sở thí nghiệm ép nước tại các hố phụt thí nghiệm, trong giới hạn phạm vi '
         'khoan của TCVN 8216:2018 mục 9.2.2.9: phương án thử tại mặt cắt đại diện là đáy màng thấp hơn mặt lớp II '
         '3,0 m (+69,4 m, so với +68,3 m của TKCS) nếu lớp II có lượng mất nước q ≤ 5 Lu ngay từ mặt lớp, nhưng không '
         'sâu hơn đáy màng TKCS – đây là phương án được đánh giá bằng tính toán (Bảng 8.8), không phải cao trình bắt '
         'buộc của tiêu chuẩn;')
    p = tim_doan(body, 'Quan trắc sau xử lý: bố trí 2 giếng đo áp phía hạ lưu màng')
    thay(p, 'bố trí 2 giếng đo áp phía hạ lưu màng', 'bố trí 2 giếng đo áp phía hạ lưu màng (bổ sung ngoài 14 giếng '
                                                       'quan trắc thân đập – Bảng 7.1, Bảng 7.3)')
    p = tim_doan(body, 'Mô hình thấm phần tử hữu hạn tái hiện được hiện tượng thấm ra trên mái hạ lưu')
    thay(p, 'Mô hình thấm phần tử hữu hạn tái hiện được hiện tượng thấm ra trên mái hạ lưu mang trái tràn sự cố khi mực '
            'nước hồ cao, xác nhận lớp I',
         'Mô hình thấm phần tử hữu hạn hai chiều cho thấy cơ chế thấm ra trên mái hạ lưu mang trái tràn sự cố khi mực '
         'nước lũ cao (chưa tái hiện được ngưỡng xuất hiện thấm quan sát từ khoảng +74,0 m – hiệu ứng tiếp xúc đất – '
         'bê tông và dòng thấm ba chiều cần được khảo sát, hiệu chỉnh trước khi dùng kết quả để chốt thiết kế), với '
         'lớp I')
    p = tim_doan(body, 'Có thể hợp lý hóa thiết kế (3 hàng ở đoạn 15 m sát tràn')
    thay(p, 'để giảm khoảng 28% khối lượng khoan, phụt mà hiệu quả chỉ giảm khoảng 6 điểm phần trăm,',
         'để giảm 27,85% chiều dài khoan và 29,01% chiều dài phụt (tỷ lệ khối lượng, chưa phải tỷ lệ chi phí) mà hiệu '
         'quả chỉ giảm khoảng 6 điểm phần trăm,')
    thay(p, 'với điều kiện phụt thí nghiệm khẳng định K màng ≤ 1·10-5 cm/s và ép nước cho lớp II q ≤ 1 Lu;',
         'với ba điều kiện ở mục 8.6 (màng 2 hàng đạt K ≤ 1·10^{-5} cm/s; lớp II q ≤ 1 Lu và kiểm tra lại gradien, '
         'đường thấm vòng; nếu không đạt thì giữ 3 hàng);')
    ghi('Chương 8: "tái hiện" → "cho thấy cơ chế, chưa tái hiện ngưỡng +74,0 m" (A5); "chưa bị xói ngầm" → chỉ loại '
        'trừ mất ổn định thấm dạng đẩy trồi, chưa bao quát xói tiếp xúc (A6); "+3 m" là giới hạn phạm vi khoan, +69,4 m '
        'là phương án thử; 1 ÷ 1,5 m là kết quả độ nhạy; Lu – K gần đúng (A7); ba điều kiện chuyển phương án hợp lý '
        'hóa thành phương án chọn; 27,85%/29,01% là tỷ lệ khối lượng (A8); 2 giếng ngoài 14 giếng (B10).')


# =========================================================================== KẾT LUẬN, KIẾN NGHỊ
def ket_luan(d, body, KQ):
    ne = KQ['huong']['NE']
    cn = KQ['cap_nuoc']['TH1A|khong_Qtt']
    neo = KQ['neo']
    L = KQ['luoi']
    M = KQ['mua']
    dp = DU_PHONG_LUN
    p = tim_doan(body, 'Rà soát và hiệu chỉnh số liệu cơ bản:')
    dat_noi_dung(p, 'Số liệu cơ bản được cập nhật, thống nhất trước khi tính toán: quan hệ Z ~ F ~ W theo bình đồ 2022 '
                    '(diện tích theo km^{2}, dung tích tính lại từ diện tích); bề rộng tràn chính hiện trạng 21,6 m '
                    '(9 × 2,4 m) thay cho bề rộng thiết kế ban đầu 25 m; lũ tính theo tung độ đường quá trình; hệ số an '
                    'toàn [K] = 1,17 cho tổ hợp đặc biệt của đập cấp II (TCVN 8216:2018); vận tốc gió tính sóng quy đổi '
                    'về mặt nước và xét các hướng gió thổi vào mái thượng lưu (TCVN 8421:2010); kiểm tra trượt kết cấu bê '
                    'tông trên nền đá kể hệ số điều kiện làm việc γ_{c} = 0,95 (QCVN 04-05:2022 Bảng B.2).')
    p = tim_doan(body, 'Hồ giữ nguyên MNDBT +74,70 m, MNC +67,50 m;')
    them_cuoi(p, ' Theo quy tắc ưu tiên (cắt giảm tưới trước), cấp nước sinh hoạt đạt mức bảo đảm '
                 + so(cn['P_sinh_hoat'], 1) + '% (≥ 90%).')
    p = tim_doan(body, 'Với tràn hiện trạng, MNLTK = +77,63 m, MNLKT = +78,54 m (lũ BCNCKT);')
    thay(p, '(Zmax = +78,97 m, gió Đông Bắc)', '(Z_{max} = +' + so(ne['Z2']) + ' m, gió Đông Bắc)')
    thay(p, '(cao hơn TKCS 0,10 m; là cao trình tối thiểu sau lún, dự phòng lún 2 cm).',
         '(cao hơn TKCS 0,10 m; là cao trình thiết kế sau lún, thi công cộng dự phòng lún ' + so(dp) + ' m ≥ 0,5%H).')
    p = tim_doan(body, 'Thân, nền đập bảo đảm an toàn thấm')
    thay(p, 'Thân, nền đập bảo đảm an toàn thấm', 'Tại mặt cắt lòng sông đại diện, với các tổ hợp và chỉ tiêu cơ lý đã '
         'nêu (một số chỉ tiêu cường độ lấy theo trị tiêu chuẩn do hồ sơ thiếu trị tính toán), thân, nền đập bảo đảm an '
         'toàn thấm')
    thay(p, 'kết quả Bishop được kiểm tra chéo bằng Spencer (chênh không quá 0,2%), lời giải thấm hội tụ theo lưới (q '
            'thay đổi không quá 0,2%);',
         'tìm lại cung trượt bằng lưới lồng nhau và tối ưu hóa cho hệ số ổn định nhỏ hơn '
         + so(min(v['giam_pt'] for v in L.values()), 2) + ' ÷ ' + so(max(v['giam_pt'] for v in L.values()), 2)
         + '%, phương pháp Spencer nhỏ hơn Bishop không quá 0,5% (riêng '
         'TH8: ' + so(100 * (1 - L['TH8']['spencer'] / L['TH8']['K_toi_uu']), 1) + '%); lời giải thấm hội tụ theo lưới '
         '(q thay đổi khoảng 0,24%, gradien cục bộ hội tụ về ≈ 1,09 < 2,0); xét bất đẳng hướng k_{h}/k_{v} = 5 hệ số '
         'ổn định giảm không quá 2,4% – vẫn đạt;')
    p = tim_doan(body, 'Hệ thống rãnh chia mái thành ô 4 × 4 m giảm ứng suất tiếp do dòng chảy tràn')
    thay(p, 'từ ≈ 6,6 Pa xuống ≈ 1,8 Pa (1,4 ÷ 2,2 Pa tùy cường độ mưa)',
         'từ ≈ 6,6 Pa xuống ≈ 1,8 ÷ 2,2 Pa (theo cạnh và đường chéo ô; 1,4 ÷ 2,7 Pa tùy cường độ mưa)')
    p = tim_doan(body, 'Tràn chính đủ khả năng tháo, tường bên dốc nước đủ cao')
    thay(p, 'ngưỡng tràn chỉ ổn định khi có lực dính c ≥ 9,1 kPa trên mặt tiếp xúc với nền đá – đồ án bổ sung neo D25 '
            'thụ động vào nền (lực kéo cần 40 ÷ 57 kN/thanh, không quá 33% khả năng thiết kế; cần thí nghiệm kéo nhổ xác '
            'nhận).',
         'ngưỡng tràn chỉ ổn định khi có lực dính c ≥ ' + so(neo['MNLKT']['c_can'], 1) + ' kPa trên mặt tiếp xúc với '
         'nền đá ([K] = 1,21 / 1,09 với γ_{c} = 0,95) – đồ án đề xuất neo D25 thụ động vào nền (lực kéo cần '
         + so(neo['MNLTK']['N_can'], 0) + ' ÷ ' + so(neo['MNLKT']['N_can'], 0) + ' kN/thanh, không quá '
         + so(100 * neo['MNLKT']['N_can'] / 171.8, 0) + '% khả năng thiết kế) – phương án có điều kiện, chốt sau thí '
         'nghiệm kéo nhổ và hồ sơ hoàn công.')
    thay(p, 'trục vít dài 6 m cần bổ sung giá đỡ trung gian để bảo đảm ổn định khi đóng cưỡng bức.',
         'trục vít dài 6 m cần bổ sung hai giá đỡ dẫn hướng trung gian (đoạn L ≤ 2 m, μ = 1,0) để bảo đảm ổn định khi '
         'đóng cưỡng bức.')
    thay(p, 'cho cốt thép yêu cầu 8 ÷ 16 cm2/m', 'cho cốt thép yêu cầu 8 ÷ 17 cm^{2}/m')
    p = tim_doan(body, 'Chuyên đề: mô hình thấm PTHH cho thấy màng khoan phụt thiết kế')
    thay(p, 'giảm khoảng 28% khối lượng khoan, phụt mà hiệu quả chỉ giảm khoảng 6 điểm phần trăm, trên cơ sở phụt thí '
            'nghiệm;',
         'giảm 27,85% chiều dài khoan và 29,01% chiều dài phụt mà hiệu quả chỉ giảm khoảng 6 điểm phần trăm, với ba '
         'điều kiện kiểm tra được bằng phụt thí nghiệm, ép nước (mục 8.6);')
    kq = {(r['nguon'], r['P'], round(100 * r['k'])): r for r in M['dieu_tiet']}
    p2 = doan_moi('Lượng mưa lưu vực sau khi công trình vận hành (2000 ÷ 2024) không có xu thế tăng có ý nghĩa thống kê '
                  '(mưa năm ' + so(M['mua_nam']['thay_doi_pt'], 1) + '%, mưa 1 ngày lớn nhất '
                  + so(M['x1max']['thay_doi_pt'], 1) + '%) nhưng mưa cực trị phân tán hơn; lũ thiết kế cập nhật lớn hơn '
                  'thiết kế năm 1996 khoảng 8 ÷ 9% và MNLTK hiện nay (+77,63 m) cao hơn MNLTK thiết kế ban đầu (+77,24 '
                  'm) 0,39 m. Với tràn hiện trạng, khi lũ tăng 30% mực nước lũ thiết kế vẫn thấp hơn đỉnh đập (Z_{max} ≤ +'
                  + so(kq[('BCNCKT', '1%', 30)]['Zmax']) + ' m), nhưng dự trữ theo các điều kiện cao trình chỉ tương ứng '
                  'mức tăng lũ khoảng 9 ÷ 12% (lũ kiểm tra chạm đỉnh đập +78,85 m; sóng gió Đông Bắc khi có lũ thiết kế '
                  'chạm đỉnh tường +79,40 m); các giải pháp sửa chữa nâng mức an toàn của đập và '
                  'công trình tháo lũ (Bảng 4.9, Bảng 5.7) chứ không làm tăng khả năng tháo.', p)
    chen_sau(p, p2)
    # kiến nghị
    p = tim_doan(body, 'Bổ sung khảo sát: hồ sơ hoàn công tràn chính, cống lấy nước')
    them_cuoi(p, ' Đo quan hệ Q ~ Z của tràn chính khi có lũ lớn để xác định hệ số lưu lượng m (đỉnh +78,85 m chỉ đạt '
                 'khi m ≥ ' + so(KQ['m']['m_gh'], 3) + '); quan trắc lún mặt đập để kiểm chứng dự phòng lún; phân tích '
                 'thành phần hóa học của nước hồ phục vụ kiểm tra khe nứt kết cấu BTCT (TCVN 4116:2023 mục 9.2).')
    p = tim_doan(body, 'Thực hiện phụt thí nghiệm trước, điều chỉnh số hàng, chiều sâu màng')
    p3 = doan_moi('Bổ sung vào quy trình vận hành quy định xả trước theo dự báo mưa lớn (hạ hồ xuống khoảng +73,8 m '
                  'trong 1 ÷ 2 ngày qua tràn sự cố) và cập nhật tính toán thủy văn định kỳ khoảng 5 năm (QCVN 04-05:2022 '
                  'điều 2.3.2.2, 2.6.3.3); nếu lũ kiểm tra cập nhật tăng quá khoảng 9% thì cần nghiên cứu mở rộng tràn '
                  'hoặc nâng đỉnh đập.', p)
    chen_sau(p, p3)
    p = tim_doan(body, 'Bổ sung vào quy trình vận hành: kiểm tra cửa van tràn sự cố')
    thay(p, 'lún mặt đập vượt 2 cm', 'lún mặt đập vượt dự phòng lún ' + so(dp) + ' m')
    ghi('Kết luận, kiến nghị: cập nhật theo các nội dung trên (γc = 0,95; Vw; sinh hoạt %s%%; Z2 = +%s; dự phòng lún '
        '%s m; giới hạn phạm vi kết luận an toàn; kiểm tra chéo ổn định; neo có điều kiện; 2 gối trục vít; 8 ÷ 17 cm²/m; '
        '27,85/29,01%%); thêm ý kết luận về mưa – lũ – khả năng tháo; kiến nghị đo Q ~ Z, quan trắc lún, phân tích nước, '
        'xả trước theo dự báo, cập nhật thủy văn.' % (so(cn['P_sinh_hoat'], 1), so(ne['Z2']), so(dp)))


# =========================================================================== TÀI LIỆU THAM KHẢO
def tai_lieu(d, body):
    for i in range(41, 30, -1):
        n = C.doi_tham_chieu(body, '[%d]' % i, '[%d]' % (i + 1))
        if n == 0:
            raise ValueError('Không thấy [%d]' % i)
    n = 0
    for p in body.iter(qn('w:p')):
        k = van_ban(p).count('[@31]')
        if k:
            thay(p, '[@31]', '[31]', lan=k)
            n += k
    assert n >= 2, n
    tbl = None
    for t in body.iter(qn('w:tbl')):
        if van_ban(hang(t)[0].findall(qn('w:tc'))[0]).strip() == '[1]':
            tbl = t
    r = tim_hang(tbl, '[30]', dung=True)
    them_hang(tbl, r, ['[31]', 'Lê Thị Hải Yến, Ngô Lê An, Nguyễn Thị Thu Hà và Ngô Lê Long, “Đánh giá tác động của biến '
                               'đổi khí hậu đến mưa lớn vùng Nam Trung Bộ và Tây Nguyên của Việt Nam theo các mô hình khí '
                               'hậu toàn cầu khác nhau,” Tạp chí Khoa học Kỹ thuật Thủy lợi và Môi trường, số 57, 2017.'])
    ghi('Tài liệu tham khảo: thêm [31] Lê Thị Hải Yến và cs. (2017) – trích ở mục 4.6; đánh số lại [31] ÷ [41] → [32] '
        '÷ [42] trong lời văn và danh mục (%d chỗ dùng tài liệu mới).' % n)


# =========================================================================== PHỤ LỤC
def phu_luc(d, body, KQ):
    an, L, M = KQ['anis'], KQ['luoi'], KQ['mua']
    cap, tbl = tim_bang(body, 'Bảng PL5')
    r = tim_hang(tbl, 'cao_trinh_dinh.py')
    dat_hang(tbl, r, [None, 'Tính sóng, sóng leo, nước dềnh và cao trình đỉnh đập theo TCVN 8216:2018, TCVN 8421:2010 '
                            '(quy đổi gió A.3.3, hệ số k_{α} Bảng 9)'])
    r = tim_hang(tbl, 'bo_sung_giai_tich.py')
    for ten, nd in reversed([
            ('kt_cap_nuoc_bo_sung.py', 'Mức bảo đảm theo từng mục tiêu (TCVN 10778:2024 mục 8.2.2.2), số tháng mực nước '
                                       'dưới +67,94 m (mục 2.2.3, 6.3)'),
            ('kt_mua_lu_bien_doi.py', 'Biến động mưa từ 2000 (Mann – Kendall, Sen, Welch), điều tiết lũ khi lũ tăng, mức '
                                      'tăng giới hạn, mực nước đón lũ (mục 4.6)'),
            ('kt_bo_sung_ket_cau.py', 'Ổn định trục vít theo sơ đồ gối, neo ngưỡng (phía bê tông, nhổ khối đá), ứng suất '
                                      'tiếp theo đường chéo ô rãnh'),
            ('kt_kiem_chung_12_10.py', 'Gradien lớn nhất và vị trí, bất đẳng hướng k_{h}/k_{v} = 5, tìm lại cung trượt '
                                       '(miền mở rộng, lưới lồng nhau, Nelder – Mead, Spencer) – Bảng PL10, PL11')]):
        them_hang(tbl, r, [ten, nd])
    # PL6
    cap, tbl = tim_bang(body, 'Bảng PL6')
    r = tim_hang(tbl, 'K Bishop – Spencer, cùng cung trượt')
    ths = ['TH%d' % i for i in range(1, 11)]
    dat_hang(tbl, r, ['K Bishop / Spencer trên cung trượt tìm lại (sau SC, TH1 ÷ TH10)',
                      '; '.join(so(L[t]['K_toi_uu'], 3) for t in ths), '; '.join(so(L[t]['spencer'], 3) for t in ths),
                      '≤ 0,5%; TH8: ' + so(100 * (1 - L['TH8']['spencer'] / L['TH8']['K_toi_uu']), 1) + '%',
                      'Spencer [23]; TH8 vẫn > [K] = 1,30'])
    r = tim_hang(tbl, 'K khi tìm lại trên lưới tâm mịn gấp đôi')
    dat_hang(tbl, r, ['Tìm lại cung trượt: miền mở rộng (đáy cung tới +36 m), lưới lồng nhau, Nelder – Mead (10 điểm '
                      'xuất phát)', 'K Bảng 5.2', 'K nhỏ nhất tìm lại (Bảng PL11)',
                      'giảm ' + so(min(v['giam_pt'] for v in L.values()), 2) + ' ÷ '
                      + so(max(v['giam_pt'] for v in L.values()), 2) + '%', 'Mọi trường hợp vẫn ≥ [K]'])
    r = tim_hang(tbl, 'Hội tụ lưới PTHH')
    dat_hang(tbl, r, [None, None, None, '≈ 0,24% (< 0,3%)', 'Bảng PL8; J tiếp xúc 95% thay đổi 0,94 ÷ 1,09'])
    # PL7
    cap, tbl = tim_bang(body, 'Bảng PL7')
    mm = KQ['m']
    r = tim_hang(tbl, 'Hệ số lưu lượng tràn m')
    dat_hang(tbl, r, [None, '0,467 (BCNCKT); 0,40 (riêng tràn chính)', 'MNLKT +78,49 / +' + so(mm['0.40'][1]) + ' m',
                      'Đỉnh +78,85 m đạt khi m ≥ ' + so(mm['m_gh'], 3)])
    ne, nw = KQ['huong']['NE'], KQ['huong']['NW']
    r = tim_hang(tbl, 'Hướng gió (Bắc)')
    dat_hang(tbl, r, [None, 'NE, NW; đà gió lớn nhất; V_{w} = k_{l}·V_{l}', 'Z_{2} = +' + so(ne['Z2']) + ' m (NE); +'
                      + so(nw['Z2']) + ' m (NW)', None])
    r = tim_hang(tbl, 'Tham số hàm thấm α')
    rr = them_hang(tbl, r, ['Bất đẳng hướng thấm k_{h}/k_{v} (1)', '5 (TCVN 8216:2018, 11.2.2)',
                            'q tăng 44 ÷ 61%; K giảm ≤ 2,4% (Bảng PL10)', 'Vẫn đạt'])
    r = tim_hang(tbl, 'Tỷ lệ mưa 30 phút / ngày')
    xm = KQ['kc']['xoi_mai']
    dat_hang(tbl, r, [None, None, 'τ (ô 4 × 4 m, L = 4 ÷ 5,66 m) ' + so(xm['20%_L4.00']['tau'], 1) + ' ÷ '
                      + so(xm['40%_L5.66']['tau'], 1) + ' Pa; τ (L = 34 m) ' + so(xm['20%_L34.00']['tau'], 1) + ' ÷ '
                      + so(xm['40%_L34.00']['tau'], 1) + ' Pa', None])
    r = tim_hang(tbl, 'Mức bảo đảm cấp nước')
    them_hang(tbl, r, ['Lũ đến tăng k (0)', '10 ÷ 30%', 'Z_{max} P = 0,2%: +78,89 ÷ +79,56 m (BCNCKT)',
                       'Vượt đỉnh đập khi k > 9% (mục 4.6)'])
    p = tim_doan(body, 'Phụ lục này tổng hợp kết quả kiểm chứng các chương trình tính (Bảng PL6)')
    thay(p, '(Bảng PL6) và phân tích độ nhạy các giả thiết (Bảng PL7 ÷ Bảng PL9).',
         '(Bảng PL6, Bảng PL11) và phân tích độ nhạy các giả thiết (Bảng PL7 ÷ Bảng PL10).')
    # PL10, PL11 – sau Bảng PL9
    cap9, tbl9 = tim_bang(body, 'Bảng PL9')
    neo_chen = tbl9.getnext()
    hang10 = []
    for pha, tp in (('HT', 'Hiện trạng'), ('SC', 'Sau sửa chữa')):
        for k, tk in (('MNDBT', 'MNDBT'), ('MNLTK', 'MNLTK'), ('MNLKT', 'MNLKT'), ('MNDBT_tac', 'MNDBT, thoát nước tắc')):
            a1, a5 = an['%s_%s_r1' % (pha, k)], an['%s_%s_r5' % (pha, k)]
            hang10.append([tp + ' – ' + tk, so(a1['q'] * 1e6, 2) + ' / ' + so(a5['q'] * 1e6, 2),
                           so(a1['z_tim']) + ' / ' + so(a5['z_tim']), so(a1['J_tiep95']) + ' / ' + so(a5['J_tiep95']),
                           so(a1['J_ra95'], 3) + ' / ' + so(a5['J_ra95'], 3)])
    c10, t10, k10 = bang_moi('Bảng PL10 Ảnh hưởng của bất đẳng hướng thấm k_{h}/k_{v} = 5 (giá trị k_{h}/k_{v} = 1 / = 5)',
                             ['Trường hợp', 'q (10^{-6} m^{3}/s/m)', 'Z bão hòa tại tim (m)', 'J tiếp xúc 95%',
                              'J cửa ra 95%'], hang10, [2700, 1700, 1700, 1450, 1450], cap9)
    p10 = doan_moi('Hệ số thấm theo phương ngang lấy bằng 5 lần phương đứng cho các lớp đất đắp và lớp tàn tích (TCVN '
                   '8216:2018 mục 11.2.2), giữ nguyên hệ số thấm tương đương k = (k_{h}·k_{v})^{1/2}. Hệ số ổn định sau '
                   'sửa chữa TH1 ÷ TH10 khi k_{h}/k_{v} = 5: ' + '; '.join(so(an['SC_TH%d_r5_K' % i], 3) for i in range(1, 11))
                   + ' (so với ' + '; '.join(so(an['SC_TH%d_r1_K' % i], 3) for i in range(1, 11)) + ').', p)
    hang11 = []
    kcp = {1: 1.30, 2: 1.30, 3: 1.17, 4: 1.17, 5: 1.17, 6: 1.17, 7: 1.30, 8: 1.30, 9: 1.17, 10: 1.17}
    for i in range(1, 11):
        v = L['TH%d' % i]
        hang11.append(['TH%d' % i, so(v['K_co_so'], 3), so(v['K_mo_rong'], 3), so(v['K_luoi_long'], 3),
                       so(v['K_toi_uu'], 3) + ' (−' + so(v['giam_pt'], 2) + '%)',
                       so(v['spencer'], 3) + ' (θ = ' + so(v['spencer_theta'], 1) + '°)', so(kcp[i], 2)])
    c11, t11, k11 = bang_moi('Bảng PL11 Kiểm tra tìm cung trượt nguy hiểm nhất sau sửa chữa (Bishop; Spencer trên cung '
                             'tìm lại)',
                             ['TH', 'K cơ sở (Bảng 5.2)', 'Miền mở rộng', 'Lưới lồng nhau', 'Nelder – Mead',
                              'Spencer', '[K]'], hang11, [700, 1350, 1350, 1350, 1700, 1700, 850], cap9)
    p11 = doan_moi('Miền mở rộng: đáy cung tới cao trình +36 m (sâu hơn đáy mô hình cũ khoảng 7 m); lưới lồng nhau: chứa '
                   'toàn bộ nút tâm, bán kính của lưới cơ sở và chia mịn gấp đôi; Nelder – Mead: tối ưu tọa độ tâm và bán '
                   'kính từ 10 điểm xuất phát tốt nhất (chương trình kt_kiem_chung_12_10.py). K cơ sở là kết quả sau bước '
                   'tinh chỉnh quanh tâm tốt nhất của lưới cơ sở; lưới lồng nhau chứa mọi nút của lưới cơ sở nhưng không '
                   'chứa điểm đã tinh chỉnh nên có thể cho K lớn hơn K cơ sở (không quá 0,3%); giá trị nhỏ nhất của các '
                   'cách tìm (cột Nelder – Mead) luôn nhỏ hơn K cơ sở và được dùng để kết luận.', p)
    chen_sau(neo_chen, c10, t10, k10, p10, c11, t11, k11, p11)
    # Phụ lục 7 – số liệu mưa
    h_pl6 = tim_doan(body, 'Phụ lục 6. Kiểm chứng chương trình tính và phân tích độ nhạy')
    h7 = doan_moi('Phụ lục 7. Số liệu mưa lưu vực hồ Cà Giây', h_pl6)
    mua = KQ['mua_so_lieu']
    nam = sorted(mua['mua_nam'].keys(), key=int)
    nua = (len(nam) + 1) // 2
    hang12 = []
    for i in range(nua):
        dong = []
        for j in (i, i + nua):
            if j < len(nam):
                y = nam[j]
                x1 = mua['x1max'].get(y)
                dong += [y, so(mua['mua_nam'][y][12], 1), so(x1, 1) if x1 is not None else '–']
            else:
                dong += ['', '', '']
        hang12.append(dong)
    c12, t12, k12 = bang_moi('Bảng PL12 Lượng mưa năm X_{năm} và lượng mưa 1 ngày lớn nhất X_{1max} của lưu vực (mm)',
                             ['Năm', 'X_{năm}', 'X_{1max}', 'Năm', 'X_{năm}', 'X_{1max}'], hang12,
                             [1000, 1500, 1500, 1000, 1500, 1500], cap9, can=['center'] * 6)
    p12 = doan_moi('Nguồn: lượng mưa năm – QTVH [1] Bảng 2-9 (quy về lưu vực từ các trạm Đại Ninh, Sông Lũy); X_{1max} – '
                   'báo cáo KTTV [7] Bảng 3-13 (1978 ÷ 2020). Kết quả phân tích ở mục 4.6.1.', p)
    chen_sau(p11, h7, c12, t12, k12, p12)
    ghi('Phụ lục: Bảng PL5 thêm 4 chương trình; Bảng PL6 (kiểm tra chéo Spencer TH1 ÷ TH10, tìm lại cung trượt, hội '
        'tụ 0,24%%), Bảng PL7 (m = 0,40; gió Vw; bất đẳng hướng; lũ tăng); thêm Bảng PL10 (bất đẳng hướng), Bảng PL11 '
        '(tìm cung trượt) vào Phụ lục 6; thêm Phụ lục 7 – Bảng PL12 số liệu mưa (%d năm).' % len(nam))


# =========================================================================== CHẠY
def tai_ket_qua(dir_tt):
    sys.path.insert(0, os.path.abspath(dir_tt))
    import so_lieu as S
    import cao_trinh_dinh as CT
    import kt_mua_lu_bien_doi as KM
    kq = os.path.join(dir_tt, 'ket_qua')

    def doc(ten):
        return json.load(open(os.path.join(kq, ten), encoding='utf-8'))
    bs = doc('bo_sung_giai_tich.json')
    KQ = dict(bs=bs, cap_nuoc=doc('kt_cap_nuoc_bo_sung.json'), kc=doc('kt_bo_sung_ket_cau.json'),
              mua=doc('kt_mua_lu_bien_doi.json'), gradien=doc('kiem_chung_gradien.json'),
              anis=doc('kiem_chung_bat_dang_huong.json'), luoi=doc('kiem_chung_luoi_tam.json'))
    KQ['mua_so_lieu'] = json.load(open(os.path.join(dir_tt, 'du_lieu', 'mua_luu_vuc.json'), encoding='utf-8'))
    V4 = S.GIO[4][S.GIO_HUONG.index('N')]
    V50 = S.GIO[50][S.GIO_HUONG.index('N')]
    ct = CT.cao_trinh(S.MNDBT, 77.63, 78.54, V4, V50, 2970, 3025, 3047)
    KQ['ct'] = ct
    huong = {r['huong']: r for r in bs['gio']['bang'] if 'Z1' in r}
    KQ['huong'] = huong
    b1 = CT.cao_trinh(S.MNDBT, 77.43, 78.25, V4, V50, 2970, 3025, 3047)
    b3 = CT.cao_trinh(S.MNDBT, 77.50, 78.44, V4, V50, 2970, 3025, 3047)
    KQ['bang25'] = dict(bcnckt=(b1['Z1'], b1['Z2'], b1['Z3']), da=(ct['Z1'], ct['Z2'], ct['Z3']),
                        ne=(huong['NE']['Z1'], huong['NE']['Z2'], ct['Z3']), qtvh=(b3['Z1'], b3['Z2'], b3['Z3']))
    KQ['Z_sl_DBT'] = max(r['Z1'] - 0.70 for r in huong.values())
    mt = bs['do_nhay_m_tran_chinh']
    KQ['m'] = {'0.40': (mt['0.400']['MNLTK'], mt['0.400']['MNLKT']), 'm_gh': mt['m_gh']}
    KQ['neo'] = {k: dict(R_can=v['R_can'], N_can=v['luc_keo_can_1_thanh'], K_cp=v['K_cp'], c_can=v['c_can'])
                 for k, v in bs['neo']['chi_tiet'].items()}
    KQ['k_qtvh_dinh'] = KM.k_gioi_han(S.Q_QTVH['0.2%'], 78.85)
    z0 = KQ['mua']['Z0_don_lu']
    KQ['dV_don_lu'] = {k: S.V_of_Z(S.MNDBT) - S.V_of_Z(v) for k, v in z0.items()}
    t = 78.60 + DU_PHONG_LUN - 78.20
    a_dat = (t * 4.30 - 0.01 * 4.30 ** 2) + (t * 0.20 - 0.01 * (4.50 ** 2 - 4.30 ** 2)) + ((0.21 + DU_PHONG_LUN) * 0.20 - 0.20 ** 2 / 6)
    KQ['v_dat_dinh'] = round(986 * a_dat / 10.0) * 10.0
    KQ['v_dat_tkcs'] = round(986 * 4.70 * 0.30 / 10.0) * 10.0
    return KQ



# =========================================================================== CỬA VAN TRÀN CHÍNH (13/10)
def cua_van_tran_chinh(body):
    """Tràn chính có 9 cửa van (QTVH): giả thiết mở hết; trường hợp 1 ÷ 2 cửa không mở được; kiểm tra trước mùa lũ."""
    p = tim_doan(body, 'Kết quả điều tiết lũ cho trong Bảng 2.3')
    them_cuoi(p, ' Các kết quả trên giả thiết 9 cửa van tràn chính được mở hết từ đầu mùa lũ như QTVH quy định. Nếu có '
                 'cửa không mở được (kẹt, mất điện, vật cản), bề rộng tràn chính giảm còn 19,2 m (1 cửa) hoặc 16,8 m '
                 '(2 cửa): với lũ thiết kế MNLTK = +77,69 m (+77,75 m), với lũ kiểm tra MNLKT = +78,62 m (+78,70 m) – '
                 'cao hơn điều kiện MNLKT + 0,3 m ≤ +78,85 m 0,07 ÷ 0,15 m nhưng vẫn thấp hơn đỉnh đập; các trường hợp '
                 'này được phân loại cùng các tình huống kẹt van ở Bảng 4.2.')
    cap, tbl = tim_bang(body, 'Bảng 4.2')
    r = tim_hang(tbl, 'P = 1% + kẹt van tràn sự cố')
    rr = them_hang(tbl, r, ['P = 1%, 1 (2) cửa van tràn chính không mở', '+77,69 (+77,75)', 'thấp hơn 1,16 (1,10) m',
                            'thấp hơn 1,71 (1,65) m', 'Tổ hợp đặc biệt (CHÚ THÍCH điều 2.6.3.10)',
                            'Đỉnh đập ≥ Zmax + 0,3 m: đạt (dư 0,86 / 0,80 m)'])
    r = tim_hang(tbl, 'P = 0,2% + kẹt van (BCNCKT / QTVH)')
    them_hang(tbl, r, ['P = 0,2%, 1 (2) cửa van tràn chính không mở', '+78,62 (+78,70)', 'thấp hơn 0,23 (0,15) m',
                       'thấp hơn 0,78 (0,70) m', 'Ngoài các trường hợp bắt buộc (điều 2.6.3.10b) – đánh giá rủi ro',
                       'Không tràn đỉnh; mở hết, kiểm tra 9 cửa trước mùa lũ (QTVH)'])
    p = tim_doan(body, 'Bổ sung vào quy trình vận hành: kiểm tra cửa van tràn sự cố')
    thay(p, 'Bổ sung vào quy trình vận hành: kiểm tra cửa van tràn sự cố và nguồn điện dự phòng trước mùa lũ;',
         'Bổ sung vào quy trình vận hành: kiểm tra cửa van tràn sự cố và nguồn điện dự phòng trước mùa lũ; kiểm tra, '
         'mở hết 9 cửa van tràn chính trước mùa lũ và không giữ cửa đóng để tích nước khi có dự báo lũ;')
    p = tim_doan(body, 'Tràn chính đủ khả năng tháo, tường bên dốc nước đủ cao')
    thay(p, 'Tràn chính đủ khả năng tháo,', 'Tràn chính (9 khoang có cửa van, mở hết trong mùa lũ) đủ khả năng tháo,')
    ghi('Tràn chính có 9 cửa van (QTVH): mục 2.3 thêm trường hợp 1 ÷ 2 cửa không mở được (MNLTK +77,69 ÷ +77,75; '
        'MNLKT +78,62 ÷ +78,70 m), Bảng 4.2 thêm 2 dòng; kiến nghị quy trình vận hành kiểm tra, mở hết 9 cửa trước mùa lũ; '
        'mục 2.3.2: hệ số lưu lượng theo đúng QTVH (m = 0,40; ε = Σb/(Σb + Σd) = 0,82 – tỷ lệ bề rộng thông thủy).')


# =========================================================================== GÓP Ý BỔ SUNG 12/10 (tối)
def bo_sung_doi_chieu(body):
    """Giải thích K sau sửa chữa giảm nhẹ; chênh lệch mức bảo đảm cấp nước với QTVH; độ nhạy tổn thất thấm."""
    p = tim_doan(body, 'Tất cả các trường hợp đều thỏa mãn; hệ số ổn định nhỏ nhất của mái hạ lưu K = 1,349')
    them_cuoi(p, ' So với hiện trạng, hệ số ổn định sau sửa chữa giảm 0,9 ÷ 1,6% ở mọi trường hợp (Bảng 5.2) vì phần '
                 'đắp nâng đỉnh đập (dày 0,4 ÷ 0,5 m cùng tấm bê tông mặt đập) làm tăng tải trọng ở đầu các cung trượt '
                 'sâu, trong khi các hạng mục sửa chữa chủ yếu tác động lên lớp mặt và không thay đổi hệ thống chống thấm '
                 '– thoát nước của thân đập; mức giảm nhỏ, mọi trường hợp vẫn lớn hơn [K] ít nhất 15% (TH5: 1,349 so với '
                 '1,17; TH8: 1,496 so với 1,30). Hiệu quả của sửa chữa đối với ổn định thể hiện ở lớp mặt mái hạ lưu '
                 '(K = 0,52 → 4,41) và cao trình đỉnh đập (mục 5.6, Bảng 5.7).')
    p = tim_doan(body, 'Kết quả điều tiết (Bảng 2.2, Hình 2.1) cho thấy:')
    thay(p, 'Mức bảo đảm do đồ án tính thấp hơn kết quả của QTVH chủ yếu do đồ án kể tổn thất thấm kt = 1%/tháng và '
            'dùng công thức P = m/(n + 1); nếu tính theo P = m/n như QTVH thì TH1 đạt 40,9%, TH1A đạt 88,6%, TH2 đạt '
            '84,1% – kết luận không đổi.',
         'Mức bảo đảm do đồ án tính thấp hơn kết quả của QTVH do hai giả thiết: (i) công thức tần suất – nếu tính theo '
         'P = m/n như QTVH thì TH1 đạt 40,9%, TH1A đạt 88,6%, TH2 đạt 84,1%; (ii) tổn thất thấm lòng hồ – nếu lấy '
         'k_{t} = 0,5%/tháng (cận dưới ứng với địa chất lòng hồ tốt, TCVN 10778:2024 Bảng 1) thì số năm đủ nước của '
         'TH1A và TH2 là 42/44 và 39/44, chỉ chênh 1 năm so với QTVH (41/44 và 38/44 năm, tức 93% và 86,4% theo '
         'P = m/n). Với TH1, số năm đủ nước của đồ án (18/44) không phụ thuộc k_{t}, nên chênh lệch với QTVH (61%) '
         'xuất phát từ khác biệt về số liệu đầu vào hoặc cách tính mà hồ sơ chưa trình bày đủ chi tiết để đối chiếu; '
         'cả hai đều cho kết luận TH1 không đạt 85%. Phương án TH2 nằm sát ngưỡng: theo công thức (14) của TCVN '
         '10778:2024, 38/44 năm của QTVH tương ứng 84,4% (< 85%), còn đồ án cho 82,2% với k_{t} = 1%/tháng và 86,7% '
         'với k_{t} = 0,5%/tháng – vì vậy cần xác định tổn thất thấm từ số liệu quan trắc cân bằng nước của hồ trước '
         'khi kết luận về nhiệm vụ mở rộng (Bảng PL7).')
    p = tim_doan(body, 'Bổ sung khảo sát: hồ sơ hoàn công tràn chính, cống lấy nước')
    them_cuoi(p, ' Quan trắc cân bằng nước của hồ để xác định tổn thất thấm lòng hồ (mức bảo đảm của nhiệm vụ mở rộng '
                 'TH2 phụ thuộc giả thiết này – mục 2.2.3).')
    cap, tbl = tim_bang(body, 'Bảng PL7')
    r = tim_hang(tbl, 'Mức bảo đảm cấp nước')
    them_hang(tbl, r, ['Tổn thất thấm lòng hồ k_{t} (1%/tháng)', '0,5%; 0',
                       'TH1A: 93,3% (0,5% và 0); TH2: 86,7% (0,5%), 91,1% (0)',
                       'TH1A luôn đạt; TH2 đạt khi k_{t} ≤ 0,5%/tháng – cần quan trắc tổn thất'])
    ghi('Bổ sung (góp ý 12/10 tối): giải thích hệ số ổn định sau sửa chữa giảm 0,9 ÷ 1,6% (mục 5.3); phân tích chênh '
        'lệch mức bảo đảm với QTVH theo công thức tần suất và tổn thất thấm k_t, TH2 sát ngưỡng (mục 2.2.3, Bảng PL7 '
        'thêm dòng k_t); xu thế mưa theo độ dốc Sen (%/10 năm) toàn chuỗi và từ 2000 (mục 4.6.1).')


# =========================================================================== GIỌNG VĂN VỚI HỒ SƠ DỰ ÁN
def giong_van_ho_so(body):
    """So sánh với hồ sơ của đơn vị tư vấn theo hướng "cập nhật, thống nhất số liệu", không dùng từ đánh giá đúng/sai."""
    ds = [
        ('Quan hệ đặc tính lòng hồ được lập từ bình đồ khảo sát năm 2022.',
         'Khi kiểm tra lại bảng số liệu, đồ án nhận thấy hai điểm cần hiệu chỉnh: (i) đơn vị của cột diện tích được ghi là '
         '“ha” nhưng giá trị thực là km2 (một hồ dung tích 37 triệu m3 không thể có mặt thoáng khoảng 6 ha; chính biểu đồ '
         'trong TKCS cũng ghi F theo km2); (ii) từ cao trình +70 m trở lên, cột dung tích không khớp với cột diện tích (ví '
         'dụ từ +70 đến +71 m dung tích chỉ tăng 2,81 triệu m3 trong khi diện tích trung bình là 3,70 km2). Vì vậy, toàn '
         'bộ đồ án sử dụng dung tích tính lại',
         'Khi sử dụng bảng số liệu, đồ án thống nhất hai điểm: (i) cột diện tích được hiểu theo đơn vị km^{2} như biểu đồ '
         'đặc tính lòng hồ trong TKCS (bảng số ghi đơn vị “ha”); (ii) từ cao trình +70 m trở lên, dung tích được tính lại '
         'từ diện tích để quan hệ Z ~ F ~ W nhất quán (ví dụ từ +70 đến +71 m diện tích trung bình 3,70 km^{2}, trong bảng '
         'dung tích tăng 2,81 triệu m^{3}). Toàn bộ đồ án sử dụng dung tích tính lại'),
        ('Hai bộ số liệu (Bảng 1.5) có lưu lượng đỉnh',
         'Ngoài ra, tổng lượng ghi trong báo cáo BCNCKT (32,43 và 41,95 triệu m3) nhỏ hơn tổng lượng tích phân từ chính '
         'đường quá trình (36,27 và 47,69 triệu m3). Đồ án sử dụng trực tiếp',
         'Tổng lượng tích phân từ đường quá trình (36,27 và 47,69 triệu m^{3}) khác giá trị ghi trong bảng đặc trưng lũ '
         '(32,43 và 41,95 triệu m^{3}); đồ án sử dụng trực tiếp'),
        ('Tràn chính có 9 khoang × 2,4 m',
         'Lưu ý: hồ sơ BCNCKT năm 2022 mô tả tràn có tổng bề rộng 25 m chia 3 khoang (9 + 8 + 8 m) và dùng bề rộng này để '
         'tính điều tiết lũ, trong khi hồ sơ QTVH 2026 cập nhật hiện trạng là 9 khoang × 2,4 m = 21,6 m. Đồ án sử dụng số '
         'liệu hiện trạng 21,6 m.',
         'Hồ sơ BCNCKT năm 2022 tính điều tiết lũ với bề rộng tràn theo thiết kế ban đầu 25 m (3 khoang 9 + 8 + 8 m); hồ '
         'sơ QTVH 2026 cập nhật hiện trạng 9 khoang × 2,4 m = 21,6 m. Đồ án sử dụng số liệu hiện trạng cập nhật 21,6 m.'),
        ('Các bài toán được lập trình tính toán bằng ngôn ngữ Python',
         'để chủ động kiểm soát giả thiết và kiểm tra chéo kết quả của hồ sơ tư vấn:',
         'để chủ động kiểm soát giả thiết và đối chiếu với kết quả của hồ sơ dự án:'),
        ('Trình tự thực hiện và kiểm chứng kết quả',
         '(i) tính lại với đúng thông số của hồ sơ tư vấn rồi so sánh với kết quả đã công bố',
         '(i) tính lại với cùng thông số của hồ sơ dự án và so sánh với kết quả trong hồ sơ'),
        ('Lưu ý: hồ sơ BCNCKT dùng hệ số [K] = 1,13',
         'Lưu ý: hồ sơ BCNCKT dùng hệ số [K] = 1,13 cho các tổ hợp đặc biệt; giá trị này tương ứng công trình cấp III. Đồ '
         'án dùng [K] = 1,17 cho tổ hợp đặc biệt của đập cấp II.',
         'Đồ án dùng [K] = 1,17 cho tổ hợp đặc biệt của đập cấp II theo TCVN 8216:2018 (BCNCKT dùng 1,13; các hệ số ổn '
         'định tính được đều lớn hơn cả hai trị số nên kết luận đánh giá không đổi).'),
        ('Kết quả tính đường mặt nước cho trong Bảng 4.4',
         'Kết quả tính của tư vấn [26] cho độ vượt cao tường bên nhỏ hơn (0,16 m tại mặt cắt thứ ba với lũ kiểm tra) do '
         'dùng lưu lượng tràn chính 311 m3/s ứng với giả thiết tràn rộng 25 m.',
         'Phụ lục tính toán của BCNCKT [26] cho độ vượt cao tường bên nhỏ hơn (0,16 m tại mặt cắt thứ ba với lũ kiểm '
         'tra) vì tính với lưu lượng tràn chính 311 m^{3}/s ứng với bề rộng tràn 25 m.'),
        ('Trong đó: Bb – khoảng cách giữa hai hàng biên',
         '(hồ sơ ghi Bb = 1,5 m nhưng thay số 3,0 m – cần ghi lại cho đúng)', '(trị số TKCS dùng khi tính)'),
    ]
    for dau, cu, moi in ds:
        thay(tim_doan(body, dau), cu, moi)
    ghi('Giọng văn khi so sánh với hồ sơ dự án (đơn vị tư vấn là nơi SV thực tập, hướng dẫn ĐATN): viết theo hướng '
        '"đồ án cập nhật, thống nhất số liệu theo hiện trạng và tiêu chuẩn hiện hành", bỏ các từ đánh giá đúng/sai '
        '(mục 1.1.3, 1.2.3, 1.4.3, 1.6.2, 1.7, 2.1, 4.2.3, 8.2.2 và Kết luận).')

def main(vao, ra, dir_tt):
    KQ = tai_ket_qua(dir_tt)
    d = docx.Document(vao)
    body = d.element.body
    if 'Cấp công trình thay đổi theo các thế hệ tiêu chuẩn' in ''.join(body.itertext()):
        sys.exit('Tệp vào đã qua đợt 6 – dùng bản trước đợt 6: git show 0276cd6:DATN_HoCaGiay/thuyet_minh/'
                 'DATN_TranQuocThanh_HoCaGiay.docx > vao.docx')
    n_dong, n_bm = C.go_muc_luc(body)
    ghi('Gỡ mục lục cũ (%d dòng, %d bookmark) để điền lại sau khi sửa.' % (n_dong, n_bm))
    chuong1_2(body, KQ)
    chuong3(d, body, KQ)
    chuong4_on_dinh(d, body, KQ)
    chuong4_muc46(d, body, KQ)
    chuong5(d, body, KQ)
    chuong5_muc56(d, body, KQ)
    chuong6(d, body, KQ)
    chuong7(d, body, KQ)
    chuong8(d, body, KQ)
    ket_luan(d, body, KQ)
    phu_luc(d, body, KQ)
    tai_lieu(d, body)
    bo_sung_doi_chieu(body)
    cua_van_tran_chinh(body)
    giong_van_ho_so(body)
    d.save(ra)
    with open(os.path.splitext(ra)[0] + '_nhat_ky.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(C.NHAT_KY) + '\n')
    print('Đã lưu', ra, '–', len(C.NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(*sys.argv[1:4])
