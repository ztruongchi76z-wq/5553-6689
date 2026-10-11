# -*- coding: utf-8 -*-
"""
TIẾP THU GÓP Ý, HOÀN THIỆN THUYẾT MINH (bản tối 10/10/2026)

  python3 tiep_thu_gop_y.py <vào.docx> <ra.docx> <thư mục hình mới> <tệp kết quả JSON...>
  python3 chen_muc_luc.py <ra.docx> <bản cuối.docx> <bản cuối.pdf>

Sửa trực tiếp trên tệp Word bản 10/10 (đã có mục lục) theo các góp ý đã chọn lọc (RA_SOAT_THUYET_MINH.md mục 3,
mục 6 và các góp ý "astra", "gem"), số liệu lấy từ các lần tính lại:
  - run_7885: mô hình thấm, ổn định sau sửa chữa với đỉnh +78,85 m (thay +78,75 m);
  - phan_tich_bo_sung.json: Spencer, lưới tâm mịn, tải trọng thi công;
  - bo_sung_giai_tich.json: hướng gió, độ nhạy m tràn, cống i = 3 ‰, neo ngưỡng, khung cống, mưa, m/n;
  - kt_hoi_tu.json: hội tụ lưới PTHH, tham số α, hệ số động đất k_h;
  - kt_hai_hang.json: màng 2 hàng kề nhau A – B, đáy màng theo TCVN 8216:2018 mục 9.2.2.9.
Mỗi thay đổi được ghi vào nhật ký (in ra màn hình và tệp *_nhat_ky.txt).
"""
import json
import os
import sys

import docx
from docx.oxml.ns import qn

import cong_cu_docx as C
from cong_cu_docx import ghi, thay, them_cuoi, dat_noi_dung, doan_moi, chen_sau, tim_doan, tim_bang, dat_o, \
    them_hang, bang_moi, van_ban, hang, o

D_DAP_BU = None    # đất đắp bù mặt đập (m3) – lấy theo BV-04, gán ở main()


def so(x, n=2):
    """Số kiểu Việt Nam: dấu phẩy thập phân, dấu trừ U+2212."""
    return ('%.*f' % (n, x)).replace('.', ',').replace('-', '−')


def so_mu(q):
    """6,06·10^{-6}"""
    return '%s·10^{-6}' % so(q * 1e6)


# =========================================================================== PHẦN ĐẦU, CHƯƠNG 1
def phan_dau_chuong1(d, body, hinh_moi):
    p = tim_doan(body, 'LỜI CÁM ƠN', style='CoverB')
    thay(p, 'LỜI CÁM ƠN', 'LỜI CẢM ƠN')
    ghi('Tiêu đề "LỜI CÁM ƠN" -> "LỜI CẢM ƠN" (thống nhất với lời văn "cảm ơn").')

    # 1.2.3 – câu so sánh dung tích (dung tích hữu ích tính lại LỚN hơn thiết kế)
    p = tim_doan(body, 'Với bảng tính lại (Bảng 1.2, Hình 1.1)')
    thay(p, '; các giá trị này nhỏ hơn số liệu thiết kế ban đầu (8,41; 37,24 và 28,83 triệu m3) do bồi lắng và sai '
            'khác giữa hai lần đo vẽ địa hình.',
         '. Dung tích chết và dung tích toàn bộ nhỏ hơn số liệu thiết kế ban đầu (8,41 và 37,24 triệu m^{3}) do bồi '
         'lắng và sai khác giữa hai lần đo vẽ địa hình; riêng dung tích hữu ích lớn hơn thiết kế (28,83 triệu m^{3}) '
         'khoảng 0,77 triệu m^{3} (2,7%) vì phần dung tích dưới MNC giảm nhiều hơn (1,59 so với 0,83 triệu m^{3}). '
         'Ảnh hưởng của sai khác này đến mức bảo đảm cấp nước được xét ở mục 2.2.3.')
    ghi('Mục 1.2.3: sửa câu "các giá trị này nhỏ hơn thiết kế" – dung tích hữu ích tính lại 29,60 > 28,83 triệu m3 '
        '(lớn hơn 2,7%); nêu rõ nguyên nhân và chỗ xét ảnh hưởng.')

    p = tim_doan(body, 'Lưu vực tự nhiên 141 km2')
    thay(p, 'chuỗi dòng chảy tháng đến hồ 44 năm (1981 – 2025)',
         'chuỗi dòng chảy tháng đến hồ 44 năm thủy văn (1981-1982 ÷ 2024-2025)')
    ghi('Mục 1.4.2, 2.2.2: ghi rõ "44 năm thủy văn (1981-1982 ÷ 2024-2025)" để tránh hiểu nhầm 45 năm.')

    # 1.6.4 – số máy đóng mở
    p = tim_doan(body, 'Cống hộp BTCT 1,8 × 2,0 m, dài 67 m')
    thay(p, 'Máy đóng mở đã hư hỏng,', 'Hai máy đóng mở (02 máy V20 theo BCNCKT [3], TKCS [2]) đã hư hỏng,')
    _, t11 = tim_bang(body, 'Bảng 1.1 ')
    for tr in hang(t11):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) == 4 and van_ban(tcs[1]) == 'Máy đóng mở':
            dat_o(tcs[3], '02 máy VĐ20, trục vít D85 dài 6,0 m')
    ghi('Mục 1.6.4, Bảng 1.1: ghi rõ cống có 02 máy đóng mở (theo BCNCKT, TKCS) – trước đây chỉ ghi chung.')

    # danh mục từ viết tắt: thêm Lu, TC (dùng ở các chương 5, 7, 8)
    h = [x for x in body.iterchildren(qn('w:p')) if van_ban(x).startswith('DANH MỤC CÁC TỪ VIẾT TẮT')
         and not (C.kieu(x) or '').startswith('TOC')][0]
    tvt = h.getnext()
    while tvt.tag != qn('w:tbl'):
        tvt = tvt.getnext()

    def _vt(ma):
        for i, tr in enumerate(hang(tvt)):
            if van_ban(tr.findall(qn('w:tc'))[0]) == ma:
                return i
        raise ValueError(ma)
    them_hang(tvt, _vt('KTTV'), ['Lu', 'Lugeon – đơn vị lượng mất nước khi ép nước (1 Lu = 1 l/phút trên 1 m hố khoan ở '
                                       'áp lực 1 MPa)'])
    them_hang(tvt, _vt('TB'), ['TC', 'Thi công (tổ hợp thi công – sửa chữa khi kiểm tra ổn định)'])
    ghi('Danh mục từ viết tắt: thêm Lu (Lugeon), TC (thi công).')

    # 1.7 – phương pháp, kiểm chứng, giả thiết, khai báo sử dụng AI
    p = tim_doan(body, 'Các bài toán được tác giả lập trình tính toán bằng ngôn ngữ Python')
    thay(p, 'Các bài toán được tác giả lập trình tính toán', 'Các bài toán được lập trình tính toán')
    thay(p, 'ổn định mái theo phương pháp Bishop đơn giản với lưới tâm trượt;',
         'ổn định mái theo phương pháp Bishop đơn giản với lưới tâm trượt, kiểm tra chéo bằng phương pháp Spencer;')
    p1 = doan_moi('Trình tự thực hiện và kiểm chứng kết quả được tóm tắt trên Hình 1.2. Kết quả của các chương trình '
                  'được kiểm chứng theo ba cách: (i) tính lại với đúng thông số của hồ sơ tư vấn rồi so sánh với kết quả '
                  'đã công bố (điều tiết lũ, thấm, ổn định); (ii) giải lại bằng phương pháp độc lập (ổn định mái theo '
                  'Spencer – thỏa mãn đồng thời cân bằng lực và cân bằng mô men; tìm lại cung trượt trên lưới tâm mịn '
                  'gấp đôi); (iii) kiểm tra hội tụ lưới phần tử hữu hạn. Các số liệu không có trong hồ sơ được giả '
                  'thiết theo cấu tạo thông thường hoặc thiên an toàn; kết quả kiểm chứng và phân tích độ nhạy các giả thiết '
                  'tổng hợp ở Phụ lục 6 (Bảng PL6 ÷ PL9).', p)
    mau_anh = tim_doan(body, 'Hình 1.1 ', style='Chuthichhinh').getprevious()
    anh = C.doan_anh(d, mau_anh, os.path.join(hinh_moi, 'hinh_1_2_trinh_tu.jpg'), 4680000, 'hinh_1_2_trinh_tu.jpg')
    cap_h = doan_moi('Hình 1.2 Trình tự thực hiện và kiểm chứng kết quả tính toán của đồ án',
                     tim_doan(body, 'Hình 1.1 ', style='Chuthichhinh'))
    p2 = doan_moi('Trong quá trình thực hiện, tác giả có sử dụng công cụ trí tuệ nhân tạo (AI) để hỗ trợ soạn và kiểm '
                  'tra mã nguồn tính toán, rà soát trình bày và đối chiếu điều khoản tiêu chuẩn. Toàn bộ số liệu đầu '
                  'vào, giả thiết, kết quả tính toán và kết luận đã được tác giả chạy lại, kiểm tra, đối chiếu với hồ sơ '
                  'gốc; tác giả chịu trách nhiệm về nội dung của đồ án.', p)
    chen_sau(p, p1, anh, cap_h, p2)
    ghi('Mục 1.7: thêm đoạn trình tự thực hiện – kiểm chứng (Hình 1.2 mới), đoạn khai báo việc sử dụng công cụ AI (RA_SOAT mục 3 ý 12); "được tác giả lập trình" -> '
        '"được lập trình".')


# =========================================================================== CHƯƠNG 2
def chuong2(d, body, KQ):
    _, t21 = tim_bang(body, 'Bảng 2.1 ')
    rows = hang(t21)
    sua = {
        3: (None, None, 'TCVN 8216:2018, Bảng 4 [11]'),
        4: (None, None, 'TCVN 8216:2018, Bảng 3'),
        5: ('Hệ số an toàn ổn định mái nhỏ nhất cho phép [K]: tổ hợp cơ bản / đặc biệt / thi công – sửa chữa',
            '1,30 / 1,17 / 1,24', 'TCVN 8216:2018, Bảng 1 (5.2.2)'),
        8: (None, None, 'TCVN 8216:2018, 7.2.2'),
        9: (None, None, 'TCVN 8216:2018, Bảng 6 (0,85); hồ sơ [3] (2,0)'),
        10: (None, 'k_{h} = 0,05 (thiên an toàn)', 'TCVN 9386:2012 Phụ lục H; TCVN 9386-1, -5:2025 [12]'),
    }
    for r, (c1, c2, c3) in sua.items():
        tcs = rows[r].findall(qn('w:tc'))
        assert van_ban(tcs[0]) == str(r), van_ban(tcs[0])
        for tc, gt in zip(tcs[1:], (c1, c2, c3)):
            if gt is not None:
                dat_o(tc, gt)
    ghi('Bảng 2.1: ghi số điều, bảng của TCVN 8216:2018 vào cột "Căn cứ"; bổ sung [K] = 1,24 cho điều kiện thi công – '
        'sửa chữa; nguồn trị số 2,0 (hồ sơ [3]); tiêu chuẩn động đất hiện hành TCVN 9386-1:2025, 9386-5:2025.')

    p = tim_doan(body, 'Lưu ý: hồ sơ BCNCKT dùng hệ số [K] = 1,13')
    C.thay_tu(p, 'Hệ số động đất k',
              'Hệ số động đất: theo bảng gia tốc nền tham chiếu của TCVN 9386:2012 (Phụ lục H), tại thị trấn Chợ Lầu '
              '(huyện Bắc Bình cũ) a_{gR} = 0,0222g; với hệ số tầm quan trọng γ_{I} ≤ 1,25, gia tốc nền thiết kế '
              'a_{g} = γ_{I}·a_{gR} ≤ 0,028g < 0,04g – vùng động đất rất yếu. Đồ án vẫn kiểm tra tổ hợp động đất với '
              'k_{h} = 0,05 (lớn hơn khoảng 2 lần a_{g}/g, thiên an toàn) và phân tích độ nhạy k_{h} = 0 ÷ 0,125 ở mục '
              '5.3. TCVN 8216:2018 dẫn chiếu TCVN 9386 không ghi năm nên khi lập thiết kế bản vẽ thi công cần đối chiếu '
              'với phiên bản hiện hành TCVN 9386-1:2025, TCVN 9386-5:2025 (thay thế TCVN 9386:2012) [12] và QCVN '
              '02:2022/BXD. Giai đoạn thi công – sửa chữa được kiểm tra với [K] = 1,24 (TH11, TH12 ở mục 5.3).')
    ghi('Ghi chú sau Bảng 2.1: căn cứ k_h – a_gR = 0,0222g (TCVN 9386:2012 Phụ lục H, Chợ Lầu), a_g ≤ 0,028g: vùng '
        'động đất rất yếu, k_h = 0,05 thiên an toàn; đối chiếu TCVN 9386-1/-5:2025, QCVN 02:2022/BXD.')

    p = tim_doan(body, 'Khả năng cấp nước được đánh giá bằng điều tiết toàn chuỗi')
    thay(p, 'với chuỗi dòng chảy đến 44 năm (1981 – 2025)',
         'với chuỗi dòng chảy đến 44 năm thủy văn (1981-1982 ÷ 2024-2025)')
    p = tim_doan(body, 'Tháng nào V chạm dung tích chết')
    thay(p, 'Mức bảo đảm P = m/(n + 1)·100%, với m – số năm đủ nước, n = 44.',
         'Mức bảo đảm tính theo công thức tần suất kinh nghiệm P = m/(n + 1)·100% [13], với m – số năm đủ nước, '
         'n = 44 (QTVH [1] dùng P = m/n·100%; kết quả theo cả hai cách được so sánh ở mục 2.2.3).')
    p = tim_doan(body, 'Kết quả điều tiết (Bảng 2.2, Hình 2.1) cho thấy')
    thay(p, 'Mức bảo đảm do đồ án tính thấp hơn kết quả của QTVH do đồ án kể tổn thất thấm 1%/tháng và dùng dung '
            'tích tính lại từ bình đồ 2022 (nhỏ hơn).',
         'Mức bảo đảm do đồ án tính thấp hơn kết quả của QTVH chủ yếu do đồ án kể tổn thất thấm k_{t} = 1%/tháng và '
         'dùng công thức P = m/(n + 1); nếu tính theo P = m/n như QTVH thì TH1 đạt 40,9%, TH1A đạt 88,6%, TH2 đạt '
         '84,1% – kết luận không đổi. Quan hệ Z ~ W 2022 có dung tích hữu ích lớn hơn thiết kế (mục 1.2.3) nên cho '
         'mức bảo đảm cao hơn khoảng 4,5% so với khi dùng quan hệ Z ~ W thiết kế.')
    ghi('Mục 2.2: dẫn nguồn công thức P = m/(n+1) [13]; so sánh với P = m/n của QTVH (40,9 / 88,6 / 84,1%); sửa lý '
        'do sai "dung tích tính lại nhỏ hơn" (thực tế dung tích hữu ích 2022 lớn hơn).')

    p = tim_doan(body, 'Quy tắc vận hành theo QTVH')
    thay(p, 'đồ án tính thêm trường hợp sự cố cửa van tràn sự cố bị kẹt (không mở được).',
         'đồ án tính thêm trường hợp sự cố cửa van tràn sự cố bị kẹt (không mở được) – đây là tình huống kiểm tra sự '
         'cố thiết bị, không thuộc tổ hợp thiết kế; kết quả dùng để đánh giá mức dự trữ an toàn và yêu cầu quản lý '
         'vận hành (phân loại ở Bảng 4.2).')
    p = tim_doan(body, 'Kết quả điều tiết lũ cho trong Bảng 2.3, Hình 2.2 và Hình 2.3')
    thay(p, 'Khi cửa van tràn sự cố bị kẹt, mực nước lũ P = 1% lên tới +78,21 m.',
         'Khi cửa van tràn sự cố bị kẹt, mực nước lũ P = 1% lên tới +78,21 m (BCNCKT: +78,09 m với giả thiết tràn chính '
         'rộng 25 m).')
    m45, m467 = KQ['giai_tich']['do_nhay_m']['0.45'], KQ['giai_tich']['do_nhay_m']['0.467']
    p = tim_doan(body, 'Đối chiếu với hồ sơ: BCNCKT 2022 tính được MNLTK')
    them_cuoi(p, ' Với hệ số lưu lượng m = 0,467 như BCNCKT (thay cho m = 0,45 thiên an toàn), MNLTK = +%s m và '
                 'MNLKT = +%s m, thấp hơn %s ÷ %s m – giả thiết m = 0,45 chỉ ảnh hưởng nhỏ và thiên về an toàn.'
              % (so(m467['MNLTK']), so(m467['MNLKT']), so(m45['MNLTK'] - m467['MNLTK']),
                 so(m45['MNLKT'] - m467['MNLKT'])))
    ghi('Mục 2.3: phân loại tình huống kẹt van (kiểm tra sự cố, Bảng 4.2); so với BCNCKT +78,09 m; độ nhạy hệ số lưu '
        'lượng m = 0,467: MNLTK +%s, MNLKT +%s m.' % (so(m467['MNLTK']), so(m467['MNLKT'])))

    # 2.4 – hướng gió, điều 10.1.5
    p = tim_doan(body, 'Trong đó: V – vận tốc gió (m/s); D – đà gió (m); H – chiều sâu nước trước đập')
    thay(p, 'kα = 1 (gió thổi thẳng góc).',
         'k_{α} – hệ số kể đến góc giữa hướng sóng và pháp tuyến mái (gió hướng Bắc lệch pháp tuyến mái 6,2° nên '
         'k_{α} ≈ 1).')
    gio = {r['huong']: r for r in KQ['giai_tich']['gio']['bang']}
    ne, nw = gio['NE'], gio['NW']
    _, t25 = tim_bang(body, 'Bảng 2.5 ')
    them_hang(t25, 2, ['Đồ án – lũ BCNCKT, gió Đông Bắc (NE)', so(ne['Z1']), so(ne['Z2']), '78,74', so(ne['Z2']),
                       '78,84'])
    p = tim_doan(body, 'Kết quả tính chi tiết cho trong Bảng 2.4 và tổng hợp trong Bảng 2.5.')
    dat_noi_dung(p,
                 'Kết quả tính chi tiết cho trong Bảng 2.4 và tổng hợp trong Bảng 2.5. Tuyến đập có phương vị N96,2°E '
                 'nên pháp tuyến mái thượng lưu hướng N6,2°E (lòng hồ ở phía Bắc). Ngoài hướng Bắc, đồ án kiểm tra thêm '
                 'các hướng gió thổi vào mái thượng lưu với góc tới không quá 60°, dùng đà gió lớn nhất cho mọi hướng '
                 '(thiên an toàn): hướng Đông Bắc (góc tới 38,8°, V_{4%%} = 18,6 m/s, V_{50%%} = 11,9 m/s) cho '
                 'Z_{1} = +%s m, Z_{2} = +%s m; hướng Tây Bắc (51,2°) cho Z_{1} = +%s m, Z_{2} = +%s m; các hướng '
                 'Đông, Tây gần song song tuyến đập (góc tới 84 ÷ 96°), các hướng Nam thổi từ phía hạ lưu nên không '
                 'gây sóng leo lên mái thượng lưu. Như vậy trị số lớn nhất là Z_{2} = +%s m (gió Đông Bắc). Theo '
                 'TCVN 8216:2018 mục 10.1.5, đập có tường chắn sóng thẳng đứng liên kết với bộ phận chống thấm thì độ '
                 'vượt cao được tính đến đỉnh tường: max(Z_{1}, Z_{2}, Z_{3}) = +%s m < +79,40 m – tường chắn sóng '
                 'hiện trạng đạt yêu cầu (dự trữ %s m); đồng thời đỉnh đập (không kể tường) phải cao hơn MNLKT tối '
                 'thiểu 0,3 m, tức +78,84 m – đây là điều kiện quyết định cao trình đỉnh đập. Cao trình đỉnh đập hiện '
                 'trạng +78,35 m **thấp hơn yêu cầu này 0,49 m và thấp hơn cả MNLKT (+78,54 m)**; như vậy khi xảy ra '
                 'lũ kiểm tra, nước hồ chỉ còn được ngăn bởi tường chắn sóng. Đây là cơ sở để nâng cao trình đỉnh đập '
                 'trong dự án sửa chữa (Chương 3, Chương 5).'
                 % (so(ne['Z1']), so(ne['Z2']), so(nw['Z1']), so(nw['Z2']), so(ne['Z2']), so(ne['Z2']),
                    so(79.40 - ne['Z2'])))
    ghi('Mục 2.4: kiểm tra các hướng gió NE, NW (đà gió lớn nhất, k_α theo góc tới): Z2 lớn nhất +%s m (NE) – thêm '
        'một hàng vào Bảng 2.5; áp dụng đúng điều 10.1.5 (độ vượt cao tính đến đỉnh tường +79,40; đỉnh đập ≥ MNLKT + '
        '0,3 = +78,84 quyết định); sửa "thấp hơn yêu cầu 0,39 m" -> 0,49 m.' % so(ne['Z2']))

    p = tim_doan(body, 'Tâm trượt nguy hiểm được tìm trên lưới tâm 28 × 26 điểm')
    thay(p, 'loại bỏ các cung trượt cắt ra ngoài biên mô hình.',
         'loại bỏ các cung trượt cắt ra ngoài biên mô hình. Kết quả được kiểm tra chéo bằng phương pháp Spencer (thỏa '
         'mãn đồng thời cân bằng lực và cân bằng mô men – nhóm phương pháp cân bằng giới hạn mà TCVN 8216:2018 mục '
         '11.3.1 yêu cầu) và bằng lưới tâm mịn gấp đôi (mục 5.3, Phụ lục 6).')
    ghi('Mục 2.5.4: nêu kiểm tra chéo Spencer, lưới tâm mịn (TCVN 8216:2018 mục 11.3.1).')

    # 2.7 – Bảng 2.8 tổng hợp và kết luận
    p = tim_doan(body, 'Mực nước lũ kiểm tra +78,54 m cao hơn đỉnh đập hiện trạng +78,35 m')
    dat_noi_dung(p, 'Mực nước lũ kiểm tra +78,54 m cao hơn đỉnh đập hiện trạng +78,35 m; đỉnh đập phải cao hơn MNLKT '
                    'tối thiểu 0,3 m (+78,84 m, TCVN 8216:2018 mục 10.1.5) – cần khôi phục và nâng cao đỉnh đập.')
    h27 = [x for x in body.iterchildren(qn('w:p')) if van_ban(x) == 'Kết luận chương 2' and C.kieu(x) == 'Heading2'][0]
    mau_cap = tim_doan(body, 'Bảng 2.7 ', style='Chuthichbang')
    dan = doan_moi('Kết quả đánh giá an toàn công trình hiện trạng được tổng hợp trong Bảng 2.8.',
                   tim_doan(body, 'Kết quả kiểm tra ổn định cho trong Bảng 2.7'))
    cap, tbl, kc = bang_moi(
        'Bảng 2.8 Tổng hợp kết quả đánh giá an toàn công trình hiện trạng',
        ['Hạng mục – chỉ tiêu', 'Kết quả', 'Yêu cầu', 'Đánh giá'],
        [
            ['Cấp nước tưới, TH1 (tưới truyền thống)', 'P = 40,0%', '≥ 85%', 'Không đạt'],
            ['Cấp nước tưới, TH1A (50% cây trồng cạn tưới tiết kiệm)', 'P = 86,7%', '≥ 85%', 'Đạt (có điều kiện)'],
            ['Mực nước lũ kiểm tra so với đỉnh tường chắn sóng', '+78,54 m', '< +79,40 m', 'Đạt'],
            ['Cao trình đỉnh đập (không kể tường)', '+78,35 m', '≥ MNLKT + 0,3 = +78,84 m', 'Không đạt'],
            ['Đỉnh tường chắn sóng so với sóng leo (mọi hướng gió)', 'Z_{max} = +%s m' % so(ne['Z2']), '≤ +79,40 m',
             'Đạt'],
            ['Gradien thấm cửa ra / tiếp xúc ống khói (phân vị 95%)', '≤ 0,17 / ≤ 1,19', '0,85 / 2,0', 'Đạt'],
            ['Ổn định tổng thể mái hạ lưu / thượng lưu', 'K_{min} = 1,364 / 1,510', '1,17 / 1,30', 'Đạt'],
            ['Ổn định lớp mặt mái hạ lưu đã mất lực dính, bão hòa', 'K = 0,52 ÷ 0,76', '≥ 1,17',
             'Không đạt (đang sạt, xói)'],
            ['Thấm mang trái tràn sự cố', 'Thấm ra trên mái khi Z > +74,0 m', 'Không thấm ra mái', 'Không đạt'],
            ['Tường bên dốc tràn chính (độ vượt cao, lũ kiểm tra)', '0,73 m', '≥ 0,5 m', 'Đạt (mục 4.2.3)'],
            ['Cống lấy nước: khả năng tháo ở MNC / máy đóng mở', '7,31 m^{3}/s / hư hỏng', '≥ 6,54 m^{3}/s / vận hành '
             'an toàn', 'Đạt / Không đạt'],
            ['Quan trắc, giám sát', 'Thủ công, thiếu lún, mưa, SCADA', 'TCVN 8414:2010, NĐ 114/2018', 'Chưa đạt'],
        ],
        [3742, 2268, 1928, 1134], mau_cap, can=['left', 'center', 'center', 'center'], co_chu=20)
    chen_sau(h27, dan, cap, tbl, kc)
    ghi('Mục 2.7: thêm Bảng 2.8 tổng hợp kết quả đánh giá an toàn hiện trạng (đạt / không đạt theo từng chỉ tiêu); '
        'sửa ý kết luận về cao trình đỉnh đập theo điều 10.1.5.')


# =========================================================================== CHƯƠNG 3
def chuong3(d, body, KQ):
    z2 = so(KQ['ne_Z2'])
    p = tim_doan(body, 'Hồ sơ TKCS [2] đề xuất giữ nguyên tường chắn sóng')
    thay(p, 'Theo kết quả tính lại ở Chương 2: cao trình đỉnh yêu cầu Zđ = +78,74 m và điều kiện đỉnh đập cao hơn MNLKT '
            'tối thiểu 0,3 m khi có tường chắn sóng cho +78,84 m.',
         'Theo kết quả tính lại ở Chương 2: tường chắn sóng hiện trạng (+79,40 m) đủ chắn sóng với mọi hướng gió '
         '(Z_{max} = +%s m); cao trình đỉnh đập do điều kiện cao hơn MNLKT tối thiểu 0,3 m khi có tường chắn sóng '
         'quyết định: Z_{đ} ≥ +78,84 m (TCVN 8216:2018, mục 10.1.5).' % z2)
    p = tim_doan(body, 'Phương án Đ1 – theo TKCS')
    dat_noi_dung(p, '**Phương án Đ1** – theo TKCS: đỉnh mặt đập +78,75 m, chỉ cao hơn MNLKT 0,21 m – thấp hơn điều kiện '
                    'MNLKT + 0,3 m khoảng 0,09 m khi dùng lũ BCNCKT (đạt nếu dùng lũ QTVH, MNLKT = +78,44 m);')
    p = tim_doan(body, 'Phương án Đ3 – như Đ1 nhưng đắp bù đất đến +78,60 m')
    dat_noi_dung(p, '**Phương án Đ3** – như Đ1 nhưng đắp bù đất đến +78,60 m, đỉnh mặt đập sau gia cố **+78,85 m**: '
                    'thỏa mãn điều kiện với cả hai bộ số liệu lũ; vì với lũ BCNCKT chỉ dư 0,01 m nên +78,85 m được quy '
                    'định là cao trình tối thiểu sau lún tại mép thượng lưu, khi thi công lấy dự phòng lún 2 cm (mục '
                    '5.1.1); khối lượng đất đắp tăng khoảng 986 × 5 × 0,1 ≈ 490 m^{3}, chi phí tăng không đáng kể; '
                    'đỉnh tường chắn sóng +79,40 m vẫn cao hơn mặt đập 0,55 m (không quá 1,2 m theo mục 10.1.5).')
    p = tim_doan(body, 'Đồ án chọn phương án Đ3 vì bảo đảm an toàn')
    thay(p, 'lớp đất đắp bù đầm chặt K ≥ 0,97; lớp nilon tái sinh; bê tông M250 đá 1×2 dày 20 cm.',
         'lớp đất đắp bù đầm chặt K ≥ 0,97; lớp cấp phối đá dăm đệm dày 5 cm; lớp nilon tái sinh; bê tông M250 đá 1×2 '
         'dày 20 cm.')
    thay(p, 'tạo dốc ngang 2% về phía hạ lưu', 'tạo dốc ngang 2% về phía hạ lưu (TCVN 8216:2018, mục 10.2.3: 2 ÷ 3%)')
    ghi('Mục 3.2.1: cập nhật căn cứ cao trình đỉnh (10.1.5, gió mọi hướng); Đ1 thấp hơn điều kiện 0,09 m; Đ3: +78,85 '
        'là cao trình tối thiểu sau lún, dự phòng lún 2 cm; bổ sung lớp cấp phối đá dăm đệm 5 cm (khép kín +78,60 + '
        '0,05 + 0,20 = +78,85); dốc ngang 2% theo 10.2.3.')

    _, t31 = tim_bang(body, 'Bảng 3.1 ')
    for tr in hang(t31):
        tcs = tr.findall(qn('w:tc'))
        if van_ban(tcs[0]) == 'Chi phí':
            dat_o(tcs[1], 'Thấp hơn: 14,38 tỷ đồng [3]')
            dat_o(tcs[2], 'Cao hơn 11%: 15,96 tỷ đồng [3]')
    ghi('Bảng 3.1: ghi chi phí hai phương án bảo vệ mái theo BCNCKT (14,38 / 15,96 tỷ đồng).')
    p = tim_doan(body, 'Kết quả Chương 2 cho thấy khả năng tháo của hai tràn hiện trạng')
    thay(p, 'thông các lỗ thoát nước dưới bản đáy.',
         'thông các lỗ thoát nước dưới bản đáy; khoan neo thép ngưỡng tràn vào nền đá để ổn định trượt của ngưỡng '
         'không phụ thuộc lực dính (mục 4.5.3; số lượng điều chỉnh theo hồ sơ hoàn công).')
    p = tim_doan(body, 'Phương án 1 được chọn. Về công nghệ, phương pháp phụt truyền thống')
    thay(p, 'được phân tích chi tiết trong chuyên đề (Chương 8)',
         'được phân tích chi tiết trong chuyên đề (Chương 8), trong đó đề xuất hợp lý hóa: giữ 3 hàng ở đoạn 15 m sát '
         'tràn, đoạn còn lại 2 hàng kề nhau A, B và xác định đáy màng theo TCVN 8216:2018 mục 9.2.2.9 trên cơ sở thí '
         'nghiệm ép nước')
    p = tim_doan(body, 'Kết cấu cống còn tốt nên giữ nguyên.')
    them_cuoi(p, ' Trong cùng đợt kiểm tra, đo cao độ đáy cống tại cửa vào và cửa ra để xác định độ dốc đáy (hồ sơ ghi '
                 '3% hoặc 3‰ – ảnh hưởng đến khả năng tháo không áp, mục 6.3.2).')
    _, t32 = tim_bang(body, 'Bảng 3.2 ')
    for tr in hang(t32):
        tcs = tr.findall(qn('w:tc'))
        ten = van_ban(tcs[0])
        if ten == 'Tràn chính':
            dat_o(tcs[1], 'Giữ nguyên quy mô; sửa chữa bề mặt đáy dốc, khe co giãn; nạo vét kênh dẫn; khoan neo ngưỡng '
                          'D25 vào đá lớp II (điều chỉnh theo hồ sơ hoàn công)')
        elif ten == 'Tràn sự cố':
            dat_o(tcs[1], 'Màng khoan phụt vữa xi măng – sét chống thấm mang trái (L = 50 m, 3 hàng theo TKCS; đề xuất '
                          'hợp lý hóa: 3 hàng đoạn 15 m sát tràn + 2 hàng A, B – Chương 8); lát lại tấm BT mái')
        elif ten == 'Cống lấy nước':
            dat_o(tcs[1], 'Thay 2 máy đóng mở VĐ20; thay gioăng, tấm trượt; bổ sung giá đỡ trục vít; kiểm tra thân cống '
                          'bằng camera, đo cao độ đáy cống')
    ghi('Chương 3: tràn chính bổ sung neo ngưỡng (hạng mục thiết kế); tràn sự cố nêu đề xuất hợp lý hóa của chuyên đề; '
        'cống bổ sung đo cao độ đáy; cập nhật Bảng 3.2.')


# =========================================================================== CHƯƠNG 4
def chuong4(d, body, KQ):
    # đánh số lại bảng 4.2 – 4.4 -> 4.3 – 4.5 (chèn bảng phân loại tình huống lũ làm Bảng 4.2)
    for cu, moi in (('Bảng 4.4', 'Bảng 4.5'), ('Bảng 4.3', 'Bảng 4.4'), ('Bảng 4.2', 'Bảng 4.3')):
        n1 = C.doi_so_chu_thich(body, cu, moi)
        n2 = C.doi_tham_chieu(body, cu, moi)
        assert n1 == 1 and n2 >= 1, (cu, n1, n2)
    ghi('Chương 4: đánh số lại Bảng 4.2 ÷ 4.4 thành 4.3 ÷ 4.5 (kể cả lời dẫn) để chèn bảng phân loại tình huống lũ.')
    p = tim_doan(body, 'Trường hợp cửa van tràn sự cố bị kẹt với lũ kiểm tra cho mực nước +79,14 m')
    dat_noi_dung(p, 'Các tình huống lũ được phân loại trong Bảng 4.2. Khi cửa van tràn sự cố bị kẹt trong lũ kiểm tra '
                    '(sự cố chồng sự cố), mực nước hồ vượt đỉnh đập sau sửa chữa nhưng còn thấp hơn đỉnh tường; tường '
                    'chắn sóng khi đó làm việc như tường chắn nước nhưng chưa được kiểm tra chịu lực, chống thấm cho '
                    'trạng thái này nên **không được coi là điều kiện bảo đảm an toàn**. QTVH cần quy định kiểm tra, '
                    'chạy thử cửa van và nguồn điện dự phòng trước mùa lũ hằng năm, có phương án mở cửa van bằng tay khi '
                    'mất điện.')
    mau_cap = tim_doan(body, 'Bảng 4.1 ', style='Chuthichbang')
    cap, tbl, kc = bang_moi(
        'Bảng 4.2 Phân loại các tình huống lũ và yêu cầu đối với đỉnh đập, tường chắn sóng',
        ['Tình huống', 'Z_{max} (m)', 'So với đỉnh đập +78,85', 'So với đỉnh tường +79,40', 'Phân loại',
         'Yêu cầu, biện pháp'],
        [
            ['Lũ thiết kế P = 1%, vận hành đúng QTVH', '+77,63', 'thấp hơn 1,22 m', 'thấp hơn 1,77 m',
             'Tổ hợp cơ bản (thiết kế)', 'Z_{2} = +%s ≤ đỉnh tường: đạt' % so(KQ['ne_Z2'])],
            ['Lũ kiểm tra P = 0,2%, vận hành đúng QTVH', '+78,54', 'thấp hơn 0,31 m', 'thấp hơn 0,86 m',
             'Tổ hợp đặc biệt (kiểm tra)', 'Đỉnh đập ≥ MNLKT + 0,3 m: đạt'],
            ['P = 1% + kẹt van tràn sự cố', '+78,21', 'thấp hơn 0,64 m', 'thấp hơn 1,19 m',
             'Kiểm tra sự cố thiết bị', 'Đạt, không cần tường giữ nước'],
            ['P = 0,2% + kẹt van (BCNCKT / QTVH)', '+79,14 / +79,18', 'cao hơn 0,29 / 0,33 m', 'thấp hơn 0,26 / 0,22 m',
             'Sự cố chồng sự cố – đánh giá rủi ro', 'Không coi tường là kết cấu giữ nước; phòng ngừa bằng vận hành, '
             'nguồn điện dự phòng'],
            ['P = 0,1% + kẹt van (QTVH)', '+79,59', 'cao hơn 0,74 m', 'cao hơn 0,19 m',
             'Vượt tần suất kiểm tra – kịch bản khẩn cấp', 'Phương án ứng phó, bản đồ ngập lụt hạ du'],
        ],
        [2268, 907, 1247, 1247, 1588, 1814], mau_cap, can=['left', 'center', 'center', 'center', 'left', 'left'],
        co_chu=19)
    chen_sau(p, cap, tbl, kc)
    ghi('Mục 4.1: thêm Bảng 4.2 phân loại tình huống lũ (thiết kế / kiểm tra / sự cố / sự cố chồng sự cố / khẩn cấp); '
        'không coi tường chắn sóng là kết cấu giữ nước khi chưa kiểm tra (RA_SOAT mục 6 ý 7, mục 3 ý 8).')

    # 4.5 – neo ngưỡng thành hạng mục thiết kế
    neo = KQ['giai_tich']['neo']
    lt, lk = neo['chi_tiet']['MNLTK'], neo['chi_tiet']['MNLKT']
    p = tim_doan(body, 'Nhận xét (Bảng 4.5): khi bỏ qua hoàn toàn lực dính')
    C.thay_tu(p, 'Kết luận: ',
              'Kết luận: ngưỡng tràn chính **chỉ bảo đảm ổn định trượt khi có lực dính c ≥ 9,1 kPa** trên mặt tiếp xúc '
              'với nền đá – điều kiện phù hợp với thực tế công trình làm việc bình thường hơn 20 năm nhưng chưa được '
              'kiểm chứng bằng hồ sơ hoàn công hay khoan lấy lõi. Để ổn định không phụ thuộc vào lực dính, đồ án đưa '
              'neo thép ngưỡng vào nền đá thành hạng mục thiết kế (mục 4.5.3); khi có hồ sơ hoàn công cần kiểm tra lại '
              'kích thước bản đáy, chân khay và thép neo hiện có để điều chỉnh số lượng neo.')
    h_mau = tim_doan(body, 'Kết quả kiểm tra', style='Heading3', so_luong=None)
    h_mau = [x for x in h_mau if x.getnext() is not None and van_ban(x.getnext()).startswith('Bảng 4.5')][0]
    h = doan_moi('Kiểm tra ổn định khi có neo ngưỡng vào nền', h_mau)
    mau_nd = tim_doan(body, 'Điều kiện ổn định trượt phẳng theo QCVN 04-05:2022')
    mau_ct = [x for x in body.iterchildren(qn('w:p')) if van_ban(x).strip() == '(4-5)'][0]
    q1 = doan_moi('Neo thép CB400-V đường kính 25 mm bố trí 2 hàng song song tim tràn trong phạm vi bản đáy ngưỡng, '
                  'khoảng cách 1,5 m theo mỗi hàng (n = 1,33 thanh/m dài ngưỡng); lỗ khoan D76 qua khối ngưỡng, ngàm '
                  '2,0 m vào đá lớp II, bơm vữa xi măng không co ngót. Khi bỏ qua hoàn toàn lực dính (c = 0), neo làm '
                  'việc theo nguyên lý ma sát – cắt: khi mặt tiếp xúc có xu hướng trượt, lực kéo trong neo tạo lực ép '
                  'bổ sung lên mặt trượt:', mau_nd)
    ct6 = C.doan_cong_thuc(mau_ct, ['K_{neo} = ', C.phan_so('f·(ΣV + n·N_{n})', 'ΣH'), ' ≥ [K] ;  N_{n} = A_{s}·R_{s}'],
                           '4-6')
    q2 = doan_moi('Trong đó: A_{s} = 4,91 cm^{2} (D25); R_{s} = 350 MPa → N_{n} = %s kN/thanh; f = 0,60; ΣV, ΣH như Bảng '
                  '4.5. Đoạn ngàm phải đủ để neo không bị nhổ – lực bám dính giữa vữa và đá lớn hơn lực kéo tính toán:'
                  % so(neo['N1'], 1), mau_nd)
    ct7 = C.doan_cong_thuc(mau_ct, 'N_{bd} = π·D·L_{ng}·τ_{bd} ≥ N_{n}', '4-7')
    q3 = doan_moi('với D = 0,076 m, L_{ng} = 2,0 m, τ_{bd} = 0,40 MPa (lấy thận trọng cho đá phong hóa vừa) → '
                  'N_{bd} = %s kN > N_{n} = %s kN: đạt. Kết quả kiểm tra trượt cho trong Bảng 4.6.'
                  % (so(neo['N_bam_da'], 0), so(neo['N1'], 1)), mau_nd)
    mau_cap = tim_doan(body, 'Bảng 4.5 ', style='Chuthichbang')
    cap, tbl, kc = bang_moi(
        'Bảng 4.6 Kiểm tra ổn định trượt ngưỡng tràn chính khi có neo (c = 0, dải rộng 1 m)',
        ['Mực nước', 'ΣV (kN)', 'ΣH (kN)', 'K không neo', 'Lực chống trượt cần bổ sung (kN/m)',
         'Lực kéo cần trong 1 thanh (kN)', 'K có neo', '[K]', 'Đánh giá'],
        [
            ['MNLTK (+77,63)', so(lt['SV'], 1), so(lt['SH'], 1), so(lt['K_c0']), so(lt['R_can'], 1),
             so(lt['luc_keo_can_1_thanh'], 1), so(lt['K_neo']), '1,150', 'Đạt'],
            ['MNLKT (+78,54)', so(lk['SV'], 1), so(lk['SH'], 1), so(lk['K_c0']), so(lk['R_can'], 1),
             so(lk['luc_keo_can_1_thanh'], 1), so(lk['K_neo']), '1,035', 'Đạt'],
        ],
        [1474, 794, 794, 850, 1361, 1247, 794, 794, 794], mau_cap, co_chu=19)
    q4 = doan_moi('Neo làm việc với hệ số sử dụng thấp (lực kéo cần %s ÷ %s kN so với khả năng %s kN/thanh), nên khi hồ sơ '
                  'hoàn công cho thấy ngưỡng đã có thép neo hoặc chân khay ngàm vào đá thì có thể giảm số lượng hoặc bỏ '
                  'hạng mục này. Khối lượng dự kiến: 36 thanh neo, khoảng 144 m khoan và 0,6 t thép (Bảng 7.1).'
                  % (so(lt['luc_keo_can_1_thanh'], 0), so(lk['luc_keo_can_1_thanh'], 0), so(neo['N1'], 0)), mau_nd)
    chen_sau(p, h, q1, ct6, q2, ct7, q3, cap, tbl, kc, q4)
    ghi('Mục 4.5: kết luận ổn định ngưỡng viết lại có điều kiện (c ≥ 9,1 kPa); thêm mục 4.5.3 "Kiểm tra ổn định khi '
        'có neo" – công thức (4-6), (4-7), Bảng 4.6: K = %s / %s với c = 0 (neo D25, 2 hàng, a = 1,5 m, ngàm 2 m), '
        'kiểm tra bám dính %s kN > %s kN.' % (so(lt['K_neo']), so(lk['K_neo']), so(neo['N_bam_da'], 0),
                                              so(neo['N1'], 1)))


# =========================================================================== CHƯƠNG 5
def chuong5(d, body, KQ, hinh_7885):
    z2 = so(KQ['ne_Z2'])
    ne = KQ['ne']
    p = tim_doan(body, 'Từ kết quả mục 2.4 và phương án chọn ở mục 3.2.1')
    dat_noi_dung(p, 'Từ kết quả mục 2.4 và phương án chọn ở mục 3.2.1: theo TCVN 8216:2018 mục 10.1.5, độ vượt cao tính '
                    'đến đỉnh tường chắn sóng: max(Z_{1}, Z_{2}, Z_{3}) = max(+%s; +%s; +78,74) = +%s m (gió Đông Bắc) '
                    '< đỉnh tường +79,40 m; đỉnh đập (không kể tường) Z_{đ} ≥ MNLKT + 0,3 = +78,84 m. **Chọn cao trình '
                    'đỉnh đập sau sửa chữa +78,85 m** – là cao trình tối thiểu sau lún của mặt bê tông tại mép thượng '
                    'lưu (chân tường chắn sóng, nơi khối đắp tiếp xúc với nước nếu tường hư hỏng); mép hạ lưu tấm '
                    '+78,76 m vẫn cao hơn MNLKT 0,22 m. Điều 11.4.4 TCVN 8216:2018 (dự phòng lún ≥ 0,5%%H) viết cho đập '
                    'đắp mới; với lớp đắp bù dày không quá 0,4 m trên thân đập đã lún ổn định hơn 25 năm, độ lún bổ '
                    'sung ước tính chỉ khoảng 0,5 ÷ 1 cm, vì vậy đồ án quy định dự phòng lún 2 cm khi thi công (cao độ '
                    'hoàn thiện +78,85 ÷ +78,87 m, dung sai chỉ cho phép dương), nghiệm thu bằng thủy chuẩn và theo dõi '
                    'bằng mốc lún mặt đập (mục 7.5). Đỉnh tường chắn sóng hiện trạng +79,40 m được giữ nguyên; sau sửa '
                    'chữa tường còn nhô cao 0,55 m so với mặt đập (không quá 1,2 m theo mục 10.1.5).'
                    % (so(ne['Z1']), z2, z2))
    p = tim_doan(body, 'Bề rộng đỉnh đập giữ nguyên B = 5,0 m')
    C.thay_tu(p, 'Mô hình tính lấy đỉnh +78,75 m',
              'Mô hình tính thấm và ổn định sau sửa chữa lấy đỉnh +78,85 m (đã tính lại toàn bộ; so với mô hình đỉnh '
              '+78,75 m theo TKCS, hệ số ổn định mái hạ lưu giảm 0,01 ÷ 0,03 và lưu lượng thấm giảm không quá 1%).')
    p = tim_doan(body, 'Cào bóc lớp mặt đá dăm láng nhựa hư hỏng')
    thay(p, 'đến cao trình +78,60 m;', 'đến cao trình +78,60 m tại mép thượng lưu, mặt đất đắp dốc ngang 2% như mặt bê '
                                       'tông;')
    p = tim_doan(body, 'Trải nilon tái sinh chống mất nước xi măng')
    dat_noi_dung(p, 'Rải lớp cấp phối đá dăm đệm dày 5 cm, lu lèn chặt; trải nilon tái sinh chống mất nước xi măng; đổ bê '
                    'tông M250 đá 1×2 dày 20 cm rộng 4,3 m, đỉnh +78,85 m tại mép thượng lưu (+78,76 m tại mép hạ lưu '
                    'tấm), dốc ngang 2% về phía hạ lưu (TCVN 8216:2018, mục 10.2.3); dự phòng lún 2 cm; khe co giãn '
                    'khoảng 5 m/khe chèn ma tít;')
    ghi('Mục 5.1: cao trình đỉnh theo 10.1.5 (gió NE, Z2 = +%s < +79,40); +78,85 là cao trình tối thiểu sau lún tại mép '
        'thượng lưu, giải trình điều 11.4.4, dự phòng lún 2 cm; mô hình sau sửa chữa đã tính lại với +78,85; cấu tạo mặt '
        'đập khép kín +78,60 + đá dăm 0,05 + BT 0,20 = +78,85 (mép hạ lưu +78,76).' % z2)

    # 5.2 – thấm
    T = KQ['tham']
    p = tim_doan(body, 'Phương pháp, sơ đồ vật liệu và điều kiện biên như mục 2.5.')
    thay(p, 'Mặt cắt sau sửa chữa có đỉnh +78,75 m,', 'Mặt cắt sau sửa chữa có đỉnh +78,85 m,')
    _, t51 = tim_bang(body, 'Bảng 5.1 ')
    for r, k in ((1, 'SC_MNDBT'), (2, 'SC_MNLTK'), (3, 'SC_MNLKT'), (4, 'SC_MNDBT_tac')):
        t = T[k]
        tcs = hang(t51)[r].findall(qn('w:tc'))
        dat_o(tcs[2], so_mu(t['q']))
        dat_o(tcs[3], '%s / %s' % (so(t['J_tiep']), so(t['J_tiep95'])))
        dat_o(tcs[4], '%s / %s' % (so(t['J_ra'], 3), so(t['J_ra95'], 3)))
        dat_o(tcs[5], '%s / %s' % (so(t['J_ck']), so(t['J_ck95'])))
        dat_o(tcs[6], so(t['z_tim']))
    H = KQ['hoi_tu']
    giam = [1 - T['SC_' + k]['q'] / T['HT_' + k]['q'] for k in ('MNDBT', 'MNLTK', 'MNLKT', 'MNDBT_tac')]
    p = tim_doan(body, 'So với hiện trạng, lưu lượng thấm giảm khoảng 2%')
    thay(p, 'lưu lượng thấm giảm khoảng 2% do mái hạ lưu được đắp trả; gradien thấm ở mọi bộ phận nhỏ hơn giá trị cho '
            'phép (tiếp xúc ống khói ≤ 0,93 < 2,0; cửa ra ≤ 0,17 < 0,85).',
         'lưu lượng thấm giảm khoảng %d ÷ %d%% do mái hạ lưu được đắp trả; gradien thấm ở mọi bộ phận nhỏ hơn giá trị '
         'cho phép (tiếp xúc ống khói ≤ %s < 2,0; cửa ra ≤ %s < 0,85).'
         % (int(min(giam) * 100), round(max(giam) * 100), so(max(T['SC_' + k]['J_tiep95'] for k in ('MNDBT', 'MNLTK',
                                                                                                   'MNLKT'))),
            so(T['SC_MNDBT_tac']['J_ra95'])))
    l1, l2, l4 = H['luoi_1.00'], H['luoi_0.50'], H['luoi_0.25']
    q_lech = max(abs(H[k]['q'] / l4['q'] - 1) for k in ('luoi_2.00', 'luoi_1.00', 'luoi_0.50'))
    z_lech = max(H[k]['ztim'] for k in ('luoi_2.00', 'luoi_1.00', 'luoi_0.50', 'luoi_0.25')) - \
        min(H[k]['ztim'] for k in ('luoi_2.00', 'luoi_1.00', 'luoi_0.50', 'luoi_0.25'))
    a05, a2 = H['alpha_0.5'], H['alpha_2.0']
    q_a = (a05['q'] / l1['q'] - 1, a2['q'] / l1['q'] - 1)
    z_a = (a05['ztim'] - l1['ztim'], a2['ztim'] - l1['ztim'])
    mau = p
    pm = doan_moi('Kiểm tra độ tin cậy của lời giải (Bảng PL8): chia mịn lưới đến 4 lần (%s nút) làm lưu lượng thấm '
                  'thay đổi không quá %s%%, cao độ đường bão hòa tại tim không quá %s m, gradien cục bộ tại vùng tiếp '
                  'xúc ống khói hội tụ về khoảng %s (< 2,0); thay tham số hàm thấm α = 0,5 ÷ 2 m^{−1} làm lưu lượng thấm '
                  'thay đổi %s ÷ +%s%% và cao độ đường bão hòa tại tim thay đổi %s ÷ +%s m – sai số này được kể đến '
                  'khi đặt ngưỡng cảnh báo quan trắc (mục 7.5).'
                  % ('{:,}'.format(l4['nodes']).replace(',', '.'), so(q_lech * 100, 1), so(z_lech), so(l4['Jtiep95']),
                     so(min(q_a) * 100, 1), so(max(q_a) * 100, 1), so(min(z_a), 1), so(max(z_a), 1)), mau)
    chen_sau(p, pm)
    ghi('Mục 5.2: Bảng 5.1 cập nhật theo mô hình đỉnh +78,85; thêm kiểm tra hội tụ lưới (q lệch ≤ %s%%) và độ nhạy α '
        '(q %s ÷ +%s%%, Z tim %s ÷ +%s m).' % (so(q_lech * 100, 1), so(min(q_a) * 100, 1), so(max(q_a) * 100, 1),
                                               so(min(z_a), 1), so(max(z_a), 1)))

    # 5.3 – ổn định
    O = KQ['on_dinh']
    B = KQ['bo_sung']
    p = tim_doan(body, 'Các trường hợp tính toán, tổ hợp tải trọng và phương pháp như mục 2.5. Kết quả trong Bảng 5.2.')
    dat_noi_dung(p, 'Các trường hợp tính toán, tổ hợp tải trọng và phương pháp như mục 2.5; bổ sung điều kiện thi công – '
                    'sửa chữa (TCVN 8216:2018, mục 11.3.2 và Bảng 1): hồ ở MNDBT, tải trọng xe máy thi công trên toàn bề '
                    'rộng đỉnh đập q = 20 kPa (lu rung 9 ÷ 16 T, ô tô vận chuyển), [K] = 1,24 (TH11, TH12). Kết quả '
                    'trong Bảng 5.2.')
    _, t52 = tim_bang(body, 'Bảng 5.2 ')
    for r in range(1, 11):
        tcs = hang(t52)[r].findall(qn('w:tc'))
        ma = van_ban(tcs[0])
        dat_o(tcs[5], so(O['SC_' + ma]['K'], 3))
    them_hang(t52, 10, ['TH11', 'Thi công: MNDBT + tải thi công 20 kPa trên đỉnh', 'HL', 'TC', '–',
                        so(B['TC_HL_q20']['K_bishop'], 3), '1,24', 'Đạt'])
    them_hang(t52, 11, ['TH12', 'Thi công: MNDBT + tải thi công 20 kPa trên đỉnh', 'TL', 'TC', '–',
                        so(B['TC_TL_q20']['K_bishop'], 3), '1,24', 'Đạt'])
    p = tim_doan(body, 'Tất cả các trường hợp đều thỏa mãn; hệ số ổn định nhỏ nhất của mái hạ lưu')
    thay(p, 'K = 1,360 (động đất, [K] = 1,17), của mái thượng lưu K = 1,498 (hạ mực nước, [K] = 1,30)',
         'K = %s (động đất, [K] = 1,17), của mái thượng lưu K = %s (hạ mực nước, [K] = 1,30)'
         % (so(O['SC_TH5']['K'], 3), so(O['SC_TH8']['K'], 3)))
    sp = {k: B[k] for k in ('TH1', 'TH4', 'TH5', 'TH7', 'TH10')}
    lech_sp = max(1 - v['spencer']['F'] / v['K_bishop'] for v in sp.values())
    lech_min = [v['K_bishop_luoi_min'] / v['K_bishop'] - 1 for v in sp.values()]
    K = KQ['hoi_tu']
    kh = [0.0, 0.05, 0.075, 0.1, 0.125]
    k5 = [K['TH5_kh%.3f' % x]['K'] for x in kh]
    k10 = [K['TH10_kh%.3f' % x]['K'] for x in kh]

    def nguong(ks):
        for (x0, a0), (x1, a1) in zip(zip(kh, ks), zip(kh[1:], ks[1:])):
            if a1 < 1.17 <= a0:
                return x0 + (a0 - 1.17) * (x1 - x0) / (a0 - a1)
        x0, x1, a0, a1 = kh[-2], kh[-1], ks[-2], ks[-1]
        return x1 + (a1 - 1.17) * (x1 - x0) / (a0 - a1)
    kh5, kh10 = nguong(k5), nguong(k10)
    KQ['kh_nguong'] = (kh5, kh10)
    q_sp = doan_moi('Kiểm tra chéo (Bảng PL6): với cung trượt nguy hiểm của các trường hợp TH1, TH4, TH5, TH7, TH10, '
                    'phương pháp Spencer (thỏa mãn đồng thời cân bằng lực và mô men) cho hệ số ổn định nhỏ hơn Bishop '
                    'không quá %s%%; tìm lại trên lưới tâm mịn gấp đôi không phát hiện cung trượt nguy hiểm hơn. Trong '
                    'giai đoạn thi công (TH11, TH12), tải trọng xe máy trên đỉnh chỉ làm hệ số ổn định giảm %s ÷ %s, vẫn '
                    'lớn hơn nhiều [K] = 1,24.'
                    % (so(lech_sp * 100, 1), so(O['SC_TH1']['K'] - B['TC_HL_q20']['K_bishop'], 2),
                       so(O['SC_TH7']['K'] - B['TC_TL_q20']['K_bishop'], 2)), p)
    q_kh = doan_moi('Độ nhạy theo hệ số động đất (Bảng PL9): mái hạ lưu thỏa mãn [K] = 1,17 khi k_{h} ≤ %s, mái thượng '
                    'lưu đạt với mọi giá trị đã tính (k_{h} ≤ 0,125). Với a_{gR} = 0,0222g của khu vực (mục 2.1), gia tốc '
                    'nền thiết kế không quá 0,028g, nên hệ số tính toán k_{h} = 0,05 thiên an toàn và mái đập còn dự trữ '
                    'đến khoảng 3 lần giá trị tương ứng với gia tốc nền.' % so(kh5 - 0.0005, 2), p)
    chen_sau(p, q_sp, q_kh)
    ghi('Mục 5.3: Bảng 5.2 cập nhật theo đỉnh +78,85 (TH5 = %s; TH8 = %s), thêm TH11, TH12 (thi công, [K] = 1,24); '
        'kiểm tra chéo Spencer (lệch ≤ %s%%), lưới tâm mịn; độ nhạy k_h: mái HL đạt khi k_h ≤ %s, TL khi k_h ≤ %s.'
        % (so(O['SC_TH5']['K'], 3), so(O['SC_TH8']['K'], 3), so(lech_sp * 100, 1), so(kh5 - 0.0005, 2),
           so(kh10 - 0.0005, 2)))

    # ảnh sau sửa chữa (Hình 5.1 – 5.4, PL1 – PL10)
    for ten in ('tham_SC_MNLKT', 'tham_SC_MNDBT_tac', 'od_SC_TH5', 'od_SC_TH8', 'od_SC_TH1', 'od_SC_TH2', 'od_SC_TH3',
                'od_SC_TH4', 'od_SC_TH6', 'od_SC_TH7', 'od_SC_TH9', 'od_SC_TH10', 'tham_SC_MNDBT', 'tham_SC_MNLTK'):
        C.thay_anh(d, ten, os.path.join(hinh_7885, ten + '.jpg'))
    for ma, so_pl in (('TH1', 1), ('TH2', 2), ('TH3', 3), ('TH4', 4), ('TH6', 5), ('TH7', 6), ('TH9', 7), ('TH10', 8)):
        cap = tim_doan(body, 'Hình PL%d ' % so_pl, style='Chuthichhinh')
        t = van_ban(cap)
        cu = t[t.index('(K'):]
        thay(cap, cu, '(K_{min} = %s)' % so(O['SC_' + ma]['K'], 3))
    ghi('Hình 5.1 ÷ 5.4, Hình PL1 ÷ PL10: thay bằng hình của mô hình đỉnh +78,85 m; cập nhật K_min trong chú thích '
        'Hình PL1 ÷ PL8.')

    # 5.5 – mưa, rãnh dọc
    M = KQ['giai_tich']['mua']
    p = tim_doan(body, 'Đặc trưng dòng chảy tràn theo chiều dài đường chảy cho trong Bảng 5.5.')
    them_cuoi(p, ' Với tỷ lệ mưa 30 phút / mưa ngày thay đổi 20 ÷ 40%% (i = %s ÷ %s mm/h), ứng suất tiếp trên ô 4 × 4 m '
                 'là %s ÷ %s Pa – vẫn nhỏ hơn giới hạn 2 ÷ 5 Pa, trong khi đường chảy dài cả mái cho %s ÷ %s Pa; kết '
                 'luận về kích thước ô 4 × 4 m không phụ thuộc giả thiết này.'
              % (so(M['20%_L4']['i'], 0), so(M['40%_L4']['i'], 0), so(M['20%_L4']['tau'], 1), so(M['40%_L4']['tau'], 1),
                 so(M['20%_L34']['tau'], 1), so(M['40%_L34']['tau'], 1)))
    R = KQ['giai_tich']['ranh_doc']
    p = tim_doan(body, 'Kết quả kiểm tra (Bảng 5.6) cho thấy các rãnh đều có khả năng thoát')
    thay(p, 'Rãnh dọc chạy theo mái 1:3 có lưu tốc lớn (≈ 9 m/s khi đầy)',
         'Rãnh dọc chạy theo mái 1:3 có lưu tốc lớn: với lưu lượng thiết kế lớp nước trong rãnh chỉ dày khoảng %s cm '
         'nhưng lưu tốc đã đạt %s m/s (≈ 9 m/s khi đầy),' % (so(R['h'] * 100, 1), so(R['v'], 1)))
    ghi('Mục 5.5: độ nhạy tỷ lệ mưa 30 phút/ngày 20 ÷ 40%% (τ ô 4 × 4 m = %s ÷ %s Pa); lưu tốc rãnh dọc với lưu lượng '
        'thiết kế %s m/s.' % (so(M['20%_L4']['tau'], 1), so(M['40%_L4']['tau'], 1), so(R['v'], 1)))


# =========================================================================== CHƯƠNG 6
def chuong6(d, body, KQ):
    p = tim_doan(body, 'Cống lấy nước là cống hộp BTCT chảy không áp')
    thay(p, 'điều tiết bằng cửa van phẳng đặt trong tháp van phía thượng lưu;',
         'điều tiết bằng cửa van phẳng đặt trong tháp van phía thượng lưu (theo BCNCKT, TKCS cống có 02 máy đóng mở V20 '
         'cần thay thế);')
    _, t61 = tim_bang(body, 'Bảng 6.1 ')
    for tr in hang(t61):
        tcs = tr.findall(qn('w:tc'))
        ten = van_ban(tcs[0])
        if ten.startswith('Chiều dài thân cống'):
            dat_o(tcs[3], '67 ; 0,03 (kiểm tra thêm 0,003)')
        elif ten == 'Máy đóng mở':
            dat_o(tcs[3], '02 máy VĐ20 (20 T), trục vít D85 dài 6,0 m')
    G = KQ['giai_tich']['cong_do_doc']['0.003']
    p = tim_doan(body, 'Với độ dốc đáy i = 0,03, độ sâu dòng đều')
    thay(p, '(Nếu độ dốc đáy chỉ là 3‰ như một số tài liệu ghi, h0 = 1,39 ÷ 1,96 m > hk, dòng êm; đồ án lấy i = 3% theo '
            'QTVH [1].)',
         '(BCNCKT [3] và QTVH [1] ghi i = 3%, một số tài liệu khác ghi 3‰; đồ án lấy i = 3% và kiểm tra thêm trường hợp '
         '3‰ ở cuối mục này.)')
    p = tim_doan(body, 'Kết quả tính (Bảng 6.2, Hình 6.1) cho mức đầy lớn nhất')
    pm = doan_moi('Trường hợp độ dốc đáy chỉ là 3‰: độ sâu dòng đều h_{0} = %s m (Q = 6,54 m^{3}/s) và %s m (Q = 10 '
                  'm^{3}/s), lớn hơn h_{k} – dòng êm; với Q = 6,54 m^{3}/s khoảng lưu không còn %s m (mức đầy %d%%) – '
                  'vẫn đạt, nhưng với Q = 10 m^{3}/s khoảng lưu không chỉ còn %s m, dòng chảy chuyển sang bán áp không '
                  'ổn định. Lưu lượng lớn nhất còn bảo đảm chảy không áp ổn định (mức đầy 75%%) khi i = 3‰ là %s '
                  'm^{3}/s. Vì vậy, nếu kết quả đo cao độ đáy cống (mục 3.4) cho i = 3‰ thì nhiệm vụ Q = 10 m^{3}/s '
                  'của QTVH chỉ thực hiện được sau khi kiểm tra cống làm việc có áp (độ kín khớp nối, áp lực trong thân '
                  'cống), nếu không phải giới hạn lưu lượng ≤ %s m^{3}/s.'
                  % (so(G['6.54']['h0']), so(G['10.0']['h0']), so(G['6.54']['luu_khong']),
                     round(G['6.54']['h0'] / 2.0 * 100), so(G['10.0']['luu_khong']), so(G['Qmax_75'], 1),
                     so(G['Qmax_75'], 1)), p)
    chen_sau(p, pm)
    ghi('Mục 6.3.2: nêu nguồn i = 3%% (BCNCKT, QTVH); thêm kịch bản i = 3‰: Q = 10 m3/s gần đầy cống (lưu không %s m), '
        'Q không áp lớn nhất %s m3/s; yêu cầu đo cao độ đáy cống.' % (so(G['10.0']['luu_khong']), so(G['Qmax_75'], 1)))

    p = tim_doan(body, 'Cửa van phẳng kéo trượt (b × h = 1,8 × 2,0 m)')
    thay(p, 'Cửa van phẳng kéo trượt (b × h = 1,8 × 2,0 m),',
         'Do hồ sơ không có bản vẽ cửa van, đồ án giả thiết hai cửa van có cùng kích thước lỗ cống và kiểm tra cho cửa '
         'chịu cột nước lớn nhất: cửa van phẳng kéo trượt (b × h = 1,8 × 2,0 m),')
    ghi('Mục 6.1, 6.4, Bảng 6.1: ghi 02 máy đóng mở, giả thiết hai cửa cùng kích thước.')

    F = KQ['giai_tich']['khung_cong']['kq']
    p = tim_doan(body, 'Áp lực đất thẳng đứng trên nắp cống')
    dat_noi_dung(p, 'Áp lực đất thẳng đứng trên nắp cống: q = K_{đ}·γ·H = 1,0 × 20,6 × 11,75 = 242 kPa, trong đó K_{đ} – '
                    'hệ số tập trung áp lực đất; với cống cứng đặt trên nền đá, đất đắp hai bên lún nhiều hơn nên K_{đ} '
                    'có thể lớn hơn 1 – đồ án xét K_{đ} = 1,0 ÷ 1,3 (q = 242 ÷ 315 kPa);')
    p = tim_doan(body, 'Áp lực đất ngang lên tường cống')
    thay(p, 'tại đáy tường e2 = 202 kPa;', 'tại đáy tường e_{2} = 202 kPa; khi tính nội lực xét thêm trường hợp áp lực '
                                           'ngang chỉ bằng 0,5·K_{0} (bất lợi cho mô men giữa nhịp bản nắp, bản đáy);')
    p = tim_doan(body, 'Việc nâng đỉnh đập từ +78,35 m lên +78,85 m làm áp lực đất đứng trên cống tăng')
    s0, s5, s13 = F['sau_SC_Kd1.0_K0x1.0'], F['sau_SC_Kd1.0_K0x0.5'], F['sau_SC_Kd1.3_K0x0.5']
    t13 = F['truoc_SC_Kd1.3_K0x0.5']
    dat_noi_dung(p, 'Việc nâng đỉnh đập từ +78,35 m lên +78,85 m làm áp lực đất đứng trên cống tăng khoảng 10,3 kPa '
                    '(≈ 4,4%). Do hồ sơ không có bản vẽ kết cấu, đồ án tính nội lực tham khảo của mặt cắt ngang cống như '
                    'một khung kín đối xứng (nhịp tim 2,2 m, chiều cao tim 2,425 m; giả thiết bản nắp, tường dày 0,40 m, '
                    'bản đáy dày 0,45 m), giải bằng phương pháp lực với hai ẩn là mô men góc trên và góc dưới, phản lực '
                    'nền phân bố đều; cốt thép yêu cầu ước tính theo A_{s} ≈ K_{n}·n_{c}·M/(R_{s}·0,9·h_{0}) với '
                    'K_{n} = 1,15, n_{c} = 1,0, thép CB300-V R_{s} = 260 MPa, lớp bảo vệ a = 5 cm (nguyên tắc tính kết cấu '
                    'BTCT thủy công theo TCVN 4116:2023 [30]). Kết quả cho trong Bảng 6.4.')
    mau_cap = tim_doan(body, 'Bảng 6.3 ', style='Chuthichbang')

    def dong(ten, x):
        return [ten, so(x['MA'], 1), so(x['MB'], 1), so(x['Mt_giua'], 1), so(x['Mb_giua'], 1), so(x['As_nap_giua'], 1),
                so(x['As_goc'], 1), so(x['As_day_giua'], 1)]
    cap, tbl, kc = bang_moi(
        'Bảng 6.4 Nội lực và cốt thép yêu cầu của khung kín cống hộp (tham khảo, dải rộng 1 m)',
        ['Trường hợp tải trọng', 'M góc trên (kNm)', 'M góc dưới (kNm)', 'M giữa nắp (kNm)', 'M giữa đáy (kNm)',
         'A_{s} giữa nắp (cm^{2})', 'A_{s} góc (cm^{2})', 'A_{s} giữa đáy (cm^{2})'],
        [dong('Sau SC: K_{đ} = 1,0; K_{0} = 0,674', s0), dong('Sau SC: K_{đ} = 1,0; 0,5·K_{0}', s5),
         dong('Sau SC: K_{đ} = 1,3; 0,5·K_{0}', s13), dong('Trước SC: K_{đ} = 1,3; 0,5·K_{0} (so sánh)', t13)],
        [2608, 794, 794, 794, 794, 850, 794, 850], mau_cap, co_chu=19)
    tang = s13['As_nap_giua'] / t13['As_nap_giua'] - 1
    q2 = doan_moi('Cốt thép yêu cầu sau sửa chữa khoảng %s ÷ %s cm^{2}/m (tương đương 5Ø16 ÷ 5Ø20 mỗi mét), lớn hơn trước '
                  'sửa chữa khoảng %d%%. Đây là kết quả tham khảo, dùng làm tiêu chí đối chiếu khi kiểm tra cống: kết '
                  'hợp camera (robot) quan sát vết nứt, rò nước tại khớp nối khi đóng cửa van giữa hai vụ và khoan lõi, '
                  'dò cốt thép tại cửa ra để xác định chiều dày, cốt thép thực tế; nếu cốt thép thực tế nhỏ hơn yêu cầu '
                  'hoặc phát hiện nứt, rò nước thì tính toán kết cấu chi tiết và xử lý khớp nối, vết nứt bằng vữa, keo '
                  'epoxy hoặc ống lót. Kết luận về thân cống vì vậy là "chưa kiểm tra chịu lực theo hồ sơ thực tế".'
                  % (so(min(s0['As_day_giua'], s0['As_nap_giua']), 0), so(s13['As_nap_giua'], 0), round(tang * 100)),
                  tim_doan(body, 'Nhận xét (Bảng 6.3)'))
    chen_sau(p, cap, tbl, kc, q2)
    ghi('Mục 6.5: thay câu "mức tăng nằm trong hệ số an toàn" bằng bài tính khung kín tham khảo (Bảng 6.4, A_s = %s ÷ '
        '%s cm2/m, tăng %d%% so với trước SC) làm tiêu chí kiểm tra; K_đ = 1,0 ÷ 1,3; dẫn TCVN 4116:2023.'
        % (so(min(s0['As_day_giua'], s0['As_nap_giua']), 1), so(s13['As_nap_giua'], 1), round(tang * 100)))


# =========================================================================== CHƯƠNG 7
def chuong7(d, body, KQ):
    p = tim_doan(body, 'Trình tự đại trà: phụt hai hàng biên (A, C) trước')
    thay(p, 'phụt hai hàng biên (A, C) trước, hàng giữa (B) sau;',
         'đoạn 3 hàng phụt hàng hạ lưu C, hàng thượng lưu A trước, hàng giữa B sau; nếu áp dụng phương án hợp lý hóa '
         '(Chương 8), đoạn 2 hàng phụt hàng B (hạ lưu) trước, hàng A sau;')
    p = tim_doan(body, 'Kiểm tra chất lượng: khoan kiểm tra (3 hố)')
    thay(p, 'màng đạt yêu cầu khi hệ số thấm K ≤ 1·10-5 cm/s.',
         'màng đạt yêu cầu khi hệ số thấm K ≤ 1·10^{-5} cm/s (đoạn trong đá tương ứng lượng mất nước q ≲ 0,8 Lu theo quy '
         'đổi 1 Lu ≈ 1,3·10^{-5} cm/s [35]); đồng thời ép nước lớp II tại các hố phụt thí nghiệm để xác định cao trình '
         'đáy màng theo TCVN 8216:2018 mục 9.2.2.9.')
    p = tim_doan(body, 'Bê tông mặt đập, rãnh, gờ chắn bánh và mặt đường quản lý')
    them_cuoi(p, ' Khoan neo ngưỡng tràn chính: khoan lỗ D76 bằng máy khoan xoay – đập, thổi rửa lỗ, đặt thanh neo D25 '
                 'có định vị tâm, bơm vữa xi măng không co ngót từ đáy lỗ lên; thí nghiệm kéo kiểm tra tối thiểu 3 neo '
                 'với lực kéo thử 100 kN (≈ 1,7 lần lực kéo yêu cầu lớn nhất 57 kN).')
    p = tim_doan(body, 'Tháo dỡ máy đóng mở cũ, kiểm tra cửa van')
    thay(p, 'Tháo dỡ máy đóng mở cũ, kiểm tra cửa van;', 'Tháo dỡ 02 máy đóng mở cũ, kiểm tra cửa van;')
    thay(p, 'lắp máy VĐ20 mới,', 'lắp 02 máy VĐ20 mới,')
    them_cuoi(p, ' Trong đợt ngừng cấp nước, kiểm tra bên trong thân cống bằng camera (robot) và đo cao độ đáy cống tại '
                 'cửa vào, cửa ra (mục 6.3.2, 6.5).')
    p = tim_doan(body, 'Không tập kết vật liệu, máy móc nặng trên mái đập')
    them_cuoi(p, ' Tải trọng xe máy trên đỉnh đập trong thi công đã được kiểm tra với [K] = 1,24 (TH11, TH12 – mục 5.3).')
    ghi('Mục 7.2: trình tự phụt (đoạn 3 hàng C → A → B; đoạn 2 hàng B → A); tiêu chí Lugeon; ép nước xác định đáy màng '
        '(9.2.2.9); biện pháp khoan neo ngưỡng; 02 máy đóng mở; kiểm tra thân cống, đo cao độ đáy; tải thi công trên '
        'đỉnh.')

    # Bảng 7.1
    _, t71 = tim_bang(body, 'Bảng 7.1 ')

    def tim_hang(dau):
        for i, tr in enumerate(hang(t71)):
            if van_ban(tr.findall(qn('w:tc'))[1]).startswith(dau):
                return i
        raise ValueError(dau)
    i = tim_hang('Đất đắp bù mặt đập')
    dat_o(o(t71, i, 1), 'Đất đắp bù mặt đập K ≥ 0,97 (đến +78,60 m, dốc 2%)')
    dat_o(o(t71, i, 3), '{:,}'.format(D_DAP_BU).replace(',', '.'))
    i = tim_hang('Nilon tái sinh lót móng')
    dat_o(o(t71, i, 3), '4.240')
    them_hang(t71, i - 1, ['3', 'Cấp phối đá dăm đệm dày 5 cm (rộng 4,3 m)', 'm^{3}', '212'])
    i = tim_hang('Bê tông M250 mặt đập dày 20 cm')
    dat_o(o(t71, i, 1), 'Bê tông M250 mặt đập dày 20 cm (rộng 4,3 m)')
    dat_o(o(t71, i, 3), '848')
    i = tim_hang('Cào bóc mặt đập đá dăm láng nhựa cũ')
    dat_o(o(t71, i, 1), 'Cào bóc mặt đập đá dăm láng nhựa cũ (dày 15 cm, rộng 4,70 m)')
    dat_o(o(t71, i, 3), '695')
    i = tim_hang('Bê tông M250 gờ chắn bánh')
    dat_o(o(t71, i, 1), 'Bê tông M250 gờ chắn bánh (20 × 50 cm, đoạn 1,0 m cách 0,30 m)')
    dat_o(o(t71, i, 3), '76')
    # đánh lại số TT phần I
    rows = hang(t71)
    a = tim_hang('Cào bóc mặt đập')
    b = tim_hang('Trồng cỏ mái hạ lưu')
    for k, r in enumerate(range(a, b + 1), start=1):
        dat_o(rows[r].findall(qn('w:tc'))[0], str(k))
    i = tim_hang('Hố khoan kiểm tra; thí nghiệm đổ nước, ép nước')
    them_hang(t71, i, ['8', 'Khoan neo ngưỡng tràn chính: thanh neo D25 (khoan D76, ngàm 2,0 m vào đá lớp II); chiều '
                            'dài khoan; thép neo', 'thanh ; m ; t', '36 ; 144 ; 0,6'])
    i = tim_hang('Thay gioăng, tấm trượt, giá đỡ trục trung gian')
    them_hang(t71, i, ['3', 'Kiểm tra thân cống bằng camera (robot); đo cao độ đáy cống tại cửa vào, cửa ra', 'lần',
                       '1'])
    p = tim_doan(body, 'Khối lượng khoan phụt tổng hợp từ Bảng 2-2 TKCS')
    p_von = doan_moi('Theo BCNCKT [3], vốn đầu tư sửa chữa, nâng cấp hồ Cà Giây khoảng 48,69 tỷ đồng. Các điều chỉnh của '
                     'đồ án làm thay đổi chi phí không đáng kể (tăng khoảng 490 m^{3} đất đắp mặt đập và 36 neo ngưỡng '
                     'tràn); nếu áp dụng phương án hợp lý hóa màng khoan phụt (Chương 8), khối lượng khoan, phụt giảm '
                     'khoảng 28%. Khái toán chi tiết cần đơn giá năm 2026 của địa phương, thực hiện ở bước thiết kế bản '
                     'vẽ thi công.', tim_doan(body, 'Khối lượng chủ yếu được tác giả tính từ kích thước thiết kế'))
    chen_sau(p, p_von)
    them_cuoi(p, ' //Khối lượng mặt đập tính theo chi tiết bản vẽ BV-04 (tấm bê tông rộng 4,3 m, gờ chắn bánh, đất đắp bù '
                 'từ +78,20 m sau cào bóc đến +78,60 m dốc 2%). Nếu áp dụng phương án hợp lý hóa màng khoan phụt '
                 '(Chương 8, Bảng 8.7), tổng chiều dài khoan, phụt giảm khoảng 28%.//')
    ghi('Bảng 7.1: sửa khối lượng mặt đập theo chi tiết BV-04 (BT 848 m3 thay 986; nilon 4.240 m2 thay 4.930; đất đắp '
        'bù %s m3 thay 1.972; cào bóc 695 m3 (rộng 4,70 m) thay 740; gờ chắn bánh 76 m3 thay 61; thêm đá dăm đệm '
        '212 m3); thêm khoan neo ngưỡng (36 thanh, 144 m, 0,6 t) và kiểm tra thân cống.'
        % '{:,}'.format(D_DAP_BU).replace(',', '.'))

    # Bảng 7.2 – tiến độ
    _, t72 = tim_bang(body, 'Bảng 7.2 ')
    for tr in hang(t72):
        tc = tr.findall(qn('w:tc'))[0]
        t = van_ban(tc)
        if t.startswith('Đắp bù và đổ bê tông mặt đập'):
            dat_o(tc, 'Đắp bù và đổ bê tông mặt đập (đoạn mang tràn sự cố sau khi khoan phụt xong), gờ chắn, điện chiếu '
                      'sáng')
        elif t.startswith('Sửa chữa đáy dốc tràn chính'):
            dat_o(tc, 'Sửa chữa đáy dốc, khoan neo ngưỡng tràn chính, nạo vét kênh dẫn')
        elif t.startswith('Thay máy đóng mở'):
            dat_o(tc, 'Thay máy đóng mở, phục hồi cửa van; kiểm tra thân cống, đo cao độ đáy')
    p = tim_doan(body, 'Tiến độ thi công dự kiến thể hiện trên Bảng 7.2')
    p_tt = doan_moi('Trình tự bắt buộc: (i) khoan phụt mang tràn sự cố hoàn thành trước khi đổ bê tông mặt đập đoạn mang '
                 'tràn (các hố khoan đi qua mặt đập) và trước khi lát lại tấm mái tràn sự cố; (ii) xử lý hố xói, đắp trả '
                 'và rãnh mái hạ lưu hoàn thành trước mùa mưa (tháng 5); trồng cỏ vào đầu mùa mưa để cỏ bén rễ; trong '
                 'tháng 5 – 6 phải che phủ mặt đất mới đắp, khơi thông rãnh tạm và theo dõi bản tin dự báo để ứng phó lũ '
                 'sớm; (iii) sửa chữa đáy dốc, khoan neo ngưỡng tràn chính khi mực nước hồ thấp hơn ngưỡng +74,70 m.',
                 tim_doan(body, 'Khối lượng chủ yếu được tác giả tính từ kích thước thiết kế'))
    chen_sau(p, p_tt)
    ghi('Mục 7.4: Bảng 7.2 ghi rõ trình tự (mặt đập đoạn mang tràn sau khoan phụt; neo ngưỡng; kiểm tra cống); thêm '
        'các ràng buộc trình tự và biện pháp cho tháng 5 – 6 đầu mùa mưa (RA_SOAT mục 6 ý 8).')

    # 7.5 – ngưỡng cảnh báo
    T = KQ['tham']
    p = tim_doan(body, 'Công tác quan trắc thực hiện theo TCVN 8414:2010')
    dat_noi_dung(p, 'Công tác quan trắc thực hiện theo TCVN 8414:2010 [33] và Nghị định 114/2018/NĐ-CP. Kết quả tính thấm '
                    'của đồ án được dùng làm “giá trị dự báo” để so sánh với số liệu quan trắc: cao độ đường bão hòa tại '
                    'tim đập mặt cắt lòng sông khoảng +%s m khi hồ ở MNDBT, +%s m khi hồ ở MNLTK và +%s m khi hồ ở MNLKT. '
                    'Trong năm đầu sau sửa chữa cần lập đường cơ sở (quan hệ giữa cao độ đường bão hòa tại từng giếng và '
                    'mực nước hồ), sau đó cảnh báo theo độ lệch so với đường cơ sở, tốc độ tăng và lưu lượng, độ đục '
                    'của nước thấm. Ngưỡng đề xuất: (1) **cấp I** – tăng tần suất quan trắc, kiểm tra hiện trường khi cao '
                    'độ đo được vượt giá trị dự báo từ 1,0 m (lớn hơn sai số của mô hình khoảng −1,2 ÷ +0,5 m do tham số '
                    'hàm thấm, mục 5.2); (2) **cấp II** – báo cáo cấp có thẩm quyền, kiểm tra thiết bị thoát nước khi '
                    'vượt từ 2,0 m (khi ống khói, đệm tiêu nước bị tắc, đường bão hòa tại tim dâng khoảng %s m, Bảng '
                    '5.1) hoặc khi xuất hiện dòng thấm đục ở mái, chân đập, mang tràn sự cố; (3) **lún mặt đập** – khi '
                    'lún so với cao độ nghiệm thu vượt dự phòng lún 2 cm hoặc lún lệch giữa hai mốc liền kề trên 1 cm thì '
                    'kiểm tra lại cao trình đỉnh đập so với điều kiện MNLKT + 0,3 m. Nội dung và tần suất quan trắc, '
                    'kiểm tra đề xuất cho trong Bảng 7.3.'
                    % (so(T['SC_MNDBT']['z_tim'], 1), so(T['SC_MNLTK']['z_tim'], 1), so(T['SC_MNLKT']['z_tim'], 1),
                       so(T['SC_MNDBT_tac']['z_tim'] - T['SC_MNDBT']['z_tim'], 1)))
    p = tim_doan(body, 'Cập nhật quy trình vận hành: quy định rõ thao tác mở cửa van tràn sự cố')
    thay(p, 'giới hạn vận hành cống khi mực nước hồ cao (mục 6.4);',
         'giới hạn vận hành cống khi mực nước hồ cao (mục 6.4) và giới hạn lưu lượng cống ≤ 7,2 m^{3}/s nếu đo kiểm tra '
         'cho độ dốc đáy 3‰ (mục 6.3.2);')
    ghi('Mục 7.5: cập nhật giá trị dự báo đường bão hòa (mô hình +78,85); thay ngưỡng "vượt 1,5 m" chưa có căn cứ bằng '
        'đường cơ sở + ngưỡng cấp I (1,0 m – theo sai số mô hình), cấp II (2,0 m – theo trạng thái tắc thoát nước), '
        'ngưỡng lún (2 cm) (RA_SOAT mục 6 ý 9).')


# =========================================================================== CHƯƠNG 8
def chuong8(d, body, KQ, hinh_moi):
    p = tim_doan(body, 'Trong đó: ΔHm – tổn thất cột nước qua màng (m)')
    thay(p, 'Theo TKCS, TCVN 8216:2018 yêu cầu chiều dày màng phụt trong thân đập không nhỏ hơn 1/10 cột nước công tác, '
            'tức T ≥ 1,65 m [11].',
         'TCVN 8216:2018 (mục 10.4.2) yêu cầu chiều dày lõi màng phụt trong thân đập (chỉ áp dụng cho đập từ cấp II trở '
         'xuống) không nhỏ hơn 1/10 cột nước làm việc, tức T ≥ 1,65 m [11]; mục 9.2.3.2 cho gradien cho phép của màng '
         'vữa xi măng – sét J = 3 ÷ 4, phù hợp với [J_{m}] ≈ 3,6 mà TKCS sử dụng.')
    p = tim_doan(body, 'Trong đó: Bb – khoảng cách giữa hai hàng biên (m)')
    thay(p, 'với 2 hàng: T = 1,5 + 1,8 = 3,3 m;',
         'với 2 hàng kề nhau (khoảng cách 1,5 m, hố so le): T = 1,5 + 1,8 = 3,3 m;')
    them_cuoi(p, ' Công thức (8-3) chỉ đúng khi các hàng kề nhau đủ gần để vùng vữa lan của các hàng giao nhau; nếu bỏ '
                 'hàng giữa và giữ hai hàng biên A, C (cách nhau 3,0 m) thì giữa hai hàng còn dải đất chưa được phụt và '
                 'không thể coi màng có chiều dày 3,0 + 1,8 = 4,8 m.')
    p = tim_doan(body, '(3) Chiều sâu màng: màng phải cắt qua hết các lớp thấm mạnh')
    thay(p, 'dưới chân màng [35], [36], [37].',
         'dưới chân màng [35], [36], [37]. TCVN 8216:2018 (mục 9.2.2.9) quy định độ sâu màng xác định theo tính toán '
         'thấm; khi lớp cách nước nằm nông thì màng cắm vào lớp này khoảng 5 m, nhưng với đập từ cấp II trở xuống phạm '
         'vi khoan phụt không vượt quá vị trí nền có lượng mất nước (5 ÷ 7) Lu cộng thêm 3 m, và trong mọi trường hợp '
         'không sâu quá 1H. Với quy đổi 1 Lu ≈ 1,3·10^{-5} cm/s [35], lớp II có K ≈ 1·10^{-5} cm/s (≈ 0,8 Lu) thuộc vùng '
         'ít thấm, nên đáy màng hợp lý nằm khoảng 3 m dưới mặt lớp II nếu thí nghiệm ép nước xác nhận q ≤ 5 Lu.')
    ghi('Mục 8.2.2: dẫn đúng TCVN 8216:2018 mục 10.4.2 (T ≥ H/10), 9.2.3.2 (J = 3 ÷ 4), 9.2.2.9 (giới hạn độ sâu màng '
        'đập cấp II: vị trí (5 ÷ 7) Lu + 3 m); công thức (8-3) chỉ áp dụng cho các hàng kề nhau.')

    p = tim_doan(body, 'Mặt cắt tính toán lấy theo “Cắt ngang đại diện khoan phụt”')
    thay(p, '(+78,75 m trong mô hình)', '(+78,75 m trong mô hình; tính lại với +78,85 m, lưu lượng thấm và hiệu quả màng '
                                       'thay đổi không quá 0,4%)')

    H2 = KQ['hai_hang']
    p = tim_doan(body, 'Chiều dày màng có ảnh hưởng vừa phải.')
    them_cuoi(p, ' Vị trí dải màng 2 hàng (A – B, B – C hay đối xứng qua tim) hầu như không ảnh hưởng: η = %s ÷ %s%% '
                 '(MNLTK) và %s ÷ %s%% (MNLKT).'
              % (so(min(H2['77.63_' + k]['eta'] for k in ('2h_doi_xung', '2h_A_B', '2h_B_C')) * 100, 1),
                 so(max(H2['77.63_' + k]['eta'] for k in ('2h_doi_xung', '2h_A_B', '2h_B_C')) * 100, 1),
                 so(min(H2['78.54_' + k]['eta'] for k in ('2h_doi_xung', '2h_A_B', '2h_B_C')) * 100, 1),
                 so(max(H2['78.54_' + k]['eta'] for k in ('2h_doi_xung', '2h_A_B', '2h_B_C')) * 100, 1)))
    pm = doan_moi('Mức độ ảnh hưởng của các thông số đến hiệu quả màng được tổng hợp trên Hình 8.6: chất lượng phụt '
                  '(K_{màng}) và việc cắt hết lớp I là hai yếu tố quyết định, tiếp theo là tính thấm lớp II; số hàng '
                  'khoan và cao trình đỉnh màng ảnh hưởng ít.', tim_doan(body, 'Các phân tích sau thực hiện với mực nước'))
    mau_anh = tim_doan(body, 'Hình 8.5 ', style='Chuthichhinh').getprevious()
    anh = C.doan_anh(d, mau_anh, os.path.join(hinh_moi, 'hinh_8_6_loc_xoay.jpg'), 5580000, 'hinh_8_6_loc_xoay.jpg')
    cap = doan_moi('Hình 8.6 Phạm vi thay đổi hiệu quả giảm lưu lượng thấm η theo từng thông số của màng (MNLTK)',
                   tim_doan(body, 'Hình 8.5 ', style='Chuthichhinh'))
    chen_sau(p, pm, anh, cap)

    # 8.6 – đề xuất hợp lý hóa
    p = tim_doan(body, 'Đoạn còn lại (khoảng 35 m) có thể giảm còn 2 hàng (A, C)')
    dat_noi_dung(p, '**Đoạn còn lại (khoảng 35 m) có thể giảm còn 2 hàng kề nhau A, B** (khoảng cách hàng 1,5 m, hố so le, '
                    'A = 3,0 m; bỏ hàng hạ lưu C) nếu kết quả phụt thí nghiệm cho thấy màng 2 hàng đạt K ≤ 1·10^{-5} '
                    'cm/s; không bỏ hàng giữa B vì hai hàng biên A, C cách nhau 3,0 m không tạo được màng liên tục; tại '
                    'mặt cắt đại diện hiệu quả chống thấm chỉ giảm khoảng 6%% (η = %s%% so với %s%%, MNLTK);'
                    % (so(H2['77.63_2h_A_B']['eta'] * 100, 0), so(H2['77.63_3h_TK']['eta'] * 100, 0)))
    p = tim_doan(body, 'Cao trình đáy màng xác định theo thí nghiệm ép nước lớp II')
    dat_noi_dung(p, '**Cao trình đáy màng xác định theo TCVN 8216:2018 mục 9.2.2.9** trên cơ sở thí nghiệm ép nước tại các '
                    'hố phụt thí nghiệm: nếu lớp II có lượng mất nước q ≤ 5 Lu ngay từ mặt lớp thì đáy màng thấp hơn mặt '
                    'lớp II 3,0 m (+69,4 m tại mặt cắt đại diện, so với +68,3 m của TKCS; η không đổi – %s%% với 3 hàng, '
                    '%s%% với 2 hàng), nhưng không sâu hơn đáy màng TKCS; nếu lớp II nứt nẻ (q > 5 Lu) thì đáy màng thấp '
                    'hơn vị trí đạt q = 5 Lu 3,0 m;'
                    % (so(H2['77.63_3h_day_69_4']['eta'] * 100, 1), so(H2['77.63_2h_A_B_day_69_4']['eta'] * 100, 1)))
    p = tim_doan(body, 'Quan trắc sau xử lý: bố trí 2 giếng đo áp')
    thay(p, '(sau hàng C khoảng 5 m)', '(sau hàng hạ lưu của màng khoảng 5 m)')
    kl = KQ['kl_mang']
    pb = doan_moi('Khối lượng và hiệu quả của màng theo TKCS và theo phương án hợp lý hóa được so sánh trong Bảng 8.7.',
                  tim_doan(body, 'Theo kết quả mô hình, tổn thất cột nước qua màng'))
    mau_cap = tim_doan(body, 'Bảng 8.6 ', style='Chuthichbang')

    def eta(k):
        return '%s%% / %s%%' % (so(H2['77.63_' + k]['eta'] * 100, 0), so(H2['78.54_' + k]['eta'] * 100, 0))

    def jra(k):
        return '%s / %s' % (so(H2['77.63_' + k]['Jra']), so(H2['78.54_' + k]['Jra']))
    cap, tbl, kc = bang_moi(
        'Bảng 8.7 So sánh khối lượng và hiệu quả của màng theo TKCS và theo phương án hợp lý hóa',
        ['Phương án', 'Số hố', 'Khoan (m)', 'Phụt (m)', 'Trong đó phụt trong đá II (m)', 'η MNLTK / MNLKT',
         'J_{ra} MNLTK / MNLKT'],
        [
            ['TKCS: 3 hàng toàn tuyến, đáy theo TKCS', str(kl['tk'][0]), so(kl['tk'][1], 1), so(kl['tk'][2], 1),
             so(kl['tk'][3], 1), eta('3h_TK'), jra('3h_TK')],
            ['Hợp lý hóa: 3 hàng đoạn 15 m + 2 hàng A, B; đáy theo TKCS', str(kl['hl_tk'][0]), so(kl['hl_tk'][1], 1),
             so(kl['hl_tk'][2], 1), so(kl['hl_tk'][3], 1), eta('2h_A_B'), jra('2h_A_B')],
            ['Hợp lý hóa + đáy màng theo mục 9.2.2.9 (mặt lớp II − 3 m)', str(kl['hl_9229'][0]),
             so(kl['hl_9229'][1], 1), so(kl['hl_9229'][2], 1), so(kl['hl_9229'][3], 1), eta('2h_A_B_day_69_4'),
             jra('2h_A_B_day_69_4')],
        ],
        [3061, 624, 907, 907, 1134, 1247, 1191], mau_cap, co_chu=19)
    ghi_chu = doan_moi('//Khối lượng tính theo hình học cắt dọc tuyến màng (BV-05), cùng một cách tính cho các phương án, '
                       'không kể hố phụt thí nghiệm và hố kiểm tra; η, J_{ra} tại mặt cắt đại diện (thuộc đoạn 2 hàng).//',
                       tim_doan(body, 'Khối lượng khoan phụt tổng hợp từ Bảng 2-2 TKCS'))
    p_cuoi = tim_doan(body, 'Quan trắc sau xử lý: bố trí 2 giếng đo áp')
    chen_sau(p_cuoi, pb, cap, tbl, ghi_chu)

    p = tim_doan(body, 'Có thể hợp lý hóa thiết kế (3 hàng ở đoạn sát tràn')
    dat_noi_dung(p, 'Có thể hợp lý hóa thiết kế (3 hàng ở đoạn 15 m sát tràn, 2 hàng kề nhau A, B ở đoạn còn lại; đáy màng '
                    '3 m dưới vị trí có q ≤ 5 Lu theo TCVN 8216:2018) để giảm khoảng %d%% khối lượng khoan, phụt mà hiệu '
                    'quả chỉ giảm khoảng 6%%, với điều kiện được khẳng định bằng phụt thí nghiệm.'
                    % round((1 - kl['hl_9229'][1] / kl['tk'][1]) * 100))
    ghi('Chương 8: đoạn 2 hàng đổi từ A, C (cách 3,0 m – không liên tục) sang 2 hàng kề nhau A, B; đáy màng theo 9.2.2.9 '
        '(mặt lớp II − 3 m, +69,4 thay +70,5 ÷ +71,0); Bảng 8.7 so sánh khối lượng (khoan %s -> %s m); Hình 8.6 tổng hợp '
        'độ nhạy; mô hình +78,85 thay đổi ≤ 0,4%%.' % (so(kl['tk'][1], 0), so(kl['hl_9229'][1], 0)))


# =========================================================================== KẾT LUẬN, TLTK, PHỤ LỤC
def ket_luan_tltk_phu_luc(d, body, KQ):
    z2 = so(KQ['ne_Z2'])
    O = KQ['on_dinh']
    T = KQ['tham']
    p = tim_doan(body, 'Với tràn hiện trạng, MNLTK = +77,63 m, MNLKT = +78,54 m (lũ BCNCKT)')
    dat_noi_dung(p, 'Với tràn hiện trạng, MNLTK = +77,63 m, MNLKT = +78,54 m (lũ BCNCKT); tường chắn sóng +79,40 m đủ chắn '
                    'sóng với mọi hướng gió (Z_{max} = +%s m, gió Đông Bắc), cao trình đỉnh đập do điều kiện MNLKT + '
                    '0,3 m = +78,84 m quyết định (TCVN 8216:2018, mục 10.1.5). Đỉnh đập hiện trạng +78,35 m không đạt; đồ '
                    'án chọn đỉnh đập sau sửa chữa +78,85 m (cao hơn TKCS 0,10 m; là cao trình tối thiểu sau lún, dự '
                    'phòng lún 2 cm).' % z2)
    p = tim_doan(body, 'Thân, nền đập bảo đảm an toàn thấm')
    qmin = min(T['SC_' + k]['q'] for k in ('MNDBT', 'MNLTK', 'MNLKT'))
    qmax = max(T['SC_' + k]['q'] for k in ('MNDBT', 'MNLTK', 'MNLKT'))
    dat_noi_dung(p, 'Thân, nền đập bảo đảm an toàn thấm (gradien nhỏ hơn cho phép, lưu lượng thấm sau sửa chữa khoảng '
                    '%s ÷ %s·10^{-6} m^{3}/s/m) và ổn định tổng thể (K_{min} = %s ÷ 2,33 ≥ [K], kể cả giai đoạn thi '
                    'công) ở cả trạng thái hiện tại và sau sửa chữa; kết quả Bishop được kiểm tra chéo bằng Spencer '
                    '(chênh không quá 0,2%%), lời giải thấm hội tụ theo lưới (q thay đổi không quá 0,2%%); tổ hợp động '
                    'đất tính với k_{h} = 0,05, thiên an toàn so với a_{gR} = 0,0222g của khu vực (mái hạ lưu đạt khi '
                    'k_{h} ≤ %s). Hư hỏng mái hạ lưu là '
                    'hư hỏng lớp mặt – phân tích mái dốc vô hạn giải thích cơ chế sạt lở khi lớp mặt mất lực dính và '
                    'bão hòa.' % (so(qmin * 1e6, 1), so(qmax * 1e6, 1), so(O['SC_TH5']['K'], 2),
                                  so(KQ['kh_nguong'][0] - 0.0005, 2)))
    p = tim_doan(body, 'Tràn chính đủ khả năng tháo, tường bên dốc nước đủ cao')
    thay(p, 'ngưỡng tràn ổn định khi kể lực dính nhỏ trên mặt tiếp xúc với nền đá.',
         'ngưỡng tràn chỉ ổn định khi có lực dính c ≥ 9,1 kPa trên mặt tiếp xúc với nền đá – đồ án bổ sung neo D25 vào '
         'nền (K = 1,81 ÷ 2,23 khi c = 0).')
    them_cuoi(p, ' Nếu độ dốc đáy cống là 3‰ thì cống chỉ tháo không áp ổn định Q ≤ 7,2 m^{3}/s; nội lực khung cống '
                 '(tham khảo) cho cốt thép yêu cầu 8 ÷ 16 cm^{2}/m – tiêu chí đối chiếu khi kiểm tra cống.')
    kl = KQ['kl_mang']
    p = tim_doan(body, 'Chuyên đề: mô hình thấm PTHH cho thấy màng khoan phụt thiết kế')
    thay(p, 'có thể hợp lý hóa số hàng và chiều sâu màng trên cơ sở phụt thí nghiệm.',
         'có thể hợp lý hóa (2 hàng kề nhau A, B ngoài đoạn 15 m sát tràn; đáy màng theo TCVN 8216:2018 mục 9.2.2.9), '
         'giảm khoảng %d%% khối lượng khoan, phụt trên cơ sở phụt thí nghiệm.'
         % round((1 - kl['hl_9229'][1] / kl['tk'][1]) * 100))
    p = tim_doan(body, 'Bổ sung khảo sát: hồ sơ hoàn công tràn chính')
    dat_noi_dung(p, 'Bổ sung khảo sát: hồ sơ hoàn công tràn chính, cống lấy nước (bản đáy, chân khay, thép neo, cốt thép '
                    'thân cống); đo cao độ đáy cống để xác định độ dốc (3%% hay 3‰); thí nghiệm ép nước lớp II tại mang '
                    'tràn sự cố; đối chiếu gia tốc nền a_{gR} theo QCVN 02:2022/BXD, TCVN 9386-1:2025 khi lập thiết kế '
                    'bản vẽ thi công.')
    p = tim_doan(body, 'Bổ sung vào quy trình vận hành: kiểm tra cửa van tràn sự cố')
    thay(p, 'dùng các giá trị đường bão hòa dự báo của đồ án làm ngưỡng cảnh báo quan trắc.',
         'dùng các giá trị đường bão hòa dự báo của đồ án và đường cơ sở năm đầu làm căn cứ cảnh báo quan trắc (cấp I: '
         'vượt 1,0 m; cấp II: vượt 2,0 m hoặc nước thấm đục; lún mặt đập vượt 2 cm).')
    ghi('Kết luận, kiến nghị: cập nhật số liệu cao trình đỉnh (10.1.5), ổn định (+78,85; kiểm tra chéo), kết luận có '
        'điều kiện (k_h, ngưỡng tràn, thân cống); kiến nghị đo độ dốc cống, tiêu chuẩn động đất 2025, ngưỡng cảnh báo.')

    # Tài liệu tham khảo
    cap = [x for x in body.iterchildren(qn('w:p')) if van_ban(x) == 'TÀI LIỆU THAM KHẢO'][0]
    tbl = cap.getnext()
    while tbl.tag != qn('w:tbl'):
        tbl = tbl.getnext()

    def hang_tl(so_tl):
        for tr in hang(tbl):
            if van_ban(tr.findall(qn('w:tc'))[0]) == '[%d]' % so_tl:
                return tr
        raise ValueError(so_tl)
    dat_o(hang_tl(12).findall(qn('w:tc'))[1],
          'Bộ Khoa học và Công nghệ, //TCVN 9386-1:2025 – Thiết kế công trình chịu động đất – Phần 1: Quy định chung, tác '
          'động động đất và quy định đối với kết cấu nhà// và //TCVN 9386-5:2025 – Thiết kế công trình chịu động đất – Phần '
          '5: Nền móng, tường chắn và các vấn đề địa kỹ thuật// (thay thế TCVN 9386:2012), Hà Nội, 2025.')
    dat_o(hang_tl(30).findall(qn('w:tc'))[1],
          'Bộ Khoa học và Công nghệ, //TCVN 4116:2023 – Công trình thủy lợi – Kết cấu bê tông và bê tông cốt thép thủy '
          'công – Yêu cầu thiết kế// (thay thế TCVN 4116:1985), Hà Nội, 2023.')
    tr37 = hang_tl(37)
    r38 = them_hang(tbl, hang(tbl).index(tr37), ['[38]', 'Bộ Xây dựng, //QCVN 02:2022/BXD – Quy chuẩn kỹ thuật quốc gia '
                                                         'về Số liệu điều kiện tự nhiên dùng trong xây dựng//, Hà Nội, '
                                                         '2022.'])
    them_hang(tbl, hang(tbl).index(r38), ['[39]', 'S. Spencer, “A method of analysis of the stability of embankments '
                                                  'assuming parallel inter-slice forces,” //Géotechnique//, vol. 17, no. '
                                                  '1, pp. 11–26, 1967.'])
    ghi('Tài liệu tham khảo: [12] -> TCVN 9386-1:2025 và 9386-5:2025; [30] -> TCVN 4116:2023; thêm [38] QCVN '
        '02:2022/BXD, [39] Spencer (1967) – trích dẫn lần đầu ở Phụ lục 6.')

    # Phụ lục 5 – danh mục chương trình
    _, t_pl5 = tim_bang(body, 'Bảng PL5 ')
    i = len(hang(t_pl5)) - 2
    for ten, nd in (('phan_tich_bo_sung.py', 'Kiểm tra chéo ổn định theo Spencer, lưới tâm mịn, tải trọng thi công '
                                             '(TH11, TH12)'),
                    ('bo_sung_giai_tich.py', 'Hướng gió, độ nhạy hệ số lưu lượng tràn, cống khi i = 3‰, neo ngưỡng '
                                             'tràn, khung kín cống hộp, độ nhạy mưa, mức bảo đảm m/n'),
                    ('kt_hoi_tu.py', 'Hội tụ lưới PTHH, độ nhạy tham số α, độ nhạy hệ số động đất k_{h}'),
                    ('kt_hai_hang.py, kt_dinh.py', 'Màng 2 hàng kề nhau, đáy màng theo mục 9.2.2.9, ảnh hưởng đỉnh đập '
                                                   '+78,85 m (Chương 8)')):
        them_hang(t_pl5, i, [ten, nd])
        i += 1
    ghi('Bảng PL5: thêm 4 dòng chương trình tính bổ sung.')

    # Phụ lục 6
    pl5 = [x for x in body.iterchildren(qn('w:p')) if van_ban(x).startswith('Phụ lục 5. Chương trình tính toán')][0]
    h6 = doan_moi('Phụ lục 6. Kiểm chứng chương trình tính và phân tích độ nhạy', pl5)
    mau_nd = tim_doan(body, 'Ghi chú: cột V lấy theo quan hệ V(Z)')
    i6 = doan_moi('Phụ lục này tổng hợp kết quả kiểm chứng các chương trình tính (Bảng PL6) và phân tích độ nhạy các '
                  'giả thiết (Bảng PL7 ÷ Bảng PL9). Phương pháp Spencer [39] dùng để kiểm tra chéo Bishop vì thỏa mãn '
                  'đồng thời cân bằng lực và mô men; gia tốc nền tham chiếu cần đối chiếu theo QCVN 02:2022/BXD [38].',
                  mau_nd)
    mau_cap = tim_doan(body, 'Bảng PL5 ', style='Chuthichbang')
    B = KQ['bo_sung']
    H = KQ['hoi_tu']
    sp = ['TH1', 'TH4', 'TH5', 'TH7', 'TH10']
    cap6, t6, kc6 = bang_moi(
        'Bảng PL6 Kiểm chứng kết quả tính toán của đồ án',
        ['Nội dung kiểm chứng', 'Kết quả đồ án', 'Đối chứng', 'Chênh lệch', 'Ghi chú'],
        [
            ['MNLTK / MNLKT, dùng thông số BCNCKT (B = 25 m)', '+77,48 / +78,32 m', '+77,43 / +78,25 m (BCNCKT)',
             '0,05 / 0,07 m', 'Mục 2.3.2'],
            ['MNLTK / MNLKT, dùng thông số QTVH (B = 21,6 m)', '+77,53 / +78,44 m', '+77,43 / +78,31 m (QTVH)',
             '0,10 / 0,13 m', 'Dung tích tính lại nhỏ hơn'],
            ['Lưu lượng thấm hiện trạng MNDBT / MNLTK (m^{3}/s/m)', '6,22 / 7,73·10^{-6}',
             '5,71 / 7,61·10^{-6} (SEEP/W)', '+9 / +2%', 'Bảng 2.6'],
            ['K mái hạ lưu hiện trạng TH1 / TH2 / TH3 / TH4', '1,667 / 1,647 / 1,637 / 1,519',
             '1,69 / 1,59 / 1,55 / 1,54 (BCNCKT)', '−1,4 ÷ +5,6%', 'Bảng 2.7'],
            ['K Bishop – Spencer, cùng cung trượt (sau SC: TH1, TH4, TH5, TH7, TH10)',
             '; '.join(so(B[k]['K_bishop'], 3) for k in sp), '; '.join(so(B[k]['spencer']['F'], 3) for k in sp),
             '≤ %s%%' % so(max(1 - B[k]['spencer']['F'] / B[k]['K_bishop'] for k in sp) * 100, 1), 'Spencer [39]'],
            ['K khi tìm lại trên lưới tâm mịn gấp đôi', 'như trên', '; '.join(so(B[k]['K_bishop_luoi_min'], 3)
                                                                             for k in sp),
             '+%s ÷ +%s%%' % (so(min(B[k]['K_bishop_luoi_min'] / B[k]['K_bishop'] - 1 for k in sp) * 100, 1),
                              so(max(B[k]['K_bishop_luoi_min'] / B[k]['K_bishop'] - 1 for k in sp) * 100, 1)),
             'Không có cung nguy hiểm hơn'],
            ['Hội tụ lưới PTHH (MNLKT, sau SC): q với 4.229 ÷ 31.794 nút',
             '%s·10^{-6}' % so(H['luoi_1.00']['q'] * 1e6), '%s ÷ %s·10^{-6}' % (
                 so(min(H[k]['q'] for k in H if k.startswith('luoi')) * 1e6),
                 so(max(H[k]['q'] for k in H if k.startswith('luoi')) * 1e6)), '≤ 0,2%', 'Bảng PL8'],
            ['Cân bằng nước toàn trận lũ', 'sai số < 0,02·10^{6} m^{3}', '–', '–', 'Mục 2.3.1'],
        ],
        [2722, 1701, 2041, 1077, 1531], mau_cap, can=['left', 'center', 'center', 'center', 'left'], co_chu=19)
    G = KQ['giai_tich']
    m467 = G['do_nhay_m']['0.467']
    gio = {r['huong']: r for r in G['gio']['bang']}
    cap7, t7, kc7 = bang_moi(
        'Bảng PL7 Phân tích độ nhạy các giả thiết tính toán',
        ['Thông số (giá trị dùng)', 'Phạm vi kiểm tra', 'Kết quả', 'Kết luận'],
        [
            ['Hệ số lưu lượng tràn m (0,45)', '0,467 (BCNCKT)', 'MNLTK +%s m; MNLKT +%s m' % (so(m467['MNLTK']),
                                                                                         so(m467['MNLKT'])),
             'm = 0,45 thiên an toàn'],
            ['Hướng gió (Bắc)', 'NE, NW; đà gió lớn nhất', 'Z_{2} = +%s m (NE); +%s m (NW)' % (
                so(gio['NE']['Z2']), so(gio['NW']['Z2'])), '< đỉnh tường +79,40 m'],
            ['Tham số hàm thấm α (1 m^{−1})', '0,5 ÷ 2 m^{−1} (MNLKT)',
             'q −1,9 ÷ +3,5%; Z tim −1,2 ÷ +0,5 m', 'Kể vào ngưỡng quan trắc (7.5)'],
            ['Hệ số động đất k_{h} (0,05)', '0 ÷ 0,125', 'K HL 1,647 → 1,056; K TL 2,296 → 1,259 (Bảng PL9)',
             'Đạt khi k_{h} ≤ %s; a_{gR} = 0,0222g' % so(KQ['kh_nguong'][0] - 0.0005, 2)],
            ['Tải thi công trên đỉnh (0)', '10 ÷ 20 kPa', 'K HL %s ÷ %s; K TL %s ÷ %s' % (
                so(B['TC_HL_q10']['K_bishop'], 3), so(B['TC_HL_q20']['K_bishop'], 3), so(B['TC_TL_q10']['K_bishop'], 3),
                so(B['TC_TL_q20']['K_bishop'], 3)), '> [K] = 1,24'],
            ['Tỷ lệ mưa 30 phút / ngày (30%)', '20 ÷ 40%', 'τ (L = 4 m) 1,4 ÷ 2,2 Pa; τ (L = 34 m) 5,1 ÷ 7,8 Pa',
             'Ô 4 × 4 m đủ chống xói'],
            ['Độ dốc đáy cống (3%)', '3‰', 'h_{0} = 1,39 ÷ 1,96 m; Q không áp ≤ 7,2 m^{3}/s', 'Cần đo cao độ đáy cống'],
            ['Lực dính ngưỡng tràn – nền (50 kPa)', '0 ÷ 50 kPa; có neo', 'K = 0,65 ÷ 3,41; có neo 1,81 ÷ 2,23',
             'Bổ sung neo D25'],
            ['Mức bảo đảm cấp nước (m/(n+1))', 'm/n (QTVH)', 'TH1 40,9%; TH1A 88,6%; TH2 84,1%', 'Kết luận không đổi'],
        ],
        [2495, 1814, 2948, 1814], mau_cap, can=['left', 'left', 'left', 'left'], co_chu=19)
    luoi = [('×2 (thô)', 'luoi_2.00'), ('×1 (cơ sở)', 'luoi_1.00'), ('×0,5', 'luoi_0.50'), ('×0,25', 'luoi_0.25')]
    dong8 = [['Kích thước phần tử ' + t, '{:,}'.format(H[k]['nodes']).replace(',', '.'), so(H[k]['q'] * 1e6, 3),
              so(H[k]['ztim']), so(H[k]['Jtiep95']), so(H[k]['Jra95'], 3), so(H[k]['Jck95'], 3)] for t, k in luoi]
    dong8 += [['α = %s m^{−1} (lưới cơ sở)' % t, '{:,}'.format(H['luoi_1.00']['nodes']).replace(',', '.'),
               so(H[k]['q'] * 1e6, 3), so(H[k]['ztim']), so(H[k]['Jtiep95']), so(H[k]['Jra95'], 3), '–']
              for t, k in (('0,5', 'alpha_0.5'), ('2,0', 'alpha_2.0'))]
    cap8, t8, kc8 = bang_moi(
        'Bảng PL8 Kiểm tra hội tụ lưới và độ nhạy tham số α (MNLKT, mặt cắt sau sửa chữa)',
        ['Trường hợp', 'Số nút', 'q (10^{-6} m^{3}/s/m)', 'Z tim (m)', 'J tiếp xúc 95%', 'J cửa ra 95%',
         'J chân khay 95%'],
        dong8, [2608, 907, 1361, 1020, 1077, 1077, 1020], mau_cap, co_chu=19)
    K = KQ['hoi_tu']
    kh = [0.0, 0.05, 0.075, 0.1, 0.125]
    cap9, t9, kc9 = bang_moi(
        'Bảng PL9 Hệ số ổn định mái đập sau sửa chữa theo hệ số động đất k_{h} (MNDBT)',
        ['Trường hợp'] + [so(x, 3) for x in kh] + ['[K]'],
        [['TH5 – mái hạ lưu'] + [so(K['TH5_kh%.3f' % x]['K'], 3) for x in kh] + ['1,17'],
         ['TH10 – mái thượng lưu'] + [so(K['TH10_kh%.3f' % x]['K'], 3) for x in kh] + ['1,17']],
        [2608, 1077, 1077, 1077, 1077, 1077, 1077], mau_cap, co_chu=19)
    # chèn sau bảng PL5
    t_pl5_cuoi = t_pl5
    chen_sau(t_pl5_cuoi, h6, i6, cap6, t6, kc6, cap7, t7, kc7, cap8, t8, kc8, cap9, t9, kc9)
    ghi('Thêm Phụ lục 6: Bảng PL6 kiểm chứng (so với BCNCKT, QTVH, SEEP/W, Spencer, lưới mịn, hội tụ lưới), Bảng PL7 '
        'độ nhạy giả thiết, Bảng PL8 hội tụ lưới – α, Bảng PL9 hệ số ổn định theo k_h.')


# =========================================================================== CHÍNH
def doc_kq(duong_dan):
    with open(duong_dan, encoding='utf-8') as f:
        return json.load(f)


def main(vao, ra, hinh_moi, hinh_7885, dir_7885, dir_kp, d_dap_bu):
    global D_DAP_BU
    D_DAP_BU = int(d_dap_bu)
    to = doc_kq(os.path.join(dir_7885, 'tham_on_dinh.json'))
    KQ = dict(tham=to['tham'], on_dinh=to['on_dinh'],
              bo_sung=doc_kq(os.path.join(dir_7885, 'phan_tich_bo_sung_on_dinh.json')),
              giai_tich=doc_kq(os.path.join(dir_7885, 'bo_sung_giai_tich.json')),
              hoi_tu=doc_kq(os.path.join(dir_7885, 'kt_hoi_tu.json')),
              hai_hang=doc_kq(os.path.join(dir_kp, 'kt_hai_hang.json')),
              kl_mang=doc_kq(os.path.join(dir_kp, 'khoi_luong_mang.json')))
    gio = {r['huong']: r for r in KQ['giai_tich']['gio']['bang']}
    KQ['ne'] = gio['NE']
    KQ['ne_Z2'] = gio['NE']['Z2']
    d = docx.Document(vao)
    body = d.element.body
    n_dong, n_bm = C.go_muc_luc(body)
    ghi('Gỡ mục lục cũ (%d dòng, %d bookmark) để điền lại sau khi sửa.' % (n_dong, n_bm))
    phan_dau_chuong1(d, body, hinh_moi)
    chuong2(d, body, KQ)
    chuong3(d, body, KQ)
    chuong4(d, body, KQ)
    chuong5(d, body, KQ, hinh_7885)
    chuong6(d, body, KQ)
    chuong7(d, body, KQ)
    chuong8(d, body, KQ, hinh_moi)
    ket_luan_tltk_phu_luc(d, body, KQ)
    d.save(ra)
    with open(os.path.splitext(ra)[0] + '_nhat_ky.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(C.NHAT_KY) + '\n')
    print('Đã lưu', ra, '–', len(C.NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(*sys.argv[1:8])
