# -*- coding: utf-8 -*-
"""
TIẾP THU GÓP Ý ĐỢT 2 (11/10/2026, tệp "Gpt_gop_y.txt") – sửa trực tiếp tệp Word thuyết minh.

  python3 tiep_thu_gop_y_2.py <vào.docx> <ra.docx> <thư mục ket_qua của tinh_toan>
  python3 chen_muc_luc.py <ra.docx> <bản cuối.docx> <bản cuối.pdf>

Số liệu mới lấy từ tinh_toan/ket_qua: kt_tran_su_co.json (thủy lực tràn sự cố bắt đầu từ hk),
kt_lop_II_lu.json (hiệu quả màng theo hệ số thấm lớp II). Nhật ký: nhat_ky_tiep_thu_gop_y_2.txt.
"""
import json
import os
import sys

import docx
from docx.oxml.ns import qn

import cong_cu_docx as C
from cong_cu_docx import ghi, tim_doan, tim_bang, thay, dat_noi_dung, doan_moi, chen_sau, dat_o, them_hang, \
    bang_moi, van_ban, hang


def so(x, n=1):
    return ('%.*f' % (n, x)).replace('.', ',')


def o_theo_chu(body, chu):
    kq = [tc for tc in body.iter(qn('w:tc')) if van_ban(tc).strip() == chu]
    if len(kq) != 1:
        raise ValueError('Ô %r: thấy %d' % (chu, len(kq)))
    return kq[0]


# =========================================================================== CHƯƠNG 2
def chuong2(body):
    p = tim_doan(body, 'Kết quả điều tiết (Bảng 2.2, Hình 2.1) cho thấy')
    thay(p, '(iii) phương án tiềm năng TH2 đạt 82,2%, xấp xỉ yêu cầu.',
         '(iii) phương án tiềm năng TH2 đạt 82,2% < 85% – chưa đạt mức bảo đảm yêu cầu.')
    p = tim_doan(body, 'Quy tắc vận hành theo QTVH')
    thay(p, '(phân loại ở Bảng 4.3)', '(phân loại ở Bảng 4.2)')
    p = tim_doan(body, 'Kết quả tính thấm hiện trạng cho trong Bảng 2.6')
    thay(p, '(1 ÷ 3)% dung tích hồ mỗi tháng theo TCVN 10778.', '(1 ÷ 3)% dung tích hồ mỗi tháng [13].')
    ghi('Chương 2: TH2 (82,2% < 85%) ghi rõ "chưa đạt"; dẫn chiếu phân loại tình huống lũ sửa thành Bảng 4.2; '
        'tổn thất thấm cho phép dẫn [13] (TCVN 10778 đã có bản 2024).')


# =========================================================================== CHƯƠNG 4
def chuong4(body, kq_tsc):
    a, b = kq_tsc['P=1%'], kq_tsc['P=0,2%']
    p = tim_doan(body, 'Tràn sự cố rộng 5 m, ngưỡng +71,00 m', chua='không có bậc hạ sau ngưỡng')
    dat_noi_dung(p, 'Tràn sự cố rộng 5 m, ngưỡng +71,00 m, không có bậc hạ sau ngưỡng (P = 0), dốc nước dài 32 m với '
                    'i = 0,08 > i_{k}: mặt cắt đầu dốc là mặt cắt khống chế, độ sâu đầu dốc bằng độ sâu phân giới h_{k} '
                    'và đường mặt nước trên dốc là đường nước hạ b_{II} bắt đầu từ h_{k} [22]. Phương trình (4-1) – dùng '
                    'cho ngưỡng có bậc hạ phía sau – không dùng để tìm độ sâu đầu dốc ở đây: với P = 1%% (E_{0} = H_{0} = '
                    '6,63 m, φ = 0,95) lưu lượng lớn nhất theo (4-1) chỉ là %s m^{3}/s, xấp xỉ Q = 138,6 m^{3}/s (hệ số '
                    'εm = %s so với giá trị lớn nhất %s ứng với φ = 0,95), tức dòng chảy ở đầu dốc ở trạng thái phân '
                    'giới. Lấy độ sâu đầu dốc h = h_{k} = %s m (P = 1%%, Q = 138,6 m^{3}/s) và %s m (P = 0,2%%, '
                    'Q = 162,7 m^{3}/s); đường mặt nước tính theo (4-2) bắt đầu từ 0,99h_{k}. Trên dốc dòng chảy tăng '
                    'tốc, cuối dốc h = %s ÷ %s m, v = %s ÷ %s m/s. Chiều cao tường bên yêu cầu (h_{hk} + 0,5 m) giảm từ '
                    'khoảng %s m ở đầu dốc (lũ kiểm tra; %s m với lũ thiết kế) xuống %s m ở cuối dốc. Hồ sơ hiện có không '
                    'ghi chiều cao tường bên dốc nước tràn sự cố; kết cấu dốc còn tốt và đã làm việc qua các mùa lũ nên '
                    'đồ án không thay đổi kết cấu dốc nước tràn sự cố, nhưng cần đối chiếu chiều cao tường thực tế (hồ sơ '
                    'hoàn công, đo đạc) với yêu cầu trên.'
                 % (so(a['Qmax_41'], 2), so(a['m_tuong_duong'], 3), so(a['m_max'], 3), so(a['hk'], 2), so(b['hk'], 2),
                    so(a['h_cuoi'], 2), so(b['h_cuoi'], 2), so(a['v_cuoi']), so(b['v_cuoi']), so(b['tuong_dau']),
                    so(a['tuong_dau']), so(b['tuong_cuoi'])))
    ghi('Mục 4.3: tràn sự cố – độ sâu đầu dốc lấy bằng h_k (P = 0, mặt cắt khống chế); nêu rõ (4-1) không có nghiệm '
        'với P = 1%% (Q_max = %s m3/s); cuối dốc h = %s ÷ %s m; tường yêu cầu %s → %s m.'
        % (so(a['Qmax_41']), so(a['h_cuoi'], 2), so(b['h_cuoi'], 2), so(b['tuong_dau']), so(b['tuong_cuoi'])))

    # --- neo ngưỡng tràn chính (mục 4.5.3)
    p = tim_doan(body, 'Neo thép CB400-V đường kính 25 mm bố trí 2 hàng')
    dat_noi_dung(p, 'Neo thép CB400-V đường kính 25 mm bố trí 2 hàng song song tim tràn trong phạm vi bản đáy ngưỡng, '
                    'khoảng cách 1,5 m theo mỗi hàng (n = 1,33 thanh/m dài ngưỡng); lỗ khoan D76 xuyên qua khối ngưỡng, '
                    'ngàm 2,0 m vào đá lớp II, bơm vữa xi măng không co ngót kín toàn bộ lỗ – thanh neo được ngàm cả '
                    'trong bê tông ngưỡng và trong đá. Đây là neo thụ động (không căng trước): lực kéo trong thanh chỉ '
                    'xuất hiện khi mặt tiếp xúc bê tông – đá bắt đầu chuyển vị; do mặt tiếp xúc nhám, chuyển vị trượt kéo '
                    'theo chuyển vị tách, thanh neo bị kéo và tạo lực ép bổ sung lên mặt trượt (cơ chế ma sát – cắt). '
                    'Cơ chế này là cơ sở của mô hình chịu cắt tại mặt tiếp giáp có cốt thép xuyên qua trong EN 1992-1-1, '
                    'mục 6.2.5 [40], với điều kiện mặt tiếp xúc nhám và thanh được neo đủ ở cả hai phía mặt trượt; khả '
                    'năng chịu cắt chốt của thanh được bỏ qua – thiên an toàn. Khi bỏ qua hoàn toàn lực dính (c = 0), '
                    'điều kiện ổn định:')
    p = tim_doan(body, 'Trong đó: As = 4,91 cm')
    dat_noi_dung(p, 'Trong đó: n = 1,33 thanh/m; f = 0,60 – hệ số ma sát bê tông – đá như trường hợp không neo; ΣV, ΣH như '
                    'Bảng 4.5; N_{tt} – khả năng chịu kéo thiết kế của một neo, lấy giá trị nhỏ nhất của ba điều kiện: '
                    'thép chảy dẻo, bám dính vữa – đá và bám dính thép – vữa trên đoạn ngàm:')
    ct6 = [x for x in body.iterchildren(qn('w:p')) if van_ban(x).strip() == '(4-6)']
    assert len(ct6) == 1
    ct6_moi = C.doan_cong_thuc(ct6[0], ['K_{neo} = ', C.phan_so('f·(ΣV + n·N_{tt})', 'ΣH'), ' ≥ [K]'], '4-6')
    ct6[0].addprevious(ct6_moi)
    ct6[0].getparent().remove(ct6[0])
    ct_cu = [x for x in body.iterchildren(qn('w:p')) if van_ban(x).strip() == '(4-7)']
    assert len(ct_cu) == 1
    ct_moi = C.doan_cong_thuc(ct_cu[0], 'N_{tt} = min(A_{s}·R_{s} ; N_{bd} ; N_{bs}) ;  N_{bd} = π·D·L_{ng}·τ_{bd} ;  '
                                        'N_{bs} = π·d·L_{ng}·τ_{bs}', '4-7')
    ct_cu[0].addprevious(ct_moi)
    ct_cu[0].getparent().remove(ct_cu[0])
    p = tim_doan(body, 'với D = 0,076 m, Lng = 2,0 m')
    dat_noi_dung(p, 'với A_{s} = 4,91 cm^{2} (D25), R_{s} = 350 MPa – cường độ chịu kéo tính toán của thép CB400-V → '
                    'A_{s}·R_{s} = 171,8 kN; D = 0,076 m – đường kính lỗ khoan; d = 0,025 m; L_{ng} = 2,0 m. Cường độ '
                    'bám dính giới hạn vữa – '
                    'đá theo [41]: đá cát kết phong hóa 0,7 ÷ 0,8 MPa, đá phiến sét mềm 0,2 ÷ 0,8 MPa; với đá cát bột kết '
                    'phong hóa vừa lấy giới hạn 0,8 MPa và hệ số an toàn 2 → τ_{bd} = 0,40 MPa → N_{bd} = 191 kN. Bám dính '
                    'thép gân – vữa xi măng lấy thận trọng τ_{bs} = 2,0 MPa → N_{bs} = 314 kN. Như vậy N_{tt} = '
                    'min(171,8; 191; 314) = 171,8 kN (thép quyết định). Lực kéo cần trong mỗi thanh để hệ số ổn định '
                    'đạt [K] và hệ số ổn định khi mọi neo làm việc đến N_{tt} cho trong Bảng 4.6.')
    p = tim_doan(body, 'Neo làm việc với hệ số sử dụng thấp')
    dat_noi_dung(p, 'Lực kéo cần trong mỗi thanh để đạt [K] chỉ 40 ÷ 57 kN, bằng 23 ÷ 33% N_{tt}, tương ứng ứng suất '
                    'trong thép 82 ÷ 117 MPa và cường độ bám dính vữa – đá cần huy động 0,08 ÷ 0,12 MPa – nhỏ hơn cả cận '
                    'dưới 0,2 MPa của đá yếu [41], nên kết luận không nhạy với giả thiết τ_{bd}. Với neo thụ động, lực kéo '
                    'này huy động được khi mặt tiếp xúc mở chỉ khoảng 0,1 ÷ 0,2 mm (giãn dài đàn hồi của thanh trên chiều '
                    'dài truyền lực ≈ 10d mỗi phía mặt trượt) – chuyển vị rất nhỏ, không ảnh hưởng làm việc của ngưỡng. '
                    'Hệ số K có neo 2,23 / 1,81 trong Bảng 4.6 ứng với trạng thái giới hạn khi mọi neo đạt N_{tt}, dùng '
                    'để đánh giá dự trữ. Giải pháp neo là đề xuất có điều kiện: trước khi thi công đại trà phải thí '
                    'nghiệm kéo nhổ tối thiểu 3 neo đến 100 kN (≈ 1,7 lần lực kéo cần lớn nhất) để xác nhận bám dính và '
                    'khoan lõi kiểm tra mặt tiếp xúc bê tông – đá; khi hồ sơ hoàn công cho thấy ngưỡng đã có thép neo '
                    'hoặc chân khay ngàm vào đá thì có thể giảm số lượng hoặc bỏ hạng mục này. Khối lượng dự kiến: 36 '
                    'thanh neo, khoảng 144 m khoan và 0,6 t thép (Bảng 7.1).')
    ghi('Mục 4.5.3: neo D25 – nêu rõ neo thụ động, cơ chế ma sát – cắt (EN 1992-1-1 mục 6.2.5 [40]); khả năng thiết kế '
        'N_tt = min(thép, bám dính vữa – đá theo FHWA [41], bám dính thép – vữa) – công thức (4-7) mới; lực cần 40 ÷ 57 '
        'kN = 23 ÷ 33% N_tt, bám dính cần 0,08 ÷ 0,12 MPa, độ mở 0,1 ÷ 0,2 mm; K 2,23/1,81 là trạng thái giới hạn; thí '
        'nghiệm kéo nhổ trước thi công đại trà.')


# =========================================================================== CHƯƠNG 5
def chuong5(body):
    p = tim_doan(body, 'Từ kết quả mục 2.4 và phương án chọn ở mục 3.2.1')
    thay(p, 'vì vậy đồ án quy định dự phòng lún 2 cm khi thi công (cao độ hoàn thiện +78,85 ÷ +78,87 m, dung sai chỉ '
            'cho phép dương), nghiệm thu bằng thủy chuẩn và theo dõi bằng mốc lún mặt đập (mục 7.5).',
         'vì vậy đồ án quy định dự phòng lún 2 cm khi thi công: cao trình ghi trên bản vẽ là cao trình thiết kế sau lún '
         '(mép thượng lưu +78,85 m, mép hạ lưu tấm +78,76 m); cao độ hoàn thiện khi nghiệm thu (trước lún) không thấp hơn '
         '+78,87 m tại mép thượng lưu và +78,78 m tại mép hạ lưu tấm (đất đắp bù đến +78,62 m), dung sai chỉ cho phép '
         'dương; nghiệm thu bằng thủy chuẩn tại vị trí khống chế (mép thượng lưu, chân tường chắn sóng) và theo dõi bằng '
         'mốc lún mặt đập (mục 7.5).')
    p = tim_doan(body, 'Cào bóc lớp mặt đá dăm láng nhựa hư hỏng, san gạt')
    thay(p, 'đầm đạt K ≥ 0,97 đến cao trình +78,60 m tại mép thượng lưu,',
         'đầm đạt K ≥ 0,97 đến cao trình thiết kế +78,60 m tại mép thượng lưu (khi thi công đắp đến +78,62 m – kể dự '
         'phòng lún 2 cm),')
    p = tim_doan(body, 'Rải lớp cấp phối đá dăm đệm dày 5 cm')
    thay(p, '; dự phòng lún 2 cm;', '; cao độ nghiệm thu trước lún ≥ +78,87 m tại mép thượng lưu (dự phòng lún 2 cm);')
    ghi('Mục 5.1.1, 5.1.3: tách cao trình thiết kế sau lún (+78,85 / +78,76) và cao độ nghiệm thu trước lún (≥ +78,87 / '
        '+78,78; đất đắp đến +78,62).')

    p = tim_doan(body, 'Đặc trưng dòng chảy tràn theo chiều dài đường chảy cho trong Bảng 5.5')
    dat_noi_dung(p, 'Đặc trưng dòng chảy tràn theo chiều dài đường chảy cho trong Bảng 5.5. Khi chưa có rãnh, đường chảy '
                    'dài bằng cả mái (≈ 34 m) cho ứng suất tiếp khoảng 6,6 Pa, vượt ứng suất tiếp cho phép của đất không '
                    'có lớp phủ – phù hợp với hiện tượng xói rãnh trên mái. Khi chia mái thành ô 4 × 4 m bằng rãnh xiên, '
                    'đường chảy dài nhất ≈ 4 m, ứng suất tiếp còn 1,8 Pa – giảm khoảng 3,6 lần; với tỷ lệ mưa 30 phút / '
                    'mưa ngày thay đổi 20 ÷ 40% (i = 77 ÷ 153 mm/h), ứng suất tiếp trên ô 4 × 4 m là 1,4 ÷ 2,2 Pa, trong khi '
                    'đường chảy dài cả mái cho 5,1 ÷ 7,8 Pa. Ứng suất tiếp cho phép của đất không có lớp phủ phụ thuộc loại '
                    'đất và độ chặt: khoảng 1,8 ÷ 2,3 Pa với đất pha cát, đất bụi và khoảng 3,6 Pa với đất thịt chặt [23]; '
                    'khi cỏ đã phủ kín, mái chịu được ứng suất tiếp lớn hơn nhiều lần. Vì vậy ô 4 × 4 m đủ chống xói khi '
                    'cỏ đã phủ kín; trong giai đoạn cỏ chưa phát triển, ứng suất tiếp khi mưa lớn (≈ 2,2 Pa) xấp xỉ hoặc '
                    'vượt giới hạn dưới của lớp đất màu mới rải, nên phải bảo vệ tạm: trồng cỏ bằng vầng (hoặc gieo hạt '
                    'có phủ thảm xơ dừa cố định bằng ghim) vào đầu mùa mưa, kiểm tra và đắp bù chỗ xói sau mỗi trận mưa lớn '
                    'cho đến khi cỏ phủ kín. Kích thước ô 4 × 4 m là cơ sở định lượng của giải pháp: mức giảm ứng suất tiếp '
                    '(khoảng 3,6 lần) không phụ thuộc giả thiết cường độ mưa.')
    ghi('Mục 5.5.1: kết luận chống xói có điều kiện – τ = 1,4 ÷ 2,2 Pa so với [τ] ≈ 1,8 ÷ 2,3 Pa (đất pha cát, đất màu) '
        'và 3,6 Pa (đất thịt chặt) theo Chow [23]; bảo vệ tạm bằng cỏ vầng/thảm xơ dừa khi cỏ chưa phủ kín.')

    p = tim_doan(body, 'Kiểm tra chéo (Bảng PL6)')
    thay(p, 'tìm lại trên lưới tâm mịn gấp đôi không phát hiện cung trượt nguy hiểm hơn.',
         'tìm lại trên lưới tâm mịn gấp đôi chưa tìm thấy cung trượt bất lợi hơn trong phạm vi lưới tâm và bán kính đã '
         'khảo sát.')
    dat_o(o_theo_chu(body, 'Không có cung nguy hiểm hơn'), 'Chưa thấy cung bất lợi hơn trong phạm vi khảo sát')
    ghi('Mục 5.3, Bảng PL6: lưới tâm mịn – "chưa tìm thấy cung bất lợi hơn trong phạm vi khảo sát".')


# =========================================================================== CHƯƠNG 6
def chuong6(body):
    p = tim_doan(body, 'Khi cửa van mở hoàn toàn và mực nước hồ thấp')
    dat_noi_dung(p, 'Khi cửa van mở hoàn toàn và cửa vào chưa bị ngập, cửa vào cống làm việc như đập tràn đỉnh rộng chảy '
                    'không ngập. Theo [22], [27], cống chảy không áp ở cửa vào khi cột nước trước cống H ≤ (1,2 ÷ '
                    '1,4)·h_{c}, với h_{c} = 2,0 m – chiều cao cống (trị số lớn ứng với cửa vào có mép lượn), tức H ≤ '
                    '2,4 ÷ 2,8 m:')
    p = tim_doan(body, 'Ở MNC: H0 = 1,90 m')
    thay(p, 'Để tháo được 10 m3/s cần H0 = 2,34 m, tức mực nước hồ ≥ +67,94 m.',
         'Để tháo được 10 m^{3}/s cần H_{0} = 2,34 m (= 1,17h_{c}, vẫn trong giới hạn chảy không áp ở cửa vào nên công '
         'thức (6-1) áp dụng được), tức mực nước hồ ≥ +67,94 m; kiểm tra thêm theo sơ đồ cửa vào ngập (chảy qua lỗ, '
         'Q = μ·b·h_{c}·√(2g(H_{0} − ε·h_{c})), μ = 0,59, ε = 0,62 – thiên an toàn) cho +67,97 m, sai khác 0,03 m nên '
         'kết luận không đổi.')
    ghi('Mục 6.3.1: điều kiện áp dụng công thức (6-1) sửa thành H ≤ (1,2 ÷ 1,4)h_c = 2,4 ÷ 2,8 m; H0 = 2,34 m = 1,17h_c '
        'vẫn chảy không áp; kiểm tra sơ đồ chảy qua lỗ cho +67,97 m.')


# =========================================================================== CHƯƠNG 7
def chuong7(body):
    _, t71 = tim_bang(body, 'Bảng 7.1 ')
    r = [tr for tr in hang(t71) if van_ban(tr.findall(qn('w:tc'))[1]).startswith('Đất đắp bù mặt đập')]
    assert len(r) == 1
    tcs = r[0].findall(qn('w:tc'))
    assert van_ban(tcs[3]).strip() == '1.610'
    dat_o(tcs[1], 'Đất đắp bù mặt đập K ≥ 0,97 (đến +78,62 m, kể dự phòng lún 2 cm; dốc 2%)')
    dat_o(tcs[3], '1.700')
    p = tim_doan(body, 'Khối lượng khoan phụt tổng hợp từ Bảng 2-2 TKCS')
    thay(p, 'đất đắp bù từ +78,20 m sau cào bóc đến +78,60 m dốc 2%)',
         'đất đắp bù từ +78,20 m sau cào bóc đến +78,62 m – kể dự phòng lún 2 cm, dốc 2%)')
    p = tim_doan(body, 'Bê tông mặt đập, rãnh, gờ chắn bánh và mặt đường quản lý')
    thay(p, 'thí nghiệm kéo kiểm tra tối thiểu 3 neo với lực kéo thử 100 kN (≈ 1,7 lần lực kéo yêu cầu lớn nhất 57 kN).',
         'trước khi thi công đại trà, thí nghiệm kéo nhổ tối thiểu 3 neo với lực kéo thử 100 kN (≈ 1,7 lần lực kéo yêu '
         'cầu lớn nhất 57 kN) để xác nhận bám dính vữa – đá.')
    p = tim_doan(body, 'Kiểm tra chất lượng: khoan kiểm tra (3 hố)')
    thay(p, 'q ≲ 0,8 Lu', 'q ≤ 0,8 Lu')
    p = tim_doan(body, 'Trình tự bắt buộc: (i) khoan phụt mang tràn sự cố')
    thay(p, 'trồng cỏ vào đầu mùa mưa để cỏ bén rễ;',
         'trồng cỏ bằng vầng (hoặc gieo hạt có phủ thảm xơ dừa) vào đầu mùa mưa để cỏ bén rễ và bảo vệ mặt ô khi cỏ chưa '
         'phủ kín (mục 5.5.1);')
    ghi('Chương 7: Bảng 7.1 đất đắp bù 1.610 → 1.700 m3 (kể dự phòng lún 2 cm); thí nghiệm kéo nhổ neo trước thi công '
        'đại trà; trồng cỏ vầng/thảm xơ dừa.')


# =========================================================================== CHƯƠNG 8
def chuong8(body, kq_lu):
    p = tim_doan(body, 'Chiều dày màng có ảnh hưởng vừa phải')
    thay(p, 'η tăng khoảng 10%;', 'η tăng khoảng 10 điểm phần trăm;')
    thay(p, 'η tăng thêm khoảng 6%', 'η tăng thêm khoảng 6 điểm phần trăm')
    p = tim_doan(body, 'Đoạn còn lại (khoảng 35 m) có thể giảm còn 2 hàng')
    thay(p, 'hiệu quả chống thấm chỉ giảm khoảng 6% (η = 47% so với 54%, MNLTK);',
         'hiệu quả chống thấm chỉ giảm khoảng 6 điểm phần trăm (η = 47,4% so với 53,6%, MNLTK, K lớp II = 1·10^{-5} '
         'cm/s);')

    def eta(K, Zw, pa):
        return kq_lu['%.1e_%.2f_%s' % (K, Zw, pa)]['eta'] * 100
    e65_tk, e65_hl = eta(6.5e-5, 77.63, '3h_TK'), eta(6.5e-5, 77.63, '2h_A_B_day_69_4')
    p = tim_doan(body, 'Cao trình đáy màng xác định theo TCVN 8216:2018 mục 9.2.2.9')
    dat_noi_dung(p, 'Cao trình đáy màng xác định theo TCVN 8216:2018 mục 9.2.2.9 trên cơ sở thí nghiệm ép nước tại các '
                    'hố phụt thí nghiệm: nếu lớp II có lượng mất nước q ≤ 5 Lu ngay từ mặt lớp thì đáy màng thấp hơn mặt '
                    'lớp II 3,0 m (+69,4 m tại mặt cắt đại diện, so với +68,3 m của TKCS), nhưng không sâu hơn đáy màng '
                    'TKCS; nếu lớp II nứt nẻ (q > 5 Lu) thì đáy màng thấp hơn vị trí đạt q = 5 Lu 3,0 m. Các trị số η '
                    'của chuyên đề ứng với giả thiết K lớp II = 1·10^{-5} cm/s (≈ 0,8 Lu); khi đó nâng đáy màng lên '
                    '+69,4 m không làm giảm η (%s%% với 3 hàng, %s%% với 2 hàng, MNLTK). Tiêu chí 5 Lu tương ứng K lớp II '
                    'tới khoảng 6,5·10^{-5} cm/s; tính lại (Bảng 8.8) cho thấy khi lớp II thấm hơn, η của mọi phương án '
                    'giảm mạnh vì dòng thấm vòng qua chân màng trong lớp II – với K = 6,5·10^{-5} cm/s màng TKCS còn %s%%, '
                    'phương án hợp lý hóa còn %s%% (MNLTK); nâng đáy màng làm η giảm thêm 1 ÷ 2 điểm phần trăm, bỏ hàng C '
                    'giảm 4 ÷ 6 điểm phần trăm. Vì vậy phụt thí nghiệm phải đo được q (K) của lớp II: nếu q ≤ 1 Lu thì '
                    'giữ các kết luận hiệu quả của chuyên đề; nếu 1 Lu < q ≤ 5 Lu thì đáy màng theo mục 9.2.2.9 vẫn phù '
                    'hợp quy phạm nhưng hiệu quả chống thấm chỉ còn khoảng 26 ÷ 42%%, phải tính lại với K đo được và '
                    'cân nhắc giữ 3 hàng trên toàn tuyến;'
                 % (so(eta(1e-5, 77.63, '3h_day_69_4')), so(eta(1e-5, 77.63, '2h_A_B_day_69_4')), so(e65_tk, 0),
                    so(e65_hl, 0)))

    # Bảng 8.8 – sau ghi chú của Bảng 8.7
    gc87 = tim_doan(body, 'Khối lượng tính theo hình học cắt dọc tuyến màng (BV-05)')
    mau_doan = tim_doan(body, 'Khối lượng và hiệu quả của màng theo TKCS và theo phương án hợp lý hóa')
    pdan = doan_moi('Ảnh hưởng của tính thấm lớp II đến hiệu quả các phương án màng – tính lại theo tiêu chí 5 Lu của '
                    'mục 9.2.2.9 (1 Lu ≈ 1,3·10^{-5} cm/s [35]) – cho trong Bảng 8.8.', mau_doan)
    mau_cap = tim_doan(body, 'Bảng 8.7 ', style='Chuthichbang')
    pa = ['3h_TK', '3h_day_69_4', '2h_A_B', '2h_A_B_day_69_4']
    dong = []
    for K, ten_K in ((1e-5, '1·10^{-5}'), (3e-5, '3·10^{-5}'), (6.5e-5, '6,5·10^{-5}'), (1e-4, '1·10^{-4}')):
        dong.append([ten_K, so(K / 1.3e-5)] + ['%s / %s' % (so(eta(K, 77.63, x)), so(eta(K, 78.54, x))) for x in pa])
    cap, tbl, kc = bang_moi(
        'Bảng 8.8 Hiệu quả giảm lưu lượng thấm η (%) của các phương án màng theo hệ số thấm lớp II (mặt cắt đại diện, '
        'MNLTK / MNLKT)',
        ['K lớp II (cm/s)', 'q (Lu)', '3 hàng, đáy TKCS', '3 hàng, đáy theo mục 9.2.2.9', '2 hàng A, B, đáy TKCS',
         '2 hàng A, B, đáy theo mục 9.2.2.9'],
        dong, [1531, 794, 1701, 1701, 1672, 1672], mau_cap, co_chu=19)
    gc = doan_moi('//Đáy TKCS +68,3 m, đáy theo mục 9.2.2.9 +69,4 m (mặt lớp II − 3 m); đỉnh màng +77,00 m, K màng = '
                  '1·10^{-5} cm/s; chương trình kt_lop_II_lu.py.//', gc87)
    chen_sau(gc87, pdan, cap, tbl, kc, gc)

    p = tim_doan(body, 'Có thể hợp lý hóa thiết kế (3 hàng ở đoạn 15 m sát tràn')
    thay(p, 'mà hiệu quả chỉ giảm khoảng 6%, với điều kiện được khẳng định bằng phụt thí nghiệm.',
         'mà hiệu quả chỉ giảm khoảng 6 điểm phần trăm, với điều kiện phụt thí nghiệm khẳng định K màng ≤ 1·10^{-5} '
         'cm/s và ép nước cho lớp II q ≤ 1 Lu; nếu lớp II thấm tới 5 Lu thì hiệu quả của mọi phương án chỉ còn khoảng '
         '26 ÷ 32% (Bảng 8.8) và phải tính lại với K đo được.')
    ghi('Chương 8: "6%%" → "6 điểm phần trăm"; nêu rõ η ứng với K lớp II = 1e-5 cm/s (≈ 0,8 Lu); thêm Bảng 8.8 η theo K '
        'lớp II (1e-5 ÷ 1e-4): với 5 Lu màng TKCS %s%%, hợp lý hóa %s%%; điều kiện q ≤ 1 Lu để giữ kết luận.'
        % (so(e65_tk, 0), so(e65_hl, 0)))


# =========================================================================== KẾT LUẬN, TLTK, PHỤ LỤC
def ket_luan_tltk(body):
    p = tim_doan(body, 'Tràn chính đủ khả năng tháo, tường bên dốc nước đủ cao')
    thay(p, 'đồ án bổ sung neo D25 vào nền (K = 1,81 ÷ 2,23 khi c = 0).',
         'đồ án bổ sung neo D25 thụ động vào nền (lực kéo cần 40 ÷ 57 kN/thanh, không quá 33% khả năng thiết kế; cần '
         'thí nghiệm kéo nhổ xác nhận).')
    p = tim_doan(body, 'Hệ thống rãnh chia mái thành ô 4 × 4 m')
    dat_noi_dung(p, 'Hệ thống rãnh chia mái thành ô 4 × 4 m giảm ứng suất tiếp do dòng chảy tràn từ ≈ 6,6 Pa xuống ≈ 1,8 '
                    'Pa (1,4 ÷ 2,2 Pa tùy cường độ mưa) – đủ chống xói khi cỏ đã phủ kín; trong mùa mưa đầu, khi cỏ chưa '
                    'phát triển, cần bảo vệ tạm (trồng cỏ vầng hoặc phủ thảm xơ dừa) vì ứng suất tiếp xấp xỉ giới hạn '
                    'dưới của đất chưa có lớp phủ; các rãnh có khả năng thoát nước lớn hơn nhiều lần yêu cầu.')
    p = tim_doan(body, 'Chuyên đề: mô hình thấm PTHH cho thấy')
    thay(p, 'giảm khoảng 28% khối lượng khoan, phụt trên cơ sở phụt thí nghiệm.',
         'giảm khoảng 28% khối lượng khoan, phụt mà hiệu quả chỉ giảm khoảng 6 điểm phần trăm, trên cơ sở phụt thí '
         'nghiệm; các trị số hiệu quả ứng với lớp II ít thấm (K ≈ 1·10^{-5} cm/s, ≈ 1 Lu) – nếu lớp II thấm tới 5 Lu '
         'thì hiệu quả của mọi phương án chỉ còn khoảng 26 ÷ 32%.')
    p = tim_doan(body, 'Bổ sung khảo sát: hồ sơ hoàn công tràn chính')
    thay(p, 'thí nghiệm ép nước lớp II tại mang tràn sự cố;',
         'thí nghiệm ép nước lớp II tại mang tràn sự cố (hiệu quả màng phụ thuộc K lớp II); thí nghiệm kéo nhổ neo ngưỡng '
         'tràn chính; đo chiều cao tường bên dốc nước tràn sự cố;')
    ghi('Kết luận, kiến nghị: neo thụ động (lực cần ≤ 33% khả năng); chống xói có điều kiện; hiệu quả màng phụ thuộc K '
        'lớp II; bổ sung thí nghiệm kéo nhổ neo, đo tường bên tràn sự cố.')

    cap = [x for x in body.iterchildren(qn('w:p')) if van_ban(x) == 'TÀI LIỆU THAM KHẢO'][0]
    tbl = cap.getnext()
    while tbl.tag != qn('w:tbl'):
        tbl = tbl.getnext()

    def hang_tl(so_tl):
        for tr in hang(tbl):
            if van_ban(tr.findall(qn('w:tc'))[0]) == '[%d]' % so_tl:
                return tr
        raise ValueError(so_tl)
    dat_o(hang_tl(14).findall(qn('w:tc'))[1],
          'Bộ Khoa học và Công nghệ, //TCVN 10778:2024 – Hồ chứa nước – Xác định các mực nước đặc trưng// (thay thế '
          'TCVN 10778:2015), Hà Nội, 2024.')
    r40 = them_hang(tbl, hang(tbl).index(hang_tl(39)),
                    ['[40]', 'CEN, //EN 1992-1-1:2004 – Eurocode 2: Design of concrete structures – Part 1-1: General '
                             'rules and rules for buildings//, Brussels: European Committee for Standardization, 2004.'])
    them_hang(tbl, hang(tbl).index(r40),
              ['[41]', 'P. J. Sabatini, D. G. Pass and R. C. Bachus, //Ground Anchors and Anchored Systems//, '
                       'Geotechnical Engineering Circular No. 4, FHWA-IF-99-015, Washington, DC: Federal Highway '
                       'Administration, 1999.'])
    ghi('Tài liệu tham khảo: [14] → TCVN 10778:2024; thêm [40] EN 1992-1-1 (mặt tiếp giáp có cốt thép), [41] FHWA GEC-4 '
        '(bám dính vữa – đá).')

    _, t_pl5 = tim_bang(body, 'Bảng PL5 ')
    them_hang(t_pl5, len(hang(t_pl5)) - 2,
              ['kt_tran_su_co.py, kt_lop_II_lu.py', 'Thủy lực tràn sự cố bắt đầu từ h_{k}; hiệu quả màng theo hệ số '
                                                     'thấm lớp II (tiêu chí 5 Lu, Bảng 8.8)'])
    dat_o(o_theo_chu(body, 'Ô 4 × 4 m đủ chống xói'), 'Đủ khi cỏ phủ kín; mùa mưa đầu cần bảo vệ tạm')
    ghi('Phụ lục: Bảng PL5 thêm kt_tran_su_co.py, kt_lop_II_lu.py; dòng độ nhạy mưa ghi kết luận có điều kiện.')


def main(vao, ra, dir_kq):
    kq_tsc = json.load(open(os.path.join(dir_kq, 'kt_tran_su_co.json'), encoding='utf-8'))
    kq_lu = json.load(open(os.path.join(dir_kq, 'kt_lop_II_lu.json'), encoding='utf-8'))
    d = docx.Document(vao)
    body = d.element.body
    n_dong, n_bm = C.go_muc_luc(body)
    ghi('Gỡ mục lục cũ (%d dòng, %d bookmark) để điền lại sau khi sửa.' % (n_dong, n_bm))
    chuong2(body)
    chuong4(body, kq_tsc)
    chuong5(body)
    chuong6(body)
    chuong7(body)
    chuong8(body, kq_lu)
    ket_luan_tltk(body)
    d.save(ra)
    with open(os.path.splitext(ra)[0] + '_nhat_ky.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(C.NHAT_KY) + '\n')
    print('Đã lưu', ra, '–', len(C.NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(*sys.argv[1:4])
