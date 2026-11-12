# -*- coding: utf-8 -*-
"""
RÀ SOÁT ĐỘC LẬP ĐỢT 3 (11/10/2026, sau khi tiếp thu "Gpt góp y.txt") – sửa trực tiếp tệp Word thuyết minh.

  python3 tiep_thu_gop_y_3.py <vào.docx> <ra.docx> <thư mục ket_qua của tinh_toan>
  python3 chen_muc_luc.py <ra.docx> <bản cuối.docx> <bản cuối.pdf>

Nội dung: sửa các chỗ lệch số liệu giữa lời văn và bảng (mục 6.4, 6.5, 8.5.4, 8.6, Bảng 1.1), làm rõ giả thiết
(áp lực đẩy ngược ngưỡng tràn, ước tính lún), thêm dòng K lớp II = 1 Lu vào Bảng 8.8 (kt_lop_II_lu.json),
sửa tên TCVN 9386-1, -5:2025, đánh số lại tài liệu tham khảo theo thứ tự xuất hiện (IEEE – Hướng dẫn trình bày
của Trường), sửa định dạng một số đoạn, bảng. Nhật ký: nhat_ky_tiep_thu_gop_y_3.txt.
"""
import copy
import json
import os
import re
import sys

import docx
from docx.oxml.ns import qn

import cong_cu_docx as C
from cong_cu_docx import ghi, tim_doan, tim_bang, thay, van_ban, hang, them_hang, dat_o


def so(x, n=1):
    return ('%.*f' % (n, x)).replace('.', ',')


# =========================================================================== CHƯƠNG 1
def chuong1(body):
    _, t = tim_bang(body, 'Bảng 1.1 ')
    for tr in hang(t):
        tcs = tr.findall(qn('w:tc'))
        if van_ban(tcs[1]).startswith('Chiều cao đập lớn nhất'):
            dat_o(tcs[1], 'Chiều cao đập lớn nhất: hiện trạng / sau sửa chữa (đáy đập +54,0 m)')
            dat_o(tcs[3], '24,35 / 24,85')
            break
    else:
        raise ValueError('Bảng 1.1: không thấy dòng chiều cao đập')
    ghi('Bảng 1.1: chiều cao đập lớn nhất ghi 24,35 / 24,85 m (hiện trạng / sau sửa chữa, đỉnh +78,35 / +78,85 m, '
        'đáy +54,0 m) thay cho 24,75 m – trị số của TKCS ứng với đỉnh +78,75 m.')

    _, t = tim_bang(body, 'Bảng 1.2 ')
    n = 0
    for tc in t.iter(qn('w:tc')):
        if van_ban(tc).strip() == '-0,00':
            dat_o(tc, '0,00')
            n += 1
    ghi('Bảng 1.2: %d ô "-0,00" ghi lại là "0,00".' % n)


# =========================================================================== CHƯƠNG 2
def chuong2(body):
    p = tim_doan(body, 'Mực nước chết của hồ MNC = +67,50 m')
    thay(p, 'còn trên MNC khoảng 0,45 m', 'còn trên MNC khoảng 0,44 m')

    p = tim_doan(body, 'Lưu ý: hồ sơ BCNCKT dùng hệ số [K] = 1,13')
    thay(p, '[12] và QCVN 02:2022/BXD.', '[12] và QCVN 02:2022/BXD [38].')

    p = tim_doan(body, 'Tháng nào V chạm dung tích chết')
    ppr = p.find(qn('w:pPr'))
    for tag in ('w:ind', 'w:spacing'):
        for e in ppr.findall(qn(tag)):
            ppr.remove(e)

    p = tim_doan(body, 'Tâm trượt nguy hiểm được tìm trên lưới tâm')
    thay(p, 'kiểm tra chéo bằng phương pháp Spencer (thỏa mãn', 'kiểm tra chéo bằng phương pháp Spencer [39] (thỏa mãn')
    ghi('Chương 2: mục 2.2.1 mực nước còn trên MNC 0,44 m (+67,94 − 67,50); đoạn quy ước năm thiếu nước ở mục 2.2.2 '
        'trả về kiểu đoạn văn thường (bỏ thụt lề); dẫn nguồn QCVN 02:2022/BXD ở mục 2.1 và phương pháp Spencer ở '
        'mục 2.5.4 ngay lần nêu đầu tiên.')


# =========================================================================== CHƯƠNG 4
def chuong4(body):
    p = tim_doan(body, 'Ngưỡng tràn chính có mặt cắt hình thang bạt góc')
    thay(p, '(Zdốc + hc − Zđáy) ở mũi (không kể hiệu quả của màng chống thấm, sân phủ – thiên an toàn).',
         '(Z_{dốc} + h − Z_{đáy}) ở mũi, trong đó h – chiều sâu nước trên đáy dốc tại chân ngưỡng: h = 0 khi hồ ở '
         'MNDBT, lấy tròn lên 1,0 m (MNLTK) và 1,2 m (MNLKT) – lớn hơn độ sâu đầu dốc kể cả hàm khí ở mục 4.2.2, '
         '4.2.3; không kể hiệu quả của màng chống thấm, sân phủ (thiên an toàn).')

    _, t = tim_bang(body, 'Bảng 4.5 ')
    rong = [1304, 680, 737, 794, 737, 737, 737, 737, 737, 680, 1192]
    _dat_rong_cot(t, rong)
    ghi('Mục 4.5.1: ghi rõ chiều sâu nước tại chân ngưỡng dùng tính áp lực đẩy ngược (0 / 1,0 / 1,2 m – đúng như '
        'chương trình tính Bảng 4.5) thay cho "h_c"; Bảng 4.5 nới cột c, [K] để số không bị xuống dòng.')


def _dat_rong_cot(t, rong):
    grid = t.find(qn('w:tblGrid')).findall(qn('w:gridCol'))
    assert len(grid) == len(rong), (len(grid), len(rong))
    for g, w in zip(grid, rong):
        g.set(qn('w:w'), str(w))
    for tr in hang(t):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) != len(rong):
            continue
        for tc, w in zip(tcs, rong):
            tw = tc.find(qn('w:tcPr') + '/' + qn('w:tcW'))
            if tw is not None:
                tw.set(qn('w:w'), str(w))
                tw.set(qn('w:type'), 'dxa')


# =========================================================================== CHƯƠNG 5
def chuong5(body):
    p = tim_doan(body, 'Từ kết quả mục 2.4 và phương án chọn ở mục 3.2.1')
    thay(p, 'độ lún bổ sung ước tính chỉ khoảng 0,5 ÷ 1 cm,',
         'độ lún bổ sung ước tính chỉ khoảng 0,5 ÷ 1 cm (tải trọng tăng thêm ≈ 11 kPa trên dải rộng 5 m, ứng suất '
         'giảm dần theo chiều sâu nên ∫Δσ_{z}·dz ≈ 100 kPa·m trên chiều cao đập 24 m; với mô đun biến dạng của đất '
         'đắp đã cố kết E_{oed} ≈ 10 ÷ 20 MPa: s ≈ ∫Δσ_{z}·dz/E_{oed} ≈ 0,5 ÷ 1 cm),')
    ghi('Mục 5.1.1: bổ sung căn cứ ước tính độ lún bổ sung 0,5 ÷ 1 cm (tải ≈ 11 kPa, ∫Δσ·dz ≈ 100 kPa·m, '
        'E_oed ≈ 10 ÷ 20 MPa) cho dự phòng lún 2 cm.')


# =========================================================================== CHƯƠNG 6
def chuong6(body):
    p = tim_doan(body, 'Nhận xét (Bảng 6.3)')
    thay(p, '(Kođ tăng lên 3,4 ÷ 6,5)',
         '(K_{ođ} tăng lên 3,4 ÷ 6,5 khi f ≤ 0,4; trường hợp bất lợi nhất f = 0,5 ở MNLTK vẫn đạt 2,7 > 2)')

    p = tim_doan(body, 'Áp lực nước ngoài')
    thay(p, 'ở cao trình +65,7 ÷ +73,4 m (mục 5.2)', 'ở cao trình +65,6 ÷ +73,0 m (Bảng 5.1)')

    _, t = tim_bang(body, 'Bảng 6.3 ')
    _dat_rong_cot(t, [1247, 624, 737, 907, 794, 737, 794, 850, 1077, 1077])
    ghi('Chương 6: K_ođ khi chia trục hai nhịp ghi đủ phạm vi Bảng 6.3 (3,4 ÷ 6,5 khi f ≤ 0,4; 2,7 khi f = 0,5 ở '
        'MNLTK); đường bão hòa tại tim đập ở mục 6.5 sửa thành +65,6 ÷ +73,0 m theo Bảng 5.1 (số cũ +65,7 ÷ +73,4 '
        'là của mô hình trước khi tính lại với đỉnh +78,85); Bảng 6.3 nới cột f.')


# =========================================================================== CHƯƠNG 7
def chuong7(body):
    p = tim_doan(body, 'Tiến độ thi công dự kiến thể hiện trên Bảng 7.2')
    for r in p.iter(qn('w:r')):
        rp = r.find(qn('w:rPr'))
        if rp is not None:
            for tag in ('w:i', 'w:iCs', 'w:sz', 'w:szCs'):
                for e in rp.findall(qn(tag)):
                    rp.remove(e)
    p = tim_doan(body, 'Theo BCNCKT [3], vốn đầu tư sửa chữa')
    thay(p, '(tăng khoảng 490 m3 đất đắp mặt đập', '(tăng khoảng 490 m^{3} đất đắp mặt đập – khoảng 590 m^{3} nếu kể dự phòng lún 2 cm')
    p = tim_doan(body, 'Không tập kết vật liệu, máy móc nặng trên mái đập')
    thay(p, 'khi chưa hoàn thành lớp bê tông; Tải trọng xe máy', 'khi chưa hoàn thành lớp bê tông; tải trọng xe máy')
    ghi('Chương 7: mục 7.3 ghi thêm khối lượng đất đắp tăng ≈ 590 m³ khi kể dự phòng lún 2 cm (986 × 5 × 0,12); đoạn '
        'tiến độ (sau Bảng 7.2) đồng nhất cỡ chữ, bỏ chữ nghiêng lẫn vào; mục 7.6 sửa chữ hoa sau dấu chấm phẩy.')


# =========================================================================== CHƯƠNG 8
def tach_doan(p, moc):
    """Tách đoạn p tại chuỗi `moc` (đầu đoạn mới), đoạn mới giữ pPr; trả về đoạn mới."""
    runs = [r for r in p.iterchildren(qn('w:r'))]
    chu = [''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in runs]
    toan = ''.join(chu)
    a = toan.index(moc)
    vt = 0
    for i, c in enumerate(chu):
        if a < vt + len(c):
            o = a - vt
            break
        vt += len(c)
    r = runs[i]
    if o > 0:
        r2 = copy.deepcopy(r)
        C._dat_chu(r, chu[i][:o].rstrip())
        C._dat_chu(r2, chu[i][o:])
        r.addnext(r2)
        con = [r2] + runs[i + 1:]
    else:
        con = runs[i:]
    moi = copy.deepcopy(p)
    for e in list(moi):
        if e.tag != qn('w:pPr'):
            moi.remove(e)
    for e in con:
        moi.append(e)          # lxml: append = chuyển run sang đoạn mới
    p.addnext(moi)
    return moi


def chuong8(body, kq_lu, kq_1lu):
    p = tim_doan(body, '(3) Chiều sâu màng')
    tach_doan(p, '(4) Chất lượng màng')

    p = tim_doan(body, 'Chiều dày màng có ảnh hưởng vừa phải')
    thay(p, 'chỉ tăng η thêm 3 ÷ 5% do', 'chỉ tăng η thêm 3 ÷ 5 điểm phần trăm do')

    k = kq_1lu
    e = {pa: (100 * k['1.3e-05_77.63_%s' % pa]['eta'], 100 * k['1.3e-05_78.54_%s' % pa]['eta'])
         for pa in ('3h_TK', '3h_day_69_4', '2h_A_B', '2h_A_B_day_69_4')}
    lo3 = min(e['3h_TK'] + e['3h_day_69_4']); hi3 = max(e['3h_TK'] + e['3h_day_69_4'])
    lo2 = min(e['2h_A_B'] + e['2h_A_B_day_69_4']); hi2 = max(e['2h_A_B'] + e['2h_A_B_day_69_4'])
    p = tim_doan(body, 'Cao trình đáy màng xác định theo TCVN 8216:2018 mục 9.2.2.9')
    z_ra = kq_lu['6.5e-05_78.54_3h_TK']['zseep']
    assert abs(z_ra - 74.95) < 0.01, z_ra
    thay(p, '(MNLTK); nâng đáy màng làm η giảm thêm',
         '(MNLTK), và ở MNLKT dòng thấm vẫn thoát ra trên mái hạ lưu ở cao trình %s m (sát chân mái) với mọi phương '
         'án màng; nâng đáy màng làm η giảm thêm' % ('+' + so(z_ra, 2)))
    thay(p, 'nếu q ≤ 1 Lu thì giữ các kết luận hiệu quả của chuyên đề; nếu 1 Lu < q ≤ 5 Lu thì đáy màng theo mục 9.2.2.9 '
            'vẫn phù hợp quy phạm nhưng hiệu quả chống thấm chỉ còn khoảng 26 ÷ 42%,',
         'nếu q ≤ 1 Lu thì giữ các kết luận hiệu quả của chuyên đề (với q = 1 Lu: η = %s ÷ %s%% với 3 hàng, %s ÷ %s%% với '
         '2 hàng – Bảng 8.8); nếu 1 Lu < q ≤ 5 Lu thì đáy màng theo mục 9.2.2.9 vẫn phù hợp quy phạm nhưng hiệu quả '
         'chống thấm giảm dần, còn 26 ÷ 32%% khi q = 5 Lu,' % (so(lo3), so(hi3), so(lo2), so(hi2)))

    cap, t = tim_bang(body, 'Bảng 8.8 ')
    i1 = [i for i, tr in enumerate(hang(t)) if van_ban(tr.findall(qn('w:tc'))[0]).startswith('1·10')][0]
    them_hang(t, i1, ['1,3·10^{-5}', '1,0'] + ['%s / %s' % (so(e[pa][0]), so(e[pa][1]))
                                              for pa in ('3h_TK', '3h_day_69_4', '2h_A_B', '2h_A_B_day_69_4')])

    p = tim_doan(body, 'Khối lượng tính theo hình học cắt dọc tuyến màng (BV-05)')
    thay(p, 'cùng một cách tính cho các phương án, không kể hố phụt thí nghiệm và hố kiểm tra;',
         'cùng một cách tính cho các phương án (51 hố của TKCS gồm cả 10 hố phụt thí nghiệm; không kể 3 hố kiểm tra); '
         'tổng khoan, phụt theo cách tính này lớn hơn Bảng 2-2 TKCS (479,9 m; 403,4 m) khoảng 5 ÷ 7%, chỉ dùng để so '
         'sánh tương đối giữa các phương án;')
    p = tim_doan(body, 'Có thể hợp lý hóa thiết kế (3 hàng')
    thay(p, 'chỉ còn khoảng 26 ÷ 32% (Bảng 8.8) và phải tính lại với K đo được.',
         'chỉ còn khoảng 26 ÷ 32% (Bảng 8.8), ở MNLKT dòng thấm vẫn thoát ra trên mái hạ lưu, và phải tính lại với K '
         'đo được.')
    ghi('Chương 8: tách "(4) Chất lượng màng" thành đoạn riêng; "3 ÷ 5%%" → "3 ÷ 5 điểm phần trăm"; Bảng 8.8 thêm dòng '
        'K lớp II = 1,3·10⁻⁵ cm/s (1 Lu): η = %s ÷ %s%% (3 hàng), %s ÷ %s%% (2 hàng) – chương trình kt_lop_II_lu.py; '
        'mục 8.6 sửa khoảng hiệu quả khi 1 Lu < q ≤ 5 Lu (trước ghi 26 ÷ 42%%, chưa gồm q gần 1 Lu) và nêu: với 5 Lu, ở '
        'MNLKT dòng thấm vẫn thoát ra trên mái hạ lưu (+74,95 m) dù có màng; ghi chú Bảng 8.7 '
        'nêu rõ 51 hố gồm hố phụt thí nghiệm và so với tổng khối lượng Bảng 2-2 TKCS.'
        % (so(lo3), so(hi3), so(lo2), so(hi2)))


# =========================================================================== KẾT LUẬN, TLTK
def ket_luan_tltk(body):
    p = tim_doan(body, 'Chuyên đề: mô hình thấm PTHH cho thấy')
    thay(p, '(K ≈ 1·10-5 cm/s, ≈ 1 Lu)', '(K ≈ 1·10^{-5} cm/s, ≈ 0,8 Lu)')

    cap = [x for x in body.iterchildren(qn('w:p')) if van_ban(x) == 'TÀI LIỆU THAM KHẢO'][0]
    tbl = cap.getnext()
    for tr in hang(tbl):
        tcs = tr.findall(qn('w:tc'))
        if van_ban(tcs[0]) == '[12]':
            dat_o(tcs[1], 'Bộ Khoa học và Công nghệ, //TCVN 9386-1:2025 – Thiết kế kết cấu chịu động đất – Phần 1: Quy '
                          'định chung, tác động động đất và quy định cho nhà// và //TCVN 9386-5:2025 – Thiết kế kết cấu '
                          'chịu động đất – Phần 5: Nền móng, kết cấu chắn và các vấn đề địa kỹ thuật// (thay thế '
                          'TCVN 9386:2012), Hà Nội, 2025.')
    for tr in hang(tbl):
        tcs = tr.findall(qn('w:tc'))
        if van_ban(tcs[1]).startswith('S. Spencer'):
            thay(tcs[1].find(qn('w:p')), 'S. Spencer,', 'E. Spencer,')
    ghi('Tài liệu Spencer (1967): sửa tên tác giả "S. Spencer" thành "E. Spencer".')
    ghi('Kết luận: lớp II ít thấm ghi ≈ 0,8 Lu (thống nhất với mục 8.2.2, 8.6). Tài liệu [12]: sửa tên đúng của '
        'TCVN 9386-1:2025, TCVN 9386-5:2025 ("Thiết kế kết cấu chịu động đất").')


# =========================================================================== ĐÁNH SỐ LẠI TLTK (IEEE)
RX_TRICH = re.compile(r'\[(\d{1,2}(?:\s*[,–-]\s*\d{1,2})*)\]')


def _so_trong(nhom):
    out = []
    for phan in re.split(r'\s*,\s*', nhom):
        if re.search(r'[–-]', phan):
            a, b = [int(x) for x in re.split(r'\s*[–-]\s*', phan)]
            out += list(range(a, b + 1))
        else:
            out.append(int(phan))
    return out


def danh_so_lai(body):
    cap = [x for x in body.iterchildren(qn('w:p')) if van_ban(x) == 'TÀI LIỆU THAM KHẢO'][0]
    tbl = cap.getnext()
    assert tbl.tag == qn('w:tbl')
    so_tl = len(hang(tbl))
    # thứ tự xuất hiện lần đầu (bỏ bảng TLTK; đề mục/chú thích tính theo vị trí trong văn bản)
    thu_tu = []
    for t in body.iter(qn('w:t')):
        anc = t
        trong_tltk = False
        while anc is not None:
            if anc is tbl:
                trong_tltk = True
                break
            anc = anc.getparent()
        if trong_tltk:
            continue
        for m in RX_TRICH.finditer(t.text or ''):
            for n in _so_trong(m.group(1)):
                if n not in thu_tu:
                    thu_tu.append(n)
    thieu = [n for n in range(1, so_tl + 1) if n not in thu_tu]
    if thieu or len(thu_tu) != so_tl:
        raise ValueError('Tài liệu chưa được trích: %s / thứ tự %s' % (thieu, thu_tu))
    map_ = {cu: i + 1 for i, cu in enumerate(thu_tu)}
    doi = {k: v for k, v in map_.items() if k != v}

    def _sub(m):
        nums = _so_trong(m.group(1))
        moi = [map_[n] for n in nums]
        if len(nums) == 1:
            return '[%d]' % moi[0]
        return '[' + ', '.join(str(x) for x in moi) + ']'
    n_cho = 0
    for t in body.iter(qn('w:t')):
        anc = t; trong_tltk = False
        while anc is not None:
            if anc is tbl:
                trong_tltk = True
                break
            anc = anc.getparent()
        if trong_tltk or not RX_TRICH.search(t.text or ''):
            continue
        moi = RX_TRICH.sub(_sub, t.text)
        if moi != t.text:
            n_cho += len(RX_TRICH.findall(t.text))
            t.text = moi
    # sắp xếp lại bảng TLTK
    rows = hang(tbl)
    theo_so = {}
    for tr in rows:
        tc0 = tr.findall(qn('w:tc'))[0]
        cu = int(van_ban(tc0).strip('[] '))
        theo_so[map_[cu]] = tr
        tdau = list(tc0.iter(qn('w:t')))
        tdau[0].text = '[%d]' % map_[cu]
        for x in tdau[1:]:
            x.text = ''
    neo = rows[0].getprevious()
    for tr in rows:
        tbl.remove(tr)
    for k in sorted(theo_so):
        tbl.append(theo_so[k])
    ghi('Tài liệu tham khảo: đánh số lại theo thứ tự trích dẫn lần đầu (IEEE, Hướng dẫn trình bày ĐATN của Trường) – '
        'thứ tự cũ %s; đổi số: %s (%d chỗ trích dẫn).' % (thu_tu, ', '.join('[%d]→[%d]' % kv for kv in sorted(doi.items())),
                                                       n_cho))
    return map_


def main(vao, ra, dir_kq):
    kq_lu = json.load(open(os.path.join(dir_kq, 'kt_lop_II_lu.json'), encoding='utf-8'))
    kq_1lu = kq_lu
    d = docx.Document(vao)
    body = d.element.body
    n_dong, n_bm = C.go_muc_luc(body)
    ghi('Gỡ mục lục cũ (%d dòng, %d bookmark) để điền lại sau khi sửa.' % (n_dong, n_bm))
    chuong1(body)
    chuong2(body)
    chuong4(body)
    chuong5(body)
    chuong6(body)
    chuong7(body)
    chuong8(body, kq_lu, kq_1lu)
    ket_luan_tltk(body)
    map_ = danh_so_lai(body)
    d.save(ra)
    with open(os.path.splitext(ra)[0] + '_nhat_ky.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(C.NHAT_KY) + '\n')
    with open(os.path.splitext(ra)[0] + '_map_tltk.json', 'w', encoding='utf-8') as f:
        json.dump(map_, f)
    print('Đã lưu', ra, '–', len(C.NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(*sys.argv[1:4])
