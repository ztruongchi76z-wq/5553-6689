# -*- coding: utf-8 -*-
"""
HIỆU CHỈNH THUYẾT MINH SAU RÀ SOÁT (10/10/2026)
Sửa trực tiếp trên tệp Word đã dựng (không chạy lại build.py), ghi nhật ký từng thay đổi.

  python3 hieu_chinh_docx.py <vào.docx> <ra.docx>

Nội dung sửa (chi tiết xem RA_SOAT_THUYET_MINH.md):
  1. Sửa lỗi: tham chiếu mục sai (5.1.2), mô tả độ sâu đầu dốc tràn sự cố (4.3),
     công thức (2-1) (dấu gạch trên) và (8-1) (ký tự lạ \\{}).
  2. Dẫn số hình, bảng trong lời văn cho mọi hình, bảng chưa được nhắc tới (yêu cầu mẫu Trường).
  3. Ghi chú cột V ở Bảng PL2; bổ sung, sắp xếp lại danh mục từ viết tắt theo A, B, C.
  4. Bìa: chữ "BỘ NÔNG NGHIỆP VÀ MÔI TRƯỜNG" trên một dòng; tên đề tài ngắt 3 dòng cân đối.
"""
import copy
import re
import sys

import docx
from docx.oxml.ns import qn
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
NHAT_KY = []


def ghi(msg):
    NHAT_KY.append(msg)
    print(msg)


# ---------------------------------------------------------------------------
# Tiện ích đoạn văn và run
# ---------------------------------------------------------------------------
def van_ban(p):
    """Chữ của đoạn (chỉ w:t, không kể công thức)."""
    return ''.join(t.text or '' for t in p.iter(qn('w:t')))


def tim_doan(body, dau, style=None):
    """Đoạn duy nhất (con trực tiếp của body) có chữ bắt đầu bằng `dau`."""
    kq = []
    for p in body.iterchildren(qn('w:p')):
        if van_ban(p).startswith(dau):
            if style is None or kieu(p) == style:
                kq.append(p)
    if len(kq) != 1:
        raise ValueError('Tìm thấy %d đoạn bắt đầu bằng %r' % (len(kq), dau))
    return kq[0]


def kieu(p):
    s = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
    return s.get(qn('w:val')) if s is not None else None


def tao_run(text, vert=None, dam=False):
    r = docx.oxml.OxmlElement('w:r')
    if vert or dam:
        rpr = docx.oxml.OxmlElement('w:rPr')
        if dam:
            rpr.append(docx.oxml.OxmlElement('w:b'))
        if vert:
            va = docx.oxml.OxmlElement('w:vertAlign')
            va.set(qn('w:val'), vert)
            rpr.append(va)
        r.append(rpr)
    for i, phan in enumerate(text.split('\n')):
        if i:
            r.append(docx.oxml.OxmlElement('w:br'))
        if phan:
            t = docx.oxml.OxmlElement('w:t')
            t.text = phan
            t.set(XML_SPACE, 'preserve')
            r.append(t)
    return r


def runs_tu_markup(s):
    """Chuyển chuỗi có ^{..} (chỉ số trên), _{..} (chỉ số dưới), **..** (đậm), \\n (xuống dòng)
    thành danh sách w:r."""
    out = []
    for m in re.finditer(r'\^\{([^}]*)\}|_\{([^}]*)\}|\*\*([^*]*)\*\*|([^\^_*]+|[\^_*])', s):
        if m.group(1) is not None:
            out.append(tao_run(m.group(1), 'superscript'))
        elif m.group(2) is not None:
            out.append(tao_run(m.group(2), 'subscript'))
        elif m.group(3) is not None:
            out.append(tao_run(m.group(3), dam=True))
        else:
            out.append(tao_run(m.group(4)))
    return out


def cac_run_chu(p):
    """Các w:r con trực tiếp có chữ (w:t)."""
    return [r for r in p.iterchildren(qn('w:r')) if r.find(qn('w:t')) is not None]


def dat_chu(r, s):
    ts = r.findall(qn('w:t'))
    ts[0].text = s
    ts[0].set(XML_SPACE, 'preserve')
    for t in ts[1:]:
        r.remove(t)


def thay_trong_doan(p, cu, moi_markup):
    """Thay đoạn chữ `cu` (có thể trải qua nhiều run) bằng `moi_markup`."""
    runs = cac_run_chu(p)
    chu = [''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in runs]
    toan = ''.join(chu)
    if toan.count(cu) != 1:
        raise ValueError('Chuỗi %r xuất hiện %d lần trong đoạn: %r' % (cu, toan.count(cu), toan[:80]))
    a = toan.index(cu)
    b = a + len(cu)
    # vị trí run đầu, run cuối
    vt = 0
    i0 = i1 = None
    for i, c in enumerate(chu):
        if i0 is None and a < vt + len(c):
            i0, o0 = i, a - vt
        if b <= vt + len(c):
            i1, o1 = i, b - vt
            break
        vt += len(c)
    r0, r1 = runs[i0], runs[i1]
    dau, cuoi = chu[i0][:o0], chu[i1][o1:]
    moi = runs_tu_markup(moi_markup)
    neo = r0
    if dau:
        # tách phần đầu ra run riêng giữ nguyên định dạng
        rd = copy.deepcopy(r0)
        dat_chu(rd, dau)
        r0.addprevious(rd)
    for r in moi:
        r0.addprevious(r)
    if cuoi:
        rc = copy.deepcopy(r1)
        dat_chu(rc, cuoi)
        r0.addprevious(rc)
    for r in runs[i0:i1 + 1]:
        p.remove(r)


def them_dau(p, markup):
    runs = cac_run_chu(p)
    for r in runs_tu_markup(markup):
        runs[0].addprevious(r)


def them_cuoi(p, markup):
    runs = cac_run_chu(p)
    neo = runs[-1]
    for r in runs_tu_markup(markup):
        neo.addnext(r)
        neo = r


def doan_moi_sau(phan_tu, markup, mau):
    """Chèn đoạn mới (định dạng đoạn sao chép từ `mau`) ngay sau `phan_tu`."""
    p = docx.oxml.OxmlElement('w:p')
    ppr = mau.find(qn('w:pPr'))
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    for r in runs_tu_markup(markup):
        p.append(r)
    phan_tu.addnext(p)
    return p


# ---------------------------------------------------------------------------
# Các nhóm sửa
# ---------------------------------------------------------------------------
def _m(tag, *con, **thuoc):
    e = etree.Element('{%s}%s' % (M, tag))
    for k, v in thuoc.items():
        e.set('{%s}%s' % (M, k), v)
    for c in con:
        e.append(c)
    return e


def _mr(chu, mau_r):
    """m:r có chữ `chu`, sao chép định dạng (font Cambria Math) từ run mẫu."""
    r = copy.deepcopy(mau_r)
    t = r.find('{%s}t' % M)
    t.text = chu
    t.set(XML_SPACE, 'preserve')
    return r


def _sub(goc, chi_so, mau_r):
    return _m('sSub', _m('e', _mr(goc, mau_r)), _m('sub', _mr(chi_so, mau_r)))


def _ngoac(noi_dung, mo, dong):
    dpr = _m('dPr', _m('begChr', val=mo), _m('endChr', val=dong))
    return _m('d', dpr, _m('e', noi_dung))


def sua_cong_thuc(body):
    # (8-1): dựng lại bằng đối tượng ngoặc m:d – bản cũ có ký tự lỗi "\{}H\", "\{}Q\" và dấu "]" đứng riêng
    # làm chỉ số trên (Word hiện sai, LibreOffice báo lỗi cú pháp và cắt mất nửa công thức).
    p81 = [p for p in body.iterchildren(qn('w:p')) if van_ban(p).strip().endswith('(8-1)')][0]
    om = p81.find('{%s}oMath' % M)
    mau_r = copy.deepcopy(om.find('.//{%s}r' % M))
    moi = _m('oMath',
             _ngoac(_sub('K', 'e', mau_r), '[', ']'), _mr(' = ', mau_r),
             _sub('k', 'e', mau_r), _mr('·', mau_r), _sub('A', 'e', mau_r), _mr('·', mau_r),
             _m('sSup', _m('e', _ngoac(_mr('B', mau_r), '[', ']')), _m('sup', _mr('T', mau_r))),
             _ngoac(_mr('B', mau_r), '[', ']'), _mr(' ;  ', mau_r),
             _ngoac(_mr('K', mau_r), '[', ']'), _mr('·', mau_r), _ngoac(_mr('H', mau_r), '{', '}'),
             _mr(' = ', mau_r), _ngoac(_mr('Q', mau_r), '{', '}'))
    om.addprevious(moi)
    om.getparent().remove(om)
    ghi('Công thức (8-1): dựng lại [Ke] = ke·Ae·[B]ᵀ[B] ; [K]·{H} = {Q} bằng đối tượng ngoặc '
        '(bản cũ có ký tự lỗi "\\{}H\\", "\\{}Q\\" và hiển thị sai).')
    # (2-1): run 'Z' (hoặc '·V') + sSub(base = dấu gạch trên đứng một mình) -> base 'Z̄' ('V̄')
    p21 = [p for p in body.iterchildren(qn('w:p')) if van_ban(p).strip().endswith('(2-1)')][0]
    n = 0
    for ssub in list(p21.iter('{%s}sSub' % M)):
        e = ssub.find('{%s}e' % M)
        te = e.find('.//{%s}t' % M)
        if te is not None and te.text == '̄':
            truoc = ssub.getprevious()
            tt = truoc.find('{%s}t' % M)
            chu = tt.text
            te.text = chu[-1] + '̄'
            if len(chu) == 1:
                truoc.getparent().remove(truoc)
            else:
                tt.text = chu[:-1]
            n += 1
    ghi('Công thức (2-1): ghép dấu gạch trên vào Z, V (%d chỗ) – trước đây hiển thị thành ô vuông.' % n)


def sua_loi_noi_dung(body):
    p = tim_doan(body, 'Bề rộng đỉnh đập giữ nguyên B = 5,0 m')
    thay_trong_doan(p, 'ở mục 5.3 và 5.4.', 'ở mục 5.2 và 5.3.')
    ghi('Mục 5.1.2: "kiểm tra thấm và ổn định ... ở mục 5.3 và 5.4" -> "mục 5.2 và 5.3".')
    p = tim_doan(body, 'Tràn sự cố rộng 5 m, ngưỡng +71,00 m, không có bậc hạ')
    thay_trong_doan(
        p,
        'không có bậc hạ sau ngưỡng nên dòng chảy ở đầu dốc xấp xỉ trạng thái phân giới: '
        'h ≈ hk = 4,24 m (P = 1%, Q = 138,6 m3/s) và 4,29 m (P = 0,2%, Q = 162,7 m3/s).',
        'không có bậc hạ sau ngưỡng (P = 0); độ sâu đầu dốc xác định theo phương trình (4-1), '
        'lấy nghiệm dòng xiết: h = 4,24 m (P = 1%, Q = 138,6 m^{3}/s; h_{k} = 4,28 m) và '
        'h = 4,29 m (P = 0,2%, Q = 162,7 m^{3}/s; h_{k} = 4,76 m).')
    ghi('Mục 4.3: sửa "h ≈ hk = 4,24 / 4,29 m" (với Q = 162,7 m3/s thì hk = 4,76 m, không phải 4,29 m) '
        '-> nêu rõ là nghiệm dòng xiết của (4-1), kèm hk = 4,28 / 4,76 m.')


# (tiền tố đoạn, kiểu thao tác, đối số)
DAN_HINH_BANG = [
    ('Với bảng tính lại, dung tích ứng với MNC', 'thay', ('Với bảng tính lại,', 'Với bảng tính lại (Bảng 1.2, Hình 1.1),')),
    ('Theo kết quả khoan khảo sát năm 2022 (10 hố khoan', 'thay',
     ('gồm các lớp sau:', 'gồm các lớp sau (chỉ tiêu cơ lý chủ yếu cho trong Bảng 1.3):')),
    ('Hai bộ số liệu có lưu lượng đỉnh gần như nhau', 'thay', ('Hai bộ số liệu có', 'Hai bộ số liệu (Bảng 1.5) có')),
    ('Nguồn F: bình đồ lòng hồ năm 2022', 'thay',
     ('Bảng đầy đủ (+59 ÷ +85 m) xem Phụ lục.', 'Bảng đầy đủ (+59 ÷ +85 m) xem Bảng PL1 (Phụ lục 1).')),
    ('Kết quả cho thấy: (i) với phương thức tưới', 'thay',
     ('Kết quả cho thấy:', 'Kết quả điều tiết (Bảng 2.2, Hình 2.1) cho thấy:')),
    ('Quy tắc vận hành theo QTVH', 'thay',
     ('cho trong Phụ lục.', 'cho trong Bảng PL2 (Phụ lục 1).')),
    ('Với đường quá trình lũ của BCNCKT (lên nhanh', 'dau',
     'Kết quả điều tiết lũ cho trong Bảng 2.3, Hình 2.2 và Hình 2.3; bảng tính theo thời gian xem '
     'Bảng PL3, Bảng PL4 (Phụ lục 2). '),
    ('Cao trình đỉnh đập yêu cầu do trường hợp mực nước lũ kiểm tra quyết định', 'dau',
     'Kết quả tính chi tiết cho trong Bảng 2.4 và tổng hợp trong Bảng 2.5. '),
    ('Mặt cắt tính toán là mặt cắt lòng sông có chiều cao đập lớn nhất', 'cuoi',
     ' Sơ đồ vật liệu và lưới phần tử hữu hạn thể hiện trên Hình 2.4.'),
    ('Lưu lượng thấm qua mặt cắt lòng sông ở MNDBT là', 'dau',
     'Kết quả tính thấm hiện trạng cho trong Bảng 2.6; lưới dòng thấm khi hồ ở MNLTK thể hiện trên Hình 2.5. '),
    ('Tất cả các trường hợp đều có hệ số an toàn lớn hơn giá trị cho phép', 'dau',
     'Kết quả kiểm tra ổn định cho trong Bảng 2.7; cung trượt nguy hiểm khi thiết bị thoát nước bị tắc '
     'thể hiện trên Hình 2.6. '),
    ('Đồ án chọn phương án Đ3 vì', 'thay', ('Kết cấu mặt đập:', 'Kết cấu mặt đập (Hình 3.1):')),
    ('Độ sâu lớn nhất có kể hàm khí xuất hiện ở cuối đoạn thu hẹp', 'dau',
     'Kết quả tính đường mặt nước cho trong Bảng 4.3 và Hình 4.1. '),
    ('Nhận xét: khi bỏ qua hoàn toàn lực dính', 'thay', ('Nhận xét: khi', 'Nhận xét (Bảng 4.4): khi')),
    ('Phương pháp, sơ đồ vật liệu và điều kiện biên như mục 2.5.', 'thay',
     ('Kết quả cho các trường hợp tính toán trong Bảng 5.1.',
      'Kết quả cho các trường hợp tính toán trong Bảng 5.1; lưới thấm và đường bão hòa khi hồ ở MNLKT '
      'thể hiện trên Hình 5.1, các trường hợp MNDBT và MNLTK xem Hình PL9, Hình PL10 (Phụ lục 4).')),
    ('Tất cả các trường hợp đều thỏa mãn; hệ số ổn định nhỏ nhất của mái hạ lưu', 'thay',
     ('(hạ mực nước, [K] = 1,30).',
      '(hạ mực nước, [K] = 1,30); cung trượt nguy hiểm của hai trường hợp này thể hiện trên Hình 5.3 và '
      'Hình 5.4, các trường hợp còn lại xem Phụ lục 3 (Hình PL1 ÷ Hình PL8).')),
    ('Khi lớp mặt còn giữ được lực dính của đất đắp', 'dau', 'Kết quả tính cho trong Bảng 5.3 và Bảng 5.4. '),
    ('Khi chưa có rãnh, đường chảy dài bằng cả mái', 'dau',
     'Đặc trưng dòng chảy tràn theo chiều dài đường chảy cho trong Bảng 5.5. '),
    ('Các rãnh đều có khả năng thoát lớn hơn nhiều lần', 'thay',
     ('Các rãnh đều có', 'Kết quả kiểm tra (Bảng 5.6) cho thấy các rãnh đều có')),
    ('Mức đầy lớn nhất 55 ÷ 73%', 'thay', ('Mức đầy lớn nhất', 'Kết quả tính (Bảng 6.2, Hình 6.1) cho mức đầy lớn nhất')),
    ('Nhận xét: (i) ở MNDBT, máy VĐ20', 'thay', ('Nhận xét: (i)', 'Nhận xét (Bảng 6.3): (i)')),
    ('Khối lượng chủ yếu được tác giả tính từ kích thước thiết kế', 'cuoi', ' Kết quả tổng hợp trong Bảng 7.1.'),
    ('Tổng tiến độ 06 tháng không kể thời gian', 'thay',
     ('Tổng tiến độ 06 tháng', 'Tiến độ thi công dự kiến thể hiện trên Bảng 7.2; tổng tiến độ 06 tháng')),
    ('Công tác quan trắc thực hiện theo TCVN 8414:2010', 'cuoi',
     ' Nội dung và tần suất quan trắc, kiểm tra đề xuất cho trong Bảng 7.3.'),
    ('Màng được mô hình là một dải chữ nhật đồng nhất', 'dau',
     'Hệ số thấm bão hòa của các vật liệu trong mô hình cho trong Bảng 8.1. '),
    ('Phân tích kết quả:', 'thay',
     ('Phân tích kết quả:', 'Phân tích kết quả (Bảng 8.2; Hình 8.3, Hình 8.4; trường hợp MNDBT xem Hình PL11, '
                            'Hình PL12 ở Phụ lục 4):')),
    ('Các phân tích sau thực hiện với mực nước hồ MNLTK', 'cuoi', ' Kết quả cho trong Bảng 8.3 ÷ Bảng 8.6 và Hình 8.5.'),
    ('Các bài toán được tác giả lập trình tính toán bằng ngôn ngữ Python', 'thay',
     ('để người đọc có thể kiểm tra lại.',
      'để người đọc có thể kiểm tra lại (danh mục chương trình xem Bảng PL5, Phụ lục 5).')),
]

# Đoạn dẫn mới chèn ngay sau đề mục (đề mục chưa có lời văn trước bảng)
DOAN_DAN_MOI = [
    ('Tổng hợp giải pháp được chọn',
     'Các giải pháp sửa chữa, nâng cấp được lựa chọn cho từng hạng mục được tổng hợp trong Bảng 3.2.'),
    ('Khả năng tháo qua ngưỡng',
     'Khả năng tháo của tràn chính theo cột nước, tính theo công thức (2-3) và (2-4), cho trong Bảng 4.2.'),
    ('Nhiệm vụ và các thông số tính toán',
     'Các thông số dùng để kiểm tra thủy lực cống và lực đóng mở cửa van được tổng hợp trong Bảng 6.1.'),
]


def dan_hinh_bang(body):
    for dau, loai, ts in DAN_HINH_BANG:
        p = tim_doan(body, dau)
        if loai == 'thay':
            thay_trong_doan(p, ts[0], ts[1])
        elif loai == 'dau':
            them_dau(p, ts)
        else:
            them_cuoi(p, ts)
    mau = tim_doan(body, 'Với bảng tính lại')  # đoạn kiểu Content làm mẫu định dạng
    for de_muc, chu in DOAN_DAN_MOI:
        h = [p for p in body.iterchildren(qn('w:p'))
             if van_ban(p).strip() == de_muc and (kieu(p) or '').startswith('Heading')]
        if len(h) != 1:
            raise ValueError('Đề mục %r: %d' % (de_muc, len(h)))
        doan_moi_sau(h[0], chu, mau)
    ghi('Dẫn số hình/bảng trong lời văn: %d đoạn được bổ sung, %d đoạn dẫn mới sau đề mục '
        '(mẫu Trường: "Khi đề cập đến các hình và bảng biểu, phải nêu rõ số").'
        % (len(DAN_HINH_BANG), len(DOAN_DAN_MOI)))


def ghi_chu_bang_pl2(body):
    cap = tim_doan(body, 'Bảng PL2 Quan hệ Z ~ Q xả')
    tbl = cap.getnext()
    assert tbl.tag == qn('w:tbl')
    mau = tim_doan(body, 'Nguồn F: bình đồ lòng hồ năm 2022')
    doan_moi_sau(tbl, 'Ghi chú: cột V lấy theo quan hệ V(Z) dùng trong chương trình điều tiết lũ (tích phân diện '
                      'tích trên lưới 0,01 m, F nội suy tuyến tính), lớn hơn giá trị ở Bảng PL1 (công thức (1-1), '
                      'bước 1 m) không quá 0,07·10^{6} m^{3} (≤ 0,2%); điều tiết lũ chỉ dùng hiệu số dung tích nên '
                      'sai khác này không ảnh hưởng đến kết quả.', mau)
    ghi('Bảng PL2: thêm ghi chú giải thích cột V lệch ≤ 0,07·10⁶ m³ so với Bảng PL1 (khác cách tích phân).')


VIET_TAT = [
    ('BCNCKT', 'Báo cáo nghiên cứu khả thi'),
    ('BT', 'Bê tông'),
    ('BTCT', 'Bê tông cốt thép'),
    ('CB, ĐB', 'Tổ hợp tải trọng cơ bản, đặc biệt'),
    ('ĐATN', 'Đồ án tốt nghiệp'),
    ('ĐĐ-I, ĐĐ-II', 'Lớp đất đắp thân đập I, II (theo hồ sơ khảo sát địa chất)'),
    ('FEM (PTHH)', 'Phương pháp phần tử hữu hạn'),
    ('KTTV', 'Khí tượng thủy văn'),
    ('MN', 'Mực nước'),
    ('MNC', 'Mực nước chết'),
    ('MNDBT', 'Mực nước dâng bình thường'),
    ('MNLKT', 'Mực nước lũ kiểm tra'),
    ('MNLTK', 'Mực nước lũ thiết kế'),
    ('ng.đ', 'Ngày đêm (trong đơn vị lưu lượng thấm m³/ng.đ/m)'),
    ('PA', 'Phương án'),
    ('QCVN', 'Quy chuẩn kỹ thuật quốc gia'),
    ('QTVH', 'Quy trình vận hành hồ chứa'),
    ('SC', 'Sửa chữa (“sau SC”: sau sửa chữa)'),
    ('SCADA', 'Hệ thống giám sát, điều khiển và thu thập dữ liệu'),
    ('TB', 'Trung bình'),
    ('TCVN', 'Tiêu chuẩn quốc gia'),
    ('TH', 'Trường hợp tính toán'),
    ('TKCS', 'Thiết kế cơ sở'),
    ('TL, HL', 'Thượng lưu, hạ lưu'),
    ('TSC', 'Tràn sự cố'),
    ('TTGH', 'Trạng thái giới hạn'),
]


def danh_muc_viet_tat(body):
    h = [p for p in body.iterchildren(qn('w:p')) if van_ban(p).startswith('DANH MỤC CÁC TỪ VIẾT TẮT')][0]
    tbl = h.getnext()
    assert tbl.tag == qn('w:tbl')
    rows = tbl.findall(qn('w:tr'))
    mau = rows[0]
    cu = len(rows)
    for r in rows:
        tbl.remove(r)
    for viet_tat, nghia in VIET_TAT:
        r = copy.deepcopy(mau)
        for tc, s in zip(r.findall(qn('w:tc')), (viet_tat, nghia)):
            run = tc.find('.//' + qn('w:r'))
            dat_chu(run, s)
        tbl.append(r)
    ghi('Danh mục từ viết tắt: %d -> %d mục, xếp lại theo thứ tự A, B, C (thêm BT, CB/ĐB, MN, ng.đ, PA, '
        'SC, TB; chuyển TH về đúng vị trí).' % (cu, len(VIET_TAT)))


def bia(body):
    n = 0
    for tbl in body.iterchildren(qn('w:tbl')):
        chu = ''.join(t.text or '' for t in tbl.iter(qn('w:t')))
        if 'BỘ GIÁO DỤC VÀ ĐÀO TẠO' in chu and 'BỘ NÔNG NGHIỆP' in chu:
            rong = (3969, 5102)  # 7,0 cm + 9,0 cm = 16,0 cm
            for gc, w in zip(tbl.find(qn('w:tblGrid')).findall(qn('w:gridCol')), rong):
                gc.set(qn('w:w'), str(w))
            for tc, w in zip(tbl.iter(qn('w:tc')), rong):
                tc.find(qn('w:tcPr') + '/' + qn('w:tcW')).set(qn('w:w'), str(w))
            n += 1
    ghi('Bìa chính, bìa phụ: đổi độ rộng 2 ô tiêu đề (7 cm + 9 cm) để "BỘ NÔNG NGHIỆP VÀ MÔI TRƯỜNG" '
        'nằm trên một dòng (%d bảng).' % n)
    ten = 'TÍNH TOÁN, THIẾT KẾ SỬA CHỮA NÂNG CẤP CỤM CÔNG TRÌNH ĐẦU MỐI HỒ CHỨA NƯỚC CÀ GIÂY – TỈNH LÂM ĐỒNG'
    k = 0
    for p in body.iterchildren(qn('w:p')):
        if kieu(p) == 'CoverT' and van_ban(p) == ten:
            thay_trong_doan(p, ten, 'TÍNH TOÁN, THIẾT KẾ SỬA CHỮA NÂNG CẤP\nCỤM CÔNG TRÌNH ĐẦU MỐI\n'
                                    'HỒ CHỨA NƯỚC CÀ GIÂY – TỈNH LÂM ĐỒNG')
            k += 1
    ghi('Tên đề tài trên bìa: ngắt thành 3 dòng cân đối (%d chỗ), tránh chữ "ĐỒNG" đứng riêng một dòng.' % k)


def main(vao, ra):
    d = docx.Document(vao)
    body = d.element.body
    sua_cong_thuc(body)
    sua_loi_noi_dung(body)
    dan_hinh_bang(body)
    ghi_chu_bang_pl2(body)
    danh_muc_viet_tat(body)
    bia(body)
    d.save(ra)
    print('Đã lưu', ra, '–', len(NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
