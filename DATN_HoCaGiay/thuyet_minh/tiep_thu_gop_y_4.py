# -*- coding: utf-8 -*-
"""
ĐỐI CHIẾU BẢN VẼ, SLIDE VỚI THUYẾT MINH (đêm 11/10/2026, sau rà soát đợt 3) – sửa trực tiếp tệp Word thuyết minh.

  python3 tiep_thu_gop_y_4.py <vào.docx> <ra.docx>
  python3 chen_muc_luc.py <ra.docx> <bản cuối.docx> <bản cuối.pdf>

Bản vẽ (BV-04, BV-05, BV-06) và slide (10, 12, 18, 21) đã sửa theo thuyết minh; trong thuyết minh chỉ còn một chỗ
lệch nội bộ: mục 5.5.3 ghi "cỏ Vetiver hoặc cỏ lá tre" trong khi mục 5.5.1, 7.4, Kết luận và BV-04 thống nhất trồng
cỏ bằng vầng hoặc gieo hạt có phủ thảm xơ dừa. Nhật ký: nhat_ky_tiep_thu_gop_y_4.txt.
"""
import os
import sys

import docx

import cong_cu_docx as C
from cong_cu_docx import ghi, tim_doan, thay


# =========================================================================== CHƯƠNG 5
def chuong5(body):
    p = tim_doan(body, 'Trong các ô: rải lớp đất màu dày 10 cm')
    thay(p, 'trồng cỏ (cỏ Vetiver hoặc cỏ lá tre bản địa), chăm sóc trong 3 tháng đầu.',
         'trồng cỏ bản địa (cỏ gà, cỏ lá tre) bằng vầng hoặc gieo hạt có phủ thảm xơ dừa cố định bằng ghim '
         '(mục 5.5.1); chăm sóc trong 3 tháng đầu.')
    ghi('Mục 5.5.3: lớp phủ trong ô ghi thống nhất với mục 5.5.1, 7.4, Kết luận và BV-04 – trồng cỏ bằng vầng hoặc '
        'gieo hạt có phủ thảm xơ dừa (bỏ "cỏ Vetiver": cỏ trồng theo hàng, không phủ kín mặt ô trong mùa mưa đầu).')


def main(vao, ra):
    d = docx.Document(vao)
    body = d.element.body
    n_dong, n_bm = C.go_muc_luc(body)
    ghi('Gỡ mục lục cũ (%d dòng, %d bookmark) để điền lại sau khi sửa.' % (n_dong, n_bm))
    chuong5(body)
    d.save(ra)
    with open(os.path.splitext(ra)[0] + '_nhat_ky.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(C.NHAT_KY) + '\n')
    print('Đã lưu', ra, '–', len(C.NHAT_KY), 'nhóm thay đổi')


if __name__ == '__main__':
    main(*sys.argv[1:3])
