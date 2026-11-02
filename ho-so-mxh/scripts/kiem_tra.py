#!/usr/bin/env python3
"""Kiểm tra tính toàn vẹn và nhất quán của bộ dữ liệu hồ sơ (ho-so-mxh/data/*.tsv).

Chương trình chỉ đối chiếu dữ liệu với nhau và tính lại mốc thời gian; nó KHÔNG kết luận
cơ quan hay cá nhân nào vi phạm pháp luật. Kết quả giúp người đọc tự đối chiếu.

Cách chạy:  python3 ho-so-mxh/scripts/kiem_tra.py  [--out ho-so-mxh/bao-cao/KIEM_TRA_TU_DONG.md]
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, '..', 'data')

TRANG_THAI_THU_TUC = {
    'chưa có văn bản', 'có văn bản trả lời', 'chuyển đơn', 'hướng dẫn', 'đôn đốc', 'báo tin',
    'thông báo không thụ lý', 'chưa gửi', 'chuyển đơn; thông báo không thụ lý', 'chuyển đơn; đôn đốc',
}
TRANG_THAI_NOI_DUNG = {
    'đáp ứng đầy đủ', 'đáp ứng một phần', 'chưa giải quyết', 'chưa có kết luận chuyên môn',
    'chưa đủ thông tin để đánh giá', 'không chấp nhận', 'không áp dụng',
}

# Ngày nghỉ lễ có thể rơi vào các khoảng tính hạn (2025–2026). Mục 01–02/9/2026 là giả định
# (chưa đối chiếu thông báo nghỉ lễ chính thức); không khoảng tính hạn nào trong hồ sơ đi qua các ngày này.
NGAY_LE = {dt.date(2026, 1, 1), dt.date(2026, 4, 30), dt.date(2026, 5, 1),
           dt.date(2026, 9, 1), dt.date(2026, 9, 2)}


def doc(ten):
    with open(os.path.join(DATA, ten + '.tsv'), encoding='utf-8') as f:
        return list(csv.DictReader(f, delimiter='\t'))


def ngay(s):
    m = re.match(r'\s*(\d{1,2})/(\d{1,2})/(\d{4})', s or '')
    return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def la_ngay_lam_viec(d):
    return d.weekday() < 5 and d not in NGAY_LE


def cong_ngay_lam_viec(d, n):
    """Ngày làm việc thứ n tính từ ngày liền sau d (Điều 147 BLDS 2015)."""
    k = 0
    while k < n:
        d += dt.timedelta(days=1)
        if la_ngay_lam_viec(d):
            k += 1
    return d


def dem_ngay_lam_viec(tu, den):
    """Số ngày làm việc sau 'tu' đến hết 'den'."""
    k, d = 0, tu
    while d < den:
        d += dt.timedelta(days=1)
        if la_ngay_lam_viec(d):
            k += 1
    return k


def tach_ma(s):
    return [x.strip() for x in re.split(r'[;,]', s or '') if x.strip() and x.strip() not in {'-'}]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=os.path.join(HERE, '..', 'bao-cao', 'KIEM_TRA_TU_DONG.md'))
    ap.add_argument('--ngay-tinh', default='01/11/2026')
    a = ap.parse_args()
    moc = ngay(a.ngay_tinh)

    hs = {r['id']: r for r in doc('ho_so')}
    yc = doc('yeu_cau')
    vb = {r['id']: r for r in doc('van_ban')}
    tl = {r['ma']: r for r in doc('tai_lieu')}
    vd = {r['id']: r for r in doc('van_de')}
    qh = doc('quan_he')
    ld = doc('luan_diem')
    kt = {r['id']: r for r in doc('kiem_toan')}

    loi, canh_bao, ghi = [], [], []
    ma_hop_le = set(tl) | {'2.1'}  # 2.1 = thư mục ảnh lớp học (06.04–06.14)

    def ma_ton_tai(ma):
        ma = ma.split(' ')[0]
        if ma in ma_hop_le:
            return True
        m = re.match(r'^(\d\d\.\d{1,2})[–-](\d\d\.?\d{0,2})$', ma)  # dải mã kiểu 06.01–06.09
        if m:
            return m.group(1) in ma_hop_le
        if ma in {'15.5'}:
            return '15.5a' in ma_hop_le
        return False

    # 1. Tham chiếu chéo
    for r in yc:
        if r['ho_so'] not in hs:
            loi.append(f"Yêu cầu {r['id']} trỏ tới hồ sơ không tồn tại: {r['ho_so']}")
        if r['trang_thai_thu_tuc'] not in TRANG_THAI_THU_TUC:
            loi.append(f"Yêu cầu {r['id']}: trạng thái thủ tục lạ '{r['trang_thai_thu_tuc']}'")
        if r['trang_thai_noi_dung'] not in TRANG_THAI_NOI_DUNG:
            loi.append(f"Yêu cầu {r['id']}: trạng thái nội dung lạ '{r['trang_thai_noi_dung']}'")
        for v in tach_ma(r['van_de']):
            if v not in vd:
                loi.append(f"Yêu cầu {r['id']}: mã vấn đề không tồn tại {v}")
        if r['trang_thai_noi_dung'].startswith('đáp ứng') and r['van_ban_doi_chieu'] in {'', '-'}:
            loi.append(f"Yêu cầu {r['id']} ghi '{r['trang_thai_noi_dung']}' nhưng không nêu văn bản đối chiếu")
    for h in hs.values():
        for ma in tach_ma(h['chung_tu']) + tach_ma(h['tep_don']):
            if not ma_ton_tai(ma):
                loi.append(f"Hồ sơ {h['id']}: mã tệp {ma} không có trong danh mục tài liệu")
        if h['id'] not in {r['ho_so'] for r in yc}:
            loi.append(f"Hồ sơ {h['id']} chưa được tách yêu cầu")
    for v in vb.values():
        for ma in tach_ma(v['co_trong_thu_muc'].replace(' (2 trang)', '')):
            if ma != 'không' and not ma_ton_tai(ma):
                loi.append(f"Văn bản {v['id']}: mã tệp {ma} không có trong danh mục")
    for r in qh:
        for x in (r['tu'], r['den']):
            if x not in hs and x not in vb:
                loi.append(f"Quan hệ {r['tu']}→{r['den']}: đối tượng {x} không tồn tại")
    for r in ld:
        for ma in tach_ma(r['chung_cu_truc_tiep']):
            if re.match(r'^\d', ma) and not ma_ton_tai(ma.split(' ')[0]):
                canh_bao.append(f"Luận điểm {r['id']}: mã chứng cứ {ma} chưa khớp danh mục")
        for k in tach_ma(r['kiem_toan']):
            if k.split(' ')[0] not in kt:
                loi.append(f"Luận điểm {r['id']}: mã kiểm toán {k} không tồn tại")

    # 2. Chứng từ gửi/nhận
    thieu_ct = []
    for h in hs.values():
        n = int(h['id'][2:])
        if n > 22:
            continue
        if not tach_ma(h['chung_tu']):
            thieu_ct.append(f"{h['id']} ({h['ten']}) – nguồn ngày nhận: {h['nguon_ngay_nhan']}")
    # 3. Thứ tự ngày
    for h in hs.values():
        d_de, d_gui, d_nhan = ngay(h['ngay_de']), ngay(h['ngay_gui']), ngay(h['ngay_nhan'])
        if d_de and d_gui and d_gui < d_de:
            loi.append(f"{h['id']}: ngày gửi {h['ngay_gui']} trước ngày đề {h['ngay_de']}")
        if d_gui and d_nhan and d_nhan < d_gui:
            loi.append(f"{h['id']}: ngày nhận {h['ngay_nhan']} trước ngày gửi {h['ngay_gui']}")
    van_ban_truoc_don = []
    for r in qh:
        if r['tu'] in hs and r['den'] in vb and r['loai'].startswith('được'):
            d_nhan = ngay(hs[r['tu']]['ngay_nhan']) or ngay(hs[r['tu']]['ngay_gui']) or ngay(hs[r['tu']]['ngay_de'])
            d_ky = ngay(vb[r['den']]['ngay_ky'])
            if d_nhan and d_ky and d_ky < d_nhan:
                loi.append(f"{r['den']} ({vb[r['den']]['ngay_ky']}) ban hành trước khi {r['tu']} được nhận – không thể là văn bản xử lý đơn này")
        if r['tu'] in vb and r['den'] in vb and r['loai'] == 'viện dẫn':
            for h_id, h in hs.items():
                if h_id in vb[r['tu']]['lien_quan_ho_so']:
                    d1, d2 = ngay(vb[r['den']]['ngay_ky']), ngay(h['ngay_de'])
                    if d1 and d2 and d1 < d2:
                        van_ban_truoc_don.append(f"{r['tu']} trả lời {h_id} bằng cách viện dẫn {r['den']} ({vb[r['den']]['ngay_ky']}), ban hành TRƯỚC ngày lập đơn {h['ngay_de']}")

    # 4. Tính lại mốc thời gian
    moc_tinh = []
    def them(ten, nhan, so_ngay, ky_vong, can_cu):
        d = cong_ngay_lam_viec(ngay(nhan), so_ngay)
        ok = (d == ngay(ky_vong)) if ky_vong else None
        moc_tinh.append((ten, nhan, so_ngay, d.strftime('%d/%m/%Y'), ky_vong or '-', '✔' if ok else ('✘' if ok is False else '-'), can_cu, (moc - d).days))
    them('Đơn A tại Sở (HS04) – xử lý ban đầu tối đa', '16/5/2026', 10, '29/5/2026', 'k1 Đ24 LTC')
    them('CV1436 tại phường (Đơn B) – 07 ngày', '22/5/2026', 7, '02/6/2026', 'k1 Đ24 LTC')
    them('CV1436 tại phường (Đơn B) – tối đa 10 ngày', '22/5/2026', 10, '05/6/2026', 'k1 Đ24 LTC')
    them('Hồ sơ 09/7 (HS08) – chuyển đơn 05 ngày', '09/7/2026', 5, '16/7/2026', 'k2 Đ24 LTC (theo 15.0)')
    them('Hồ sơ 09/7 (HS08) – 07 ngày', '09/7/2026', 7, '20/7/2026', 'k1 Đ24 LTC')
    them('Hồ sơ 09/7 (HS08) – tối đa 10 ngày', '09/7/2026', 10, '23/7/2026', 'k1 Đ24 LTC')
    them('Tố cáo tỉnh (HS15) – 07 ngày', '17/9/2026', 7, '28/9/2026', 'k1 Đ24 LTC')
    them('Tố cáo tỉnh (HS15) – tối đa 10 ngày', '17/9/2026', 10, '01/10/2026', 'k1 Đ24 LTC')
    them('Khiếu nại TB5553 (HS20) – 07 ngày làm việc', '06/10/2026', 7, '15/10/2026', 'Đ27 LKN sửa đổi (theo người lập; chưa đối chiếu)')
    so_ngay = [
        ('Từ 30/10/2025 đến ngày tính', (moc - ngay('30/10/2025')).days, '367 (17.20)'),
        ('Từ mốc 29/5/2026 (Đơn A)', (moc - ngay('29/5/2026')).days, '156 (17.20)'),
        ('Từ 09/7/2026 (HS08 nhận)', (moc - ngay('09/7/2026')).days, '115 (17.19, 17.20)'),
        ('Từ 17/9/2026 (HS15 nhận)', (moc - ngay('17/9/2026')).days, '45 (17.20)'),
        ('Từ mốc 01/10/2026 (HS15)', (moc - ngay('01/10/2026')).days, '31 (17.20)'),
        ('Từ 16/5/2026 (Đơn C nhận)', (moc - ngay('16/5/2026')).days, '169 (17.20)'),
        ('Từ 05/10/2026 (HS19 nhận)', (moc - ngay('05/10/2026')).days, '27 (17.20)'),
        ('Ngày làm việc sau mốc 15/10/2026 (HS20)', dem_ngay_lam_viec(ngay('15/10/2026'), moc), '11 (17.20)'),
        ('Thời gian từ khi phường nhận CV1436 đến khi ký TB5553', (ngay('28/8/2026') - ngay('22/5/2026')).days, '98 (17.20; 07.10)'),
    ]

    # 5. Đối chiếu cách đếm
    dem00 = [h['id'] for h in hs.values() if re.match(r'^\d+$', h['dem_00_27_10'] or '')]
    dem0710 = [h['id'] for h in hs.values() if re.match(r'^\d+[a-e]?(=\d+a)?$', h['dem_0710_28_10'] or '')]
    khong_dem = [h['id'] for h in hs.values() if h['dem_00_27_10'] in {'', '-'} and h['dem_0710_28_10'] in {'', '-'} and int(h['id'][2:]) <= 22]
    da_gui = [h for h in hs.values() if int(h['id'][2:]) <= 22]
    chua_gui = [h for h in hs.values() if int(h['id'][2:]) > 22]

    # 6. Tổng hợp trạng thái yêu cầu
    from collections import Counter, defaultdict
    tt_nd = Counter(r['trang_thai_noi_dung'] for r in yc if r['trang_thai_thu_tuc'] != 'chưa gửi')
    theo_hs = defaultdict(Counter)
    for r in yc:
        theo_hs[r['ho_so']][r['trang_thai_noi_dung']] += 1

    # 7. Văn bản không có trong thư mục
    vb_thieu = [f"{v['id']} ({v['so_hieu']}, {v['ngay_ky']}) – chỉ biết qua: {v['nguon_neu_khong_co']}" for v in vb.values() if v['co_trong_thu_muc'] == 'không']

    # 8. Bản trùng (theo tiêu đề và kích thước; chưa so từng chữ)
    trung = [('F7-01', 'DonB'), ('F7-02', '15.0'), ('F7-03', '15.2a'), ('F7-04', '17.05'), ('F7-05', '17.15'),
             ('07.9-27', '07.9F'), ('07.9-28', '07.9F')]

    # Ghi báo cáo
    L = []
    L.append('# KẾT QUẢ KIỂM TRA TỰ ĐỘNG BỘ DỮ LIỆU HỒ SƠ')
    L.append('')
    L.append(f'Ngày tính: {a.ngay_tinh}. Chương trình: `ho-so-mxh/scripts/kiem_tra.py`. Kết quả chỉ đối chiếu dữ liệu và tính lại mốc thời gian; không kết luận vi phạm.')
    L.append('')
    L.append(f'## 1. Lỗi dữ liệu: {len(loi)}')
    L += [f'- {x}' for x in loi] or ['- Không có.']
    L.append('')
    L.append(f'## 2. Hồ sơ đã gửi chưa có chứng từ gửi/nhận trong thư mục: {len(thieu_ct)}')
    L += [f'- {x}' for x in thieu_ct]
    L.append('')
    L.append('## 3. Văn bản viện dẫn ban hành trước ngày lập đơn')
    L += [f'- {x}' for x in van_ban_truoc_don] or ['- Không có.']
    L.append('')
    L.append('## 4. Tính lại mốc thời hạn (ngày làm việc tính từ ngày liền sau ngày nhận; không tính thứ Bảy, Chủ nhật, ngày lễ)')
    L.append('')
    L.append('| Trường hợp | Ngày nhận | Số ngày làm việc | Mốc tính lại | Mốc hồ sơ ghi | Khớp | Căn cứ (theo hồ sơ) | Số ngày lịch đã qua mốc đến ngày tính |')
    L.append('|---|---|---|---|---|---|---|---|')
    for t in moc_tinh:
        L.append('| ' + ' | '.join(str(x) for x in t) + ' |')
    L.append('')
    L.append('Lưu ý: căn cứ của các mốc sau 01/7/2026 phụ thuộc nội dung Luật số 136/2025/QH15 sửa đổi Luật Tố cáo, Luật Khiếu nại – chưa được đối chiếu nguyên văn trong hồ sơ (KT25). Đơn 17.02-03 tính cả ngày nhận nên ra mốc sớm hơn 2 ngày (KT16).')
    L.append('')
    L.append('| Số ngày | Tính lại | Hồ sơ ghi |')
    L.append('|---|---|---|')
    for t in so_ngay:
        L.append(f'| {t[0]} | {t[1]} | {t[2]} |')
    L.append('')
    L.append('## 5. Đối chiếu cách đếm hồ sơ')
    L.append(f'- Tệp 00 (đến 27/10/2026, hồ sơ CHƯA giải quyết): {len(dem00)} → {", ".join(dem00)}')
    L.append(f'- Tệp 07.10 (đến 28/10/2026, đơn có chứng từ, trừ báo cáo và tài liệu bổ sung; dòng 7 = 12a): {len(dem0710)} → {", ".join(dem0710)}')
    L.append(f'- Không có trong cả hai danh mục: {", ".join(khong_dem)}')
    L.append(f'- Bảng hợp nhất: {len(da_gui)} hồ sơ đã gửi + {len(chua_gui)} văn bản đã soạn, chưa xác nhận gửi (HS23–HS25).')
    L.append('')
    L.append('## 6. Trạng thái nội dung của từng yêu cầu (không tính văn bản chưa gửi)')
    for k, v in tt_nd.most_common():
        L.append(f'- {k}: {v}')
    L.append('')
    L.append('| Hồ sơ | Đầy đủ | Một phần | Chưa giải quyết | Chưa có kết luận chuyên môn | Không chấp nhận | Chưa đủ thông tin |')
    L.append('|---|---|---|---|---|---|---|')
    for h_id in sorted(theo_hs, key=lambda x: int(x[2:])):
        c = theo_hs[h_id]
        if hs[h_id]['ngay_gui'].startswith('chưa'):
            continue
        L.append(f"| {h_id} | {c['đáp ứng đầy đủ']} | {c['đáp ứng một phần']} | {c['chưa giải quyết']} | {c['chưa có kết luận chuyên môn']} | {c['không chấp nhận']} | {c['chưa đủ thông tin để đánh giá']} |")
    L.append('')
    L.append(f'## 7. Văn bản cơ quan được viện dẫn nhưng không có trong thư mục: {len(vb_thieu)}')
    L += [f'- {x}' for x in vb_thieu]
    L.append('')
    L.append('## 8. Bản trùng nội dung (theo tiêu đề và kích thước, chưa so từng chữ)')
    L += [f'- {x} ↔ {y}' for x, y in trung]
    L.append('')
    L.append(f'## 9. Cảnh báo: {len(canh_bao)}')
    L += [f'- {x}' for x in canh_bao] or ['- Không có.']
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with open(a.out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print(f'lỗi={len(loi)} cảnh báo={len(canh_bao)} thiếu chứng từ={len(thieu_ct)} → {a.out}')
    return 1 if loi else 0


if __name__ == '__main__':
    sys.exit(main())
