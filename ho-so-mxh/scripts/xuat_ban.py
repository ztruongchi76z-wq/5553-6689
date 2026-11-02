#!/usr/bin/env python3
"""Sinh tài liệu tra cứu từ dữ liệu ho-so-mxh/data/*.tsv.

Hai chế độ:
  python3 scripts/xuat_ban.py md   <thư_mục_ra>
      Markdown cho kho GitHub: chỉ ghi mã tài liệu, không có liên kết Drive.
  python3 scripts/xuat_ban.py html <thư_mục_ra> <thư_mục_bản_đồ_riêng>
      HTML để nhập thành Google Docs trong bộ PUBLIC_READY. Thư mục bản đồ riêng
      (KHÔNG đưa lên kho) chứa public_map.tsv (code, public_id) và folders.tsv
      (key, id, title).

Nguyên tắc: mã tài liệu trong văn bản được gắn liên kết tới bản công khai nếu có;
tài liệu chưa đưa vào bộ công khai được đánh dấu † (chỉ có trong bộ làm việc).
"""
import collections
import csv
import html
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / 'data'
NGAY_LAP = '02/11/2026'


def bang(ten):
    with open(DATA / f'{ten}.tsv', encoding='utf-8') as f:
        return list(csv.DictReader(f, delimiter='\t'))


HS = bang('ho_so')
YC = bang('yeu_cau')
VB = bang('van_ban')
VD = bang('van_de')
LD = bang('luan_diem')
KT = bang('kiem_toan')
CC = bang('can_cu')
TL = bang('tai_lieu')
VB_ID = {v['id']: v for v in VB}
TL_MA = {t['ma']: t for t in TL}

# Mã được viện dẫn theo cách viết khác trong dữ liệu
BI_DANH = {'15.5': '15.5a', '07.9': '07.9F', 'Đơn B': 'DonB'}
# Tài liệu chưa công khai ảnh gốc nhưng có bản chép đã che trong bộ công khai
BAN_CHEP = {'08.1': '08.1-2T', '08.2': '08.1-2T', '15.1': '15.1T', '15.3': '15.3T',
            '09.1': '09.1-4T', '09.4': '09.1-4T', '10.3': '10.3T',
            '14.1': '14.1-2T', '14.2': '14.1-2T'}
# Kiểm toán không đưa vào bản công khai (quyền riêng tư, liên kết nội bộ)
KT_RIENG = {'KT02', 'KT03', 'KT04', 'KT05', 'KT23'}
# Tên tệp dữ liệu trong kho được thay bằng tên dễ hiểu khi xuất cho Google Docs
NHAN_DE_DOC = [('ho_so.tsv cột dem_00_27_10, dem_0710_28_10', 'bảng hồ sơ, dòng “Đối chiếu bảng cũ” (tài liệu 01)'),
               ('ho_so.tsv', 'bảng hồ sơ (tài liệu 01)'), ('yeu_cau.tsv', 'bảng yêu cầu (tài liệu 01)'),
               ('can_cu.tsv', 'bảng căn cứ pháp lý (tài liệu 05)'), ('kiem_tra.py', 'chương trình kiểm tra tự động')]

MA_RE = re.compile(r'(?<![\w.])(' + '|'.join(
    re.escape(m) for m in sorted(list(TL_MA) + list(BI_DANH), key=len, reverse=True)
    if m not in ('thu_muc_moi',)) + r')(?![\w])(?!\.\d)')


class XuatBan:
    def __init__(self, che_do, ban_do=None):
        self.html = che_do == 'html'
        self.map, self.thu_muc = {}, {}
        if ban_do:
            for r in csv.DictReader(open(ban_do / 'public_map.tsv', encoding='utf-8'), delimiter='\t'):
                self.map[r['code']] = r['public_id']
            for r in csv.DictReader(open(ban_do / 'folders.tsv', encoding='utf-8'), delimiter='\t'):
                self.thu_muc[r['key']] = r['id']

    # ---------- định dạng ----------
    def e(self, s):
        return html.escape(s or '', quote=False)

    def url(self, file_id, folder=False):
        return (f'https://drive.google.com/drive/folders/{file_id}' if folder
                else f'https://drive.google.com/open?id={file_id}')

    def a(self, nhan, file_id, folder=False):
        if not self.html or not file_id:
            return self.e(nhan)
        return f'<a href="{self.url(file_id, folder)}">{self.e(nhan)}</a>'

    def ma(self, s):
        """Gắn liên kết cho các mã tài liệu trong chuỗi s (đã thoát HTML ở phần còn lại)."""
        out, i = [], 0
        for m in MA_RE.finditer(s or ''):
            out.append(self.e(s[i:m.start()]))
            goc = m.group(1)
            ma = BI_DANH.get(goc, goc)
            if ma in self.map:
                out.append(self.a(goc, self.map[ma]))
            elif ma in BAN_CHEP and BAN_CHEP[ma] in self.map:
                out.append(self.a(goc, self.map[BAN_CHEP[ma]]) + ' (bản chép)')
            elif ma in TL_MA and not self.map and TL_MA[ma]['cong_khai'] == 'P':
                out.append(self.e(goc))
            elif ma in TL_MA and TL_MA[ma]['cong_khai'] in ('P', 'T') and not self.html:
                out.append(self.e(goc))
            else:
                out.append(self.e(goc) + '†')
            i = m.end()
        out.append(self.e((s or '')[i:]))
        return ''.join(out)

    @staticmethod
    def sang_md(h):
        h = re.sub(r'<h1>(.*?)</h1>', r'# \1\n', h)
        h = re.sub(r'<h2>(.*?)</h2>', r'\n## \1\n', h)
        h = re.sub(r'<h3>(.*?)</h3>', r'\n### \1\n', h)
        h = re.sub(r'<li>(.*?)</li>', r'- \1', h)
        h = re.sub(r'</?ul>', '', h)
        h = re.sub(r'<p>(.*?)</p>', r'\1\n', h)
        h = h.replace('<br>', '  \n').replace('<b>', '**').replace('</b>', '**')
        h = h.replace('<i>', '_').replace('</i>', '_')
        h = re.sub(r'<a href="[^"]*">(.*?)</a>', r'\1', h)
        h = html.unescape(re.sub(r'</?(html|body)>', '', h))
        return re.sub(r'\n{3,}', '\n\n', h).strip() + '\n'

    def goi(self, khoi):
        h = '<html><body>\n' + '\n'.join(khoi) + '\n</body></html>\n'
        h = h.replace('<li>', '\n<li>').replace('</ul>', '\n</ul>')
        if self.html:
            for cu, moi in NHAN_DE_DOC:
                h = h.replace(cu, moi)
        return h if self.html else self.sang_md(h)

    def chu_thich(self):
        if self.html:
            return ('<p><i>Ký hiệu: mã tài liệu có gạch chân là liên kết tới bản trong bộ công khai; '
                    'dấu † là tài liệu chưa đưa vào bộ công khai (có thông tin cá nhân hoặc chờ người lập quyết định). '
                    f'Tài liệu sinh tự động từ dữ liệu ngày {NGAY_LAP}.</i></p>')
        return f'_Sinh tự động từ dữ liệu ngày {NGAY_LAP}. Dấu † là tài liệu chưa đưa vào bộ công khai._'

    def tep_vb(self, vb):
        t = vb['co_trong_thu_muc']
        if t in ('', '-') or t.startswith('không'):
            return 'không có trong thư mục'
        return self.ma(t)

    # ---------- 01: hồ sơ theo từng đơn ----------
    def ho_so_theo_don(self):
        yc = collections.defaultdict(list)
        for r in YC:
            yc[r['ho_so']].append(r)
        dem = collections.Counter(r['trang_thai_noi_dung'] for r in YC)
        k = [
            '<h1>Hồ sơ theo từng đơn: vòng đời, yêu cầu và tình trạng</h1>',
            self.chu_thich(),
            '<p>Mỗi hồ sơ ghi: loại đơn theo tiêu đề và theo nội dung; ngày đề, ngày gửi, ngày cơ quan nhận và chứng từ; '
            'văn bản cơ quan đã ban hành; từng yêu cầu trong đơn với hai trạng thái độc lập. '
            '<b>Trạng thái thủ tục</b> cho biết cơ quan đã có văn bản xử lý hay chưa. '
            '<b>Trạng thái nội dung</b> cho biết yêu cầu đã được đáp ứng đến đâu. '
            'Một đơn có thể đã được chuyển hoặc trả lời về thủ tục mà nội dung vẫn chưa được giải quyết.</p>',
            f'<p>Tổng cộng {len(HS)} hồ sơ, {len(YC)} yêu cầu. Trạng thái nội dung: '
            + '; '.join(f'{t}: {n}' for t, n in dem.most_common()) + '.</p>',
        ]
        for h in HS:
            k.append(f'<h2>{h["id"]}. {self.e(h["ten"])}</h2>')
            loai = h['loai_theo_tieu_de']
            if h['loai_theo_noi_dung'] and h['loai_theo_noi_dung'] != loai:
                loai += f'; theo nội dung: {h["loai_theo_noi_dung"]}'
            gui = h['ngay_gui'] or '-'
            if h['phuong_thuc'] and h['phuong_thuc'] != '-':
                gui += f', {h["phuong_thuc"]}'
            if h['ma_buu_gui'] and h['ma_buu_gui'] != '-':
                gui += f', mã bưu gửi {h["ma_buu_gui"]}'
            nhan = h['ngay_nhan'] or 'chưa xác định'
            if h['gio_nhan']:
                nhan += f' lúc {h["gio_nhan"]}'
            if h['nguon_ngay_nhan'] and h['nguon_ngay_nhan'] != '-':
                nhan += f' (nguồn: {h["nguon_ngay_nhan"]})'
            dong = [
                f'<b>Loại đơn:</b> {self.e(loai)}',
                f'<b>Ngày đề:</b> {self.e(h["ngay_de"] or "-")}. <b>Gửi:</b> {self.e(gui)}',
                f'<b>Nơi nhận:</b> {self.e(h["co_quan_nhan"])}'
                + (f' (ghi trên đơn: {self.e(h["nguoi_nhan_ghi"])})' if h['nguoi_nhan_ghi'] not in ('', '-') else ''),
                f'<b>Cơ quan nhận:</b> {self.e(nhan)}',
                f'<b>Văn bản đơn:</b> {self.ma(h["tep_don"] or "-")}. <b>Chứng từ gửi, nhận:</b> {self.ma(h["chung_tu"] or "-")}',
            ]
            if h['van_ban_xu_ly'] not in ('', '-'):
                vbs = []
                for v in re.split(r'[;,→]\s*', h['van_ban_xu_ly']):
                    v = v.strip()
                    nghi = v.endswith('?')
                    v = v.rstrip('?')
                    if v in VB_ID:
                        vb = VB_ID[v]
                        vbs.append(f'{self.e(vb["so_hieu"])} ngày {vb["ngay_ky"]} ({self.tep_vb(vb)})'
                                   + (' (chưa xác định có phải văn bản xử lý đơn này)' if nghi else ''))
                    elif v:
                        vbs.append(self.e(v))
                dong.append('<b>Văn bản cơ quan:</b> ' + '; '.join(vbs))
            dong.append(f'<b>Tình trạng chung:</b> {self.e(h["trang_thai_chung"])}')
            dem_cu = []
            if h['dem_00_27_10'] not in ('', '-'):
                dem_cu.append(f'số {h["dem_00_27_10"]} trong danh sách 16 hồ sơ của tài liệu 00')
            if h['dem_0710_28_10'] not in ('', '-'):
                dem_cu.append(f'số {h["dem_0710_28_10"]} trong bảng 07.10')
            if dem_cu:
                dong.append('<b>Đối chiếu bảng cũ:</b> ' + '; '.join(dem_cu))
            if h['ghi_chu'] not in ('', '-'):
                dong.append(f'<b>Ghi chú:</b> {self.ma(h["ghi_chu"])}')
            k.append('<ul>' + ''.join(f'<li>{d}</li>' for d in dong) + '</ul>')
            ds = yc.get(h['id'], [])
            if ds:
                k.append(f'<p><b>Yêu cầu trong đơn ({len(ds)}):</b></p>')
                truoc, items = None, []
                for r in ds:
                    dg = r['danh_gia']
                    dg_txt = '' if not dg or dg == '-' else (' Như trên.' if dg == truoc else f' {self.ma(dg)}')
                    truoc = dg
                    dc = '' if r['van_ban_doi_chieu'] in ('', '-') else f' Đối chiếu: {self.e(r["van_ban_doi_chieu"])}.'
                    muc = f' [mục {self.e(r["muc"])}]' if r['muc'] not in ('', '-') else ''
                    items.append(
                        f'<li><b>{r["id"].split("-", 1)[1]}</b>{muc} {self.e(r["noi_dung"])}. '
                        f'Thủ tục: {self.e(r["trang_thai_thu_tuc"])}; nội dung: {self.e(r["trang_thai_noi_dung"])}.'
                        f'{dc}{dg_txt}' + (f' ({self.e(r["van_de"])})' if r['van_de'] not in ('', '-') else '') + '</li>')
                k.append('<ul>' + ''.join(items) + '</ul>')
        return self.goi(k)

    # ---------- 02: tra cứu theo vấn đề ----------
    def theo_van_de(self):
        yc = collections.defaultdict(list)
        for r in YC:
            for v in re.split(r'[;,]\s*', r['van_de']):
                yc[v.strip()].append(r)
        k = ['<h1>Tra cứu theo vấn đề</h1>', self.chu_thich(),
             '<p>Mỗi vấn đề ghi câu hỏi trung tâm, chứng cứ gốc, văn bản liên quan, điều đã xác lập, điều còn tồn đọng, '
             'điều đang tranh luận và các yêu cầu trong đơn thuộc vấn đề đó (mã HSxx-Rn, tra ở tài liệu 01).</p>']
        for v in VD:
            ds = yc.get(v['id'], [])
            mo = sum(1 for r in ds if r['trang_thai_noi_dung'] in ('chưa giải quyết', 'chưa có kết luận chuyên môn', 'chưa đủ thông tin'))
            k += [f'<h2>{v["id"]}. {self.e(v["ten"])}</h2>',
                  '<ul>'
                  f'<li><b>Câu hỏi trung tâm:</b> {self.e(v["cau_hoi_trung_tam"])}</li>'
                  f'<li><b>Chứng cứ gốc:</b> {self.ma(v["chung_cu_goc"])}</li>'
                  f'<li><b>Văn bản liên quan:</b> {self.ma(v["van_ban_lien_quan"])}</li>'
                  f'<li><b>Đã xác lập:</b> {self.ma(v["da_xac_lap"])}</li>'
                  f'<li><b>Còn tồn đọng:</b> {self.ma(v["con_ton_dong"])}</li>'
                  f'<li><b>Đang tranh luận:</b> {self.ma(v["dang_tranh_luan"])}</li>'
                  f'<li><b>Yêu cầu liên quan ({len(ds)}, trong đó {mo} chưa có kết quả về nội dung):</b> '
                  + ', '.join(r['id'] for r in ds) + '</li></ul>']
        return self.goi(k)

    # ---------- 05: văn bản cơ quan ----------
    def van_ban(self):
        k = ['<h1>Văn bản của cơ quan: nội dung, chức năng pháp lý và phần còn tồn đọng</h1>', self.chu_thich(),
             '<p>Thứ tự theo ngày ký. Cột “chức năng pháp lý” phân biệt văn bản trả lời, phiếu chuyển, hướng dẫn, báo tin, '
             'thông báo không thụ lý. Văn bản không có trong thư mục được ghi nguồn biết đến nó.</p>']
        def ngay(v):
            d, m, y = (v['ngay_ky'] or '01/01/1900').split('/')
            return (y, m.zfill(2), d.zfill(2))
        for v in sorted(VB, key=ngay):
            tep = self.tep_vb(v)
            if tep == 'không có trong thư mục':
                tep += f'; nguồn: {self.ma(v["nguon_neu_khong_co"])}'
            k += [f'<h2>{self.e(v["so_hieu"])} ngày {v["ngay_ky"]}: {self.e(v["co_quan"])}</h2>',
                  '<ul>'
                  f'<li><b>Người ký:</b> {self.e(v["chuc_danh_ky"])}. <b>Kính gửi:</b> {self.e(v["kinh_gui"])}</li>'
                  f'<li><b>Trích yếu:</b> {self.e(v["trich_yeu"])}. <b>Loại:</b> {self.e(v["loai"])}</li>'
                  f'<li><b>Chức năng pháp lý:</b> {self.e(v["chuc_nang_phap_ly"])}</li>'
                  f'<li><b>Liên quan hồ sơ:</b> {self.e(v["lien_quan_ho_so"])}. <b>Tệp:</b> {tep}</li>'
                  + (f'<li><b>Nội dung, nhận xét:</b> {self.ma(v["ghi_chu"])}</li>' if v['ghi_chu'] else '')
                  + '</ul>']
        k.append('<h2>Căn cứ pháp lý đã dùng và mức độ đã đối chiếu</h2>')
        k.append('<ul>' + ''.join(
            f'<li><b>{self.e(c["van_ban"])}, {self.e(c["dieu_khoan"])}:</b> {self.e(c["noi_dung_tom_tat"])}. '
            f'<i>Hiệu lực áp dụng:</i> {self.e(c["hieu_luc_ap_dung"])}. <i>Đối chiếu:</i> {self.e(c["trang_thai_doi_chieu"])}. '
            f'<i>Dùng cho:</i> {self.e(c["dung_cho"])}</li>' for c in CC
            if 'bộ công khai' not in c['dung_cho']) + '</ul>')
        return self.goi(k)

    # ---------- 04: luận điểm và rà soát ----------
    def luan_diem(self, day_du=False):
        ten = {'A': 'A. Sai, cần sửa', 'B': 'B. Chưa chặt, cần làm rõ', 'C': 'C. Phát hiện mới có chứng cứ', 'D': 'D. Đã vững'}
        k = ['<h1>Luận điểm, chứng cứ và rà soát độc lập</h1>', self.chu_thich(),
             '<p>Phần 1 liệt kê các luận điểm chính của hồ sơ, tách sự kiện kiểm chứng được khỏi phần suy luận, kèm trạng thái. '
             'Phần 2 là kết quả rà soát độc lập các tài liệu trong hồ sơ (11/2026), chia bốn nhóm: A sai cần sửa; B chưa chặt; '
             'C phát hiện mới có chứng cứ; D đã vững. Bản đã nộp cơ quan giữ nguyên; các điểm A, B được ghi ở đây để người đọc '
             'không bị dẫn sai, không sửa ngầm vào bản đã nộp.</p>',
             '<h2>Phần 1. Luận điểm</h2>']
        for l in LD:
            k.append(f'<h3>{l["id"]}. {self.e(l["nhan_dinh"])}</h3>')
            k.append('<ul>'
                     f'<li><b>Nguồn nhận định:</b> {self.ma(l["nguon_nhan_dinh"])}</li>'
                     f'<li><b>Chứng cứ trực tiếp:</b> {self.ma(l["chung_cu_truc_tiep"])}</li>'
                     f'<li><b>Phần suy luận:</b> {self.ma(l["phan_suy_luan"])}</li>'
                     f'<li><b>Trạng thái:</b> {self.e(l["trang_thai"])}'
                     + (f' (xem {self.e(l["kiem_toan"])})' if l['kiem_toan'] not in ('', '-') else '') + '</li></ul>')
        k.append('<h2>Phần 2. Rà soát độc lập</h2>')
        for nhom in 'ABCD':
            ds = [r for r in KT if r['phan_loai'] == nhom and (day_du or r['id'] not in KT_RIENG)]
            if not ds:
                continue
            k.append(f'<h3>{ten[nhom]} ({len(ds)})</h3>')
            k.append('<ul>' + ''.join(
                f'<li><b>{r["id"]}</b> ({self.ma(r["vi_tri"])}): {self.ma(r["phat_hien"])} '
                f'<i>Chứng cứ:</i> {self.ma(r["chung_cu"])}. <i>Đề xuất:</i> {self.ma(r["de_xuat"])}. '
                f'<i>Ảnh hưởng:</i> {self.e(r["anh_huong"])}.</li>' for r in ds) + '</ul>')
        if not day_du:
            k.append(f'<p><i>{len(KT_RIENG)} điểm rà soát về quyền riêng tư, liên kết nội bộ và độ đầy đủ của thư mục chỉ có trong bộ làm việc.</i></p>')
        return self.goi(k)


def main():
    if len(sys.argv) < 3 or sys.argv[1] not in ('md', 'html'):
        print(__doc__)
        sys.exit(2)
    che_do, ra = sys.argv[1], pathlib.Path(sys.argv[2])
    xb = XuatBan(che_do, pathlib.Path(sys.argv[3]) if che_do == 'html' else None)
    ra.mkdir(parents=True, exist_ok=True)
    duoi = '.html' if che_do == 'html' else '.md'
    tep = {'01_ho_so_theo_don': xb.ho_so_theo_don(), '02_tra_cuu_theo_van_de': xb.theo_van_de(),
           '04_luan_diem_va_ra_soat': xb.luan_diem(), '05_van_ban_co_quan': xb.van_ban()}
    for ten, nd in tep.items():
        (ra / (ten + duoi)).write_text(nd, encoding='utf-8')
        print(f'{ten}{duoi}: {len(nd)} ký tự')


if __name__ == '__main__':
    main()
