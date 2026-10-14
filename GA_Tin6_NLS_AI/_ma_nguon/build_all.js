// node build_all.js <hk: 1|2> <outDir>
const fs = require('fs');
const path = require('path');
const lib = require('./lib');
const { Paragraph, TextRun, AlignmentType, PageBreak } = require('docx');
const hk = Number(process.argv[2]);
const outDir = process.argv[3];
const CFG = {
  1: { bai: [1, 2, 3, 4, 5, 6, 7, 8], nam: '2026', khac: [
        ['Kiểm tra giữa học kì I', '11'], ['Ôn tập học kì I', '16'], ['Kiểm tra cuối học kì I', '18']] },
  2: { bai: [9, 10, 11, 12, 13, 14, 15, 16, 17], nam: '2027', khac: [
        ['Kiểm tra giữa học kì II', '28'], ['Ôn tập học kì II', '32'], ['Kiểm tra cuối học kì II', '35']] },
}[hk];
const roman = hk === 1 ? 'I' : 'II';
fs.mkdirSync(outDir, { recursive: true });
const load = (n) => { const f = path.resolve(__dirname, 'lessons', `bai${String(n).padStart(2, '0')}.js`); delete require.cache[f]; return require(f); };
const fileName = (L) => `Tin 6 - Bài ${String(L.bai).padStart(2, '0')} - ${L.tenNgan} (NLS${L.ai && L.ai.length ? '-AI' : ''}).docx`;

function overview(lessons) {
  const C = lib.C;
  const out = [];
  out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: 'TRƯỜNG THCS TRẦN PHÚ – TỔ TOÁN – TIN', bold: true })] }));
  out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 120 }, children: [new TextRun({ text: 'GIÁO ÁN TIN HỌC 6', bold: true, size: 40, color: C.title })] }));
  out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: 'TÍCH HỢP NĂNG LỰC SỐ (NLS) VÀ TRÍ TUỆ NHÂN TẠO (AI)', bold: true, size: 32, color: C.title })] }));
  out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: `HỌC KÌ ${roman} – NĂM HỌC 2026 – 2027 (Bộ sách Kết nối tri thức với cuộc sống)`, bold: true, size: 28 })] }));
  out.push(lib.h(`BẢNG ĐỐI CHIẾU MÃ NLS VÀ AI THEO KẾ HOẠCH DẠY HỌC – HỌC KÌ ${roman}`, 2));
  const rows = lessons.map((L, i) => [
    String(i + 1),
    `**Bài ${L.bai}.** ${L.tenNgan}`,
    L.tiet.join(', '),
    L.nls.map((n) => ({ code: `NLS ${n.ma}` })),
    (L.ai && L.ai.length) ? L.ai.map((n) => ({ code: `AI ${n.ma}` })) : '—',
    [...new Set(L.tongHop.map((r) => r.hd.match(/^HĐ [\d.]+/)[0].replace(/\.$/, '')))].join('; '),
  ]);
  CFG.khac.forEach(([ten, t]) => rows.push(['', `*${ten}*`, t, '—', '—', '*KHDH không gắn mã NLS/AI*']));
  out.push(lib.tableBlock({
    cols: [5, 30, 9, 15, 13, 28], size: 22, center: [0, 2],
    header: ['TT', 'Bài học', 'Tiết', 'Mã NLS', 'Mã AI', 'Hoạt động có tích hợp trong giáo án'],
    rows,
  }));
  out.push(lib.h('CÁCH THỂ HIỆN VIỆC TÍCH HỢP TRONG GIÁO ÁN', 2));
  out.push(...lib.paras([
    '- Mã năng lực ghi đúng theo Phụ lục I – Kế hoạch dạy học môn Tin học khối 6 năm học 2026 – 2027 của tổ chuyên môn; mỗi bài có bảng "Mã tích hợp theo KHDH" ngay dưới tên bài.',
    '- Mục I.2.3 (NLS) và I.2.4 (AI) nêu mã, mô tả và **biểu hiện trong bài**.',
    '- Trong mục III (Tiến trình dạy học), các dòng tô nền xanh [NLS …] là việc làm cụ thể của **GV** và **HS** để phát triển năng lực số; các dòng tô nền cam [AI …] là việc làm gắn với năng lực AI; ở cột "Sản phẩm dự kiến", các dòng **Kết quả** là kết quả/minh chứng đạt được.',
    '- Mục IV của mỗi bài tổng hợp theo từng hoạt động: mã năng lực – tổ chức của GV – hoạt động của HS – kết quả đạt được (minh chứng).',
    '- Khi sử dụng công cụ AI: GV thao tác trên máy GV, HS không tự đăng kí tài khoản AI (bảo đảm quy định về độ tuổi và an toàn thông tin cá nhân).',
  ]));
  return out;
}

(async () => {
  const lessons = CFG.bai.map(load);
  for (const L of lessons) {
    const doc = lib.makeDoc([lib.lessonSection(L)], `Tin học 6 – Bài ${L.bai}: ${L.tenNgan}`);
    const buf = await lib.Packer.toBuffer(doc);
    fs.writeFileSync(path.join(outDir, fileName(L)), buf);
    console.log('OK', fileName(L), buf.length);
  }
  const sections = [{ properties: lib.SECTION_PROPS, footers: { default: lib.footer(`Giáo án Tin học 6 – Học kì ${roman} – Tích hợp NLS & AI theo KHDH 2026 – 2027`) }, children: overview(lessons) }];
  lessons.forEach((L) => sections.push(lib.lessonSection(load(L.bai))));
  const all = lib.makeDoc(sections, `Giáo án Tin học 6 – Học kì ${roman} – Tích hợp NLS & AI`);
  const buf = await lib.Packer.toBuffer(all);
  const name = `Tin 6 - HK${roman} - Giáo án tích hợp NLS-AI (Bài ${CFG.bai[0]}-${CFG.bai[CFG.bai.length - 1]}) - TỔNG HỢP.docx`;
  fs.writeFileSync(path.join(outDir, name), buf);
  console.log('OK', name, buf.length);
})().catch((e) => { console.error(e); process.exit(1); });
