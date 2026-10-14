// Bộ dựng giáo án Tin học 6 tích hợp NLS & AI (định dạng theo Phụ lục IV – CV 5512)
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, BorderStyle, Footer, PageNumber, TabStopType, PositionalTab,
  PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader, VerticalAlign, PageBreak,
} = require('docx');

const FONT = 'Times New Roman';
const BODY = 26; // 13pt
const TBL = 24; // 12pt
const SMALL = 22; // 11pt
const PAGE_W = 11906, PAGE_H = 16838; // A4
const M = { top: 1134, bottom: 1134, left: 1418, right: 851, header: 567, footer: 567 };
const CW = PAGE_W - M.left - M.right; // 9637 DXA ≈ 17 cm

const C = {
  title: '1F3864',
  nls: '1F4E79', nlsFill: 'DEEAF6',
  ai: 'C00000', aiFill: 'FCE4D6',
  mixFill: 'E8E0F0',
  head: 'D9E2F3', label: 'F2F2F2', banner: 'FFF2CC',
};

const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDERS = { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

// ---------- inline: **đậm**, *nghiêng*, [NLS x] / [AI y] tô màu ----------
function inline(text, o = {}) {
  text = String(text).replace(/"([^"\n]*)"/g, '\u201C$1\u201D');
  const base = {};
  if (o.size) base.size = o.size;
  if (o.bold) base.bold = true;
  if (o.italics) base.italics = true;
  if (o.color) base.color = o.color;
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|\[(?:NLS|AI) [^\]]+\])/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ text: t.slice(2, -2), ...base, bold: true }));
    else if (t.startsWith('*')) out.push(new TextRun({ text: t.slice(1, -1), ...base, italics: true }));
    else out.push(new TextRun({ text: t.replace(/ /g, '\u00A0'), ...base, bold: true, color: t.startsWith('[AI') ? C.ai : C.nls }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}

const LABEL_RE = /^(GV|HS|GV và HS|Kết quả|Minh chứng|Tiêu chí đạt|Lưu ý|Đáp án|Gợi ý|Câu \d+[a-z]?|Bài \d+|Nhiệm vụ \d+|Nhóm \d+(?:, \d+)*|Bước \d+|B\d+)( \([^)]*\))?:\s*/;

// ---------- một dòng nội dung -> Paragraph (hoặc Table) ----------
// Quy ước: "#..." tiêu đề bước; "- " gạch đầu dòng; "+ " ý con; "> " nghiêng;
// "[NLS x] ..." / "[AI y] ..." dòng tích hợp (tô nền); "GV:", "HS:", "Kết quả:" ... in đậm nhãn.
function para(item, o = {}) {
  if (item && typeof item === 'object') {
    if (item.table) return tableBlock(item.table, o);
    if (item.pageBreak) return new Paragraph({ children: [new PageBreak()] });
    if (item.code) {
      const [k, v] = item.code.split(' ');
      return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, shading: { type: ShadingType.CLEAR, color: 'auto', fill: k === 'AI' ? C.aiFill : C.nlsFill }, children: [new TextRun({ text: k, bold: true, size: o.size, color: k === 'AI' ? C.ai : C.nls }), new TextRun({ break: 1 }), new TextRun({ text: v, bold: true, size: o.size, color: k === 'AI' ? C.ai : C.nls })] });
    }
    if (item.center) return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: inline(item.center, { size: o.size, bold: item.bold }) });
  }
  let s = String(item);
  const size = o.size || BODY;
  const spacing = { after: o.after ?? 20, line: o.line ?? 264 };
  const pre = [];
  let indent, fill = null, bold = false, italics = false, keepNext = o.keepNext;
  const tags = [];
  let tm;
  while ((tm = s.match(/^\[(NLS|AI) ([^\]]+)\]\s*/))) {
    tags.push(tm[1]);
    pre.push(new TextRun({ text: `[${tm[1]}\u00A0${tm[2]}] `, bold: true, size, color: tm[1] === 'AI' ? C.ai : C.nls }));
    s = s.slice(tm[0].length);
  }
  if (tags.length) fill = tags.includes('NLS') && tags.includes('AI') ? C.mixFill : (tags[0] === 'AI' ? C.aiFill : C.nlsFill);
  if (s.startsWith('#')) {
    s = s.replace(/^#\s*/, ''); bold = true; spacing.before = 50; keepNext = !o.inTable;
  } else if (s.startsWith('- ')) {
    s = s.slice(2); indent = { left: 227, hanging: 170 }; pre.push(new TextRun({ text: '– ', size }));
  } else if (s.startsWith('+ ')) {
    s = s.slice(2); indent = { left: 454, hanging: 170 }; pre.push(new TextRun({ text: '+ ', size }));
  } else if (s.startsWith('> ')) {
    s = s.slice(2); italics = true;
  }
  const lm = s.match(LABEL_RE);
  if (lm) { pre.push(new TextRun({ text: lm[0], bold: true, size, italics })); s = s.slice(lm[0].length); }
  return new Paragraph({
    children: [...pre, ...inline(s, { size, bold, italics })],
    spacing, indent, keepNext,
    alignment: o.align || (o.inTable ? AlignmentType.LEFT : AlignmentType.JUSTIFIED),
    shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined,
  });
}

function paras(v, o = {}) {
  const arr = Array.isArray(v) ? v : (v === undefined || v === null ? [] : [v]);
  return arr.map((x) => para(x, o));
}

function cell(children, width, o = {}) {
  const kids = children.slice();
  if (!kids.length || kids[kids.length - 1] instanceof Table) kids.push(new Paragraph({ children: [] }));
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 40, bottom: 40, left: 85, right: 85 },
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    columnSpan: o.span,
    verticalAlign: o.vAlign || VerticalAlign.TOP,
    children: kids,
  });
}

function widthsOf(rel, W) {
  const sum = rel.reduce((a, b) => a + b, 0);
  const w = rel.map((r) => Math.floor((W * r) / sum));
  w[w.length - 1] = W - w.slice(0, -1).reduce((a, b) => a + b, 0);
  return w;
}

// Bảng nội dung tổng quát: { cols:[tỉ lệ], header:[...], rows:[[...]], size }
function tableBlock(t, o = {}) {
  const W = t.width || o.width || CW;
  const widths = widthsOf(t.cols, W);
  const size = t.size || (o.size ? Math.min(o.size, TBL) : TBL);
  const rows = [];
  if (t.header) {
    rows.push(new TableRow({
      tableHeader: true,
      children: t.header.map((h, i) => cell([new Paragraph({ alignment: AlignmentType.CENTER, children: inline(h, { size, bold: true }) })], widths[i], { fill: C.head, vAlign: VerticalAlign.CENTER })),
    }));
  }
  t.rows.forEach((r) => {
    rows.push(new TableRow({
      children: r.map((v, i) => {
        const fill = t.firstColFill && i === 0 ? C.label : undefined;
        const align = t.center && t.center.includes(i) ? AlignmentType.CENTER : undefined;
        return cell(paras(v, { size, width: widths[i] - 170, align, inTable: true }), widths[i], { fill });
      }),
    }));
  });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows, borders: BORDERS });
}

// ---------- các khối định dạng ----------
const P = (children, opts = {}) => new Paragraph({ children, ...opts });
const R = (text, opts = {}) => new TextRun({ text, ...opts });

function h(text, level, extra = {}) {
  const map = { 1: 'Heading1', 2: 'Heading2', 3: 'Heading3', 4: 'Heading4' };
  return new Paragraph({ style: map[level], keepNext: true, children: inline(text), ...extra });
}

function dotLines(n) {
  const out = [];
  for (let i = 0; i < n; i++) {
    out.push(new Paragraph({ spacing: { after: 80 }, tabStops: [{ type: TabStopType.RIGHT, position: CW, leader: 'dot' }], children: [new TextRun({ text: '\t' })] }));
  }
  return out;
}

function headerBlock(L) {
  const w = widthsOf([46, 54], CW);
  const left = [
    P([R('TRƯỜNG THCS TRẦN PHÚ', { bold: true, size: 24 })], { alignment: AlignmentType.CENTER }),
    P([R('TỔ: TOÁN – TIN', { bold: true, size: 24 })], { alignment: AlignmentType.CENTER }),
  ];
  const right = [
    P([R('Họ và tên giáo viên: ', { size: 24 }), R('……………………………', { size: 24 })], { alignment: AlignmentType.CENTER }),
    P([R(`Ngày soạn: …/…/${L.nam}`, { size: 24, italics: true })], { alignment: AlignmentType.CENTER }),
  ];
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: w, borders: NO_BORDERS,
    rows: [new TableRow({ children: [cell(left, w[0]), cell(right, w[1])] })],
  });
}

function dateTable(L) {
  const n = L.tiet.length;
  const w = widthsOf(Array(n).fill(1), CW);
  const lop = L.lop || ['6A', '6B'];
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: w, borders: BORDERS,
    rows: [new TableRow({
      children: L.tiet.map((t, i) => cell([
        P([R(`Ngày dạy – Tiết ${t} (theo KHDH)`, { bold: true, size: 22 })], { alignment: AlignmentType.CENTER }),
        ...lop.map((l) => P([R(`Lớp ${l}: …/…/${L.nam}    Sĩ số: …/…`, { size: 22 })], { alignment: AlignmentType.CENTER })),
      ], w[i])),
    })],
  });
}

function tietText(tiet) {
  return tiet.length === 1 ? `TIẾT ${tiet[0]}` : `TIẾT ${tiet.join(', ')}`;
}

function badge(L) {
  const w = widthsOf([24, 76], CW);
  const rows = [];
  const add = (labels, items, fill) => rows.push(new TableRow({
    children: [
      cell(labels.map((lb) => P(inline(lb, { size: 23, bold: true }), { alignment: AlignmentType.CENTER })), w[0], { fill, vAlign: VerticalAlign.CENTER }),
      cell(items.map((x) => para(x, { size: 23, inTable: true })), w[1]),
    ],
  }));
  add(['Năng lực số (NLS)', 'theo KHDH'], L.nls.map((n) => `[NLS ${n.ma}] ${n.mota}`), C.nlsFill);
  if (L.ai && L.ai.length) add(['Năng lực AI', 'theo KHDH'], L.ai.map((n) => `[AI ${n.ma}] ${n.mota}`), C.aiFill);
  add(['Căn cứ'], [`Phụ lục I – Kế hoạch dạy học môn Tin học, khối 6, năm học 2026 – 2027 (Tổ Toán – Tin): Bài ${L.bai} – Tiết ${L.tiet.join(', ')}.`], C.label);
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: w, rows, borders: BORDERS });
}

function labeled(label, lines) {
  if (lines === undefined || lines === null) return [];
  const arr = [].concat(lines);
  const out = [];
  const first = arr[0];
  if (typeof first === 'string' && !/^(\[|- |\+ |#)/.test(first)) {
    out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 20, line: 264 }, children: [R(label + ' ', { bold: true }), ...inline(first)] }));
    arr.slice(1).forEach((l) => out.push(para(l)));
  } else {
    out.push(new Paragraph({ children: [R(label, { bold: true })] }));
    arr.forEach((l) => out.push(para(l)));
  }
  return out;
}

function activityTable(rows) {
  const w = widthsOf([58, 42], CW);
  const head = new TableRow({
    tableHeader: true,
    children: [
      cell([P([R('Hoạt động của GV và HS', { bold: true, size: TBL })], { alignment: AlignmentType.CENTER })], w[0], { fill: C.head }),
      cell([P([R('Sản phẩm dự kiến', { bold: true, size: TBL })], { alignment: AlignmentType.CENTER })], w[1], { fill: C.head }),
    ],
  });
  const body = rows.map((r) => new TableRow({
    children: [
      cell(paras(r.left, { size: TBL, width: w[0] - 170, inTable: true }), w[0]),
      cell(paras(r.right, { size: TBL, width: w[1] - 170, inTable: true }), w[1]),
    ],
  }));
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: w, rows: [head, ...body], borders: BORDERS });
}

function tietBanner(text) {
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], borders: BORDERS,
    rows: [new TableRow({ children: [cell([new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40 }, children: [R(text, { bold: true, color: C.title })] })], CW, { fill: C.banner })] })],
  });
}

function renderActivity(a) {
  const out = [];
  if (a.group) out.push(h(a.group, 3, { spacing: { before: 200, after: 60 } }));
  out.push(new Paragraph({
    keepNext: true, spacing: { before: 160, after: 60 },
    children: [
      R(String(a.ten).replace(/"([^"]*)"/g, '\u201C$1\u201D'), { bold: true, color: C.title }),
      ...(a.tg ? [R(` (${a.tg})`, { bold: true, italics: true, color: C.title })] : []),
      ...((a.tags || []).flatMap((t) => [R('  '), R(`[${t}]`, { bold: true, color: t.startsWith('AI') ? C.ai : C.nls })])),
    ],
  }));
  out.push(...labeled('a) Mục tiêu:', a.mucTieu));
  out.push(...labeled('b) Nội dung:', a.noiDung));
  out.push(...labeled('c) Sản phẩm:', a.sanPham));
  out.push(new Paragraph({ keepNext: true, children: [R('d) Tổ chức thực hiện:', { bold: true })] }));
  out.push(activityTable(a.rows));
  return out;
}

function summaryTable(L) {
  const hasAI = L.ai && L.ai.length;
  return tableBlock({
    cols: [5, 16, 11, 22, 22, 24], size: SMALL,
    header: ['TT', 'Hoạt động', hasAI ? 'Mã NLS / AI' : 'Mã NLS', 'Tổ chức của giáo viên', 'Hoạt động của học sinh', 'Kết quả đạt được (minh chứng)'],
    center: [0],
    rows: L.tongHop.map((r, i) => [String(i + 1), `**${r.hd}**`, r.ma.map((m) => ({ code: m })), r.gv, r.hs, r.kq]),
  });
}

function renderLesson(L) {
  const out = [];
  out.push(headerBlock(L));
  out.push(P([], { spacing: { after: 60 } }));
  out.push(dateTable(L));
  out.push(P([R(L.chuDe, { bold: true, size: 24 })], { alignment: AlignmentType.CENTER, spacing: { before: 160, after: 20 } }));
  out.push(new Paragraph({ style: 'Heading1', children: [R(`${tietText(L.tiet)} – BÀI ${L.bai}: ${L.ten}`)] }));
  out.push(P([R(`Môn học: Tin học – Lớp 6 (Kết nối tri thức với cuộc sống) | Thời gian thực hiện: ${String(L.tiet.length).padStart(2, '0')} tiết`, { italics: true, size: 24 })], { alignment: AlignmentType.CENTER, spacing: { after: L.ghiChuTiet ? 20 : 100 } }));
  if (L.ghiChuTiet) out.push(P([R(`(${L.ghiChuTiet})`, { italics: true, size: 22, color: C.ai })], { alignment: AlignmentType.CENTER, spacing: { after: 100 } }));
  out.push(badge(L));

  out.push(h('I. MỤC TIÊU', 2));
  out.push(h('1. Về kiến thức', 3));
  out.push(...paras(L.kienThuc.map((x) => '- ' + x)));
  out.push(h('2. Về năng lực', 3));
  out.push(h('2.1. Năng lực chung', 4));
  out.push(...paras(L.nlChung.map((x) => '- ' + x)));
  out.push(h('2.2. Năng lực Tin học', 4));
  out.push(...paras(L.nlTinHoc.map((x) => '- ' + x)));
  out.push(h('2.3. Năng lực số (NLS) – theo Kế hoạch dạy học', 4));
  L.nls.forEach((n) => {
    out.push(para(`[NLS ${n.ma}] **${n.mota}**`));
    out.push(para('- **Biểu hiện trong bài:** ' + n.bieuHien));
  });
  if (L.ai && L.ai.length) {
    out.push(h('2.4. Năng lực trí tuệ nhân tạo (AI) – theo Kế hoạch dạy học', 4));
    L.ai.forEach((n) => {
      out.push(para(`[AI ${n.ma}] **${n.mota}**`));
      out.push(para('- **Biểu hiện trong bài:** ' + n.bieuHien));
    });
  }
  out.push(h('3. Về phẩm chất', 3));
  out.push(...paras(L.phamChat.map((x) => '- ' + x)));

  out.push(h('II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 2));
  out.push(h('1. Giáo viên', 3));
  out.push(...paras(L.gv.map((x) => '- ' + x)));
  out.push(h('2. Học sinh', 3));
  out.push(...paras(L.hs.map((x) => '- ' + x)));

  out.push(h('III. TIẾN TRÌNH DẠY HỌC', 2));
  L.hoatDong.forEach((a) => {
    if (a.tietBanner) { out.push(P([], { spacing: { before: 120, after: 0 } })); out.push(tietBanner(a.tietBanner)); }
    else out.push(...renderActivity(a));
  });
  if (L.huongDan) {
    out.push(h('* Hướng dẫn về nhà', 4));
    out.push(...paras(L.huongDan.map((x) => '- ' + x)));
  }

  out.push(h(`IV. BẢNG TỔNG HỢP TÍCH HỢP NĂNG LỰC SỐ${L.ai && L.ai.length ? ' VÀ AI' : ''} TRONG BÀI HỌC`, 2));
  out.push(para('> Bảng thể hiện cụ thể việc tổ chức của GV, hoạt động của HS và kết quả/minh chứng đạt được gắn với từng mã năng lực theo KHDH.', { size: 24 }));
  out.push(summaryTable(L));

  out.push(h('V. RÚT KINH NGHIỆM, ĐIỀU CHỈNH SAU TIẾT DẠY', 2));
  out.push(...dotLines(3));

  if (L.phuLuc && L.phuLuc.length) {
    out.push(h('PHỤ LỤC – PHIẾU HỌC TẬP / HỌC LIỆU', 2));
    L.phuLuc.forEach((pl) => {
      out.push(P(inline(pl.ten, { bold: true }), { alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60 } }));
      out.push(...paras(pl.noiDung));
    });
  }
  return out;
}

function footer(text) {
  return new Footer({
    children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: CW }],
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: '808080', space: 2 } },
      children: [
        R(text, { size: 18, italics: true, color: '595959' }),
        new TextRun({ children: ['\tTrang ', PageNumber.CURRENT], size: 18, color: '595959' }),
      ],
    })],
  });
}

const STYLES = {
  default: {
    document: { run: { font: FONT, size: BODY }, paragraph: { spacing: { line: 264, after: 20 } } },
  },
  paragraphStyles: [
    { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 30, bold: true, color: C.title }, paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40 }, outlineLevel: 0, keepNext: true } },
    { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: BODY, bold: true, color: '000000' }, paragraph: { spacing: { before: 200, after: 60 }, outlineLevel: 1, keepNext: true } },
    { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: BODY, bold: true, color: '000000' }, paragraph: { spacing: { before: 100, after: 40 }, outlineLevel: 2, keepNext: true } },
    { id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: BODY, bold: true, italics: true, color: '000000' }, paragraph: { spacing: { before: 80, after: 40 }, outlineLevel: 3, keepNext: true } },
  ],
};

const SECTION_PROPS = {
  page: { size: { width: PAGE_W, height: PAGE_H }, margin: M },
};

function lessonSection(L) {
  return {
    properties: SECTION_PROPS,
    footers: { default: footer(`Tin học 6 – Bài ${L.bai}: ${L.tenNgan || L.ten} | Tích hợp NLS${L.ai && L.ai.length ? ' & AI' : ''} theo KHDH 2026 – 2027`) },
    children: renderLesson(L),
  };
}

function makeDoc(sections, title) {
  return new Document({
    creator: 'Tổ Toán – Tin, Trường THCS Trần Phú',
    title,
    description: 'Giáo án Tin học 6 tích hợp năng lực số và AI theo KHDH năm học 2026 – 2027',
    styles: STYLES,
    sections,
  });
}

module.exports = {
  Packer, Paragraph, TextRun, Table, TableRow, AlignmentType, PageBreak,
  inline, para, paras, tableBlock, h, P, R, CW, C, BODY, TBL, SMALL, SECTION_PROPS,
  lessonSection, makeDoc, footer, widthsOf, cell, BORDERS,
};
