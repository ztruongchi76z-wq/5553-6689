// Đối chiếu từng giáo án với Phụ lục I – KHDH Tin 6 năm học 2026-2027
const KHDH = {
  1: { tiet: [1, 2], nls: ['1.1.TC1a'], ai: [] },
  2: { tiet: [3, 4], nls: ['5.2.TC1b'], ai: ['6.A1.3'] },
  3: { tiet: [5, 6], nls: ['1.3.TC1a'], ai: [] },
  4: { tiet: [7, 8], nls: ['2.1.TC1a'], ai: [] },
  5: { tiet: [9, 10], nls: ['2.3.TC1a'], ai: ['6.A3.3'] },
  6: { tiet: [12, 13], nls: ['1.1.TC1c'], ai: [] },
  7: { tiet: [14, 15], nls: ['1.1.TC1b'], ai: [] },
  8: { tiet: [17], nls: ['2.2.TC1a'], ai: [] },
  9: { tiet: [19], nls: ['4.1.TC1a'], ai: ['6.B2.1'] },
  10: { tiet: [20, 21], nls: ['3.1.TC1a'], ai: ['6.A1.2'] },
  11: { tiet: [22, 23], nls: ['5.2.TC1b'], ai: [] },
  12: { tiet: [24, 25], nls: ['1.3.TC1b'], ai: [] },
  13: { tiet: [26], nls: ['3.2.TC1a'], ai: [] },
  14: { tiet: [27], nls: ['3.1.TC1a'], ai: [] },
  15: { tiet: [29, 30], nls: ['3.4.TC1a'], ai: [] },
  16: { tiet: [31, 33], nls: ['3.4.TC1a'], ai: [] },
  17: { tiet: [34], nls: ['3.4.TC1a'], ai: [] },
};
let bad = 0;
for (let n = 1; n <= 17; n++) {
  const L = require(`./lessons/bai${String(n).padStart(2, '0')}.js`);
  const k = KHDH[n];
  const errs = [];
  if (JSON.stringify(L.tiet) !== JSON.stringify(k.tiet)) errs.push(`tiết ${L.tiet} ≠ ${k.tiet}`);
  const nls = L.nls.map((x) => x.ma), ai = (L.ai || []).map((x) => x.ma);
  if (JSON.stringify(nls) !== JSON.stringify(k.nls)) errs.push(`NLS ${nls} ≠ ${k.nls}`);
  if (JSON.stringify(ai) !== JSON.stringify(k.ai)) errs.push(`AI ${ai} ≠ ${k.ai}`);
  // mọi mã xuất hiện trong nội dung phải thuộc mã của bài
  const text = JSON.stringify(L);
  const used = new Set([...text.matchAll(/\[(NLS|AI) ([^\]]+)\]|"(NLS|AI) ([0-9A-Z.]+[a-z0-9]*)"/g)].map((m) => (m[2] || m[4])));
  const allowed = new Set([...k.nls, ...k.ai]);
  for (const u of used) if (!allowed.has(u)) errs.push(`mã lạ trong nội dung: ${u}`);
  // tổng thời lượng mỗi tiết = 45 phút
  let cur = 0, mins = [0];
  for (const a of L.hoatDong) {
    if (a.tietBanner) { if (mins[cur] > 0) { cur++; mins[cur] = 0; } continue; }
    const m = String(a.tg || '').match(/^(\d+)/);
    mins[cur] += m ? Number(m[1]) : 0;
  }
  if (mins.length !== L.tiet.length || mins.some((x) => x !== 45)) errs.push(`thời lượng theo tiết: ${mins.join(' + ')} phút`);
  // mỗi mã của bài phải có trong bảng tổng hợp và trong tiến trình
  for (const c of [...k.nls.map((x) => 'NLS ' + x), ...k.ai.map((x) => 'AI ' + x)]) {
    if (!L.tongHop.some((r) => r.ma.includes(c))) errs.push(`bảng tổng hợp thiếu ${c}`);
    const inRows = L.hoatDong.some((a) => (a.rows || []).some((r) => JSON.stringify(r).includes('[' + c + ']')));
    if (!inRows) errs.push(`tổ chức thực hiện thiếu ${c}`);
  }
  console.log(`Bài ${String(n).padStart(2)}: tiết ${L.tiet.join(',').padEnd(6)} NLS ${nls.join(',').padEnd(9)} AI ${(ai.join(',') || '—').padEnd(7)} phút/tiết ${mins.join('/')} ${errs.length ? 'LỖI: ' + errs.join('; ') : 'ĐÚNG KHDH'}`);
  if (errs.length) bad++;
}
console.log(bad ? `${bad} bài có lỗi` : 'Tất cả 17 bài khớp KHDH.');
