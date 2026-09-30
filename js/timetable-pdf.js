/*
 * 나이스 '학생별 시간표' PDF 읽기 (pdf.js, 브라우저 안에서만 처리).
 *
 * 한 페이지에 학생 한 명:  머리글 "2026학년도 2학기 2학년 7반" / "- 7 주차(2026.09.28.) 1번 강민지"
 * 표: 한쪽 축에 "1교시 … 7교시", 다른 축에 "월요일(09/28) … 금요일(10/02)"
 * 칸: "음악 감상과 비평-다"(반/분반 줄) + "음악 감상과 비평(심희진)"(줄바꿈될 수 있음)
 *
 * 페이지가 회전되어 있어도 되도록 PDF 좌표계에서 교시 축·요일 축 방향을 직접 계산한다.
 */
(function () {
  'use strict';

  let pdfjsPromise = null;
  function loadPdfjs() {
    if (!pdfjsPromise) {
      const base = new URL('vendor/pdfjs/', document.baseURI).href;
      pdfjsPromise = import(base + 'pdf.min.mjs').then(lib => {
        lib.GlobalWorkerOptions.workerSrc = base + 'pdf.worker.min.mjs';
        return { lib, base };
      });
    }
    return pdfjsPromise;
  }

  const sub = (p, q) => [p[0] - q[0], p[1] - q[1]];
  const dot = (p, q) => p[0] * q[0] + p[1] * q[1];
  const unit = p => { const n = Math.hypot(p[0], p[1]) || 1; return [p[0] / n, p[1] / n]; };

  function toItems(tc) {
    const out = [];
    for (const it of tc.items) {
      const str = (it.str || '').trim();
      if (!str) continue;
      const [a, b, c, d, e, f] = it.transform;
      const size = Math.hypot(a, b) || Math.hypot(c, d) || 10;
      const dir = unit([a, b]);
      const up = unit([c, d]);
      const w = it.width || str.length * size;
      const center = [e + dir[0] * w / 2 + up[0] * size * 0.35, f + dir[1] * w / 2 + up[1] * size * 0.35];
      out.push({ str, size, dir, up, pos: [e, f], center });
    }
    return out;
  }

  const SECTION_LINE = /(^\d\s*학년\s*\d+\s*반$)|(-\s*[가-힣A-Za-z0-9]{1,2}$)/;

  function parsePage(items, parseCell) {
    const periods = [], days = [];
    for (const it of items) {
      let m = it.str.match(/^([1-9])\s*교시$/);
      if (m) { periods.push({ p: +m[1], c: it.center }); continue; }
      m = it.str.match(/^([월화수목금토일])요일\s*(?:\(\s*(\d{1,2})\s*[/.]\s*(\d{1,2})\s*\))?$/);
      if (m) days.push({ day: m[1], mm: m[2] && +m[2], dd: m[3] && +m[3], c: it.center });
    }
    if (periods.length < 2 || days.length < 2) return null;
    periods.sort((x, y) => x.p - y.p);
    const order = '월화수목금토일';
    days.sort((x, y) => order.indexOf(x.day) - order.indexOf(y.day));

    const u = unit(sub(periods[periods.length - 1].c, periods[0].c)); // 교시 축
    const v = unit(sub(days[days.length - 1].c, days[0].c));           // 요일 축
    const pGap = Math.abs(dot(sub(periods[periods.length - 1].c, periods[0].c), u)) / (periods[periods.length - 1].p - periods[0].p);
    const dGap = Math.abs(dot(sub(days[days.length - 1].c, days[0].c), v)) / (order.indexOf(days[days.length - 1].day) - order.indexOf(days[0].day));
    const dayLabelU = days.reduce((s, d) => s + dot(d.c, u), 0) / days.length;
    const perLabelV = periods.reduce((s, p) => s + dot(p.c, v), 0) / periods.length;

    // 머리글: 번호·이름·학년반·주차 날짜
    const text = items.map(i => i.str).join(' ');
    const head = {};
    let m = text.match(/(\d{1,2})\s*번\s*([가-힣]{2,6})/);
    if (m) { head.no = +m[1]; head.name = m[2]; }
    m = text.match(/(\d)\s*학년\s*(\d{1,2})\s*반/);
    if (m) { head.grade = +m[1]; head.cls = +m[2]; }
    m = text.match(/주차\s*\(\s*(20\d\d)\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,2})/);
    if (m) { head.year = +m[1]; head.month = +m[2]; }

    const cells = new Map();
    const anchorSet = new Set();
    for (const it of items) {
      if (/^([1-9])\s*교시$/.test(it.str) || /요일/.test(it.str)) { anchorSet.add(it); continue; }
      const cu = dot(it.center, u), cv = dot(it.center, v);
      // 요일 이름 칸·교시 이름 칸보다 표 안쪽에 있어야 한다
      if (Math.abs(cu - dayLabelU) < it.size * 0.9 || Math.sign(cu - dayLabelU) !== Math.sign(dot(periods[0].c, u) - dayLabelU)) continue;
      if (Math.abs(cv - perLabelV) < dGap * 0.4 || Math.sign(cv - perLabelV) !== Math.sign(dot(days[0].c, v) - perLabelV)) continue;
      let bp = null, bpd = Infinity;
      for (const p of periods) { const dd = Math.abs(cu - dot(p.c, u)); if (dd < bpd) { bpd = dd; bp = p; } }
      let bd = null, bdd = Infinity;
      for (const d of days) { const dd = Math.abs(cv - dot(d.c, v)); if (dd < bdd) { bdd = dd; bd = d; } }
      if (bpd > pGap * 0.6 || bdd > dGap * 0.6) continue;
      const key = bd.day + '|' + bp.p;
      if (!cells.has(key)) cells.set(key, { day: bd, period: bp.p, items: [] });
      cells.get(key).items.push(it);
    }

    const slots = {}, dated = {};
    for (const cell of cells.values()) {
      // 분반 줄이 "과목-" / "라" 처럼 줄바꿈된 경우 분반 기호를 다시 붙인다
      const dangling = cell.items.filter(it => /-\s*$/.test(it.str));
      for (const dg of dangling) {
        const tails = cell.items.filter(it => it !== dg && /^[가-힣A-Za-z0-9]{1,2}$/.test(it.str));
        if (!tails.length) continue;
        tails.sort((x, y) => Math.hypot(...sub(x.center, dg.center)) - Math.hypot(...sub(y.center, dg.center)));
        dg.str += tails[0].str;
        cell.items = cell.items.filter(it => it !== tails[0]);
      }
      // 줄 단위로 묶기: 글자 '아래' 방향(-up)으로 정렬
      const its = cell.items.map(it => ({ it, line: -dot(it.pos, it.up), along: dot(it.pos, it.dir) }));
      its.sort((x, y) => x.line - y.line || x.along - y.along);
      const lines = [];
      for (const o of its) {
        const last = lines[lines.length - 1];
        if (last && Math.abs(o.line - last.line) < o.it.size * 0.4) last.parts.push(o);
        else lines.push({ line: o.line, parts: [o] });
      }
      const strs = lines.map(l => l.parts.sort((x, y) => x.along - y.along).map(p => p.it.str).join(' ').trim());
      const section = strs.filter(s => SECTION_LINE.test(s));
      const rest = strs.filter(s => !SECTION_LINE.test(s));
      if (!rest.length) continue;
      const joined = rest.join('');
      const parsed = parseCell(joined);
      if (!parsed || !parsed.subject) continue;
      // 줄바꿈 때문에 과목명이 깨졌으면 분반 줄("과목-다")의 과목명을 쓴다
      const secSubj = section.map(s => s.replace(/\s*-\s*[가-힣A-Za-z0-9]{1,2}$/, '')).find(s => !/학년/.test(s));
      if (secSubj && secSubj.replace(/\s/g, '') === parsed.subject.replace(/\s/g, '')) parsed.subject = secSubj;
      slots[cell.day.day + cell.period] = parsed;
      if (head.year && cell.day.mm) {
        let y = head.year;
        if (head.month && cell.day.mm < head.month - 6) y++;
        const ds = `${y}-${String(cell.day.mm).padStart(2, '0')}-${String(cell.day.dd).padStart(2, '0')}`;
        dated[ds + '|' + cell.period] = parsed;
      }
    }
    return { head, slots, dated };
  }

  async function readPdf(file, parseCell) {
    const { lib, base } = await loadPdfjs();
    const doc = await lib.getDocument({
      data: new Uint8Array(await file.arrayBuffer()),
      cMapUrl: base + 'cmaps/', cMapPacked: true,
      isEvalSupported: false, disableFontFace: true, useSystemFonts: false,
    }).promise;
    const pages = [];
    try {
      for (let i = 1; i <= doc.numPages; i++) {
        const page = await doc.getPage(i);
        const tc = await page.getTextContent();
        const r = parsePage(toItems(tc), parseCell);
        if (r) pages.push(r);
        page.cleanup();
      }
    } finally {
      await doc.destroy();
    }
    return pages;
  }

  window.TimetablePdf = { readPdf, parsePage, toItems };
})();
