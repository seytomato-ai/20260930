/*
 * NEIS 출결 화면 캡처에서 '미마감' 칸을 찾아내는 모듈.
 * 모든 처리는 브라우저 메모리 안에서만 이루어지며, 어떤 데이터도 외부로 전송하지 않습니다.
 *
 *  1) 분홍색 테두리(미마감 칸)를 픽셀 색으로 찾는다.
 *  2) 글자 인식(OCR, 브라우저 안에서 실행)으로 'N교시' 머리글과 번호 열을 찾는다.
 *  3) 분홍 칸의 위치를 (학생 번호, 교시)로 변환한다.
 */
(function () {
  'use strict';

  // NEIS 미마감 칸 테두리는 CSS 'pink'(255,192,203)에 가깝다.
  function isPink(r, g, b) {
    return r >= 235 && g >= 140 && g <= 236 && b >= 150 && b <= 242 && r - g >= 18 && b >= g - 5;
  }

  function findPinkBoxes(imageData) {
    const { width: w, height: h, data } = imageData;
    const mask = new Uint8Array(w * h);
    for (let i = 0, p = 0; i < mask.length; i++, p += 4) {
      if (isPink(data[p], data[p + 1], data[p + 2])) mask[i] = 1;
    }
    const seen = new Uint8Array(w * h);
    const stack = new Int32Array(w * h);
    const comps = [];
    for (let start = 0; start < mask.length; start++) {
      if (!mask[start] || seen[start]) continue;
      let sp = 0;
      stack[sp++] = start;
      seen[start] = 1;
      let x0 = w, y0 = h, x1 = 0, y1 = 0, n = 0;
      while (sp) {
        const i = stack[--sp];
        const x = i % w, y = (i / w) | 0;
        n++;
        if (x < x0) x0 = x; if (x > x1) x1 = x;
        if (y < y0) y0 = y; if (y > y1) y1 = y;
        // 반경 2 이웃: 둥근 모서리의 옅은 픽셀이 끊겨도 한 덩어리로 묶는다
        for (let dy = -2; dy <= 2; dy++) {
          const yy = y + dy;
          if (yy < 0 || yy >= h) continue;
          for (let dx = -2; dx <= 2; dx++) {
            const xx = x + dx;
            if (xx < 0 || xx >= w) continue;
            const j = yy * w + xx;
            if (mask[j] && !seen[j]) { seen[j] = 1; stack[sp++] = j; }
          }
        }
      }
      comps.push({ x0, y0, x1, y1, w: x1 - x0 + 1, h: y1 - y0 + 1, n });
    }
    // 테두리 모양(가로로 긴 사각형)만 남긴다
    let boxes = comps.filter(c => c.w >= 20 && c.h >= 10 && c.w >= c.h * 1.2 && c.w <= w * 0.3);
    if (!boxes.length) return { boxes: [], partial: [] };
    const medH = median(boxes.map(b => b.h));
    const medW = median(boxes.map(b => b.w));
    const full = [], partial = [];
    for (const b of boxes) {
      const ok = b.h >= medH * 0.75 && b.w >= medW * 0.6 && b.h <= medH * 1.6;
      b.cx = (b.x0 + b.x1) / 2;
      b.cy = (b.y0 + b.y1) / 2;
      (ok ? full : partial).push(b);
    }
    return { boxes: full, partial };
  }

  /**
   * 분홍 칸 안의 글자 모양으로 '미마감'인지 가린다.
   *  - '미마감': 가로로 긴 세 글자 → 글자 영역이 높이보다 훨씬 넓다
   *  - 결석 '/': 좁은 사선 하나 → 글자 영역이 좁다
   *  - 빈 칸: 글자 없음
   */
  function classifyBox(imageData, b) {
    const { width: w, data } = imageData;
    const inset = Math.max(3, Math.round(b.h * 0.15));
    let x0 = Infinity, x1 = -1, y0 = Infinity, y1 = -1, n = 0;
    for (let y = b.y0 + inset; y <= b.y1 - inset; y++) {
      for (let x = b.x0 + inset; x <= b.x1 - inset; x++) {
        const p = (y * w + x) * 4;
        const r = data[p], g = data[p + 1], bl = data[p + 2];
        if (isPink(r, g, bl)) continue;
        const lum = 0.299 * r + 0.587 * g + 0.114 * bl;
        if (lum > 238) continue;
        n++;
        if (x < x0) x0 = x; if (x > x1) x1 = x;
        if (y < y0) y0 = y; if (y > y1) y1 = y;
      }
    }
    if (n < 4) return { kind: 'blank' };
    const iw = x1 - x0 + 1, ih = y1 - y0 + 1;
    // 세로로 잘라 글자 덩어리 수를 센다 (미마감 = 3덩어리)
    let blobs = 0, inBlob = false, gap = 0;
    for (let x = x0; x <= x1; x++) {
      let has = false;
      for (let y = y0; y <= y1 && !has; y++) {
        const p = (y * w + x) * 4;
        const lum = 0.299 * data[p] + 0.587 * data[p + 1] + 0.114 * data[p + 2];
        if (lum <= 238 && !isPink(data[p], data[p + 1], data[p + 2])) has = true;
      }
      if (has) { if (!inBlob && (gap >= 1 || blobs === 0)) blobs++; inBlob = true; gap = 0; }
      else { inBlob = false; gap++; }
    }
    if (iw >= ih * 1.8 && blobs >= 2) return { kind: 'unclosed', iw, ih, blobs };
    if (iw <= ih * 1.2 && blobs <= 1) return { kind: 'slash', iw, ih, blobs };
    return { kind: 'other', iw, ih, blobs };
  }

  function median(arr) {
    if (!arr.length) return 0;
    const s = [...arr].sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }

  function flattenWords(data, scale) {
    const words = [];
    const lines = [];
    for (const block of data.blocks || []) {
      for (const para of block.paragraphs || []) {
        for (const line of para.lines || []) {
          const lw = [];
          for (const wd of line.words || []) {
            const t = (wd.text || '').trim();
            if (!t) continue;
            const b = wd.bbox;
            const item = {
              text: t,
              conf: wd.confidence,
              x0: b.x0 / scale, y0: b.y0 / scale, x1: b.x1 / scale, y1: b.y1 / scale,
            };
            item.cx = (item.x0 + item.x1) / 2;
            item.cy = (item.y0 + item.y1) / 2;
            words.push(item);
            lw.push(item);
          }
          if (lw.length) lines.push(lw);
        }
      }
    }
    return { words, lines };
  }

  const PERIOD_RE = /([1-9])\s*[교고][시사]/;

  // 머리글 'N교시', '조회', '종례'의 x 좌표를 찾는다.
  function findHeaders(lines) {
    const found = {};
    for (const lw of lines) {
      for (let i = 0; i < lw.length; i++) {
        const w = lw[i];
        let m = w.text.match(PERIOD_RE);
        if (m) {
          addHeader(found, 'p' + m[1], w.cx, w.cy, w);
          continue;
        }
        // '1' '교시' 처럼 쪼개져 인식된 경우
        if (/^[1-9]$/.test(w.text) && lw[i + 1] && /^[교고]/.test(lw[i + 1].text)) {
          let nx = lw[i + 1];
          i++;
          if (nx.text.length === 1 && lw[i + 1] && /^[시사]/.test(lw[i + 1].text)) nx = lw[++i];
          addHeader(found, 'p' + w.text, (w.x0 + nx.x1) / 2, w.cy, w);
          continue;
        }
        // '조' '회' / '종' '례' 처럼 한 글자씩 쪼개진 경우도 붙여서 본다
        const nx = lw[i + 1];
        const pair = nx && nx.x0 - w.x1 < (w.y1 - w.y0) * 1.5 ? w.text + nx.text : '';
        if (/조회/.test(w.text)) addHeader(found, 'jo', w.cx, w.cy, w);
        else if (/종례/.test(w.text)) addHeader(found, 'jong', w.cx, w.cy, w);
        else if (/^조회/.test(pair)) { addHeader(found, 'jo', (w.x0 + nx.x1) / 2, w.cy, w); i++; }
        else if (/^종례/.test(pair)) { addHeader(found, 'jong', (w.x0 + nx.x1) / 2, w.cy, w); i++; }
        else if (/^번호/.test(w.text)) addHeader(found, 'no', w.cx, w.cy, w);
        else if (/^성명|^이름/.test(w.text)) addHeader(found, 'name', w.cx, w.cy, w);
        else if (/^마감$/.test(w.text)) addHeader(found, 'close', w.cx, w.cy, w);
      }
    }
    return found;
  }

  function addHeader(found, key, x, y, w) {
    // 같은 이름이 여러 번 나오면 가장 위(머리글 줄)에 있는 것을 쓴다
    if (!found[key] || y < found[key].y) found[key] = { x, y, word: w };
  }

  // 교시별 x 중심. 빠진 교시는 등간격으로 보간·연장한다.
  function buildColumns(headers, imageWidth) {
    const cols = [];
    const known = [];
    for (let p = 1; p <= 9; p++) if (headers['p' + p]) known.push({ p, x: headers['p' + p].x });
    if (known.length < 2) return { cols: [], spacing: 0 };
    // 최소제곱 직선 x = a + b*p
    const n = known.length;
    const mp = known.reduce((s, k) => s + k.p, 0) / n;
    const mx = known.reduce((s, k) => s + k.x, 0) / n;
    let num = 0, den = 0;
    for (const k of known) { num += (k.p - mp) * (k.x - mx); den += (k.p - mp) ** 2; }
    const b = num / den, a = mx - b * mp;
    const lastKnown = known[known.length - 1].p;
    for (let p = 1; p <= 9; p++) {
      const h = headers['p' + p];
      const x = h ? h.x : a + b * p;
      if (!h && p > lastKnown) {
        // 종례 머리글이 보이면 그 앞까지만, 안 보이면 7교시까지만 연장
        if (headers.jong ? x > headers.jong.x - b * 0.5 : p > 7) break;
        if (x > imageWidth) break;
      }
      cols.push({ key: p + '교시', period: p, x, guessed: !h });
    }
    if (headers.jo) cols.push({ key: '조회', period: 0, x: headers.jo.x });
    if (headers.jong) cols.push({ key: '종례', period: 99, x: headers.jong.x });
    return { cols, spacing: Math.abs(b) };
  }

  // 번호 열의 숫자들로 y → 번호 직선을 만든다.
  function buildRows(words, headers, headerBottom, firstColX) {
    let cand = words.filter(w => /^\d{1,2}$/.test(w.text) && w.cy > headerBottom && +w.text >= 1 && +w.text <= 60);
    if (headers.no) {
      cand = cand.filter(w => Math.abs(w.cx - headers.no.x) < 60);
    } else {
      if (firstColX) cand = cand.filter(w => w.cx < firstColX);
      if (cand.length) {
        const minX = Math.min(...cand.map(w => w.cx));
        cand = cand.filter(w => w.cx < minX + 40);
      }
    }
    cand.sort((a, b) => a.cy - b.cy);
    if (cand.length < 2) return { rows: cand.map(w => ({ no: +w.text, y: w.cy })), fit: null };
    // 인접 쌍으로 행 높이를 추정 (잘못 읽은 숫자에 강하도록 중앙값 사용)
    const slopes = [];
    for (let i = 0; i < cand.length; i++) {
      for (let j = i + 1; j < cand.length; j++) {
        const dn = +cand[j].text - +cand[i].text;
        if (dn > 0) slopes.push((cand[j].cy - cand[i].cy) / dn);
      }
    }
    const rowH = median(slopes.filter(s => s > 8));
    if (!rowH) return { rows: cand.map(w => ({ no: +w.text, y: w.cy })), fit: null };
    // 절편: 각 숫자가 가리키는 y0 = y - no*rowH 의 중앙값
    const y0 = median(cand.map(w => w.cy - +w.text * rowH));
    // 직선에서 크게 벗어난 숫자(오인식)는 버린다
    const rows = cand
      .filter(w => Math.abs(w.cy - (y0 + +w.text * rowH)) < rowH * 0.4)
      .map(w => ({ no: +w.text, y: w.cy }));
    return { rows, fit: { y0, rowH } };
  }

  // 성명 열의 글자를 행별로 모은다 (한글은 음절 단위로 쪼개져 인식되는 경우가 많다)
  function findNames(words, rowsFit, headers) {
    const names = {};
    if (!rowsFit) return names;
    const left = headers.no ? headers.no.x + 20 : -Infinity;
    const right = headers.name && headers.close ? (headers.name.x + headers.close.x) / 2
      : headers.close ? headers.close.x - 40 : Infinity;
    const byRow = {};
    for (const w of words) {
      if (w.cx <= left || w.cx >= right) continue;
      const txt = w.text.replace(/[^가-힣]/g, '');
      if (!txt) continue;
      const f = (w.cy - rowsFit.y0) / rowsFit.rowH;
      if (Math.abs(f - Math.round(f)) > 0.4) continue;
      (byRow[Math.round(f)] = byRow[Math.round(f)] || []).push({ x: w.cx, txt });
    }
    for (const no in byRow) {
      const t = byRow[no].sort((a, b) => a.x - b.x).map(o => o.txt).join('');
      if (t.length >= 2 && t.length <= 6 && !/교시|번호|성명|마감|조회|종례/.test(t)) names[no] = t;
    }
    return names;
  }

  function parseMeta(text) {
    const meta = {};
    const t = text.replace(/\s+/g, ' ');
    const d = t.match(/(20\d\d)\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/);
    if (d) meta.date = `${d[1]}-${d[2].padStart(2, '0')}-${d[3].padStart(2, '0')}`;
    const c = t.match(/([1-6])\s*[가-힣]?\s*[년넌]\s*(\d{1,2})\s*반/);
    if (c) { meta.grade = +c[1]; meta.cls = +c[2]; }
    return meta;
  }

  /**
   * @param {HTMLImageElement|ImageBitmap|HTMLCanvasElement} img
   * @param {Tesseract.Worker} worker
   */
  async function analyze(img, worker) {
    const W = img.naturalWidth || img.width, H = img.naturalHeight || img.height;
    const base = document.createElement('canvas');
    base.width = W; base.height = H;
    const bctx = base.getContext('2d', { willReadFrequently: true });
    bctx.drawImage(img, 0, 0);
    const { boxes, partial } = findPinkBoxes(bctx.getImageData(0, 0, W, H));

    // 작은 화면 캡처는 확대하면 인식률이 좋아진다
    const scale = W < 1800 ? 2 : 1;
    const oc = document.createElement('canvas');
    oc.width = W * scale; oc.height = H * scale;
    const octx = oc.getContext('2d');
    octx.imageSmoothingQuality = 'high';
    octx.drawImage(base, 0, 0, oc.width, oc.height);

    const res = await worker.recognize(oc, {}, { blocks: true, text: true });
    const { words, lines } = flattenWords(res.data, scale);
    const meta = parseMeta(res.data.text || words.map(w => w.text).join(' '));
    const headers = findHeaders(lines);
    const { cols, spacing } = buildColumns(headers, W);
    const periodHeaders = Object.keys(headers).filter(k => /^p\d$/.test(k)).map(k => headers[k]);
    const headerBottom = periodHeaders.length
      ? Math.max(...periodHeaders.map(h => h.word.y1)) + 2
      : 0;
    const firstColX = cols.length ? Math.min(...cols.map(c => c.x)) : null;
    const { rows, fit } = buildRows(words, headers, headerBottom, firstColX);
    const names = findNames(words, fit, headers);

    const warnings = [];
    if (!cols.length) warnings.push('교시 머리글(1교시, 2교시 …)을 찾지 못했습니다. 머리글 줄이 보이도록 다시 캡처해 주세요.');
    if (!fit) warnings.push('번호 열을 읽지 못했습니다. 번호·성명 열이 보이도록 캡처해 주세요.');
    if (partial.length) warnings.push(`가려지거나 잘린 칸 ${partial.length}개는 제외했습니다. 필요하면 스크롤해서 한 장 더 캡처해 주세요.`);

    const cells = [];
    const unmatched = [];
    const excluded = [];
    const imgData = bctx.getImageData(0, 0, W, H);
    for (const b of boxes) {
      if (b.cy < headerBottom) continue;
      b.cls = classifyBox(imgData, b);
      if (b.cls.kind !== 'unclosed') { excluded.push(b); continue; }
      let col = null, best = Infinity;
      for (const c of cols) {
        const d = Math.abs(c.x - b.cx);
        if (d < best) { best = d; col = c; }
      }
      if (col && spacing && best > spacing * 0.6) col = null;
      let no = null;
      if (fit) {
        const f = (b.cy - fit.y0) / fit.rowH;
        if (Math.abs(f - Math.round(f)) < 0.35) no = Math.round(f);
      }
      if (!col || !no) { unmatched.push(b); continue; }
      cells.push({ no, col: col.key, period: col.period, box: b });
    }
    if (excluded.length) {
      const slash = excluded.filter(b => b.cls.kind === 'slash').length;
      warnings.push(`'미마감' 글자가 없는 분홍 칸 ${excluded.length}개${slash ? `(결석 "/" ${slash}개 포함)` : ''}는 제외했습니다.`);
    }
    if (unmatched.length) warnings.push(`위치를 판단하지 못한 미마감 칸 ${unmatched.length}개가 있습니다. 아래 표에서 직접 확인해 주세요.`);

    const visibleNos = [];
    if (fit) {
      const top = Math.max(headerBottom, 0), bottom = H;
      for (let n = 1; n <= 60; n++) {
        const y = fit.y0 + n * fit.rowH;
        if (y - fit.rowH * 0.3 > top && y + fit.rowH * 0.3 < bottom) visibleNos.push(n);
      }
    }

    return {
      meta, cells, cols, rows, fit, names, warnings, visibleNos,
      boxes, partial, unmatched, excluded, width: W, height: H,
    };
  }

  window.NeisDetect = { analyze, findPinkBoxes, classifyBox, parseMeta, _internal: { findHeaders, buildColumns, buildRows } };
})();
