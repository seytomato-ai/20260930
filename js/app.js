(function () {
  'use strict';

  const $ = sel => document.querySelector(sel);
  const el = (tag, attrs = {}, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else if (v !== false && v != null) e.setAttribute(k, v === true ? '' : v);
    }
    for (const k of kids) if (k != null) e.append(k);
    return e;
  };

  // ── 상태 (메모리에만 존재) ─────────────────────────────
  const state = {
    tt: null,          // 시간표
    shots: [],         // 캡처 이미지와 인식 결과
    overrides: new Map(), // `${date}|${no}|${col}` → true/false (사용자가 직접 바꾼 칸)
  };
  let worker = null;
  let workerPromise = null;
  let shotSeq = 0;

  const absUrl = p => new URL(p, location.href).href;

  function getWorker() {
    if (worker) return Promise.resolve(worker);
    if (!workerPromise) {
      showStatus('글자 인식 엔진을 준비하는 중… (처음 한 번만, 수 초 걸립니다)');
      workerPromise = Tesseract.createWorker('kor', 1, {
        workerPath: absUrl('vendor/tesseract/worker.min.js'),
        corePath: absUrl('vendor/tesseract/core/'),
        langPath: absUrl('vendor/tesseract/lang'),
        workerBlobURL: false,
        cacheMethod: 'none',
        gzip: true,
      }).then(w => (worker = w)).catch(e => { workerPromise = null; throw e; });
    }
    return workerPromise;
  }

  function showStatus(msg, isError) {
    const s = $('#ocrStatus');
    if (!msg) { s.hidden = true; return; }
    s.hidden = false;
    s.textContent = msg;
    s.classList.toggle('error', !!isError);
  }

  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 2600);
  }

  // ── 드롭존 공통 ─────────────────────────────────────
  function wireDrop(zone, input, onFiles) {
    // <label>은 브라우저가 알아서 파일 창을 연다
    if (zone.tagName !== 'LABEL') zone.addEventListener('click', e => { if (e.target !== input) input.click(); });
    zone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => { if (input.files.length) onFiles([...input.files]); input.value = ''; });
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('over'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('over');
      if (e.dataTransfer.files.length) onFiles([...e.dataTransfer.files]);
    });
  }

  // ── 1. 시간표 ────────────────────────────────────────
  async function loadTimetable(files) {
    const f = files[0];
    const box = $('#ttInfo');
    box.hidden = false;
    box.className = 'info';
    box.textContent = '시간표를 읽는 중…';
    try {
      const tt = await Timetable.readFile(f);
      state.tt = tt;
      const n = tt.students.size;
      const slots = new Set();
      for (const s of tt.students.values()) Object.keys(s.slots).forEach(k => slots.add(k));
      Object.keys(tt.common).forEach(k => slots.add(k));
      const noTeacher = [];
      const check = o => { for (const k in o) if (!o[k].teacher) noTeacher.push(o[k].subject); };
      check(tt.common);
      for (const s of tt.students.values()) check(s.slots);
      box.replaceChildren(
        el('div', {}, el('b', { text: '✔ 시간표를 읽었습니다. ' }),
          `학생 ${n}명 · ${tt.format} · 과목 ${tt.subjects.size}개 · 수업 칸 ${slots.size}개`),
        el('div', { class: 'small', text: '파일 자체는 저장되지 않고, 이 탭의 메모리에만 있습니다.' }),
        ...(noTeacher.length ? [el('div', { class: 'warn', text: `교사 이름이 없는 과목: ${[...new Set(noTeacher)].slice(0, 8).join(', ')} — '과목-교사' 시트로 채울 수 있어요.` })] : []),
        ...tt.warnings.slice(0, 5).map(w => el('div', { class: 'warn', text: w })),
      );
      render();
    } catch (e) {
      state.tt = null;
      box.className = 'info error';
      box.textContent = '시간표를 읽지 못했습니다: ' + (e.message || e);
    }
  }

  async function downloadTemplate() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('학생시간표');
    const head = ['번호', '이름'];
    for (const d of ['월', '화', '수', '목', '금']) for (let p = 1; p <= 7; p++) head.push(d + p);
    ws.addRow(head);
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { horizontal: 'center' };
    ws.getRow(1).eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FF' } }; });
    const common = ['', '공통'];
    const rowA = [1, '강가온'];
    const rowB = [2, '고나래'];
    for (let i = 2; i < head.length; i++) { common.push(''); rowA.push(''); rowB.push(''); }
    const at = k => head.indexOf(k);
    common[at('수1')] = '영어Ⅱ(한서준)';
    common[at('월1')] = '문학(윤지아)';
    rowA[at('월2')] = '기하(문지후)'; rowA[at('수5')] = '역학과 에너지(정우진)'; rowA[at('수6')] = '세계사(오하린)';
    rowB[at('월2')] = '경제(백서윤)'; rowB[at('수5')] = '식품과 영양(서예린)'; rowB[at('수6')] = '인공지능 기초(김도현)';
    ws.addRow(common); ws.addRow(rowA); ws.addRow(rowB);
    ws.getColumn(1).width = 6; ws.getColumn(2).width = 10;
    for (let i = 3; i <= head.length; i++) ws.getColumn(i).width = 16;
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 1 }];

    const mt = wb.addWorksheet('과목-교사');
    mt.addRow(['과목', '교사']);
    mt.getRow(1).font = { bold: true };
    mt.addRow(['영어Ⅱ', '한서준']);
    mt.addRow(['기하', '문지후']);
    mt.getColumn(1).width = 22; mt.getColumn(2).width = 12;

    const g = wb.addWorksheet('안내');
    [
      '■ 학생 시간표 양식 안내 (예시 이름·교사는 모두 가상입니다. 지우고 사용하세요)',
      '1. [학생시간표] 시트: 한 학생이 한 줄입니다. 번호는 나이스 출석번호와 같아야 합니다.',
      '2. 칸에는 "과목(교사)" 형태로 적습니다. 예) 영어Ⅱ(한서준)',
      '3. 이름을 "공통"으로 쓴 줄은 반 전체 기본 시간표입니다. 학생 칸이 비어 있으면 공통 줄을 씁니다.',
      '4. 칸에 과목만 적었다면 [과목-교사] 시트에 과목별 교사를 적으면 자동으로 채워집니다.',
      '5. 7교시가 없는 요일은 비워 두면 됩니다.',
    ].forEach(t => g.addRow([t]));
    g.getColumn(1).width = 100;
    const buf = await wb.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = el('a', { href: url, download: '학생시간표_양식.xlsx' });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ── 2. 캡처 ──────────────────────────────────────────
  async function addImages(files) {
    const imgs = files.filter(f => f.type.startsWith('image/'));
    if (!imgs.length) { toast('이미지 파일만 넣을 수 있어요.'); return; }
    for (const f of imgs) {
      const shot = { id: ++shotSeq, name: f.name || '붙여넣은 이미지', status: 'wait', result: null, date: '', img: null };
      state.shots.push(shot);
      const url = URL.createObjectURL(f);
      const img = new Image();
      img.src = url;
      try { await img.decode(); } catch { URL.revokeObjectURL(url); shot.status = 'error'; shot.error = '이미지를 열 수 없습니다.'; continue; }
      shot.img = img;
      shot.url = url;
      renderShots();
    }
    renderShots();
    await processQueue();
  }

  let processing = false;
  async function processQueue() {
    if (processing) return;
    processing = true;
    try {
      let w;
      try { w = await getWorker(); } catch (e) {
        showStatus('글자 인식 엔진을 불러오지 못했습니다: ' + (e.message || e), true);
        return;
      }
      for (const shot of state.shots) {
        if (shot.status !== 'wait') continue;
        shot.status = 'busy';
        renderShots();
        showStatus(`캡처를 분석하는 중… (${shot.name})`);
        try {
          const r = await NeisDetect.analyze(shot.img, w);
          shot.result = r;
          shot.date = r.meta.date || lastDate() || today();
          if (!r.meta.date) r.warnings.unshift('캡처에서 날짜를 읽지 못했습니다. 날짜를 확인해 주세요.');
          if (r.meta.grade && !$('#classCode').value.trim()) $('#classCode').value = `${r.meta.grade}-${r.meta.cls}`;
          shot.status = 'done';
        } catch (e) {
          shot.status = 'error';
          shot.error = '분석 실패: ' + (e.message || e);
        }
      }
      showStatus('');
    } finally {
      processing = false;
      render();
    }
  }

  const today = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  const lastDate = () => {
    for (let i = state.shots.length - 1; i >= 0; i--) if (state.shots[i].date) return state.shots[i].date;
    return '';
  };

  function drawOverlay(canvas, shot) {
    const img = shot.img;
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const r = shot.result;
    if (!r) return;
    ctx.lineWidth = Math.max(2, img.naturalWidth / 500);
    ctx.font = `bold ${Math.max(12, Math.round(r.fit ? r.fit.rowH * 0.42 : 14))}px sans-serif`;
    for (const c of r.cells) {
      const b = c.box;
      ctx.strokeStyle = '#e0457b';
      ctx.fillStyle = 'rgba(255, 192, 203, 0.35)';
      ctx.fillRect(b.x0, b.y0, b.w, b.h);
      ctx.strokeRect(b.x0 - 1, b.y0 - 1, b.w + 2, b.h + 2);
      const label = `${c.no}번 ${c.col}`;
      const tw = ctx.measureText(label).width;
      ctx.fillStyle = 'rgba(28, 36, 51, .85)';
      ctx.fillRect(b.x0, b.y0, tw + 8, b.h);
      ctx.fillStyle = '#fff';
      ctx.textBaseline = 'middle';
      ctx.fillText(label, b.x0 + 4, b.y0 + b.h / 2);
    }
    ctx.setLineDash([6, 4]);
    ctx.strokeStyle = '#b25b00';
    for (const b of [...r.unmatched, ...r.partial]) ctx.strokeRect(b.x0 - 2, b.y0 - 2, b.w + 4, b.h + 4);
    ctx.setLineDash([]);
  }

  function renderShots() {
    const wrap = $('#shots');
    wrap.replaceChildren();
    for (const shot of state.shots) {
      const card = el('div', { class: 'shot' + (shot.status === 'busy' || shot.status === 'wait' ? ' busy' : '') });
      if (shot.img) {
        const cv = el('canvas');
        drawOverlay(cv, shot);
        card.append(cv);
      }
      const meta = el('div', { class: 'meta' });
      if (shot.status === 'done') {
        const r = shot.result;
        const dateInput = el('input', { type: 'date', value: shot.date, 'aria-label': '출결 날짜' });
        dateInput.addEventListener('change', () => { shot.date = dateInput.value; render(); });
        meta.append(
          el('label', {}, '날짜 ', dateInput),
          el('span', { text: `미마감 ${r.cells.length}칸` }),
          el('span', { class: 'hint small', text: r.visibleNos.length ? `번호 ${r.visibleNos[0]}–${r.visibleNos[r.visibleNos.length - 1]}번` : '' }),
        );
      } else if (shot.status === 'error') {
        meta.append(el('span', { class: 'warn', text: shot.error }));
      } else {
        meta.append(el('span', { text: shot.status === 'busy' ? '분석 중…' : '대기 중…' }));
      }
      meta.append(el('button', { type: 'button', class: 'btn small remove', text: '빼기', onclick: () => removeShot(shot) }));
      card.append(meta);
      if (shot.result && shot.result.warnings.length) {
        card.append(el('ul', { class: 'warns' }, ...shot.result.warnings.map(w => el('li', { text: w }))));
      }
      wrap.append(card);
    }
  }

  function removeShot(shot) {
    state.shots = state.shots.filter(s => s !== shot);
    if (shot.url) URL.revokeObjectURL(shot.url);
    render();
  }

  // ── 3. 확인 표 ───────────────────────────────────────
  const COL_ORDER = c => (c === '조회' ? 0 : c === '종례' ? 99 : parseInt(c, 10));

  function collect() {
    // date → {detected: Map(no → Set(col)), cols:Set, visible:Set, names:{}}
    const byDate = new Map();
    for (const s of state.shots) {
      if (s.status !== 'done' || !s.date) continue;
      if (!byDate.has(s.date)) byDate.set(s.date, { detected: new Map(), cols: new Set(), visible: new Set(), names: {} });
      const d = byDate.get(s.date);
      for (const c of s.result.cols) d.cols.add(c.key);
      s.result.visibleNos.forEach(n => d.visible.add(n));
      Object.assign(d.names, s.result.names);
      for (const c of s.result.cells) {
        if (!d.detected.has(c.no)) d.detected.set(c.no, new Set());
        d.detected.get(c.no).add(c.col);
        d.cols.add(c.col);
      }
    }
    return byDate;
  }

  function isOn(date, no, col, detected) {
    const k = `${date}|${no}|${col}`;
    if (state.overrides.has(k)) return state.overrides.get(k);
    return !!(detected.get(no) && detected.get(no).has(col));
  }

  function studentName(no, ocrNames) {
    const st = state.tt && state.tt.students.get(no);
    return st ? st.name : (ocrNames[no] || '');
  }

  function renderGrids(byDate) {
    const host = $('#grids');
    host.replaceChildren();
    const dates = [...byDate.keys()].sort();
    for (const date of dates) {
      const d = byDate.get(date);
      let periods = [...d.cols].filter(c => /교시$/.test(c));
      const maxP = Math.max(7, ...periods.map(c => parseInt(c, 10)));
      const cols = [];
      if (d.cols.has('조회') && [...d.detected.values()].some(s => s.has('조회'))) cols.push('조회');
      for (let p = 1; p <= maxP; p++) cols.push(p + '교시');
      if ([...d.detected.values()].some(s => s.has('종례'))) cols.push('종례');

      const nos = new Set(d.visible);
      for (const no of d.detected.keys()) nos.add(no);
      if (state.tt) for (const no of state.tt.students.keys()) nos.add(no);
      const sorted = [...nos].sort((a, b) => a - b);
      const missing = state.tt ? sorted.filter(n => !d.visible.has(n) && !d.detected.has(n)) : [];

      const dt = new Date(date + 'T00:00:00');
      const title = `${date.replace(/-/g, '.')} (${'일월화수목금토'[dt.getDay()]})`;
      const table = el('table', { class: 'grid' });
      table.append(el('tr', {}, el('th', { text: '번호' }), el('th', { text: '이름' }), ...cols.map(c => el('th', { text: c }))));
      for (const no of sorted) {
        const tr = el('tr', { class: missing.includes(no) ? 'missing' : '' });
        const nm = studentName(no, d.names);
        tr.append(el('td', { text: no }), el('td', { class: 'name' }, nm || '—',
          missing.includes(no) ? el('small', { text: ' (캡처에 없음)' }) : null));
        for (const col of cols) {
          const on = isOn(date, no, col, d.detected);
          const period = parseInt(col, 10);
          const hit = period ? Timetable.lookup(state.tt, no, date, period) : null;
          const td = el('td', {
            class: 'cell' + (on ? ' on' : '') + (on && state.tt && !hit ? ' nosubj' : ''),
            title: `${no}번 ${nm} ${col}${hit ? ` · ${hit.subject} (${hit.teacher || '교사 미확인'})` : ''} — 눌러서 바꾸기`,
            text: hit ? hit.subject : (on ? '미마감' : ''),
            role: 'button', tabindex: '0', 'aria-pressed': on ? 'true' : 'false',
          });
          const toggle = () => { state.overrides.set(`${date}|${no}|${col}`, !isOn(date, no, col, d.detected)); render(); };
          td.addEventListener('click', toggle);
          td.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); } });
          tr.append(td);
        }
        table.append(tr);
      }
      const wrap = el('div', { class: 'grid-wrap' },
        el('h3', { text: `${title} · 미마감 ${countOn(date, d, cols, sorted)}칸` }),
        missing.length ? el('p', { class: 'hint small', text: `캡처에서 보이지 않은 번호: ${compressRanges(missing)} — 해당 학생이 있는 화면도 캡처해 넣어 주세요.` }) : null,
        el('div', { class: 'table-scroll' }, table));
      host.append(wrap);
    }
  }

  function countOn(date, d, cols, nos) {
    let n = 0;
    for (const no of nos) for (const c of cols) if (isOn(date, no, c, d.detected)) n++;
    return n;
  }

  function compressRanges(arr) {
    const out = [];
    for (let i = 0; i < arr.length; i++) {
      let j = i;
      while (j + 1 < arr.length && arr[j + 1] === arr[j] + 1) j++;
      out.push(i === j ? `${arr[i]}` : `${arr[i]}–${arr[j]}`);
      i = j;
    }
    return out.join(', ');
  }

  // ── 4. 결과 행 ───────────────────────────────────────
  function buildRows(byDate) {
    const classCode = $('#classCode').value.trim();
    const msg = $('#optMsg').value;
    const status = $('#optStatus').value;
    const rows = [];
    for (const date of [...byDate.keys()].sort()) {
      const d = byDate.get(date);
      const nos = new Set(d.detected.keys());
      for (const k of state.overrides.keys()) if (k.startsWith(date + '|')) nos.add(+k.split('|')[1]);
      const allCols = new Set(d.cols);
      for (let p = 1; p <= 9; p++) allCols.add(p + '교시');
      allCols.add('조회'); allCols.add('종례');
      for (const no of [...nos].sort((a, b) => a - b)) {
        const cols = [...allCols].filter(c => isOn(date, no, c, d.detected)).sort((a, b) => COL_ORDER(a) - COL_ORDER(b));
        for (const col of cols) {
          const period = parseInt(col, 10);
          const hit = period ? Timetable.lookup(state.tt, no, date, period) : null;
          rows.push({
            classCode, no, name: studentName(no, d.names) || `${no}번`,
            date, dateText: date.replace(/-/g, '.'), periodText: col, period: period || COL_ORDER(col),
            subject: hit ? hit.subject : (period ? '' : col), teacher: hit ? hit.teacher : '',
            status, message: msg, missing: !!period && !hit,
          });
        }
      }
    }
    if ($('#optMerge').checked) return mergeRows(rows);
    return rows;
  }

  function mergeRows(rows) {
    const out = [];
    for (const r of rows) {
      const prev = out[out.length - 1];
      if (prev && prev.no === r.no && prev.date === r.date && prev.subject && prev.subject === r.subject && prev.teacher === r.teacher) {
        prev.periodText += ', ' + r.periodText;
      } else out.push({ ...r });
    }
    return out;
  }

  function renderResult(rows) {
    const t = $('#resultTable');
    t.replaceChildren(el('tr', {}, ...Exporter.HEADERS.map(h => el('th', { text: h }))));
    for (const r of rows) {
      t.append(el('tr', { class: r.missing ? 'warnrow' : '' },
        el('td', { text: r.classCode }), el('td', { text: r.name }), el('td', { text: r.dateText }),
        el('td', { class: r.subject ? '' : 'empty', text: r.subject }), el('td', { text: r.periodText }), el('td', { text: '' }),
        el('td', { text: r.teacher }), el('td', { text: r.status }), el('td', { text: r.message })));
    }
  }

  let lastRows = [];
  function render() {
    renderShots();
    const byDate = collect();
    const has = byDate.size > 0;
    $('#step3').hidden = !has;
    $('#step4').hidden = !has;
    if (!has) { lastRows = []; return; }
    renderGrids(byDate);
    if (!state.tt) {
      $('#grids').prepend(el('p', { class: 'info error', text: '⚠ 시간표를 아직 넣지 않았습니다. 지금 이름은 글자 인식 결과라 틀릴 수 있고, 과목·교사 칸은 비어 있습니다. 1단계에서 시간표를 넣어 주세요.' }));
    }
    lastRows = buildRows(byDate);
    renderResult(lastRows);
  }

  function fileName() {
    const dates = [...new Set(lastRows.map(r => r.date))].sort();
    const cc = $('#classCode').value.trim() || '학급';
    const range = dates.length > 1 ? `${dates[0]}~${dates[dates.length - 1]}` : (dates[0] || today());
    return `미마감정리_${cc}_${range}.xlsx`;
  }

  function ensureReady() {
    if (!lastRows.length) { toast('정리할 미마감 칸이 없습니다.'); return false; }
    if (!state.tt) toast('시간표 없이 만들면 과목·교사 칸이 비어 있습니다.');
    return true;
  }

  async function exportNew() {
    if (!ensureReady()) return;
    await Exporter.createNew(lastRows, {
      sheetName: $('#optSheet').value.trim() || '미마감 정리',
      subjects: state.tt ? [...state.tt.subjects] : [],
      filename: fileName(),
      withSummary: $('#optSummary').checked,
    });
    $('#exportMsg').textContent = `✔ ${lastRows.length}줄을 엑셀로 내려받았습니다.`;
  }

  async function exportAppend(files) {
    if (!ensureReady()) return;
    try {
      const res = await Exporter.appendToExisting(files[0], lastRows, { filename: files[0].name });
      $('#exportMsg').textContent = `✔ '${res.sheet}' 시트에 ${res.added}줄 추가${res.skipped ? `, 이미 있던 ${res.skipped}줄은 건너뜀` : ''}. 내려받은 파일로 기존 파일을 바꿔 주세요.`;
    } catch (e) {
      $('#exportMsg').textContent = '';
      toast('이어쓰기 실패: ' + (e.message || e));
    }
  }

  async function copyMessages() {
    if (!ensureReady()) return;
    const cc = $('#classCode').value.trim();
    const groups = Exporter.teacherSummary(lastRows);
    const byTeacher = new Map();
    for (const g of groups) {
      if (!byTeacher.has(g.teacher)) byTeacher.set(g.teacher, []);
      byTeacher.get(g.teacher).push(g);
    }
    const parts = [];
    for (const [teacher, list] of byTeacher) {
      const lines = list.map(g => {
        const d = new Date(g.dateText.replace(/\./g, '-') + 'T00:00:00');
        return `- ${d.getMonth() + 1}/${d.getDate()}(${'일월화수목금토'[d.getDay()]}) ${g.periodText} ${g.subject}: ${g.names.join(', ')}`;
      });
      parts.push(`${teacher} 선생님, ${cc ? cc + ' ' : ''}학생 교시 출결 마감 부탁드립니다.\n${lines.join('\n')}`);
    }
    const text = parts.join('\n\n');
    try {
      await navigator.clipboard.writeText(text);
      toast(`교사 ${byTeacher.size}명 분 요청 문구를 복사했습니다.`);
    } catch {
      const ta = el('textarea');
      ta.value = text;
      document.body.append(ta); ta.select(); document.execCommand('copy'); ta.remove();
      toast('요청 문구를 복사했습니다.');
    }
  }

  async function clearAll() {
    for (const s of state.shots) if (s.url) URL.revokeObjectURL(s.url);
    state.tt = null;
    state.shots = [];
    state.overrides.clear();
    lastRows = [];
    $('#ttInfo').hidden = true;
    $('#classCode').value = '';
    $('#exportMsg').textContent = '';
    if (worker) { try { await worker.terminate(); } catch { /* ignore */ } }
    worker = null; workerPromise = null;
    render();
    toast('불러온 시간표·캡처·결과를 모두 지웠습니다.');
  }

  // ── 연결 ────────────────────────────────────────────
  wireDrop($('#ttDrop'), $('#ttFile'), loadTimetable);
  wireDrop($('#imgDrop'), $('#imgFile'), addImages);
  $('#btnTemplate').addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); downloadTemplate(); });
  $('#btnNew').addEventListener('click', exportNew);
  $('#appendFile').addEventListener('change', e => { const f = [...e.target.files]; e.target.value = ''; if (f.length) exportAppend(f); });
  $('#btnCopy').addEventListener('click', copyMessages);
  $('#btnClear').addEventListener('click', clearAll);
  for (const id of ['#classCode', '#optMsg', '#optStatus', '#optMerge']) $(id).addEventListener('input', render);
  $('#optMerge').addEventListener('change', render);

  document.addEventListener('paste', e => {
    const files = [...(e.clipboardData ? e.clipboardData.items : [])]
      .filter(i => i.kind === 'file' && i.type.startsWith('image/'))
      .map(i => i.getAsFile()).filter(Boolean);
    if (files.length) { e.preventDefault(); addImages(files); }
  });

  // 테스트용 훅
  window.__app = { state, render, getRows: () => lastRows };
})();
