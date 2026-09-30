/*
 * 결과 엑셀 만들기 (ExcelJS, 브라우저 안에서 생성 → 바로 내려받기).
 */
(function () {
  'use strict';

  const HEADERS = ['학번/학급', '이름', '날짜', '과목', '교시', '출결 종류', '담당 교사', '요청중/완료', '전달 사항'];
  const WIDTHS = [10.8, 9.2, 18.1, 21.9, 10.8, 18.9, 14.7, 12.4, 25.1];
  const VALID_ROWS = 1000;

  function styleHeader(row) {
    row.eachCell(c => {
      c.font = { bold: true };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F4F7' } };
      c.border = { bottom: { style: 'thin', color: { argb: 'FF9AA4B2' } } };
    });
  }

  function addValidations(ws, subjects, fromRow, toRow) {
    const list = (col, items) => {
      const f = '"' + items.join(',') + '"';
      // 엑셀 목록 유효성 검사는 255자 제한이 있다
      if (f.length > 255) return;
      for (let r = fromRow; r <= toRow; r++) {
        ws.getCell(col + r).dataValidation = { type: 'list', allowBlank: true, formulae: [f], showErrorMessage: true };
      }
    };
    list('E', ['1교시', '2교시', '3교시', '4교시', '5교시', '6교시', '7교시']);
    list('F', ['인정', '질병', '미인정', '기타']);
    list('H', ['요청중', '완료']);
    if (subjects && subjects.length) list('D', ['동아리', ...subjects]);
  }

  function toRowValues(r) {
    return [r.classCode, r.name, r.dateText, r.subject, r.periodText, r.kind || '', r.teacher, r.status, r.message];
  }

  function teacherSummary(rows) {
    const map = new Map();
    for (const r of rows) {
      const key = [r.teacher || '(교사 미확인)', r.dateText, r.periodText, r.subject].join('\u0001');
      if (!map.has(key)) map.set(key, { teacher: r.teacher || '(교사 미확인)', dateText: r.dateText, periodText: r.periodText, subject: r.subject, names: [] });
      const g = map.get(key);
      g.names.push(r.name);
      g.total = Math.max(g.total || 0, r.groupTotal || 0, g.names.length);
      g.more = g.total - g.names.length;
    }
    return [...map.values()].sort((a, b) =>
      a.teacher.localeCompare(b.teacher, 'ko') || a.dateText.localeCompare(b.dateText) || a.periodText.localeCompare(b.periodText, 'ko', { numeric: true }));
  }

  async function download(wb, filename) {
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function buildSummarySheet(wb, rows) {
    const ws = wb.addWorksheet('교사별 요약');
    ws.columns = [
      { header: '담당 교사', width: 14 }, { header: '날짜', width: 14 }, { header: '교시', width: 9 },
      { header: '과목', width: 22 }, { header: '학생 수', width: 8 }, { header: '학생', width: 60 },
    ];
    styleHeader(ws.getRow(1));
    for (const s of teacherSummary(rows)) ws.addRow([s.teacher, s.dateText, s.periodText, s.subject, s.total, s.names.join(', ') + (s.more ? ` 외 ${s.more}명` : '')]);
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: 'A1', to: 'F1' };
  }

  /** 새 엑셀 파일 */
  async function createNew(rows, { sheetName, subjects, filename, withSummary }) {
    const wb = new ExcelJS.Workbook();
    wb.creator = '미마감 정리 도우미';
    const ws = wb.addWorksheet(sheetName || '미마감 정리');
    ws.columns = HEADERS.map((h, i) => ({ header: h, width: WIDTHS[i] }));
    styleHeader(ws.getRow(1));
    rows.forEach(r => ws.addRow(toRowValues(r)));
    ws.eachRow((row, i) => { if (i > 1) row.alignment = { vertical: 'middle' }; });
    ws.getColumn(1).alignment = { horizontal: 'center' };
    ws.getColumn(5).alignment = { horizontal: 'center' };
    ws.getColumn(8).alignment = { horizontal: 'center' };
    addValidations(ws, subjects, 2, Math.max(VALID_ROWS, rows.length + 200));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: 'A1', to: 'I1' };
    if (withSummary) buildSummarySheet(wb, rows);
    await download(wb, filename);
  }

  const normDate = v => {
    if (v instanceof Date) return `${v.getFullYear()}.${v.getMonth() + 1}.${v.getDate()}`;
    const m = String(v || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    return m ? `${+m[1]}.${+m[2]}.${+m[3]}` : String(v || '').trim();
  };
  const cellText = v => {
    if (v == null) return '';
    if (typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('');
    if (typeof v === 'object' && 'result' in v) return String(v.result ?? '');
    return String(v).trim();
  };

  /** 기존 파일(누적 기록)에 이어 쓰기. 같은 학생·날짜·교시는 건너뛴다. */
  async function appendToExisting(file, rows, { filename }) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    // '이름'과 '교시' 머리글이 있는 첫 시트
    let ws = null, hdr = null;
    for (const s of wb.worksheets) {
      for (let r = 1; r <= Math.min(5, s.rowCount); r++) {
        const vals = [];
        s.getRow(r).eachCell({ includeEmpty: true }, (c, col) => { vals[col] = cellText(c.value); });
        if (vals.includes('이름') && vals.includes('교시')) { ws = s; hdr = { row: r, vals }; break; }
      }
      if (ws) break;
    }
    if (!ws) throw new Error('기존 파일에서 "이름", "교시" 머리글이 있는 시트를 찾지 못했습니다.');

    const col = name => hdr.vals.indexOf(name);
    const c = {
      cls: col('학번/학급'), name: col('이름'), date: col('날짜'), subject: col('과목'), period: col('교시'),
      kind: col('출결 종류'), teacher: col('담당 교사'), status: col('요청중/완료'), msg: col('전달 사항'),
    };

    const existing = new Set();
    let last = hdr.row;
    ws.eachRow((row, r) => {
      if (r <= hdr.row) return;
      const name = cellText(row.getCell(c.name).value);
      const date = c.date > 0 ? normDate(row.getCell(c.date).value) : '';
      if (name || date) last = r;
      const periods = cellText(row.getCell(c.period).value).split(/\s*,\s*/);
      for (const n of name.split(/\s*,\s*/)) {
        for (const p of periods) existing.add([n, date, p].join('|'));
      }
    });

    let added = 0, skipped = 0;
    for (const r of rows) {
      const periods = r.periodText.split(/\s*,\s*/);
      if (periods.every(p => existing.has([r.name, normDate(r.dateText), p].join('|')))) { skipped++; continue; }
      last++;
      const row = ws.getRow(last);
      const set = (i, v) => { if (i > 0) row.getCell(i).value = v; };
      set(c.cls, r.classCode); set(c.name, r.name); set(c.date, r.dateText); set(c.subject, r.subject);
      set(c.period, r.periodText); set(c.kind, r.kind || null); set(c.teacher, r.teacher);
      set(c.status, r.status); set(c.msg, r.message);
      row.commit();
      added++;
    }
    if (ws.autoFilter && typeof ws.autoFilter === 'string') {
      ws.autoFilter = ws.autoFilter.replace(/(\D+)(\d+)$/, (m, a, b) => a + Math.max(+b, last));
    }
    await download(wb, filename || file.name);
    return { added, skipped, sheet: ws.name };
  }

  window.Exporter = { createNew, appendToExisting, teacherSummary, HEADERS };
})();
