/*
 * 결과 엑셀 만들기 (ExcelJS, 브라우저 안에서 생성 → 바로 내려받기).
 */
(function () {
  'use strict';

  const HEADERS = ['학번/학급', '이름', '날짜', '과목', '교시', '출결 종류', '담당 교사', '요청중/완료', '전달 사항'];
  const WIDTHS = [10.8, 24, 18.1, 25, 10.8, 18.9, 24, 12.4, 25.1];
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

  function fitRow(ws, row) {
    let lines = 1;
    row.eachCell((cell, col) => {
      cell.alignment = { ...cell.alignment, vertical: 'middle', wrapText: true };
      const width = Math.max(4, (ws.getColumn(col).width || 10) - 2);
      const length = [...String(cell.text || '')].reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 2 : 1), 0);
      lines = Math.max(lines, Math.ceil(length / width));
    });
    row.height = Math.min(409, Math.max(22, lines * 16));
  }

  function teacherSummary(rows) {
    const map = new Map();
    for (const r of rows) {
      const key = [r.teacher || '(교사 미확인)', r.dateText, r.periodText, r.subject].join('\u0001');
      if (!map.has(key)) map.set(key, { teacher: r.teacher || '(교사 미확인)', dateText: r.dateText, periodText: r.periodText, subject: r.subject, names: [] });
      const g = map.get(key);
      for (const name of r.name.split(/\s*,\s*/)) if (!g.names.includes(name)) g.names.push(name);
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
    for (const s of teacherSummary(rows)) fitRow(ws, ws.addRow([s.teacher, s.dateText, s.periodText, s.subject, s.total, s.names.join(', ') + (s.more ? ` 외 ${s.more}명` : '')]));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: 'A1', to: 'F1' };
  }

  /** 새 엑셀 파일 */
  async function createNew(rows, { sheetName, subjects, filename, withSummary }) {
    rows = RequestRows.unique(rows);
    const wb = new ExcelJS.Workbook();
    wb.creator = '미마감 정리 도우미';
    const ws = wb.addWorksheet(sheetName || '미마감 정리');
    ws.columns = HEADERS.map((h, i) => ({ header: h, width: WIDTHS[i] }));
    styleHeader(ws.getRow(1));
    rows.forEach(r => ws.addRow(toRowValues(r)));
    ws.eachRow((row, i) => { if (i > 1) fitRow(ws, row); });
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

  /** 기존 파일(누적 기록)에 이어 쓰기. 같은 날짜·교시·교사는 기존 행에 학생 이름을 합친다. */
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

    if (c.date < 1 || c.teacher < 1) throw new Error('기존 파일에 날짜와 담당 교사 열이 필요합니다.');
    const records=[];
    let last=hdr.row;
    const union=(...values)=>[...new Set(values.flatMap(v=>String(v||'').split(/\s*,\s*/)).filter(Boolean))].join(', ');
    ws.eachRow((row,i)=>{
      if(i<=hdr.row) return;
      if(!cellText(row.getCell(c.name).value) && !cellText(row.getCell(c.date).value)) return;
      last=Math.max(last,i);
      records.push({values:row.values.slice(), style:row.style, date:normDate(row.getCell(c.date).value)});
    });
    // Expand combined periods so a partial overlap never suppresses a new period.
    const atomic=records.flatMap(record=>cellText(record.values[c.period]).split(/\s*,\s*/).map(period=>{
      const values=record.values.slice();values[c.period]=period;
      return {...record,values,period,teacher:cellText(values[c.teacher])};
    }));
    const folded=[];
    function insert(record) {
      let names=Clubs.teachers(record.teacher);
      const matches=names.length ? folded.filter(x=>x.date===record.date && x.period===record.period && Clubs.teachers(x.teacher).some(t=>names.includes(t))) : [];
      if(!matches.length) {folded.push(record);return false;}
      const target=matches[0];
      for(const x of [...matches.slice(1),record]) {
        for(const key of ['cls','name','subject','teacher']) if(c[key]>0) target.values[c[key]]=union(cellText(target.values[c[key]]),cellText(x.values[c[key]]));
      }
      target.teacher=Clubs.teachers(target.values[c.teacher]).join(', ');target.values[c.teacher]=target.teacher;
      for(const x of matches.slice(1)) folded.splice(folded.indexOf(x),1);
      return true;
    }
    atomic.forEach(insert);
    let added=0, skipped=0;
    for(const r of rows) for(const period of r.periodText.split(/\s*,\s*/)) {
      const values=[];
      for(const [key,value] of Object.entries({cls:r.classCode,name:r.name,date:r.dateText,subject:r.subject,period,kind:r.kind||'',teacher:r.teacher,status:r.status,msg:r.message})) if(c[key]>0) values[c[key]]=value;
      if(insert({values,date:normDate(r.dateText),period,teacher:r.teacher})) skipped++; else added++;
    }
    // Keep custom columns from the first matching row and keep worksheet formatting.
    for(let i=hdr.row+1;i<=last;i++) ws.getRow(i).values=[];
    last=hdr.row;
    for(const record of folded) {
      const row=ws.getRow(++last);row.values=record.values;
      if(record.style) row.style=record.style;
      fitRow(ws, row);
    }
    const summary=wb.getWorksheet('교사별 요약');
    if(summary) {
      wb.removeWorksheet(summary.id);
      buildSummarySheet(wb,folded.map(x=>({teacher:x.teacher,dateText:cellText(x.values[c.date]),periodText:x.period,subject:c.subject>0?cellText(x.values[c.subject]):'',name:cellText(x.values[c.name])})));
    }
    if (ws.autoFilter && typeof ws.autoFilter === 'string') {
      ws.autoFilter = ws.autoFilter.replace(/(\D+)(\d+)$/, (m, a, b) => a + Math.max(+b, last));
    }
    await download(wb, filename || file.name);
    return { added, skipped, sheet: ws.name };
  }

  window.Exporter = { createNew, appendToExisting, teacherSummary, HEADERS };
})();
