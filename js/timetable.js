/*
 * 학생 시간표 엑셀(xlsx/xls/csv) 읽기.
 * 파일은 브라우저 메모리에서만 읽고, 어디에도 저장하거나 전송하지 않습니다.
 *
 * 지원 형식
 *  (가) 가로형: 번호 | 이름 | 월1 | 월2 | … | 금7   (칸 내용: "과목(교사)", "과목/교사", "과목" 줄바꿈 "교사")
 *  (나) 세로형: 번호 | 이름 | 요일 | 교시 | 과목 | 교사
 *  - 이름이 '공통'(또는 '전체')인 행은 모든 학생에게 기본값으로 적용됩니다.
 *  - '과목-교사' 시트(과목 | 교사)가 있으면 교사 이름이 빠진 칸을 채웁니다.
 */
(function () {
  'use strict';

  const DAYS = ['월', '화', '수', '목', '금', '토'];
  const COMMON = /^(공통|전체|학급|반공통)$/;

  const norm = v => (v == null ? '' : String(v)).replace(/ /g, ' ').trim();

  function dayPeriodFromHeader(h) {
    const t = norm(h).replace(/\s+/g, '');
    const m = t.match(/^([월화수목금토])(?:요일)?[-_.]?([1-9])(?:교시)?$/);
    return m ? { day: m[1], period: +m[2] } : null;
  }

  function parseCell(v) {
    let t = norm(v);
    if (!t) return null;
    t = t.replace(/\r/g, '');
    let m = t.match(/^(.*?)\s*[(（[]\s*([^)）\]]*?)\s*[)）\]]\s*$/s);
    if (m && m[1]) return { subject: m[1].replace(/\s*\n\s*/g, ' ').trim(), teacher: m[2].trim() };
    const parts = t.split(/\s*(?:\n|\/|\||,)\s*/).filter(Boolean);
    if (parts.length >= 2) return { subject: parts[0], teacher: parts.slice(1).join(', ') };
    return { subject: t, teacher: '' };
  }

  function findHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const r = rows[i].map(norm);
      const hasName = r.some(c => /^(이름|성명|학생명)$/.test(c));
      const hasSlots = r.filter(c => dayPeriodFromHeader(c)).length >= 3;
      const hasLong = r.some(c => /^요일$/.test(c)) && r.some(c => /^교시$/.test(c));
      if (hasName && (hasSlots || hasLong)) return i;
    }
    return -1;
  }

  function colIndex(header, re) {
    return header.findIndex(c => re.test(norm(c)));
  }

  function parseNumber(v) {
    const t = norm(v);
    if (!t) return null;
    // "20701", "2-7-1", "1" 등 → 번호만
    const m = t.match(/(\d+)\s*$/);
    if (!m) return null;
    let n = +m[1];
    if (m[1].length >= 4) n = +m[1].slice(-2);
    return n > 0 && n < 100 ? n : null;
  }

  function sheetToRows(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, blankrows: false });
  }

  /**
   * @returns {{students: Map<number,{no,name,slots:Object}>, common: Object, subjects: Set, teacherOf: Object, warnings: string[], format: string}}
   */
  function parseWorkbook(wb) {
    const warnings = [];
    const teacherOf = {};
    // 과목-교사 표
    for (const name of wb.SheetNames) {
      const rows = sheetToRows(wb.Sheets[name]);
      if (!rows.length) continue;
      const h = rows[0].map(norm);
      const si = colIndex(h, /^(과목|과목명)$/), ti = colIndex(h, /^(교사|담당\s*교사|담당|교사명)$/);
      if (si >= 0 && ti >= 0 && colIndex(h, /^(이름|성명)$/) < 0) {
        for (const r of rows.slice(1)) {
          const s = norm(r[si]), t = norm(r[ti]);
          if (s && t) teacherOf[s] = t;
        }
      }
    }

    const students = new Map();
    const common = {};
    let format = '';
    for (const name of wb.SheetNames) {
      const rows = sheetToRows(wb.Sheets[name]);
      const hi = findHeaderRow(rows);
      if (hi < 0) continue;
      const header = rows[hi].map(norm);
      const ni = colIndex(header, /^(번호|출석번호|학번)$/);
      const nmi = colIndex(header, /^(이름|성명|학생명)$/);
      const slotCols = header.map((c, i) => ({ i, dp: dayPeriodFromHeader(c) })).filter(o => o.dp);
      const di = colIndex(header, /^요일$/), pi = colIndex(header, /^교시$/);
      const si = colIndex(header, /^(과목|과목명)$/), ti = colIndex(header, /^(교사|담당\s*교사|담당|교사명)$/);
      const isLong = slotCols.length < 3 && di >= 0 && pi >= 0 && si >= 0;
      format = isLong ? '세로형' : '가로형';

      let autoNo = 0;
      for (const r of rows.slice(hi + 1)) {
        const nm = norm(r[nmi]);
        if (!nm) continue;
        let target;
        if (COMMON.test(nm)) {
          target = common;
        } else {
          let no = ni >= 0 ? parseNumber(r[ni]) : null;
          if (no == null) no = ++autoNo; else autoNo = no;
          if (!students.has(no)) students.set(no, { no, name: nm, slots: {} });
          const st = students.get(no);
          if (st.name !== nm) warnings.push(`${no}번에 이름이 둘 이상입니다: ${st.name}, ${nm}`);
          target = st.slots;
        }
        if (isLong) {
          const d = norm(r[di]).replace(/요일$/, '');
          const p = parseNumber(r[pi]);
          if (!DAYS.includes(d) || !p) continue;
          const subj = norm(r[si]);
          if (!subj) continue;
          target[d + p] = { subject: subj, teacher: ti >= 0 ? norm(r[ti]) : '' };
        } else {
          for (const { i, dp } of slotCols) {
            const c = parseCell(r[i]);
            if (c) target[dp.day + dp.period] = c;
          }
        }
      }
      break; // 첫 번째 시간표 시트만 사용
    }

    if (!students.size) {
      throw new Error('시간표에서 학생 목록을 찾지 못했습니다. 첫 줄(머리글)에 "번호", "이름"과 "월1, 월2 …" 또는 "요일, 교시, 과목, 교사" 열이 있는지 확인해 주세요.');
    }

    const subjects = new Set();
    const fill = slots => {
      for (const k in slots) {
        const s = slots[k];
        if (!s.teacher && teacherOf[s.subject]) s.teacher = teacherOf[s.subject];
        subjects.add(s.subject);
      }
    };
    fill(common);
    for (const st of students.values()) fill(st.slots);

    return { students, common, subjects, teacherOf, warnings, format };
  }

  async function readFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false, codepage: 949 });
    return parseWorkbook(wb);
  }

  /** 날짜(YYYY-MM-DD)와 교시로 학생의 과목·교사를 찾는다. */
  function lookup(tt, no, dateStr, period) {
    if (!tt) return null;
    const d = new Date(dateStr + 'T00:00:00');
    const day = '일월화수목금토'[d.getDay()];
    const key = day + period;
    const st = tt.students.get(no);
    return (st && st.slots[key]) || tt.common[key] || null;
  }

  window.Timetable = { readFile, parseWorkbook, lookup, parseCell, dayPeriodFromHeader, DAYS };
})();
