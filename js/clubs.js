(function () {
  'use strict';
  const text = v => String(v ?? '').trim();
  const teachers = v => [...new Set(text(v).split(/[,，、/\n;·]+/).map(text).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'ko'));
  function parseRows(rows) {
    const hi = rows.findIndex(r => r.some(c=>/^(성명|이름|학생명)$/.test(text(c))) && r.some(c=>/^(부서명|동아리명|동아리)$/.test(text(c))));
    if(hi<0) throw new Error('동아리 명단에 번호·성명·부서명·담당교사 머리글이 필요합니다.');
    const h=rows[hi].map(v=>text(v).replace(/\s/g,''));
    const ni=h.findIndex(v=>/^(번호|학번|출석번호)$/.test(v)), namei=h.findIndex(v=>/^(성명|이름|학생명)$/.test(v));
    const si=h.findIndex(v=>/^(부서명|동아리명|동아리)$/.test(v));
    const tis=h.map((v,i)=>/^(담당교사|담당|교사|교사명)[12]?$/.test(v)?i:-1).filter(i=>i>=0);
    if(ni<0 || !tis.length) throw new Error('동아리 명단의 번호와 담당교사 열을 확인해 주세요.');
    const out=new Map();
    for(const r of rows.slice(hi+1)) {
      const raw=text(r[ni]); if(!/^\d+$/.test(raw)) continue;
      const no=+(raw.length>=4?raw.slice(-2):raw), name=text(r[namei]), subject=text(r[si]), teacher=teachers(tis.map(i=>text(r[i])).join(',')).join(', ');
      if(!name || !subject || !teacher) throw new Error(`${no}번 동아리 이름 또는 담당교사가 비어 있습니다.`);
      if(out.has(no)) throw new Error(`${no}번 학생이 두 번 있습니다. 한 학급 명단만 넣어 주세요.`);
      out.set(no,{no,name,subject,teacher});
    }
    if(!out.size) throw new Error('동아리 명단에서 학생을 찾지 못했습니다.');
    return out;
  }
  function pdfRows(items) {
    // PDF text may be split into individual characters. Rebuild visual rows, then use header columns.
    const lines=[];
    for(const it of [...items].sort((a,b)=>b.pos[1]-a.pos[1] || a.pos[0]-b.pos[0])) {
      let line=lines.find(l=>Math.abs(l.y-it.pos[1])<Math.max(2,it.size*.35));
      if(!line) {line={y:it.pos[1],items:[]};lines.push(line);} line.items.push(it);
    }
    const joined=l=>l.items.sort((a,b)=>a.pos[0]-b.pos[0]).map(i=>i.str).join('').replace(/\s/g,'');
    const header=lines.find(l=>/번호/.test(joined(l)) && /담당교사/.test(joined(l)));
    if(!header) throw new Error('PDF에서 동아리 명단 머리글을 찾지 못했습니다. 글자를 선택할 수 있는 PDF를 사용해 주세요.');
    const hs=header.items.sort((a,b)=>a.pos[0]-b.pos[0]);
    // Find each header's starting x even when the PDF stores one glyph per item.
    let full='', coords=[];
    for(const it of hs) {for(const ch of it.str.replace(/\s/g,'')) {full+=ch;coords.push(it.pos[0]);}}
    const starts=['번호','성명','부서명','담당교사'].map(label=>coords[full.indexOf(label)]);
    if(starts.some(v=>v===undefined)) throw new Error('PDF 명단 열(번호·성명·부서명·담당교사)을 확인해 주세요.');
    // A centered header starts to the right of a two-teacher cell's left edge.
    const teacherSize=hs.find(it=>it.str.includes('담당'))?.size || 10;
    starts[3]-=teacherSize*1.8;
    const rows=[['번호','성명','부서명','담당교사']];
    for(const line of lines.filter(l=>l.y<header.y-2)) {
      const row=['','','',''];
      for(const it of line.items.sort((a,b)=>a.pos[0]-b.pos[0])) {
        let col=0; for(let i=1;i<starts.length;i++) if(it.pos[0]>=starts[i]-3) col=i;
        row[col]+=it.str;
      }
      rows.push(row);
    }
    return rows;
  }
  async function readFile(file) {
    if(/\.pdf$/i.test(file.name)) {
      const pages=await TimetablePdf.readItems(file), out=new Map();
      for(const page of pages) for(const [no,r] of parseRows(pdfRows(page))) {
        if(out.has(no)) throw new Error(`${no}번 학생이 중복됩니다. 한 학급 명단만 사용해 주세요.`);
        out.set(no,r);
      }
      return out;
    }
    const wb=XLSX.read(await file.arrayBuffer(),{type:'array',codepage:949});
    for(const name of wb.SheetNames) {
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:false});
      if(rows.some(r=>r.some(c=>/^(부서명|동아리명|동아리)$/.test(text(c))))) return parseRows(rows);
    }
    throw new Error('동아리 시트를 찾지 못했습니다.');
  }
  window.Clubs={readFile,parseRows,pdfRows,teachers};
})();
