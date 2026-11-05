// Regression checks use fictional fixtures only. Uploaded student files must never be committed.
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
const require=createRequire(import.meta.url);
const {chromium}=require('playwright-core');
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const types={'.html':'text/html','.js':'text/javascript','.mjs':'text/javascript','.css':'text/css','.png':'image/png','.svg':'image/svg+xml','.gz':'application/gzip'};
const server=http.createServer((req,res)=>{
  let f=path.join(root,decodeURIComponent(new URL(req.url,'http://x').pathname));
  if(fs.existsSync(f)&&fs.statSync(f).isDirectory()) f=path.join(f,'index.html');
  if(!f.startsWith(root+path.sep)||!fs.existsSync(f)){res.writeHead(404);return res.end();}
  res.setHeader('Content-Type',types[path.extname(f)]||'application/octet-stream');fs.createReadStream(f).pipe(res);
}).listen(0,'127.0.0.1');
await new Promise(r=>server.once('listening',r));
const base=`http://127.0.0.1:${server.address().port}/`;
const executablePath=process.env.CHROMIUM_PATH || (process.platform==='win32'?'C:/Program Files/Google/Chrome/Application/chrome.exe':'/opt/pw-browsers/chromium');
let browser;
const errors=[],external=[];
try {
  browser=await chromium.launch({executablePath,headless:true});
  const page=await browser.newPage({acceptDownloads:true});
  page.on('pageerror',e=>errors.push(e.message));
  page.on('request',r=>{if(!r.url().startsWith(base)&& !/^(blob|data):/.test(r.url())) external.push(r.url());});
  await page.goto(base);
  await page.setInputFiles('#ttFile',path.join(root,'tests/fixtures/fake-timetable.xlsx'));
  await page.locator('#weekFiles input[type=date]').fill('2026-09-21');
  await page.locator('#weekFiles input[type=date]').dispatchEvent('change');
  await page.waitForFunction(()=>!!window.__app.state.tt);
  assert.equal(await page.evaluate(()=>window.__app.state.tt.students.size),28);
  const checks=await page.evaluate(async()=>{
    const make=(teacher)=>{
      const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['번호','이름','월1','화1','수1'],[1,'가상학생',`수학(${teacher})`,'국어(교사국어)','영어(교사영어)']]),'시간표');
      return new File([XLSX.write(wb,{type:'array',bookType:'xlsx'})],'week.xlsx');
    };
    const tt=await Timetable.readFiles([make('첫주교사'),make('둘째주교사')],['2026-11-02','2026-11-09']);
    const twoWeeks=[Timetable.lookup(tt,1,'2026-11-02',1).teacher,Timetable.lookup(tt,1,'2026-11-09',1).teacher,Timetable.lookup(tt,1,'2026-11-16',1)];
    const clubs=Clubs.parseRows([['번호','성명','부서명','담당교사'],[1,'가상학생','과학부','교사갑, 교사을']]);
    tt.clubs=clubs;tt.students.get(1).dated['2026-11-06|1']={subject:'국어',teacher:'교사국어'};
    const pair=Timetable.lookup(tt,1,'2026-11-06',5);
    tt.clubs=new Map(); const missing=Timetable.lookup(tt,1,'2026-11-06',5);
    const sample=(name,teacher,periodText='5교시')=>({name,teacher,periodText,date:'2026-11-06',dateText:'2026.11.06',subject:'동아리',classCode:'2-1',status:'요청중',message:'확인 부탁드립니다.'});
    const rows=RequestRows.unique([sample('가상학생','교사갑, 교사을'),sample('가상둘','교사을, 교사갑'),sample('가상셋','교사갑, 교사을','6교시'),sample('가상넷',''),sample('가상다섯','')]);
    return {twoWeeks,pair,missing,rows};
  });
  assert.deepEqual(checks.twoWeeks,['첫주교사','둘째주교사',null]);
  assert.equal(checks.pair.teacher,'교사갑, 교사을');assert.equal(checks.missing,null);
  assert.equal(checks.rows.length,4);assert.equal(checks.rows[0].name,'가상학생, 가상둘');
  console.log('PASS weekly lookup, dual club teachers, missing roster, grouped names and unknown teachers');
  // Keep the original OCR regression: verify raw recognized cells before teacher grouping.
  await page.setInputFiles('#imgFile',['mock-part1.png','mock-part2.png'].map(n=>path.join(root,'tests/fixtures',n)));
  await page.waitForFunction(()=>window.__app.state.shots.length===2&&window.__app.state.shots.every(s=>s.status==='done'||s.status==='error'),null,{timeout:120000});
  const mockPage=await browser.newPage();await mockPage.goto(base+'tools/mock-neis.html');const mock=await mockPage.evaluate(()=>window.MOCK);await mockPage.close();
  const cells=await page.evaluate(()=>[...new Set(window.__app.state.shots.flatMap(s=>s.result.cells.map(c=>`${c.no}:${c.col}`)))].sort());
  const expected=Object.entries(mock.UNCLOSED).flatMap(([no,ps])=>ps.map(p=>`${no}:${p}교시`)).sort();
  assert.deepEqual(cells,expected);
  const unique=await page.evaluate(()=>{const rows=window.__app.getRows(),seen=new Set();return rows.every(r=>Clubs.teachers(r.teacher).every(t=>{const k=[r.date,r.periodText,t].join('|');if(seen.has(k))return false;seen.add(k);return true;}));});
  assert.ok(unique);console.log('PASS OCR fixture recognition and one row per teacher/date/period');
  const [download]=await Promise.all([page.waitForEvent('download'),page.click('#btnNew')]);
  const file=path.join(root,'tests/.out-weekly.xlsx');await download.saveAs(file);
  const bytes=[...fs.readFileSync(file)];
  const read=async file=>page.evaluate(async bytes=>{const wb=new ExcelJS.Workbook();await wb.xlsx.load(new Uint8Array(bytes));return wb.worksheets[0].getSheetValues().filter(r=>r&&r[2]);},[...fs.readFileSync(file)]);
  const before=await read(file);
  const [appended]=await Promise.all([page.waitForEvent('download'),page.setInputFiles('#appendFile',file)]);
  const file2=path.join(root,'tests/.out-weekly-append.xlsx');await appended.saveAs(file2);
  assert.equal((await read(file2)).length,before.length);
  fs.rmSync(file);fs.rmSync(file2);
  console.log('PASS generated Excel and duplicate-safe append');
  // PDF reader still handles the rotated fictional weekly fixture.
  await page.click('#btnClear');
  await page.setInputFiles('#ttFile',path.join(root,'tests/fixtures/fake-timetable.pdf'));
  await page.waitForFunction(()=>!!window.__app.state.tt);
  assert.equal(await page.evaluate(()=>window.__app.state.tt.students.size),28);
  assert.equal(await page.evaluate(()=>Timetable.lookup(window.__app.state.tt,3,'2026-09-23',1).teacher),'한서준');
  await page.click('#btnClear');
  assert.ok(await page.evaluate(()=>!window.__app.state.tt&&!window.__app.state.clubs&&!window.__app.state.timetableFiles.length));
  await page.setViewportSize({width:390,height:900});assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=390));
  assert.deepEqual(errors,[]);assert.deepEqual(external,[]);
  console.log('PASS rotated PDF, clear data, mobile layout and no external requests');
} finally {if(browser)await browser.close();server.close();}
