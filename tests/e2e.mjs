// 끝-끝 테스트: 가상 시간표 + 가상 나이스 캡처 → 결과 행·엑셀 확인, 외부 통신이 없는지 확인
//   npm test   (Chromium 경로는 CHROMIUM_PATH 환경변수로 바꿀 수 있음)
import { chromium } from 'playwright-core';
import ExcelJS from 'exceljs';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fx = f => path.join(root, 'tests/fixtures', f);
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml', '.gz': 'application/gzip' };
const server = http.createServer((req, res) => {
  let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (p.endsWith('/')) p += 'index.html';
  const f = path.join(root, p);
  if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': TYPES[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
}).listen(0);
const base = `http://127.0.0.1:${server.address().port}/`;

let failed = 0;
const ok = (cond, msg) => { console.log((cond ? '  ✔ ' : '  ✘ ') + msg); if (!cond) failed++; };

const browser = await chromium.launch({ args: ['--disable-background-networking', '--disable-component-update'], executablePath: process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium' });
const ctx = await browser.newContext({ acceptDownloads: true });
const page = await ctx.newPage();
const external = [], errors = [];
page.on('request', r => { if (!r.url().startsWith(base) && !/^(blob|data):/.test(r.url())) external.push(r.url()); });
page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
page.on('pageerror', e => errors.push(e.message));

await page.goto(base);
await page.setInputFiles('#ttFile', fx('fake-timetable.xlsx'));
await page.waitForSelector('#ttInfo:not([hidden]) b');
ok((await page.textContent('#ttInfo')).includes('학생 28명'), '시간표: 학생 28명 인식');

await page.setInputFiles('#imgFile', [fx('mock-part1.png'), fx('mock-part2.png')]);
await page.waitForFunction(() => window.__app.state.shots.length === 2 && window.__app.state.shots.every(s => s.status === 'done' || s.status === 'error'), null, { timeout: 120000 });

const mock = await (async () => {
  const p2 = await ctx.newPage();
  await p2.goto(base + 'tools/mock-neis.html');
  const m = await p2.evaluate(() => window.MOCK);
  await p2.close();
  return m;
})();
const expected = [];
for (const [no, ps] of Object.entries(mock.UNCLOSED)) for (const p of ps) expected.push(`${no}:${p}교시`);
const rows = await page.evaluate(() => window.__app.getRows());
const got = rows.map(r => `${r.no}:${r.periodText}`);
ok(JSON.stringify(got.slice().sort()) === JSON.stringify(expected.slice().sort()), `미마감 칸 ${got.length}/${expected.length}개 정확히 인식`);
if (got.length !== expected.length) console.log('    got', got.join(' '), '\n    exp', expected.join(' '));
ok(await page.inputValue('#classCode') === '2-3', '학년-반 자동 인식 (2-3)');
ok(rows.every(r => r.dateText === '2026.09.23'), '날짜 자동 인식 (2026.09.23)');
ok(rows.every(r => r.name === mock.NAMES[r.no - 1]), '이름은 시간표에서');
const r3 = rows.find(r => r.no === 3 && r.periodText === '1교시');
ok(r3 && r3.subject === '영어Ⅱ' && r3.teacher === '한서준', '공통 시간표에서 과목·교사 찾기 (영어Ⅱ/한서준)');
const r7 = rows.find(r => r.no === 7 && r.periodText === '7교시');
ok(r7 && r7.subject && r7.teacher && !r7.teacher.includes('\n'), `줄바꿈 칸 "과목↵교사" 읽기 (${r7 && r7.subject}/${r7 && r7.teacher})`);
ok(rows.every(r => r.subject && r.teacher), '모든 행에 과목·교사');

// 칸 토글
await page.click('table.grid td.cell.on');
ok((await page.evaluate(() => window.__app.getRows().length)) === expected.length - 1, '확인 표에서 칸을 눌러 끄기');
await page.click('table.grid td.cell:not(.on) >> nth=0');

// 새 엑셀
const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnNew')]);
const outPath = path.join(root, 'tests/.out-new.xlsx');
await dl.saveAs(outPath);
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(outPath);
const ws = wb.worksheets[0];
ok(ws.getRow(1).values.slice(1).join('|') === '학번/학급|이름|날짜|과목|교시|출결 종류|담당 교사|요청중/완료|전달 사항', '엑셀 머리글이 기존 양식과 같음');
let n1 = 0;
ws.eachRow((r, i) => { if (i > 1 && r.getCell(2).value) n1++; });
ok(n1 === expected.length, `엑셀 ${n1}행`);
ok(ws.getCell('H2').value === '요청중' && ws.getCell('I2').value === '교시 마감 부탁드립니다.', '상태·전달 사항 기본값');
ok(!!ws.getCell('E2').dataValidation && ws.getCell('E2').dataValidation.type === 'list', '교시 목록 유효성 검사');
ok(wb.worksheets.some(s => s.name === '교사별 요약'), '교사별 요약 시트');
console.log('    예:', ws.getRow(2).values.slice(1).join(' | '));

// 이어쓰기
const [dl2] = await Promise.all([page.waitForEvent('download'), page.setInputFiles('#appendFile', fx('fake-existing.xlsx'))]);
const appPath = path.join(root, 'tests/.out-append.xlsx');
await dl2.saveAs(appPath);
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.readFile(appPath);
const ws2 = wb2.worksheets[0];
let filled = 0;
ws2.eachRow((r, i) => { if (i > 1 && r.getCell(2).value) filled++; });
ok(filled === 2 + expected.length - 1, `이어쓰기: 기존 2행 + 새 ${filled - 2}행 (중복 1행 건너뜀)`);
ok(ws2.getCell('H40').dataValidation && ws2.getCell('H40').dataValidation.type === 'list', '이어쓰기 후에도 기존 유효성 검사 유지');
ok(/1줄은 건너뜀/.test(await page.textContent('#exportMsg')), '중복 안내 문구');

// 모두 지우기
await page.click('#btnClear');
ok(await page.evaluate(() => window.__app.state.shots.length === 0 && !window.__app.state.tt), '모두 지우기');

// 개인정보: 외부 통신 없음, 저장소 사용 없음
ok(external.length === 0, '외부 주소로 나간 요청 없음' + (external.length ? ': ' + external.join(', ') : ''));
const storage = await page.evaluate(async () => ({ ls: localStorage.length, ss: sessionStorage.length, cookie: document.cookie, idb: indexedDB.databases ? (await indexedDB.databases()).length : 0 }));
ok(storage.ls === 0 && storage.ss === 0 && !storage.cookie && storage.idb === 0, '쿠키·로컬저장소·IndexedDB 사용 없음');
ok(errors.length === 0, '콘솔 오류 없음' + (errors.length ? ': ' + errors.join(' / ') : ''));

// PDF 시간표 (나이스 학생별 시간표처럼 90° 회전된 표) → 엑셀 시간표와 같은 결과
{
  const p2 = await ctx.newPage();
  p2.on('request', r => { if (!r.url().startsWith(base) && !/^(blob|data):/.test(r.url())) external.push(r.url()); });
  p2.on('pageerror', e => errors.push(e.message));
  await p2.goto(base);
  await p2.setInputFiles('#ttFile', fx('fake-timetable.pdf'));
  await p2.waitForFunction(() => !document.querySelector('#ttInfo').hidden && !/읽는 중/.test(document.querySelector('#ttInfo').textContent), null, { timeout: 60000 });
  const info = await p2.textContent('#ttInfo');
  ok(info.includes('학생 28명') && info.includes('PDF'), 'PDF 시간표: 학생 28명 인식');
  ok(await p2.inputValue('#classCode') === '2-3', 'PDF에서 학년-반 채우기');
  await p2.setInputFiles('#imgFile', [fx('mock-part1.png'), fx('mock-part2.png')]);
  await p2.waitForFunction(() => window.__app.state.shots.length === 2 && window.__app.state.shots.every(s => s.status === 'done' || s.status === 'error'), null, { timeout: 120000 });
  const pdfRows = await p2.evaluate(() => window.__app.getRows());
  const key = r => `${r.no}|${r.name}|${r.periodText}|${r.subject}|${r.teacher}`;
  ok(JSON.stringify(pdfRows.map(key)) === JSON.stringify(rows.map(key)), `PDF 시간표로도 같은 결과 (${pdfRows.length}행)`);
  if (JSON.stringify(pdfRows.map(key)) !== JSON.stringify(rows.map(key))) console.log('    pdf', pdfRows.map(key).join(' '), '\n    xlsx', rows.map(key).join(' '));
  await p2.close();
  ok(errors.length === 0 && external.length === 0, 'PDF 처리 중에도 오류·외부 통신 없음' + (errors.length ? ': ' + errors.join(' / ') : ''));
}

await page.setViewportSize({ width: 390, height: 900 });
ok(await page.evaluate(() => document.documentElement.scrollWidth <= 390), '휴대폰 폭에서 가로 스크롤 없음');

fs.rmSync(outPath, { force: true });
fs.rmSync(appPath, { force: true });
await browser.close();
server.close();
console.log(failed ? `\n${failed}개 실패` : '\n모두 통과');
process.exit(failed ? 1 : 0);
