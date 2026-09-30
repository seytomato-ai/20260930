// tests/fixtures/fake-timetable.pdf 만들기:  node tools/make-pdf-fixture.mjs
import { chromium } from 'playwright-core';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium' });
const p = await b.newPage();
await p.goto('file://' + path.join(root, 'tools/mock-timetable-pdf.html'));
await p.pdf({ path: path.join(root, 'tests/fixtures/fake-timetable.pdf'), preferCSSPageSize: true, printBackground: true });
await b.close();
