// End-to-end UI test: scanned PDF -> OCR prompt -> convert -> download.
// usage: npm run test:e2e            (server must be running on :8000)
//        node tests/e2e_ui.cjs [outDir] [baseUrl]
// env:   CHROME_PATH        use an installed Chrome instead of Playwright's Chromium
//        PLAYWRIGHT_MODULE  path to a playwright / playwright-core package
const path = require('path');
const fs = require('fs');
const os = require('os');
const [, , outDir = path.join(os.tmpdir(), 'docxify-e2e'), base = 'http://127.0.0.1:8000/'] = process.argv;
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const launchOptions = { headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) };
const FIX = path.join(__dirname, 'fixtures');
fs.mkdirSync(outDir, { recursive: true });
const assert = (cond, msg) => { if (!cond) throw new Error('ASSERT: ' + msg); console.log('  ok -', msg); };

(async () => {
  const browser = await chromium.launch(launchOptions);
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 820 }, acceptDownloads: true, deviceScaleFactor: 1 });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('requestfailed', (r) => { if (!/aborted/i.test(r.failure()?.errorText || '')) errors.push(`${r.failure()?.errorText} ${r.method()} ${new URL(r.url()).pathname}`); });

  await page.goto(base, { waitUntil: 'load' });
  await page.screenshot({ path: path.join(outDir, '1_select.png') });
  assert(await page.getByRole('heading', { name: 'PDF to Word' }).isVisible(), 'landing heading visible');

  // 1) scanned PDF: upload -> badge -> OCR prompt -> apply OCR -> done
  await page.setInputFiles('#file-input', path.join(FIX, 'scanned.pdf'));
  await page.waitForSelector('.file-card .badge:not([hidden])', { timeout: 30000 });
  assert((await page.textContent('.file-card .badge')).trim() === 'Scanned', 'scanned badge shown');
  assert(await page.isVisible('#scan-hint'), 'OCR recommendation hint shown');
  await page.waitForSelector('.file-card .thumb img', { timeout: 10000 });
  await page.screenshot({ path: path.join(outDir, '2_files.png') });

  await page.click('#convert');
  await page.waitForSelector('#ocr-modal:not([hidden])');
  assert(await page.getByRole('dialog', { name: 'This PDF looks scanned' }).isVisible(), 'scanned-PDF OCR prompt opened');
  await page.waitForTimeout(350); // let the dialog finish its entrance animation
  await page.screenshot({ path: path.join(outDir, '3_ocr_prompt.png') });
  await page.click('#ocr-yes');

  let sawProgress = false;
  for (let i = 0; i < 40 && !sawProgress; i++) {
    const txt = await page.textContent('#status-line').catch(() => '');
    if (/Recognizing text|Rebuilding layout|Writing Word|Analyzing/.test(txt || '')) {
      sawProgress = true;
      await page.screenshot({ path: path.join(outDir, '4_progress.png') });
    } else await page.waitForTimeout(100);
  }
  await page.waitForSelector('#view-done.is-active', { timeout: 120000 });
  assert(sawProgress, 'progress view showed a live stage');
  const summary = await page.textContent('#summary');
  assert(/2 pages recognized/.test(summary), 'summary reports OCR on 2 pages');
  await page.screenshot({ path: path.join(outDir, '5_done.png') });
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
  const docx = path.join(outDir, 'ui_scanned.docx');
  await dl.saveAs(docx);
  assert(dl.suggestedFilename() === 'scanned.docx', `download named ${dl.suggestedFilename()}`);
  assert(fs.statSync(docx).size > 5000, 'downloaded DOCX is non-trivial');

  // 2) two files -> zip, digital file shows "Text" badge, no OCR prompt
  await page.click('#again');
  await page.setInputFiles('#file-input', [path.join(FIX, 'digital.pdf'), path.join(FIX, 'digital.pdf')]);
  await page.waitForFunction(() => document.querySelectorAll('.file-card .badge:not([hidden])').length === 2, null, { timeout: 30000 });
  assert((await page.textContent('.file-card .badge')).trim() === 'Text', 'digital PDF badge is "Text"');
  await page.click('#convert');
  await page.waitForSelector('#view-done.is-active', { timeout: 120000 });
  assert(await page.isHidden('#ocr-modal'), 'no OCR prompt for text PDFs');
  assert((await page.textContent('#download-label')).includes('ZIP'), 'two files -> ZIP download');
  await page.click('#delete-now');
  await page.waitForSelector('#view-select.is-active');
  assert(true, 'delete-now returned to start');

  // 3) Arabic scan: OCR + Arabic chosen up front -> the job carries ocr_lang=ar
  let processBody = '';
  page.on('request', (r) => { if (r.url().endsWith('/v1/process')) processBody = r.postData() || ''; });
  await page.setInputFiles('#file-input', path.join(FIX, 'arabic_scanned.pdf'));
  await page.waitForSelector('.file-card .badge:not([hidden])', { timeout: 30000 });
  await page.check('input[name=ocr][value="1"]', { force: true });
  await page.selectOption('#lang', 'ar');
  assert((await page.textContent('#lang option[value="ar"]')).includes('العربية'), 'Arabic offered as a text language');
  await page.waitForTimeout(250); // option cards fade to their checked state
  await page.screenshot({ path: path.join(outDir, '6_arabic_options.png') });
  await page.click('#convert');
  await page.waitForSelector('#view-done.is-active', { timeout: 120000 });
  assert(/name="ocr_lang"\r?\n\r?\nar\r?\n/.test(processBody), 'Arabic OCR language sent with the job');
  assert(/1 pages? recognized/.test(await page.textContent('#summary')), 'summary reports OCR on the Arabic page');
  const [dlAr] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
  const docxAr = path.join(outDir, 'ui_arabic.docx');
  await dlAr.saveAs(docxAr);
  assert(fs.statSync(docxAr).size > 5000, 'Arabic DOCX downloaded');
  await page.click('#delete-now');
  await page.waitForSelector('#view-select.is-active');

  // 3) mobile layout
  const m = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true });
  const mp = await m.newPage();
  await mp.goto(base, { waitUntil: 'load' });
  await mp.waitForTimeout(400);
  await mp.screenshot({ path: path.join(outDir, '6_mobile_select.png') });
  await mp.setInputFiles('#file-input', path.join(FIX, 'scanned.pdf'));
  await mp.waitForSelector('.file-card .badge:not([hidden])', { timeout: 30000 });
  await mp.waitForTimeout(300);
  await mp.screenshot({ path: path.join(outDir, '7_mobile_files.png'), fullPage: true });
  const overflow = await mp.evaluate(() => document.documentElement.scrollWidth > window.innerWidth);
  assert(!overflow, 'no horizontal overflow at 390px');

  assert(errors.length === 0, `no page/console errors${errors.length ? ': ' + errors.join(' | ') : ''}`);
  await browser.close();
  console.log('E2E PASSED');
})().catch((e) => { console.error(e.message); process.exit(1); });
