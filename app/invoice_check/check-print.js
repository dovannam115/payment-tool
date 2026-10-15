/*
 * BATCH INVOICE CHECK & PRINT  —  kiemtrahoadon.vn (M-invoice)
 * -----------------------------------------------------------------
 * Pick printer -> pick folder of XML -> for each file:
 *   upload to kiemtrahoadon.vn, wait for "Valid", open the CHECK RESULT page, print.
 * Runs files in parallel for speed (default 5).
 *
 * Usage:
 *   node check-print.js                              (ask printer + folder)
 *   node check-print.js "D:\\Invoices"               (folder given)
 *   node check-print.js "D:\\Invoices" --pdf         (print to PDF, no prompt)
 *   node check-print.js "D:\\Invoices" --jobs 8      (8 files in parallel)
 *   node check-print.js "D:\\Invoices" --printer "HPB3812A.lan (HP LaserJet Pro MFP 4103)"
 *   node check-print.js "D:\\Invoices" --simplex     (1-sided; default is 2-sided)
 */
const { chromium } = require('playwright-core'); // core = no browser download; uses system Edge
const { execFileSync } = require('child_process');
const { PDFDocument } = require('pdf-lib');
const readline = require('readline');
const path = require('path');
const os = require('os');
const fs = require('fs');

const SITE = 'https://kiemtrahoadon.vn/';
const PDF_PRINTER = 'Microsoft Print to PDF';

// ---------- PowerShell helpers ----------
function ps(cmd) {
  return execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', cmd], { encoding: 'utf8', windowsHide: true }).trim();
}
const listPrinters = () => ps('Get-Printer | Select-Object -ExpandProperty Name').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
const getDefaultPrinter = () => ps("(Get-CimInstance Win32_Printer | Where-Object {$_.Default}).Name");
const setDefaultPrinter = (name) => ps(`(New-Object -ComObject WScript.Network).SetDefaultPrinter('${name.replace(/'/g, "''")}')`);
// folder picker via wscript.exe (GUI host = NO black console window ever); path returned through a temp file
function pickFolderDialog() {
  try {
    const outFile = path.join(os.tmpdir(), `ic_folder_${process.pid}.txt`);
    const vbs = path.join(os.tmpdir(), `ic_pick_${process.pid}.vbs`);
    fs.writeFileSync(vbs,
      'Set fso=CreateObject("Scripting.FileSystemObject")\r\n' +
      'Set sh=CreateObject("Shell.Application")\r\n' +
      'Set f=sh.BrowseForFolder(0,"Select the folder with invoice XML files",&H10&,"")\r\n' +
      'p=""\r\n' +
      'If Not (f Is Nothing) Then p=f.Self.Path\r\n' +
      // 3rd arg True = write the file as Unicode (UTF-16), so Vietnamese paths survive
      'Set ts=fso.CreateTextFile("' + outFile.replace(/\\/g, '\\\\') + '",True,True)\r\n' +
      'ts.Write p\r\nts.Close\r\n');
    execFileSync('wscript.exe', ['//nologo', vbs]); // wscript = GUI host (no console); no windowsHide or the dialog gets hidden
    const p = fs.existsSync(outFile) ? fs.readFileSync(outFile, 'utf16le').replace(/^﻿/, '').trim() : '';
    try { fs.unlinkSync(vbs); fs.unlinkSync(outFile); } catch {}
    return p || null;
  } catch { return null; }
}
function ask(q) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(r => rl.question(q, a => { rl.close(); r(a.trim()); }));
}

// ---------- parallel pool, keeps result order ----------
async function runPool(items, worker, concurrency) {
  const results = new Array(items.length);
  let idx = 0;
  const run = async () => {
    while (idx < items.length) {
      const i = idx++;
      results[i] = await worker(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, run));
  return results;
}

// in-place progress bar (one line, redrawn). label is 6 chars for alignment.
function drawBar(label, done, total) {
  const w = 24;
  const filled = total ? Math.round(w * done / total) : w;
  const pct = total ? Math.round(100 * done / total) : 100;
  const bar = '\x1b[36m' + '█'.repeat(filled) + '\x1b[90m' + '░'.repeat(w - filled) + '\x1b[0m';
  const cnt = `${String(done).padStart(String(total).length)}/${total}`;
  process.stdout.write(`\r  ${label} ${bar} ${cnt} ${String(pct).padStart(3)}%`);
  if (done >= total) process.stdout.write('\n');
}

function parseArgs() {
  const a = process.argv.slice(2), o = { folder: null, printer: null, pdf: false, jobs: 5, jobsSet: false, simplex: false };
  for (let i = 0; i < a.length; i++) {
    if (a[i] === '--pdf') o.pdf = true;
    else if (a[i] === '--simplex') o.simplex = true;           // 1-sided (default is 2-sided)
    else if (a[i] === '--printer') o.printer = a[++i];
    else if (a[i] === '--jobs') { o.jobs = parseInt(a[++i], 10) || 5; o.jobsSet = true; }
    else if (!a[i].startsWith('--')) o.folder = a[i];
  }
  return o;
}

function choosePrinter(opt) {
  // Default = export PDF. Direct printing only when --printer "name" is given.
  if (opt.printer && !/print to pdf/i.test(opt.printer)) return { name: opt.printer, mode: 'physical' };
  return { name: PDF_PRINTER, mode: 'pdf' };
}

// merge the per-invoice PDFs (in file order) into one _All.pdf for one-shot printing.
// Each invoice is normally 1 page. If one still spills to an odd count > 1, pad it
// with a blank page so, when printed 2-sided, the next invoice starts on a fresh sheet.
async function mergePdfs(outFiles, outPath) {
  const merged = await PDFDocument.create();
  for (const f of outFiles) {
    const src = await PDFDocument.load(fs.readFileSync(f));
    const pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach(p => merged.addPage(p));
    if (pages.length > 1 && pages.length % 2 === 1) { // odd multi-page -> add a blank page
      const last = pages[pages.length - 1];
      merged.addPage([last.getWidth(), last.getHeight()]);
    }
  }
  fs.writeFileSync(outPath, await merged.save());
}

// seller (NBan) name + tax ID straight from the XML; the 2nd <NBan> is the signature block
function sellerFromXml(file) {
  try {
    const x = fs.readFileSync(file, 'utf8');
    for (const m of x.matchAll(/<NBan>([\s\S]*?)<\/NBan>/g)) {
      const ten = m[1].match(/<Ten>([^<]*)<\/Ten>/), mst = m[1].match(/<MST>([^<]*)<\/MST>/);
      if (ten || mst) return { name: ten ? ten[1].trim() : '', mst: mst ? mst[1].trim() : '' };
    }
  } catch {}
  return { name: '', mst: '' };
}

// "Trạng thái" of the seller -> Active / Inactive / Unknown.
// Active = "NNT đang hoạt động (...)" or "NNT đã chuyển cơ quan thuế quản lý"; every other status
// (closed tax ID, stopped, suspended, not at registered address, pending bankruptcy) = Inactive.
function companyState(status) {
  if (!status) return 'Unknown';
  return /đang hoạt động|chuyển cơ quan thuế/i.test(status) ? 'Active' : 'Inactive';
}

// ---------- upload 1 file -> get ID + status (no print) ----------
async function fetchResult(ctx, file) {
  const page = await ctx.newPage();
  const s = sellerFromXml(file);
  const rec = { file: path.basename(file), path: file, id: null, status: '?', note: '',
                seller: s.name, sellerMst: s.mst, taxStatus: '', company: 'Unknown' };
  try {
    for (let attempt = 1; attempt <= 3; attempt++) { // retry: site can be briefly slow/throttled
      try {
        await page.goto(SITE, { waitUntil: 'domcontentloaded', timeout: 60000 });
        // upload box is a hidden input -> attached is enough, setInputFiles auto-waits
        await page.locator('input[type=file][accept=".xml"]').setInputFiles(file, { timeout: 30000 });
        await page.waitForSelector('iframe[src*="xml-info?id="]', { state: 'attached', timeout: 30000 });
        const src = await page.getAttribute('iframe[src*="xml-info?id="]', 'src');
        const m = src && src.match(/id=([0-9a-zA-Z]+)/);
        if (!m) throw new Error('no invoice id');
        rec.id = m[1];
        break;
      } catch (e) {
        if (attempt === 3) throw e;
        await page.waitForTimeout(1000 * attempt);
      }
    }
  } catch (e) {
    rec.status = 'ERROR'; rec.note = e.message.split('\n')[0];
  } finally {
    await page.close();
  }
  return rec;
}

// ---------- open check-result page, read status, clean overlays ----------
async function openResultPage(ctx, id) {
  const rp = await ctx.newPage();
  try { return await readResultPage(rp, id); }
  catch (e) { await rp.close().catch(() => {}); throw e; }
}

async function readResultPage(rp, id) {
  const url = `${SITE}#/xml-info?id=${id}&excludeTawk=Y`;
  await rp.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  // load up to 3x: if the tax-code (MST) section says "couldn't fetch" (GDT throttled), retry
  for (let attempt = 1; attempt <= 3; attempt++) {
    await rp.waitForFunction(() => /KẾT QUẢ KIỂM TRA/i.test(document.body.innerText), { timeout: 30000 });
    await rp.evaluate(() => document.fonts.ready).catch(() => {}); // ensure icon font (check marks)
    await rp.waitForFunction(() => /(đang hoạt động|ngừng hoạt động)/i.test(document.body.innerText),
      { timeout: 12000 }).catch(() => {});
    const mstFail = await rp.evaluate(() => /Không lấy được thông tin mã số thuế/i.test(document.body.innerText));
    if (!mstFail || attempt === 3) break;
    await rp.waitForTimeout(1200 * attempt);
    await rp.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
  }
  await rp.waitForTimeout(300);
  await hideWidgets(rp);
  // force background colors to print (green "Valid" badge, header fills) even on a real printer
  await rp.addStyleTag({ content: '@media print{html,body,*{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important;}}' }).catch(() => {});
  const { status, taxStatus } = await rp.evaluate(() => {
    const t = document.body.innerText;
    // "TÌNH TRẠNG HOẠT ĐỘNG CỦA MST" -> "Người bán:" block -> "Trạng thái: ..."
    const sec = t.split(/TÌNH TRẠNG HOẠT ĐỘNG CỦA MST/i)[1] || '';
    const seller = (sec.split(/Người mua:/i)[0].split(/Người bán:/i)[1]) || '';
    const m = seller.match(/Trạng thái:[ \t]*([^\n]*)/i);   // blank on the site -> '' -> Unknown, never the next line
    return {
      status: /Không hợp lệ/i.test(t) ? 'INVALID' : /Hợp lệ/i.test(t) ? 'VALID' : '?',
      taxStatus: m ? m[1].trim() : '',
    };
  });
  return { rp, status, taxStatus };
}

function setResult(r, res) {
  r.status = res.status;
  r.taxStatus = res.taxStatus;
  r.company = companyState(res.taxStatus);
}

// The support chat (Subiz "M-invoice Support", bottom right) draws itself again after being
// removed, so hide it with a style rule as well and run this once more right before printing.
const WIDGETS = '[sbz-plugin-type],.widget-layout,.widget-container,.chat-button--white-bg,' +
  'iframe[src*="subiz"],iframe[src*="tawk"],[id*="tawk"],.tawk-min-container';
async function hideWidgets(rp) {
  await rp.addStyleTag({ content: WIDGETS + '{display:none !important;visibility:hidden !important}' }).catch(() => {});
  await rp.evaluate((sel) => document.querySelectorAll(sel).forEach(e => e.remove()), WIDGETS).catch(() => {});
}

async function savePdf(ctx, rp, outFile) {
  await hideWidgets(rp);          // last chance: the widget may have come back while we waited
  const client = await ctx.newCDPSession(rp);
  // ALWAYS one A4 page: shrink the scale until the PDF has a single page
  // (0.05 steps down to 0.5, then 0.1 steps down to 0.1 = Chrome's minimum)
  let buf;
  for (let scale = 0.9; ; scale = Math.round((scale - (scale > 0.5 ? 0.05 : 0.1)) * 100) / 100) {
    const { data } = await client.send('Page.printToPDF', {
      printBackground: true, scale,
      marginTop: 0.3, marginBottom: 0.3, marginLeft: 0.3, marginRight: 0.3,
      paperWidth: 8.27, paperHeight: 11.69,
    });
    buf = Buffer.from(data, 'base64');
    const pages = (await PDFDocument.load(buf)).getPageCount();
    if (pages <= 1 || scale <= 0.1) break;
  }
  fs.writeFileSync(outFile, buf);
}

(async () => {
  const opt = parseArgs();
  const printer = choosePrinter(opt);
  console.log(`> Mode: ${printer.mode === 'pdf' ? 'export PDF' : 'print -> ' + printer.name}`);

  // 1) pick folder first
  let folder = opt.folder;
  if (!folder) folder = pickFolderDialog() || await ask('XML folder path: ');
  folder = folder.replace(/^"|"$/g, '');
  if (!fs.existsSync(folder)) { console.error('Folder not found:', folder); process.exit(1); }
  const files = fs.readdirSync(folder).filter(f => f.toLowerCase().endsWith('.xml')).map(f => path.join(folder, f));
  if (!files.length) { console.error('No .xml files here.'); process.exit(1); }
  console.log(`> Folder: ${folder} (${files.length} XML)`);

  // 2) then ask how many to run in parallel
  if (!opt.jobsSet && process.stdin.isTTY) {
    const n = parseInt(await ask('Files in parallel? (5-10 recommended, Enter=5): '), 10);
    if (n >= 1) opt.jobs = n;
  }
  const MAXJOBS = 15;
  if (opt.jobs > MAXJOBS) { console.log(`  ${opt.jobs} too high - capped at ${MAXJOBS}.`); opt.jobs = MAXJOBS; }
  opt.jobs = Math.max(1, Math.min(opt.jobs, files.length));
  console.log(`> Parallel: ${opt.jobs}/batch`);

  const outDir = path.join(folder, 'CheckResults');
  if (printer.mode === 'pdf' && !fs.existsSync(outDir)) fs.mkdirSync(outDir);

  let prevDefault = null, prevDuplex = null;
  const pn = printer.name.replace(/'/g, "''");
  if (printer.mode === 'physical') {
    prevDefault = getDefaultPrinter(); setDefaultPrinter(printer.name);
    if (!opt.simplex) { // enable 2-sided (each invoice = 2 pages = 1 sheet)
      try {
        prevDuplex = ps(`(Get-PrintConfiguration -PrinterName '${pn}').DuplexingMode`);
        ps(`Set-PrintConfiguration -PrinterName '${pn}' -DuplexingMode TwoSidedLongEdge`);
        console.log('> Duplex (2-sided): ON');
      } catch { console.log('  Duplex not available - printing 1-sided.'); prevDuplex = null; }
    }
  }

  const launchOpts = printer.mode === 'physical'
    ? { channel: 'msedge', headless: false, args: ['--kiosk-printing'] }
    : { channel: 'msedge', headless: true };
  const browser = await chromium.launch(launchOpts);
  const ctx = await browser.newContext({ locale: 'vi-VN' });

  const t0 = Date.now();
  // PHASE 1: upload + get ID (parallel)
  let d1 = 0; drawBar('Check ', 0, files.length);
  const results = await runPool(files, async (file) => {
    const r = await fetchResult(ctx, file);
    drawBar('Check ', ++d1, files.length);
    return r;
  }, opt.jobs);

  // PHASE 2: print
  const todo = results.filter(r => r.id).length;
  if (printer.mode === 'pdf') {
    let d2 = 0; if (todo) drawBar('Export', 0, todo);
    await runPool(results, async (r) => {
      if (!r.id) return;
      try {
        const res = await openResultPage(ctx, r.id).catch(() => openResultPage(ctx, r.id)), rp = res.rp;
        setResult(r, res);
        const out = path.join(outDir, `${path.basename(r.file, path.extname(r.file))}.pdf`);
        await savePdf(ctx, rp, out); r.out = out;
        await rp.close();
      } catch (e) { r.status = 'ERROR'; r.note = e.message.split('\n')[0]; }
      drawBar('Export', ++d2, todo);
    }, Math.min(opt.jobs, 3)); // tax lookup: max 3 parallel to avoid throttling
    // merge all PDFs (in order) into one file for one-shot printing later
    const outs = results.filter(r => r.out && fs.existsSync(r.out)).map(r => r.out);
    if (outs.length > 1) {
      try { await mergePdfs(outs, path.join(outDir, '_All.pdf')); console.log('  Merged -> _All.pdf'); }
      catch (e) { console.log('  merge skipped:', e.message.split('\n')[0]); }
    }
  } else {
    let d2 = 0; if (todo) drawBar('Print ', 0, todo);
    for (const r of results) { // paper: sequential, in file order
      if (!r.id) continue;
      try {
        const res = await openResultPage(ctx, r.id), rp = res.rp;
        setResult(r, res);
        await rp.evaluate(() => window.print());
        await rp.waitForTimeout(2500);
        await rp.close();
        r.out = 'printed';
      } catch (e) { r.status = 'ERROR'; r.note = e.message.split('\n')[0]; }
      drawBar('Print ', ++d2, todo);
    }
  }

  await browser.close();
  if (prevDefault) setDefaultPrinter(prevDefault);
  if (prevDuplex) { try { ps(`Set-PrintConfiguration -PrinterName '${pn}' -DuplexingMode ${prevDuplex}`); } catch {} }

  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const ok = results.filter(r => r.status === 'VALID').length;
  console.log(`\n= DONE ${secs}s | Valid ${ok}/${results.length} =`);
  results.filter(r => r.status !== 'VALID').forEach(r => console.log(`  ! ${r.file}: ${r.status} ${r.note}`));
  const act = results.filter(r => r.company === 'Active').length;
  console.log(`= Seller company active ${act}/${results.length} =`);
  results.filter(r => r.company !== 'Active').forEach(r =>
    console.log(`  ! ${r.file}: ${r.company} - ${r.seller || '?'} (${r.sellerMst || '?'})${r.taxStatus ? ' - ' + r.taxStatus : ''}`));
  const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csv = '﻿File,Status,Seller,Seller tax ID,Company,Tax status,Output,Note\n' +
    results.map(r => [r.file, r.status, r.seller, r.sellerMst, r.company, r.taxStatus, r.out || '', r.note].map(q).join(','))
      .join('\n');
  fs.writeFileSync(path.join(printer.mode === 'pdf' ? outDir : folder, 'log.csv'), csv);
  if (printer.mode === 'pdf') console.log(`PDF + log: ${outDir}`);
})().catch(e => { console.error('ERROR:', e); process.exit(1); });
