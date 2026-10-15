// Tool B - dien don de nghi thanh toan tren thanhtoannew.pvi.com.vn
// Form Ant Design. Dieu khien GIAO DIEN (dien truong + import XML + dinh kem) nhu nguoi dung.
// KHONG dung API. DUNG truoc khi bam "Luu" de ban tu kiem tra va luu.
// Chay sau khi da dang nhap 1 lan (phien luu trong .edge-profile).
const { chromium } = require("playwright-core");
const fs = require("fs");
const path = require("path");
const readline = require("readline");

const URL = "https://thanhtoannew.pvi.com.vn/#/payment-requests/create?loai_ttoan=01";
const ROOT = path.join(__dirname, "..");
const PROFILE = path.join(ROOT, ".edge-profile");
const EDGE = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const DUMP = path.join(__dirname, "dump");
const DON = path.join(DUMP, "don.json");

// TB_WEB=1: started from the web UI. A question is printed as a [[ASK]] marker line, the UI
// turns it into a button, and the answer arrives as one line on stdin.
const WEB = !!process.env.TB_WEB;
function ask(q){
  if(WEB){ console.log("[[ASK]] "+String(q).trim()); q=""; }
  const rl=readline.createInterface({input:process.stdin,output:process.stdout});
  // pause stdin after the answer: a readline left attached keeps node alive forever
  return new Promise(r=>rl.question(q,a=>{rl.close();try{process.stdin.pause();}catch(e){};r(a);}));}

async function setDate(page, value){
  if(!value) return "skip";
  try{
    const m=String(value).match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    const title = m ? `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}` : null;
    await page.click("#ngay_ctu");
    await page.waitForSelector(".ant-picker-dropdown .ant-picker-cell",{timeout:4000}).catch(()=>{});
    if(title){
      const tryClick=async()=>{ const c=page.locator(`.ant-picker-dropdown .ant-picker-cell[title="${title}"]`);
        if(await c.count()){ await c.first().click(); return true; } return false; };
      if(await tryClick()) return "ok";
      for(let k=0;k<6;k++){ await page.locator(".ant-picker-dropdown .ant-picker-header-prev-btn").last().click().catch(()=>{}); await page.waitForTimeout(180); if(await tryClick()) return "ok (prev)"; }
      await page.click("#ngay_ctu").catch(()=>{}); await page.waitForTimeout(200);
      for(let k=0;k<6;k++){ await page.locator(".ant-picker-dropdown .ant-picker-header-next-btn").last().click().catch(()=>{}); await page.waitForTimeout(180); if(await tryClick()) return "ok (next)"; }
    }
    await page.keyboard.press("Escape").catch(()=>{});
    return "manual (pick date on form)";
  }catch(e){ return "FAIL:"+e.message; }
}

async function setText(page, id, value){
  if(value===undefined||value===null||value==="") return "skip";
  try{ await page.fill("#"+id, String(value), {timeout:5000}); return "ok"; }
  catch(e){ return "FAIL:"+e.message; }
}

async function setSelect(page, id, text){
  if(!text) return "skip";
  try{
    const input = page.locator("#"+id);
    const wrap = input.locator('xpath=ancestor::div[contains(@class,"ant-select")][1]');
    // skip if the select already shows the wanted value (short timeout: empty select has no item)
    const cur = (await wrap.locator('.ant-select-selection-item').first().innerText({timeout:400}).catch(()=>"")).trim();
    if(cur && cur.toLowerCase()===String(text).trim().toLowerCase()) return "skip (da co)";
    await wrap.click();
    // searchable selects load options only AFTER typing -> type first. Readonly -> options already there.
    const ro = await input.evaluate(el => el.readOnly).catch(()=>true);
    if(!ro) await input.fill(String(text)).catch(()=>{});
    // click the matching VISIBLE option as soon as it is actionable (auto-wait, no fixed delay)
    const opt = page.locator('.ant-select-item-option:visible', { hasText: String(text) });
    try{ await opt.first().click({timeout:3000}); return "ok"; }
    catch(e){
      const any = page.locator('.ant-select-item-option:visible');
      if(await any.count()){ await any.first().click().catch(()=>{}); return "ok(first)"; }
      await page.keyboard.press("Escape").catch(()=>{});
      return "FAIL: option not found '"+text+"'";
    }
  }catch(e){ return "FAIL:"+e.message; }
}

async function soHdColIndex(page){
  return await page.evaluate(()=>{
    const ths=[...document.querySelectorAll(".ant-table-thead th")].map(h=>h.innerText.trim());
    return ths.findIndex(t=>/S[ốo]\s*h[óo]a\s*đ[ơo]n/i.test(t));
  });
}
const normNum=s=>String(s||"").replace(/\D/g,"").replace(/^0+/,"");
async function hasRowFor(page, col, so){
  if(col<0 || !so) return false;
  return await page.evaluate(({c,nz})=>{
    return [...document.querySelectorAll(".ant-table-tbody tr.ant-table-row")].some(r=>{
      const td=r.children[c]; if(!td) return false;
      return (td.innerText||"").replace(/\D/g,"").replace(/^0+/,"")===nz;
    });
  }, {c:col, nz:normNum(so)});
}
// Close every portal notification and wait until they are gone, so whatever shows up next is NEW.
// (Ant re-mounts the notification stack when a new one arrives, so "seen" marks on elements are unreliable.)
async function clearMessages(page){
  for(let k=0;k<6;k++){
    const n=await page.evaluate(()=>{
      document.querySelectorAll(".ant-notification-notice-close,.ant-message-notice .anticon-close").forEach(e=>e.click());
      document.querySelectorAll(".ant-modal-confirm .ant-btn-primary").forEach(b=>b.click());
      return document.querySelectorAll(".ant-notification-notice,.ant-message-notice").length;
    }).catch(()=>0);
    if(!n) return;
    await page.waitForTimeout(250);
  }
}
async function visibleMessages(page){
  return await page.evaluate(()=>[...document.querySelectorAll(".ant-notification-notice,.ant-message-notice,.ant-modal-confirm")]
    .map(e=>({t:(e.innerText||"").trim().replace(/\s+/g," "),
              err:/notice-error|_notify-error|message-error/.test(String(e.className)) || !!e.querySelector(".ant-message-error")}))
    .filter(m=>m.t)).catch(()=>[]);
}
// Seller missing in the portal's vendor list etc. -> the portal will not accept this XML.
const REJECT=/(chưa|không) tồn tại|không tìm thấy|chưa (được )?(khai báo|đăng ký)|không có trong danh (mục|sách)|thất bại/i;
async function hasRow(page, so){
  const col=await soHdColIndex(page);          // table header may only render after the first row -> look up every time
  return hasRowFor(page,col,so);
}
// Import one XML. Returns "ok" | "rejected:<portal message>" | "none" (portal swallowed it, no message).
async function importOne(page, iv){
  if(await hasRow(page,iv.so_hd)) return "ok";
  for(let attempt=1; attempt<=3; attempt++){     // retry ONLY when the portal said nothing (swallowed upload)
    await clearMessages(page);
    const before=await page.locator('.ant-table-tbody tr.ant-table-row').count();
    try{
      const [fc]=await Promise.all([
        page.waitForEvent("filechooser",{timeout:8000}),
        page.locator('button.ant-btn').filter({hasText:/Import File XML/i}).first().click(),
      ]);
      await fc.setFiles(iv.xml);
    }catch(e){ await page.waitForTimeout(600); continue; }
    for(let t=0;t<18;t++){                        // ~9s: row appears, or the portal answers with an error
      await page.waitForTimeout(500);
      if(iv.so_hd ? await hasRow(page,iv.so_hd) : (await page.locator('.ant-table-tbody tr.ant-table-row').count())>before) return "ok";
      const rej=(await visibleMessages(page)).find(m=>!/tên người mua/i.test(m.t) && (REJECT.test(m.t) || m.err));
      if(rej) return "rejected:"+rej.t.replace(/^Thất bại\s*/i,"").split(/[,.]/)[0].trim().slice(0,120);
    }
    await page.keyboard.press("Escape").catch(()=>{});
  }
  return "none";
}
// Import every invoice. A rejected one (e.g. "Nha cung cap chua ton tai") is tried ONCE more at the end
// (the portal sometimes rejects the very first import on a fresh form), then skipped and listed on screen.
async function importXml(page, invoices){
  if(!invoices||!invoices.length) return "skip";
  let ok=0; const again=[];
  for(const iv of invoices){
    const r=await importOne(page, iv);
    if(r==="ok") ok++;
    else if(r.startsWith("rejected:")) again.push([iv, r.slice(9)]);
    else { iv.skipped="portal did not take the XML (no row after 3 tries)"; }
  }
  for(const [iv, msg] of again){
    const r=await importOne(page, iv);
    if(r==="ok") ok++;
    else iv.skipped=r.startsWith("rejected:") ? r.slice(9) : msg;
  }
  await clearMessages(page);
  const skipped=invoices.filter(iv=>iv.skipped);
  skipped.forEach(iv=>console.log(`  !! SKIP invoice ${iv.ky_hieu||""} no.${iv.so_hd} (${path.basename(iv.xml)}): ${iv.skipped}`));
  const rows=await page.locator('.ant-table-tbody tr.ant-table-row').count();
  return `${rows} row(s), ${ok}/${invoices.length} invoice(s)`+(skipped.length?`, ${skipped.length} SKIPPED (see !! above)`:"");
}

async function attach(page, files){
  if(!files||!files.length) return "skip";
  // main form "FILE DINH KEM" -> to trinh + check PDF. Use the button + file chooser.
  try{
    const [fc]=await Promise.all([
      page.waitForEvent("filechooser",{timeout:6000}),
      page.locator('button.ant-btn').filter({hasText:/Chọn file/i}).first().click(),
    ]);
    await fc.setFiles(files);
    await page.waitForTimeout(800);
    return "ok ("+files.length+" file)";
  }catch(e){
    try{ await page.locator("#file_attachs").first().setInputFiles(files); return "ok (fallback, "+files.length+" file)"; }
    catch(e2){ return "FAIL:"+e2.message; }
  }
}

// dump the invoice table (headers + every row's cells) so we can see tax-split rows / missing invoices
async function dumpTable(page){
  try{
    const data=await page.evaluate(()=>{
      const heads=[...document.querySelectorAll(".ant-table-thead th")].map(h=>h.innerText.trim());
      const rows=[...document.querySelectorAll(".ant-table-tbody tr.ant-table-row")].map(r=>
        [...r.children].map(td=>(td.innerText||"").trim().replace(/\s+/g," ")));
      return {heads, rows};
    });
    const dir=DUMP; try{fs.mkdirSync(dir,{recursive:true});}catch(_){}
    fs.writeFileSync(path.join(dir,"table.json"), JSON.stringify(data,null,2),"utf8");
    return `${data.rows.length} row(s) -> dump/table.json`;
  }catch(e){ return "dump FAIL:"+e.message; }
}
// close any warning toast/notification that may overlay the buttons
async function dismissToasts(page){
  await page.evaluate(()=>{
    document.querySelectorAll(".ant-message-notice .anticon-close, .ant-notification-notice-close").forEach(e=>e.click && e.click());
    document.querySelectorAll(".ant-message-notice,.ant-notification-notice").forEach(e=>{ try{e.style.pointerEvents="none";e.style.display="none";}catch(_){} });
  }).catch(()=>{});
}
// Per-row "Sua" in the invoice list: for EVERY row (tax-split rows too) set TEN HANG HOA (#ten_hhoa) and
// upload that invoice's PDF (row matched by invoice number + serial, not by position), then "Xac nhan".
async function editRows(page, invoices, itemName){
  await page.waitForFunction((want)=>document.querySelectorAll(".ant-table-tbody tr.ant-table-row").length>=want,
    (invoices||[]).length, {timeout:7000}).catch(()=>{});
  const rows=page.locator(".ant-table-tbody tr.ant-table-row");
  const n=await rows.count();
  if(n===0) return "FAIL: no invoice rows (import XML first?)";
  const cols=await page.evaluate(()=>{
    const ths=[...document.querySelectorAll(".ant-table-thead th")].map(h=>h.innerText.trim());
    return { so: ths.findIndex(t=>/S[ốo]\s*h[óo]a\s*đ[ơo]n/i.test(t)), kh: ths.findIndex(t=>/K[ýy]\s*hi[ệe]u\s*h/i.test(t)) };
  });
  const norm=s=>String(s||"").replace(/\D/g,"").replace(/^0+/,"");
  const cell=async(i,c)=>c>=0 ? (await rows.nth(i).locator("td").nth(c).innerText().catch(()=>"")).trim() : "";
  let named=0, pdfs=0, errs=[];
  for(let i=0;i<n;i++){
    const rowSo=await cell(i,cols.so), rowKh=await cell(i,cols.kh);
    const iv=(invoices||[]).find(v=>norm(v.so_hd)===norm(rowSo) && (!rowKh || !v.ky_hieu || v.ky_hieu.trim()===rowKh))
          || (invoices||[]).find(v=>norm(v.so_hd)===norm(rowSo));
    for(let attempt=1; attempt<=2; attempt++){
      try{
        await dismissToasts(page);
        const edit=rows.nth(i).locator("i.fa-edit, .fa-edit").first();
        await edit.scrollIntoViewIfNeeded().catch(()=>{});
        await edit.click({timeout:4000});
        const modal=page.locator(".ant-modal").filter({has:page.locator(".ant-modal-content")}).last();
        await modal.waitFor({state:"visible",timeout:8000});
        await dismissToasts(page);
        if(itemName){ await modal.locator("#ten_hhoa").fill(String(itemName),{timeout:4000}); named++; }
        if(iv && iv.pdf){
          try{
            const [fc]=await Promise.all([
              page.waitForEvent("filechooser",{timeout:6000}),
              modal.locator('button.ant-btn').filter({hasText:/Chọn file/i}).first().click(),
            ]);
            await fc.setFiles(iv.pdf);
          }catch(e){ await modal.locator('input[type=file]').first().setInputFiles(iv.pdf); }
          await page.waitForTimeout(700); pdfs++;
        } else errs.push(`row${i+1}(so=${rowSo||"?"}):no invoice PDF`);
        await modal.locator('button.ant-btn').filter({hasText:/Xác nhận/i}).first().click().catch(()=>{});
        await modal.waitFor({state:"hidden",timeout:6000}).catch(()=>{});
        await page.waitForTimeout(400);
        break;
      }catch(e){
        await page.keyboard.press("Escape").catch(()=>{});
        if(attempt===2) errs.push(`row${i+1}:${(e.message||"").slice(0,40)}`); await page.waitForTimeout(600);
      }
    }
  }
  return `${n} row(s): item name ${named}/${n}, invoice PDF ${pdfs}/${n}`+(errs.length?" | "+errs.join("; "):"");
}

// Excel column -> how to fill. Bank fields appear only when HINH THUC = Chuyen khoan.
const TEXT_IDS = { "Nguoi_Thu_Huong":"nguoi_huong", "Ma_Citad":"benBnkCd",
                   "Dien_Giai":"dien_giai", "Ghi_Chu":"ttin_lquan",
                   "Ngan_Hang":"ten_tknh", "So_Tai_Khoan":"so_tknh" };
// order: Hinh_Thuc_Thanh_Toan BEFORE Loai_Chuyen_Khoan (bank fields render after choosing Chuyen khoan)
const SELECT_IDS = { "Nguoi_De_Nghi":"ma_cbcnv", "Nhom_Ke_Toan":"loai_cphi",
                     "Chung_Tu_Kem_Theo":"ctu_ktheo", "Hinh_Thuc_Thanh_Toan":"ma_httoan",
                     "Ma_Tien_Te":"ma_tte", "Loai_Chuyen_Khoan":"hthuc_ckhoan" };

const AUTO = !!process.env.TB_AUTO;   // TB_AUTO=1 -> headless, no prompts, write report (for testing)

function label(d){ const f=d.fields||{}; const n=(d.xml||[]).length;
  return (f.Nguoi_De_Nghi||f.Nguoi_Thu_Huong||"(no name)")+` - ${n} invoice(s)`+(f.Dien_Giai?` - ${f.Dien_Giai}`:""); }

(async()=>{
  if(!fs.existsSync(DON)){ console.log("No don.json (run don_to_json.py first)."); return; }
  let dons=JSON.parse(fs.readFileSync(DON,"utf8"));
  if(!dons.length){ console.log("don.json is empty."); return; }
  // who to fill: TB_ONLY="Name A;Name B" (web UI ticks the people), else ask here
  const ONLY=(process.env.TB_ONLY||"").split(";").map(s=>s.trim()).filter(Boolean);
  if(ONLY.length){
    const norm=s=>String(s||"").trim().toLowerCase();
    const keep=dons.filter(d=>ONLY.some(o=>norm(o)===norm((d.fields||{}).Nguoi_De_Nghi)
                                          ||norm(o)===norm((d.fields||{}).Nguoi_Thu_Huong)));
    if(keep.length) dons=keep; else console.log("(nobody matched the chosen names -> fill all)");
  }
  if(!AUTO && !WEB){
    console.log("\n=== Requests to fill ===");
    dons.forEach((d,i)=>console.log(`  ${i+1}. ${label(d)}`));
    console.log("  0. All");
    const ans=(await ask("\nChoose (number, several split by space, Enter or 0 = all): ")).trim();
    if(ans!=="" && ans!=="0"){
      const pick=ans.split(/[ ,]+/).map(x=>parseInt(x,10)-1).filter(x=>x>=0&&x<dons.length);
      if(pick.length) dons=pick.map(i=>dons[i]);
      else console.log("(unrecognized choice -> fill all)");
    }
  }
  if(!AUTO){
    console.log(`-> Will fill ${dons.length} request(s):`);
    dons.forEach((d,i)=>console.log(`   ${i+1}. ${label(d)}`));
  }
  const ctx=await chromium.launchPersistentContext(PROFILE,{headless:AUTO,channel:"msedge",viewport:null,args:["--start-maximized"]});
  const page=ctx.pages()[0]||await ctx.newPage();
  // close any extra blank pages so only the working one is visible
  for(const p of ctx.pages()){ if(p!==page && (p.url()==="about:blank"||p.url()==="")) await p.close().catch(()=>{}); }
  console.log(`\n=== Tool B - ${dons.length} request(s)${AUTO?" [AUTO]":""} ===`);
  if(!AUTO) console.log("Opening the form and filling now. Watch the Edge window; do NOT close it.");
  const report=[], skippedAll=[];
  for(let i=0;i<dons.length;i++){
    const d=dons[i]; const F=d.fields||{};
    await page.bringToFront().catch(()=>{});
    await page.goto(URL,{waitUntil:"domcontentloaded"}).catch(()=>{});
    const { waitForLogin, formLoaded, HOME } = require("./login.js");
    const ok=await formLoaded(page, 30000);
    if(!ok && !AUTO){
      // not logged in: show the login page, wait for the user, then carry on by itself (no Enter needed)
      console.log("[[LOGIN]] Portal not logged in. Log in in the Edge window, the tool goes on by itself.");
      await page.goto(HOME,{waitUntil:"domcontentloaded"}).catch(()=>{});
      await page.bringToFront().catch(()=>{});
      if(!await waitForLogin(ctx, page, 10)){ console.log("  Still not logged in -> stopping here."); break; }
      console.log("  Logged in, going on.");
    }
    const rep=[];
    const T = async (label, fn) => { const s=Date.now(); let r; try{ r=await fn(); }catch(e){ r="ERR:"+e.message; } return `  ${label.padEnd(16)}: ${r}  [${Date.now()-s}ms]`; };
    // NATURAL ORDER (top to bottom of the form): header fields + beneficiary + attach, THEN import XML, THEN per-row check.
    rep.push(await T("ngay_ctu", ()=>setDate(page, F.Ngay_De_Nghi)));
    for(const [col,id] of Object.entries(SELECT_IDS))
      if(col in F) rep.push(await T(id, ()=>setSelect(page,id,F[col])));
    for(const [col,id] of Object.entries(TEXT_IDS))
      if(col in F) rep.push(await T(id, ()=>setText(page,id,F[col])));
    if(d.invoices && d.invoices.length) rep.push(await T("import XML", ()=>importXml(page,d.invoices)));
    const noneOk = (d.invoices||[]).length>0 && (d.invoices||[]).every(iv=>iv.skipped);
    if(noneOk){
      rep.push("  >> WHOLE REQUEST SKIPPED: the portal rejected every invoice. Do NOT save this one.");
    } else {
    // FILE DINH KEM = to trinh + all check PDFs. Invoice PDF is NOT attached here (it goes into each row's "Sua").
    // every PDF in the person's check\ folder (any name), minus the ones of invoices the portal rejected
    const skipStems=(d.invoices||[]).filter(iv=>iv.skipped).map(iv=>iv.stem.toLowerCase());
    const checkFiles=(d.checks||[]).filter(c=>!skipStems.some(st=>path.basename(c).toLowerCase().startsWith(st)));
    const fileAttachs=[...(d.totrinh||[]), ...checkFiles];
    if(fileAttachs.length) rep.push(await T("file attach", ()=>attach(page,fileAttachs)) + ` (${(d.totrinh||[]).length} to-trinh + ${checkFiles.length} check)`);
    if(d.invoices && d.invoices.length) rep.push(await T("invoice rows", ()=>editRows(page,d.invoices,F.Ten_Hang_Hoa||"Ăn uống")));
    }
    const warns=await page.evaluate(()=>[...document.querySelectorAll(".ant-message-notice,.ant-notification-notice")].map(e=>e.innerText.trim().replace(/\s+/g," ")).filter(Boolean));
    console.log(rep.join("\n"));
    if(warns.length) console.log("  portal messages: "+warns.join(" | "));
    const shot=path.join(DUMP,`don_${i+1}.png`);
    await page.screenshot({path:shot,fullPage:true}).catch(()=>{});
    console.log("  Screenshot: "+shot);
    console.log("  "+await dumpTable(page));
    const noBank = !(F.Nguoi_Thu_Huong || F.So_Tai_Khoan);
    if(!AUTO){
      if(noBank) console.log("  >> Click 'Chon Nguoi Thu Huong' on the form -> bank + account auto-fill.");
      console.log("  >> CHECK in Edge. Fill any missing field, then click Save yourself. The tool does NOT save.");
      if(i < dons.length-1) await ask(WEB ? `Kiểm tra và bấm Lưu đơn này trên cổng, rồi làm tiếp đơn ${i+2}/${dons.length}`
                                          : `  Save this one if OK, then press Enter for request ${i+2}/${dons.length}... `);
    }
    const sk=(d.invoices||[]).filter(iv=>iv.skipped).map(iv=>({who:F.Nguoi_De_Nghi||"", file:path.basename(iv.xml), so:iv.so_hd, reason:iv.skipped}));
    skippedAll.push(...sk);
    report.push({request:i+1, fields:rep, messages:warns, skipped:sk, screenshot:shot});
  }
  if(skippedAll.length){
    console.log(`
=== ${skippedAll.length} invoice(s) SKIPPED - portal rejected them, NOT in the request ===`);
    skippedAll.forEach(x=>console.log(`  ${x.who} | ${x.file} (no.${x.so}): ${x.reason}`));
  }
  if(AUTO) fs.writeFileSync(path.join(DUMP,"fill_report.json"), JSON.stringify(report,null,2),"utf8");
  else await ask(WEB ? "Xong. Đóng cửa sổ Edge" : "\nDone. Press Enter to close Edge... ");
  await ctx.close();
  process.exit(0);
})();
