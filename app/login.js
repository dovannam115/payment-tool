// Log in to thanhtoannew.pvi.com.vn once. The session is kept in the account's Edge profile so the
// filler (fill_don.js) works later without logging in again.
// Opens Edge, waits until the payment-request form loads, then closes.
const { chromium } = require("playwright-core");
const fs = require("fs");
const path = require("path");

const ROOT    = path.join(__dirname, "..");
// TB_PROFILE = the Edge profile of the portal account picked in the app (one login per account)
const PROFILE = process.env.TB_PROFILE || path.join(ROOT, ".edge-profile");
// TB_UNIT = the Đơn vị set for that account in the app ("PVI Digital", "PVI Connect"); empty = user picks
const UNIT    = (process.env.TB_UNIT || "").trim();
const COMPANY = "PVI Insurance";                       // both units sit under it on the login screen
// unit name -> the code the portal keeps in localStorage UNIT_KEY once logged in (learned on login)
const UNITS_FILE = process.env.TB_UNITS_FILE || path.join(__dirname, "dump", "units.json");
const HOME    = "https://thanhtoannew.pvi.com.vn/#/";
const LOGIN   = "https://thanhtoannew.pvi.com.vn/#/login";
const FORM    = "https://thanhtoannew.pvi.com.vn/#/payment-requests/create?loai_ttoan=01";
const WAIT_MIN = 10;

// The portal sends a stale session straight to #/login (pick Đơn vị + Năm, then "Đăng nhập Email"
// -> company SSO), so the URL is what tells logged in from logged out, not the page contents.
const LOGIN_URL = /[#/](login|signin|dang-nhap)/i;
const PORTAL = /thanhtoannew\.pvi\.com\.vn/i;

// The app raises the Edge window to the front on this marker, so the user sees what is filled.
function front() { console.log("[[FRONT]]"); }

// Wait for the create form, but give up at once when the portal bounced us to the login page.
async function formLoaded(page, ms) {
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (LOGIN_URL.test(page.url() || "")) return false;
    if (await page.waitForSelector("#dien_giai, #ngay_ctu", { timeout: 1200 }).then(() => true).catch(() => false)) return true;
  }
  return false;
}

// True when the user is through the login screen. Read-only (never navigates), so nobody
// gets interrupted while logging in.
async function pastLogin(page) {
  const url = page.url() || "";
  return PORTAL.test(url) && !LOGIN_URL.test(url);     // on SSO pages the host is not the portal
}

// ---------------------------------------------------------------- unit (Đơn vị)
function readUnits() {
  const d = { "PVI Digital": "45", "PVI Connect": "56" };   // read off the login screen 03/11/2026
  try { Object.assign(d, JSON.parse(fs.readFileSync(UNITS_FILE, "utf8"))); } catch (e) {}
  return d;
}
function learnUnit(name, code) {
  if (!name || !code || code === "undefined") return;
  const d = readUnits();
  if (d[name] === code) return;
  if (Object.entries(d).some(([k, v]) => k !== name && v === code)) return;   // never two names on one code
  d[name] = code;
  try { fs.mkdirSync(path.dirname(UNITS_FILE), { recursive: true });
        fs.writeFileSync(UNITS_FILE, JSON.stringify(d, null, 1)); } catch (e) {}
}
async function unitCode(page) {
  const v = await page.evaluate(() => localStorage.getItem("UNIT_KEY")).catch(() => null);
  return v ? String(v).replace(/^"|"$/g, "") : "";
}
// Logged in, but to the unit set for this account? Unknown = assume yes (nothing to compare with).
async function rightUnit(page) {
  if (!UNIT) return true;
  const got = await unitCode(page);
  if (!got || got === "undefined") return true;
  const units = readUnits();
  if (units[UNIT]) return units[UNIT] === got;
  return !Object.entries(units).some(([k, v]) => k !== UNIT && v === got);   // logged in to another known unit
}

// On the login screen: pick PVI Insurance + the account's unit + this year, then "Đăng nhập Email".
// Only the visible controls, like a user would. Returns a short note for the log.
async function pickUnit(page) {
  if (!UNIT) return "no unit set for this account, pick it on the screen";
  try {
    await page.waitForSelector("#server", { timeout: 15000 });
    await page.selectOption("#server", { label: COMPANY });
    await page.waitForTimeout(600);
    const box = page.locator(".ant-select").first();
    await box.click();
    await page.keyboard.type(UNIT, { delay: 20 });
    const opt = page.locator(`.ant-select-dropdown:visible .ant-select-item-option[title="${UNIT}"]`);
    await opt.first().waitFor({ timeout: 6000 });
    await opt.first().click();
    await page.waitForTimeout(300);
    const shown = ((await box.locator(".ant-select-selection-item").first()
                         .getAttribute("title", { timeout: 2000 }).catch(() => "")) || "").trim();
    if (shown !== UNIT) return `could not pick '${UNIT}' (screen shows '${shown}'), pick it by hand`;
    const year = String(new Date().getFullYear());
    if (await page.locator(`#year option[value="${year}"]`).count()) await page.selectOption("#year", year);
    await page.locator("button", { hasText: "Đăng nhập Email" }).first().click();
    return `picked ${COMPANY} / ${UNIT} / ${year}`;
  } catch (e) {
    return `could not pick '${UNIT}' (${String(e && e.message).split("\n")[0]}), pick it by hand`;
  }
}

// Leave the current portal session so the next login can pick another unit. Opens the login
// screen; when the portal sends a live session away from it, its saved token is dropped first
// (the same as the session running out, the SSO login itself stays).
async function leaveSession(page) {
  await page.goto(LOGIN, { waitUntil: "domcontentloaded" }).catch(() => {});
  await page.waitForTimeout(2500);
  if (LOGIN_URL.test(page.url() || "")) return;
  await page.evaluate(() => localStorage.removeItem("TOKEN_KEY")).catch(() => {});
  await page.goto(LOGIN, { waitUntil: "domcontentloaded" }).catch(() => {});
  await page.reload({ waitUntil: "domcontentloaded" }).catch(() => {});
}

// Wait for the user to log in, then open the form. Returns the page with the form up, or null.
async function waitForLogin(ctx, page, minutes) {
  const until = Date.now() + minutes * 60000;
  let streak = 0, picks = 0, lastUrl = "";
  while (Date.now() < until) {
    if (page.isClosed()) {                       // SSO can move the user to another window
      const open = ctx.pages().filter((p) => !p.isClosed());
      if (!open.length) return null;
      page = open[open.length - 1];
    }
    const url = page.url() || "";
    // fill the login screen once each time it comes up (at most 3 times, then the user picks)
    if (LOGIN_URL.test(url) && PORTAL.test(url) && url !== lastUrl && picks < 3) {
      picks++;
      console.log("  Login screen: " + (await pickUnit(page)));
    }
    lastUrl = url;
    await new Promise((r) => setTimeout(r, 2500));
    streak = (await pastLogin(page)) ? streak + 1 : 0;
    if (streak < 2) continue;                    // two checks in a row, so a half-loaded page does not count
    await page.goto(FORM, { waitUntil: "domcontentloaded" }).catch(() => {});
    if (await formLoaded(page, 8000)) {
      if (UNIT) learnUnit(UNIT, await unitCode(page));
      return page;
    }
    streak = 0; lastUrl = "";
  }
  return null;
}

// Open the create form logged in to the right unit, logging in (or switching unit) when needed.
// Returns the page with the form up, or null when nobody logged in within `minutes`.
async function ensureForm(ctx, page, minutes) {
  await page.goto(FORM, { waitUntil: "domcontentloaded" }).catch(() => {});
  front();
  if (await formLoaded(page, 30000)) {
    if (await rightUnit(page)) return page;
    console.log(`[[LOGIN]] Portal is logged in to another unit, switching to ${UNIT}. Finish the login in the Edge window if asked.`);
    await leaveSession(page);
  } else {
    console.log("[[LOGIN]] Portal not logged in. " + (UNIT ? `The tool picks ${UNIT}, ` : "") +
                "log in in the Edge window, the tool goes on by itself.");
    if (!LOGIN_URL.test(page.url() || "")) await page.goto(HOME, { waitUntil: "domcontentloaded" }).catch(() => {});
  }
  await page.bringToFront().catch(() => {});
  front();
  const p = await waitForLogin(ctx, page, minutes);
  if (p && !(await rightUnit(p))) {
    console.log(`  !! Logged in, but not to ${UNIT}. Pick ${UNIT} on the login screen.`);
    return null;
  }
  return p;
}

module.exports = { formLoaded, pastLogin, waitForLogin, ensureForm, pickUnit, front, HOME, FORM, UNIT };

if (require.main === module) (async () => {
  const ctx = await chromium.launchPersistentContext(PROFILE, {
    headless: false, channel: "msedge", viewport: null, args: ["--start-maximized"],
  });
  const page = ctx.pages()[0] || (await ctx.newPage());
  for (const p of ctx.pages()) if (p !== page && (p.url() === "about:blank" || p.url() === "")) await p.close().catch(() => {});

  console.log("Opening the portal in Edge..." + (UNIT ? " (unit: " + UNIT + ")" : ""));
  const ok = await ensureForm(ctx, page, WAIT_MIN);
  console.log(ok ? "Logged in" + (UNIT ? " to " + UNIT : "") + ". Session saved, requests can be filled now."
                 : "Still not logged in after " + WAIT_MIN + " minutes. Open it again when ready.");
  await ctx.close();
  process.exit(0);
})().catch((e) => { console.log("ERROR: " + (e && e.message)); process.exit(1); });
