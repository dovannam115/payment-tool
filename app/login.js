// Log in to thanhtoannew.pvi.com.vn once. The session is kept in .edge-profile so the
// filler (fill_don.js) works later without logging in again.
// Opens Edge, waits until the payment-request form loads, then closes.
const { chromium } = require("playwright-core");
const path = require("path");

const ROOT    = path.join(__dirname, "..");
const PROFILE = path.join(ROOT, ".edge-profile");
const HOME    = "https://thanhtoannew.pvi.com.vn/#/";
const FORM    = "https://thanhtoannew.pvi.com.vn/#/payment-requests/create?loai_ttoan=01";
const WAIT_MIN = 10;

// The portal sends a stale session straight to #/login (pick Đơn vị + Năm, then "Đăng nhập Email"
// -> company SSO), so the URL is what tells logged in from logged out, not the page contents.
const LOGIN_URL = /[#/](login|signin|dang-nhap)/i;
const PORTAL = /thanhtoannew\.pvi\.com\.vn/i;

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

// Wait for the user to log in, then open the form. Returns true when the form is up.
async function waitForLogin(ctx, page, minutes) {
  const until = Date.now() + minutes * 60000;
  let streak = 0;
  while (Date.now() < until) {
    await new Promise((r) => setTimeout(r, 2500));
    if (page.isClosed()) {                       // SSO can move the user to another window
      const open = ctx.pages().filter((p) => !p.isClosed());
      if (!open.length) return false;
      page = open[open.length - 1];
    }
    streak = (await pastLogin(page)) ? streak + 1 : 0;
    if (streak < 2) continue;                    // two checks in a row, so a half-loaded page does not count
    await page.goto(FORM, { waitUntil: "domcontentloaded" }).catch(() => {});
    if (await formLoaded(page, 8000)) return true;
    streak = 0;
  }
  return false;
}

module.exports = { formLoaded, pastLogin, waitForLogin, HOME, FORM };

if (require.main === module) (async () => {
  const ctx = await chromium.launchPersistentContext(PROFILE, {
    headless: false, channel: "msedge", viewport: null, args: ["--start-maximized"],
  });
  const page = ctx.pages()[0] || (await ctx.newPage());
  for (const p of ctx.pages()) if (p !== page && (p.url() === "about:blank" || p.url() === "")) await p.close().catch(() => {});

  console.log("Opening the portal in Edge...");
  await page.goto(FORM, { waitUntil: "domcontentloaded" }).catch(() => {});
  if (await formLoaded(page, 12000)) {
    console.log("Already logged in, nothing to do.");
    await ctx.close();
    return;
  }
  await page.goto(HOME, { waitUntil: "domcontentloaded" }).catch(() => {});
  await page.bringToFront().catch(() => {});
  console.log("[[LOGIN]] Log in to the portal in the Edge window. The tool waits and goes on by itself.");
  const ok = await waitForLogin(ctx, page, WAIT_MIN);
  console.log(ok ? "Logged in. Session saved, requests can be filled now."
                 : "Still not logged in after " + WAIT_MIN + " minutes. Open it again when ready.");
  await ctx.close();
})().catch((e) => { console.log("ERROR: " + (e && e.message)); process.exit(1); });
