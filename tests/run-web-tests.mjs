// End-to-end tests for the browser version (index.html + app.js + engine) in headless Chromium.
// Usage: npm install && npm run test:web     (set CHROMIUM_PATH to use a specific Chromium build)
// Fixtures: tests/fixtures/lease-libreoffice.docx (demo lease saved by LibreOffice) and
// tests/fixtures/psa-tracked.docx (demo PSA plus tracked changes, a table, a text box, field codes).
import { chromium } from "playwright-core";
import { createServer } from "node:http";
import { readFileSync, writeFileSync, mkdirSync, existsSync, statSync } from "node:fs";
import { join, extname, dirname, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const SHOTS = process.env.SHOTS_DIR || join(ROOT, "tests", ".shots");
mkdirSync(SHOTS, { recursive: true });
const Engine = require(join(ROOT, "engine", "clause-engine.js"));
const Demo = require(join(ROOT, "engine", "demo-docs.js"));
const FIX = join(ROOT, "tests", "fixtures");

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml" };
const server = createServer((req, res) => {
  const path = normalize(decodeURIComponent(new URL(req.url, "http://x").pathname)).replace(/^(\.\.[/\\])+/, "");
  const file = join(ROOT, path === "/" ? "index.html" : path);
  if (!file.startsWith(ROOT) || !existsSync(file) || statSync(file).isDirectory()) { res.writeHead(404); res.end("not found"); return; }
  res.writeHead(200, { "content-type": TYPES[extname(file)] || "application/octet-stream" });
  res.end(readFileSync(file));
});
await new Promise((r) => server.listen(0, "127.0.0.1", r));
const BASE = `http://127.0.0.1:${server.address().port}`;

const launchOpts = {};
if (process.env.CHROMIUM_PATH) launchOpts.executablePath = process.env.CHROMIUM_PATH;
const browser = await chromium.launch(launchOpts);

let passed = 0, failed = 0;
const errors = [];
const offsite = [];
function check(cond, label) {
  if (cond) { passed++; console.log("  ok   " + label); }
  else { failed++; console.log("  FAIL " + label); }
}

async function open(viewport = { width: 1440, height: 900 }, extra = {}) {
  const context = await browser.newContext({ viewport, acceptDownloads: true, ...extra });
  const page = await context.newPage();
  await page.route(/fonts\.(googleapis|gstatic)\.com/, (route) => route.abort());
  page.on("request", (r) => { if (!r.url().startsWith(BASE) && !/fonts\.(googleapis|gstatic)\.com/.test(r.url()) && !r.url().startsWith("blob:") && !r.url().startsWith("data:")) offsite.push(r.url()); });
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error" && !/Failed to load resource/.test(m.text())) errors.push("console: " + m.text()); });
  await page.goto(`${BASE}/index.html`);
  await page.locator("#editor p").first().waitFor();
  return { context, page };
}

const docState = (page) => page.evaluate(() => {
  const ed = document.getElementById("editor");
  const ps = Array.from(ed.children).filter((n) => /^(P|H1|H2|H3|DIV|LI)$/.test(n.tagName));
  const cur = (p) => { let out = ""; (function walk(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType === 3) out += c.nodeValue; else if (c.nodeType === 1 && c.tagName !== "DEL") walk(c); } })(p); return out; };
  const sel = window.getSelection();
  return {
    count: ps.length,
    current: ps.map(cur),
    dels: ps.map((p) => Array.from(p.querySelectorAll("del")).map((d) => d.textContent)),
    inss: ps.map((p) => Array.from(p.querySelectorAll("ins")).map((d) => d.textContent)),
    classes: ps.map((p) => p.className),
    flashed: ps.map((p, i) => (p.classList.contains("cc-flash") ? i : -1)).filter((i) => i >= 0),
    selection: sel && sel.rangeCount ? sel.toString() : "",
    selPara: sel && sel.rangeCount ? ps.findIndex((p) => p.contains(sel.anchorNode)) : -1
  };
});

const leaseTenant = Engine.review(Demo.lease.paragraphs, { perspective: "tenant" });
const titleOf = (ruleId, r = leaseTenant) => r.findings.find((f) => f.ruleId === ruleId).title;
const card = (page, title) => page.locator("#review-results .cc-card", { has: page.locator("h3", { hasText: title }) }).first();
async function cardClick(page, ruleId, button, statusText) {
  const c = card(page, titleOf(ruleId));
  await c.locator("button", { hasText: button }).click();
  if (statusText) await c.locator(".cc-status", { hasText: statusText }).waitFor({ timeout: 5000 });
}
async function hideToast(page) { await page.evaluate(() => { const t = document.getElementById("cc-toast"); if (t) t.className = "cc-toast"; }); }
/** Run an action and return the toast it raises (any earlier toast is cleared first). */
async function toastAfter(page, action) {
  await hideToast(page);
  await action();
  const t = page.locator("#cc-toast.is-visible");
  await t.waitFor({ timeout: 5000 });
  return { text: await t.textContent(), cls: await t.getAttribute("class") };
}
const clearFlash = (page) => page.evaluate(() => document.querySelectorAll(".cc-flash").forEach((e) => e.classList.remove("cc-flash")));
const noHorizontalScroll = (page) => page.evaluate(() => document.documentElement.scrollWidth <= document.documentElement.clientWidth);

/* ------------------------------------------------------------------ */
console.log("W1. Desktop: demo lease, review, redlines, Ask, Summary, export, persistence");
{
  const { context, page } = await open(undefined, { permissions: ["clipboard-read", "clipboard-write"] });
  let s = await docState(page);
  check(s.count === Demo.lease.paragraphs.length && s.current.join("\n") === Demo.lease.paragraphs.join("\n"), `demo lease loads on first visit (${s.count} paragraphs, text exact)`);
  check(s.classes[0] === "doc-title", "title styled as the document title");
  check(await page.locator("#editor p").nth(2).locator("strong").textContent() === "1. Premises.", "section captions bold, like the Word original");
  check(await page.locator("#review-stats").isHidden(), "no results before the first review");

  check(/Review done/.test((await toastAfter(page, () => page.locator("#btn-run-review").click())).text), "Run Review reports the result");
  const cards = await page.locator("#review-results .cc-card").count();
  check(cards === leaseTenant.findings.length + leaseTenant.missing.length, `all ${leaseTenant.findings.length} findings and ${leaseTenant.missing.length} missing clauses listed (${cards})`);
  check((await page.locator("#review-stats .doc-type").textContent()).includes("Lease · reviewed for the tenant side"), "detects a lease and reviews for the tenant");
  check(await page.locator("#review-stats .stat-high b").textContent() === String(leaseTenant.stats.high), "high-risk count shown");
  await hideToast(page);
  await page.screenshot({ path: join(SHOTS, "web-01-review.png") });

  await cardClick(page, "assign-sole-discretion", "Go to");
  s = await docState(page);
  check(s.selection === "sole and absolute discretion" && s.selPara === 12, "Go to selects the exact words in the document");

  await cardClick(page, "assign-sole-discretion", "Accept", "Redlined");
  s = await docState(page);
  const assign = leaseTenant.findings.find((f) => f.ruleId === "assign-sole-discretion");
  check(s.dels[12].includes("sole and absolute discretion"), "original words kept as a red strikethrough");
  check(s.inss[12].some((t) => t.startsWith("reasonable discretion (such consent")), "replacement shown as an insertion");
  check(s.current[12] === Engine.applyFixToText(assign, Demo.lease.paragraphs[12]), "paragraph now reads exactly as the engine's fix");

  await cardClick(page, "default-no-notice", "Accept", "Redlined");
  await cardClick(page, "default-short-nonmonetary", "Accept", "Redlined");
  s = await docState(page);
  check(/continues for ten \(10\) days after written notice from Landlord; or \(b\)/.test(s.current[21]) &&
        /such failure continues for thirty \(30\) days after written notice \(or such longer period/.test(s.current[21]),
        "two edits on the same phrase hit the right occurrences");

  await cardClick(page, "relocation", "Reject", "Rejected");
  s = await docState(page);
  check(s.current[19] === Demo.lease.paragraphs[19] && s.inss[19].length === 0, "rejecting leaves the document alone");

  const lowFindings = page.locator("#review-results .cc-card[data-kind=finding].sev-low");
  const lowCount = await lowFindings.count();
  await page.locator("#review-stats .stat-low").click();
  check(lowCount > 0 && await lowFindings.count() === 0, "severity filter hides low-risk items");
  await page.locator("#review-stats .stat-low").click();
  check(await lowFindings.count() === lowCount, "and shows them again");
  await page.locator("#review-stats .stat-missing").click();
  check(await page.locator("#review-results .cc-card[data-kind=missing]").count() === 0, "Missing toggle hides missing clauses");
  await page.locator("#review-stats .stat-missing").click();

  const casp = page.locator("#review-results .cc-card", { has: page.locator("h3", { hasText: "CASp" }) });
  await casp.locator("button", { hasText: "Insert clause" }).click();
  s = await docState(page);
  check(s.current[s.count - 1].startsWith("CASp Disclosure.") && s.inss[s.count - 1].length === 1, "missing CASp clause inserted at the end as an insertion");

  // Edits made between the review and a click: never redline the wrong paragraph.
  await page.evaluate(() => { const ed = document.getElementById("editor"); const p = document.createElement("p"); p.textContent = "Inserted by the user after the review."; ed.insertBefore(p, ed.children[1]); });
  let pre = await docState(page);
  await cardClick(page, "late-charge-high", "Accept", "Redlined");
  s = await docState(page);
  const late = leaseTenant.findings.find((f) => f.ruleId === "late-charge-high");
  const touched = s.current.map((t, i) => (t === pre.current[i] ? -1 : i)).filter((i) => i >= 0);
  check(touched.length === 1 && touched[0] === 23 && s.current[23] === Engine.applyFixToText(late, Demo.lease.paragraphs[22]), "paragraph moved down by an edit: the redline follows it (" + touched.join(",") + ")");
  await page.evaluate(() => { const ed = document.getElementById("editor"); ed.children[24].appendChild(document.createTextNode(" (edited)")); });
  pre = await docState(page);
  const moved = await toastAfter(page, () => card(page, titleOf("entry-any-time")).locator("button", { hasText: "Accept" }).click());
  check(/document changed since the review/.test(moved.text), "moved and edited paragraph: asks to re-run instead of guessing");
  s = await docState(page);
  check(s.current.join("\n") === pre.current.join("\n"), "nothing applied anywhere");
  await page.evaluate(() => { const ed = document.getElementById("editor"); ed.removeChild(ed.children[1]); });
  await cardClick(page, "entry-any-time", "Accept", "Redlined");
  s = await docState(page);
  check(s.dels[23].length === 1 && /\(edited\)$/.test(s.current[23]), "once the paragraph is back in place, the edit is applied and the user's own words kept");

  for (let i = 0; i < 80; i++) {
    const btn = page.locator("#review-results .cc-card button", { hasText: /^Accept$/ }).first();
    if (!(await btn.count())) break;
    await btn.click();
  }
  await page.locator("#btn-run-review-side").click();
  s = await docState(page);
  const rerun = Engine.review(s.current, { perspective: "tenant" });
  const fixable = rerun.findings.filter((f) => f.action).map((f) => f.ruleId);
  check(fixable.length === 1 && fixable[0] === "relocation", "after accepting everything and re-running, only the rejected item is still fixable (" + fixable.join(",") + ")");
  check(await page.locator("#review-results .cc-card button", { hasText: /^Accept$/ }).count() === 0, "no Accept buttons left on screen");
  check(!rerun.missing.some((m) => m.category === "casp"), "CASp no longer missing");
  check(s.current[2] === Demo.lease.paragraphs[2], "untouched paragraphs unchanged");

  await page.locator("#perspective").selectOption("landlord");
  await page.waitForTimeout(150);
  const ll = Engine.review(s.current, { perspective: "landlord" });
  check(await page.locator("#review-results .cc-card").count() === ll.findings.length + ll.missing.length, "switching to Landlord re-runs the review");
  await page.locator("#perspective").selectOption("auto");

  await page.locator(".sidebar-tabs .tab[data-tab=clauses]").click();
  check(await page.locator("#clause-list .cc-clause").count() > 20, "clause inventory listed");
  await clearFlash(page);
  await page.locator("#clause-list .cc-clause", { hasText: "Assignment" }).first().click();
  s = await docState(page);
  check(s.flashed.length === 1 && s.flashed[0] === 12, "clicking a clause jumps to it");

  await page.locator(".sidebar-tabs .tab[data-tab=ask]").click();
  await page.locator(".ask-chips .chip", { hasText: "Who can terminate?" }).click();
  const ans = page.locator("#ask-chat .cc-msg.a").last();
  check(/Landlord/.test(await ans.textContent()), "termination question answered from the document");
  await clearFlash(page);
  await ans.locator(".cc-source").first().click();
  check((await docState(page)).flashed.length === 1, "citation jumps into the document");
  await page.locator("#ask-input").fill("What is the rent?");
  await page.locator("#ask-form button").click();
  check(/\$12,442\.50/.test(await page.locator("#ask-chat .cc-msg.a").last().textContent()), "rent question answered with the figure");

  await page.locator(".sidebar-tabs .tab[data-tab=summary]").click();
  await page.locator("#btn-summary").click();
  check(await page.locator("#summary-output .cc-kv tr").count() >= 8, "summary table built");
  check(/copied/.test((await toastAfter(page, () => page.locator("#summary-output button", { hasText: "Copy summary" }).click())).text), "Copy summary confirms");
  const clip = await page.evaluate(() => navigator.clipboard.readText());
  check(clip.startsWith("CLAUSE CRAFT SUMMARY") && /Not legal advice/.test(clip), "summary copied to the clipboard");

  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#btn-export").click()]);
  const html = readFileSync(await download.path(), "utf8");
  check(/<del class="cc-del">sole and absolute discretion<\/del>/.test(html) && /<ins class="cc-ins">/.test(html), "exported HTML keeps the redlines");
  check(!/<script/i.test(html) && /not legal advice/i.test(html) && download.suggestedFilename() === "office-lease-redline.html", "export is a clean, named redline file");

  const before = await docState(page);
  await page.reload();
  await page.locator("#editor p").first().waitFor();
  const after = await docState(page);
  check(after.count === before.count && JSON.stringify(after.dels) === JSON.stringify(before.dels) && after.current.join("\n") === before.current.join("\n"), "document and redlines survive a page reload");

  await page.locator(".desk").evaluate((el) => { el.scrollTop = 1050; });
  await page.locator(".sidebar-tabs .tab[data-tab=review]").click();
  await page.locator("#btn-run-review").click();
  await hideToast(page);
  await page.screenshot({ path: join(SHOTS, "web-02-redlined.png") });

  page.once("dialog", (d) => d.dismiss());
  await page.locator("#btn-load-demo").click();
  check((await docState(page)).dels.flat().length > 0, "Load Demo Lease asks first and keeps your work when you cancel");
  page.once("dialog", (d) => d.accept());
  await page.locator("#btn-load-demo").click();
  s = await docState(page);
  check(s.dels.flat().length === 0 && s.current.join("\n") === Demo.lease.paragraphs.join("\n"), "and replaces it when you confirm");
  await context.close();
}

/* ------------------------------------------------------------------ */
console.log("W2. Upload: .txt purchase agreement");
{
  const { context, page } = await open();
  await page.locator("#file-input").setInputFiles({ name: "psa.txt", mimeType: "text/plain", buffer: Buffer.from(Demo.psa.paragraphs.join("\r\n")) });
  await page.waitForFunction(() => document.querySelectorAll("#editor p").length === 37);
  await page.locator("#btn-run-review").click();
  const psa = Engine.review(Demo.psa.paragraphs, { perspective: "buyer" });
  check((await page.locator("#review-stats .doc-type").textContent()).includes("Purchase agreement · reviewed for the buyer side"), "detects a PSA and reviews for the buyer");
  check(await page.locator("#review-results .cc-card").count() === psa.findings.length + psa.missing.length, "buyer findings listed");
  await card(page, titleOf("psa-deposit-hard", psa)).locator("button", { hasText: "Accept" }).click();
  const s = await docState(page);
  check(s.inss[5].some((t) => /fully refundable/.test(t)), "deposit protection added to the deposit paragraph");
  await context.close();
}

/* ------------------------------------------------------------------ */
console.log("W3. Upload: Word documents (.docx)");
{
  const { context, page } = await open();
  await toastAfter(page, () => page.locator("#file-input").setInputFiles(join(FIX, "lease-libreoffice.docx")));
  let s = await docState(page);
  check(s.current.join("\n") === Demo.lease.paragraphs.join("\n"), `.docx saved by LibreOffice opens with every paragraph intact (${s.count})`);
  await page.locator("#btn-run-review").click();
  check(await page.locator("#review-results .cc-card").count() === leaseTenant.findings.length + leaseTenant.missing.length, "and reviews the same as the demo");

  const expected = JSON.parse(readFileSync(join(FIX, "psa-tracked.expected.json"), "utf8"));
  await toastAfter(page, () => page.locator("#file-input").setInputFiles(join(FIX, "psa-tracked.docx")));
  s = await docState(page);
  check(s.count === expected.length, `Word .docx with tracked changes, table and text box: ${s.count} paragraphs (expected ${expected.length})`);
  const diffs = expected.map((t, i) => (t === s.current[i] ? null : i + ": " + JSON.stringify(s.current[i]) + " != " + JSON.stringify(t))).filter(Boolean);
  check(diffs.length === 0, "tracked deletions and moved-away text skipped, insertions, fields, links and tables kept " + diffs.join(" | "));
  await page.locator("#btn-run-review").click();
  const r = Engine.review(expected, { perspective: "buyer" });
  check(await page.locator("#review-results .cc-card").count() === r.findings.length + r.missing.length, "reviewed as a purchase agreement");

  let t = await toastAfter(page, () => page.locator("#file-input").setInputFiles({ name: "broken.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from("not a zip file at all") }));
  check(/isn't a valid \.docx/.test(t.text) && /is-err/.test(t.cls), "a damaged .docx gives a clear error");
  t = await toastAfter(page, () => page.locator("#file-input").setInputFiles({ name: "lease.pdf", mimeType: "application/pdf", buffer: Buffer.from("%PDF-1.4") }));
  check(/Save it as \.docx/.test(t.text), "PDF and old .doc files: says to save as .docx");
  check((await docState(page)).count === expected.length, "failed uploads leave the open document alone");
  await context.close();
}

/* ------------------------------------------------------------------ */
console.log("W4. Phone and tablet widths, print layout");
{
  const { context, page } = await open({ width: 390, height: 844 }, { isMobile: true, hasTouch: true, deviceScaleFactor: 2 });
  check(await noHorizontalScroll(page), "phone: no sideways scrolling");
  await page.screenshot({ path: join(SHOTS, "web-03-phone-top.png") });
  await page.locator("#btn-run-review").click();
  await card(page, titleOf("assign-sole-discretion")).locator("button", { hasText: "Accept" }).click();
  check(await noHorizontalScroll(page), "phone: still no sideways scrolling with results showing");
  const box = await page.locator(".sidebar").boundingBox();
  check(box && box.width <= 390 && box.x >= 0, "phone: review panel stacks under the document");
  await hideToast(page);
  await page.locator(".sidebar").evaluate((el) => el.scrollIntoView());
  await page.screenshot({ path: join(SHOTS, "web-04-phone-panel.png") });
  await page.setViewportSize({ width: 820, height: 1180 });
  check(await noHorizontalScroll(page), "tablet: no sideways scrolling");

  await page.setViewportSize({ width: 1280, height: 900 });
  await page.emulateMedia({ media: "print" });
  const hidden = await page.evaluate(() => [".sidebar", ".app-header", ".contact-strip", ".floating-toolbar", ".app-footer"].every((s) => getComputedStyle(document.querySelector(s)).display === "none"));
  check(hidden, "print: only the document prints");
  const pdf = await page.pdf({ format: "Letter" });
  check(pdf.length > 10000, "print: produces a PDF (" + Math.round(pdf.length / 1024) + " KB)");
  if (process.env.PDF_OUT) writeFileSync(process.env.PDF_OUT, pdf);
  await context.close();
}

const appErrors = errors.filter((e) => !/net::ERR_FAILED|ERR_NAME_NOT_RESOLVED|fonts\.g/.test(e));
check(appErrors.length === 0, "no JavaScript errors (" + appErrors.join(" | ") + ")");
check(offsite.length === 0, "no requests leave the page except Google Fonts (" + offsite.join(", ") + ")");
await browser.close();
server.close();
console.log(`\n${passed} passed, ${failed} failed. Screenshots: ${SHOTS}`);
process.exit(failed ? 1 : 0);
