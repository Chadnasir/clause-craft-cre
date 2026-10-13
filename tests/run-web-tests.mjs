// End-to-end tests for the browser version (index.html + app.js + engine) in headless Chromium.
// Usage: npm install && npm run test:web     (set CHROMIUM_PATH to use a specific Chromium build)
// Fixtures: tests/fixtures/lease-libreoffice.docx (demo lease saved by LibreOffice) and
// tests/fixtures/psa-tracked.docx (demo PSA plus tracked changes, a table, a text box, field codes).
import { chromium } from "playwright-core";
import { createServer } from "node:http";
import { readFileSync, writeFileSync, mkdirSync, existsSync, statSync, mkdtempSync, rmSync } from "node:fs";
import { join, extname, dirname, normalize } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { execFileSync } from "node:child_process";
import { tmpdir } from "node:os";
import zlib from "node:zlib";

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

async function open(viewport = { width: 1440, height: 900 }, extra = {}, url = `${BASE}/index.html`) {
  const context = await browser.newContext({ viewport, acceptDownloads: true, ...extra });
  const page = await context.newPage();
  const origin = url.startsWith("file:") ? pathToFileURL(ROOT).href : BASE;
  await page.route(/fonts\.(googleapis|gstatic)\.com/, (route) => route.abort());
  page.on("request", (r) => { if (!r.url().startsWith(origin) && !/fonts\.(googleapis|gstatic)\.com/.test(r.url()) && !r.url().startsWith("blob:") && !r.url().startsWith("data:")) offsite.push(r.url()); });
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error" && !/Failed to load resource/.test(m.text())) errors.push("console: " + m.text()); });
  await page.goto(url);
  await page.locator("#editor p").first().waitFor();
  return { context, page };
}

/** Put the text cursor or a selection in the document: paragraph a at offset s, to paragraph b at offset e. */
const selectInEditor = (page, a, s, e = s, b = a) => page.evaluate(([a, s, b, e]) => {
  const ps = Array.from(document.getElementById("editor").children);
  function point(p, off) {
    let count = 0, found = null;
    (function walk(n) {
      for (let c = n.firstChild; c && !found; c = c.nextSibling) {
        if (c.nodeType === 3) { if (off <= count + c.nodeValue.length) { found = [c, off - count]; return; } count += c.nodeValue.length; }
        else if (c.nodeType === 1 && c.tagName !== "DEL") walk(c);
      }
    })(p);
    return found || [p, p.childNodes.length];
  }
  const x = point(ps[a], s), y = point(ps[b], e);
  const r = document.createRange(); r.setStart(x[0], x[1]); r.setEnd(y[0], y[1]);
  const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
}, [a, s, b, e]);

/** Read a .docx (zip) in Node: every part's text, with each part's CRC checked. */
function readZip(buf) {
  const files = {};
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("not a zip");
  let off = buf.readUInt32LE(eocd + 16);
  for (let k = 0, n = buf.readUInt16LE(eocd + 10); k < n; k++) {
    const method = buf.readUInt16LE(off + 10), crc = buf.readUInt32LE(off + 16), size = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28), extra = buf.readUInt16LE(off + 30), comment = buf.readUInt16LE(off + 32), local = buf.readUInt32LE(off + 42);
    const name = buf.toString("utf8", off + 46, off + 46 + nameLen);
    const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
    const raw = buf.subarray(start, start + size);
    const data = method === 0 ? raw : zlib.inflateRawSync(raw);
    if (zlib.crc32(data) !== crc) throw new Error("bad CRC for " + name);
    files[name] = data.toString("utf8");
    off += 46 + nameLen + extra + comment;
  }
  return files;
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
  await page.locator("#btn-run-review").click();
  s = await docState(page);
  const rerun = Engine.review(s.current, { perspective: "tenant" });
  const fixable = rerun.findings.filter((f) => f.action).map((f) => f.ruleId);
  check(fixable.length === 1 && fixable[0] === "relocation", "after accepting everything and re-running, only the rejected item is still fixable (" + fixable.join(",") + ")");
  check(await page.locator("#review-results .cc-card button", { hasText: /^Accept$/ }).count() === 0, "no Accept buttons left on screen");
  check(!rerun.missing.some((m) => m.category === "casp"), "CASp no longer missing");
  check(s.current[2] === Demo.lease.paragraphs[2], "untouched paragraphs unchanged");

  await page.locator("#btn-setup-edit").click();
  await page.locator("#perspective").selectOption("landlord");
  await page.waitForTimeout(150);
  const ll = Engine.review(s.current, { perspective: "landlord" });
  check(await page.locator("#review-results .cc-card").count() === ll.findings.length + ll.missing.length, "switching to Landlord re-runs the review");
  await page.locator("#perspective").selectOption("auto");

  await page.locator(".sidebar-tabs .tab[data-tab=summary]").click();
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
console.log("W5. Spellbook-style workflow: setup, selection, playbook, drafting, saved clauses, explain, issues list, Word export");
{
  const { context, page } = await open(undefined, { permissions: ["clipboard-read", "clipboard-write"] });
  const P = Demo.lease.paragraphs;
  check(await page.locator("#review-empty").isVisible() && await page.locator("#qs-review").isVisible() && await page.locator("#qs-playbook").isVisible(), "getting started shows before the first review");
  check(await page.locator("#perspective option[value=tenant]").textContent() === "Tenant: HARBORLINE ADVISORY GROUP, INC.", "party list shows the names from the document");
  await page.locator("#qs-review").click();
  check(await page.locator("#setup-fields").isHidden() && await page.locator("#setup-summary-text").textContent() === "Tenant (auto) · Comprehensive · California (detected) · My standards", "setup folds into a one-line summary after the review");
  check(await page.locator("#review-empty").isHidden(), "getting started hides once there are results");
  check(await page.locator("#count-review").textContent() === String(leaseTenant.findings.length + leaseTenant.missing.length), "Review tab shows the item count");
  const checkTexts = await page.locator("#checks li").allTextContents();
  check(checkTexts.length === leaseTenant.checks.length && checkTexts.some((t) => /Term dates add up/.test(t)) && checkTexts.some((t) => /Rent, Additional Rent have no definition/.test(t)), `document checks listed (${checkTexts.length})`);

  // Review selection: paragraphs 12-13 (assignment, insurance)
  await selectInEditor(page, 12, 0, 20, 13);
  await page.locator("#btn-review-selection").click();
  const inSel = leaseTenant.findings.filter((f) => f.paragraphIndex >= 12 && f.paragraphIndex <= 13).length;
  check(await page.locator("#review-results .cc-card").count() === inSel && /2 paragraphs/.test(await page.locator("#scope-text").textContent()), `Review selection shows only the ${inSel} issues in the selected paragraphs`);
  await page.locator("#btn-scope-clear").click();
  check(await page.locator("#review-results .cc-card").count() === leaseTenant.findings.length + leaseTenant.missing.length, "Show all brings every issue back");

  // Playbook
  await page.locator(".sidebar-tabs .tab[data-tab=playbook]").click();
  const pb = Engine.playbook(P, { perspective: "tenant" });
  check(await page.locator("#pb-items .cc-card").count() === pb.items.length, `playbook checks ${pb.items.length} standards`);
  check(await page.locator("#count-playbook").textContent() === String(pb.stats.fail), "Playbook tab shows how many standards are off");
  const lateItem = pb.items.find((i) => i.standardId === "tn-late-charge");
  await page.locator("#pb-items [data-standard=tn-late-charge]").locator("button", { hasText: "Apply redline" }).click();
  let s = await docState(page);
  check(s.dels[22].includes("ten percent (10%)") && s.inss[22].includes("five percent (5%)") && s.current[22] === Engine.applyFixToText(lateItem, P[22]), "playbook redline: late charge 10% to your 5%");
  check(await page.locator("#pb-items [data-standard=tn-late-charge] .cc-badge").textContent() === "OK", "and that standard now passes");
  await page.locator("#btn-pb-edit").click();
  await page.locator('[data-std="tn-late-charge"]').fill("4");
  await page.locator("#btn-pb-save").click();
  const late4 = page.locator("#pb-items [data-standard=tn-late-charge]");
  check(await late4.locator(".cc-badge").textContent() === "Off" && /four percent \(4%\)/.test(await late4.textContent()), "your own number (4%) drives the check and the suggested redline");
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem("clauseCraft.standards.v1") || "{}"));
  check(stored["tn-late-charge"] && stored["tn-late-charge"].value === 4, "standards are saved in this browser");
  await page.locator("#btn-pb-edit").click();
  await page.locator("#btn-pb-reset").click();
  check(await page.locator("#pb-items [data-standard=tn-late-charge] .cc-badge").textContent() === "OK", "reset to defaults");
  await page.locator("#pb-items [data-standard=tn-renewal]").locator("button", { hasText: "Insert clause" }).click();
  s = await docState(page);
  check(s.current[s.count - 1].startsWith("Option to Extend.") && s.inss[s.count - 1].length === 1, "missing standard clause added at the end as an insertion");

  // Draft with your numbers, inserted after the paragraph the cursor is in (paragraph 4)
  await selectInEditor(page, 4, 5);
  await page.locator(".sidebar-tabs .tab[data-tab=draft]").click();
  const draft = page.locator('[data-draft="renewal_option"]');
  await draft.locator("button", { hasText: "Customize" }).click();
  await draft.locator('input[data-key="options"]').fill("2");
  await draft.locator('input[data-key="years"]').fill("3");
  check(/two \(2\) options to extend the Term for three \(3\) years each/.test(await draft.locator(".draft-preview").textContent()), "draft preview updates as you type");
  const beforeDraft = await docState(page);
  await draft.locator("button", { hasText: "Insert at cursor" }).click();
  s = await docState(page);
  check(s.count === beforeDraft.count + 1 && /Tenant shall have two \(2\) options to extend the Term for three \(3\) years each/.test(s.current[5]) && s.inss[5].length === 1 && s.current[4] === beforeDraft.current[4],
    "drafted clause inserted right after the cursor's paragraph, as an insertion");

  // My clauses: save the estoppel paragraph, insert it back
  const est = s.current.indexOf(P[18]);
  await selectInEditor(page, est, 0, P[18].length);
  await page.locator("#my-clause-title").fill("Estoppel - their form");
  await selectInEditor(page, est, 0, P[18].length);
  await page.locator("#btn-save-clause").click();
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem("clauseCraft.myClauses.v1") || "[]"));
  check(saved.length === 1 && saved[0].text === P[18] && saved[0].title === "Estoppel - their form", "selected text saved to My clauses in this browser");
  const beforeMine = (await docState(page)).count;
  await page.locator("#my-clauses li", { hasText: "Estoppel - their form" }).locator("button", { hasText: "Insert" }).click();
  s = await docState(page);
  check(s.count === beforeMine + 1 && s.current[est + 1] === P[18] && s.inss[est + 1].length === 1, "saved clause inserted back into the document");

  // Explain the selected text
  await page.locator(".sidebar-tabs .tab[data-tab=ask]").click();
  const assignIdx = s.current.indexOf(P[12]);
  await selectInEditor(page, assignIdx, 0, 40);
  await page.locator("#btn-explain").click();
  const expl = await page.locator("#ask-chat .cc-msg.a").last().textContent();
  check(/Assignment & subletting/.test(expl) && /Flagged here/.test(expl), "Explain describes the clause in plain English and lists what's flagged");

  // Issues list with deal notes, then a re-review ignores the inserted list
  await page.locator(".sidebar-tabs .tab[data-tab=review]").click();
  await page.locator("#btn-run-review").click();
  const missingBefore = await page.locator("#review-results .cc-card[data-kind=missing]").count();
  await page.locator("#btn-setup-edit").click();
  await page.locator("#deal-notes").fill("Tenant wants keys by Dec 1.");
  await page.locator(".sidebar-tabs .tab[data-tab=summary]").click();
  await page.locator("#btn-issues").click();
  const il = Engine.issuesList((await docState(page)).current, { perspective: "tenant", notes: "Tenant wants keys by Dec 1." });
  check(await page.locator("#issues li").count() === il.items.length, `issues list built (${il.items.length} items)`);
  await page.locator("#btn-issues-copy").click();
  check((await page.evaluate(() => navigator.clipboard.readText())) === il.text, "issues list copied exactly");
  await page.locator("#btn-issues-end").click();
  s = await docState(page);
  const head = s.current.indexOf("TENANT'S COMMENTS ON THE LEASE");
  check(head !== -1 && s.current[head + 2] === "Deal notes: Tenant wants keys by Dec 1." && s.current[s.count - 1] === "End of comments.", "issues list added at the end with your deal notes");
  await page.locator(".sidebar-tabs .tab[data-tab=review]").click();
  await page.locator("#btn-run-review").click();
  check(await page.locator("#review-results .cc-card[data-kind=missing]").count() === missingBefore && missingBefore > 0, "re-review ignores the inserted issues list (missing clauses unchanged)");
  await page.reload();
  await page.locator("#editor p").first().waitFor();
  check(await page.locator("#deal-notes").inputValue() === "Tenant wants keys by Dec 1.", "deal notes are kept with the document after a reload");

  // Review type, jurisdiction and playbook choice
  await page.locator("#btn-run-review").click();
  await page.locator("#btn-setup-edit").click();
  await page.locator("#review-type").selectOption("quick");
  const quick = Engine.review((await docState(page)).current, { perspective: "tenant", reviewType: "quick" });
  check(await page.locator("#review-results .cc-card[data-kind=finding]").count() === quick.findings.length && (await page.locator("#review-results .cc-card .cc-badge").allTextContents()).every((t) => t === "High"), "Quick review shows only high risks");
  check(await page.locator("#setup-fields").isVisible(), "setup stays open while you change it");
  await page.locator("#review-type").selectOption("comprehensive");
  await page.locator("#jurisdiction").selectOption("other");
  check(await page.locator("#review-results .cc-card[data-kind=missing][data-category=casp]").count() === 0 && /Another U\.S\. state/.test(await page.locator("#setup-summary-text").textContent()), "another state: California-only checks drop out");
  await page.locator("#playbook-choice").selectOption("none");
  await page.locator(".sidebar-tabs .tab[data-tab=playbook]").click();
  check(/Playbook is off/.test(await page.locator("#pb-side").textContent()) && await page.locator("#pb-items .cc-card").count() === 0, "playbook can be switched off in the setup");
  await page.locator(".sidebar-tabs .tab[data-tab=review]").click();
  await page.locator("#jurisdiction").selectOption("auto");
  await page.locator("#playbook-choice").selectOption("mine");
  await page.locator("#btn-run-review-side").click();
  check(await page.locator("#setup-fields").isHidden(), "Start review folds the setup away again");

  // Accept all, then download a Word file with the redlines as tracked changes
  await page.locator("#btn-accept-all").click();
  s = await docState(page);
  check(Engine.review(s.current, { perspective: "tenant" }).findings.filter((f) => f.action).length === 0, "Accept all applies every suggested redline");
  await page.locator(".desk").evaluate((el) => { el.scrollTop = 0; });
  await hideToast(page);
  await page.screenshot({ path: join(SHOTS, "web-05-workflow.png") });
  const [dl] = await Promise.all([page.waitForEvent("download"), page.locator("#btn-docx").click()]);
  check(dl.suggestedFilename() === "office-lease-redline.docx", "Download .docx names the file after the document");
  const docxPath = join(SHOTS, "web-redline.docx");
  await dl.saveAs(docxPath);
  let parts = null;
  try { parts = readZip(readFileSync(docxPath)); } catch (e) { parts = null; }
  check(!!parts && ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/styles.xml", "word/settings.xml", "word/_rels/document.xml.rels"].every((n) => parts[n]), "the .docx is a valid package (all parts present, CRCs check)");
  const xml = parts ? parts["word/document.xml"] : "";
  check(/<w:del w:id="\d+" w:author="Clause Craft" w:date="[^"]+"><w:r><w:delText xml:space="preserve">sole and absolute discretion<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+" w:author="Clause Craft"[^>]*><w:r><w:t xml:space="preserve">reasonable discretion/.test(xml), "redlines become Word tracked changes (deletion + insertion by Clause Craft)");
  check(/<w:trackRevisions\/>/.test(parts ? parts["word/settings.xml"] : ""), "track changes stays on for edits made in Word");
  const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
  check(ids.length > 20 && new Set(ids).size === ids.length, `each tracked change has its own id (${ids.length})`);
  // Independent readers: python-docx and LibreOffice must open it.
  let pyCount = -1;
  try { pyCount = Number(execFileSync("python3", ["-I", "-c", "import docx,sys; print(len(docx.Document(sys.argv[1]).paragraphs))", docxPath], { encoding: "utf8" }).trim()); } catch (e) { pyCount = -1; }
  check(pyCount === s.count, `python-docx opens it (${pyCount} paragraphs)`);
  let loText = "";
  const loDir = mkdtempSync(join(tmpdir(), "cc-lo-"));
  try {
    execFileSync("soffice", ["--headless", "-env:UserInstallation=file://" + join(loDir, "profile"), "--convert-to", "txt:Text", "--outdir", loDir, docxPath], { stdio: "ignore", timeout: 120000 });
    loText = readFileSync(join(loDir, "web-redline.txt"), "utf8");
  } catch (e) { loText = ""; }
  rmSync(loDir, { recursive: true, force: true });
  check(/OFFICE LEASE/.test(loText) && /reasonable discretion/.test(loText), "LibreOffice opens it and reads the text");
  // Round trip: the app reads its own .docx back as the redlined text (deletions dropped, insertions kept).
  const expectedBack = s.current.map((t) => t.trim());
  while (expectedBack.length && !expectedBack[expectedBack.length - 1]) expectedBack.pop();
  await toastAfter(page, () => page.locator("#file-input").setInputFiles(docxPath));
  const back = await docState(page);
  check(back.current.join("\n") === expectedBack.join("\n"), "re-opening the downloaded .docx gives back the redlined contract text");
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
  for (const tab of ["ask", "draft", "playbook", "summary", "review"]) {
    await page.locator(`.sidebar-tabs .tab[data-tab=${tab}]`).click();
    if (tab === "draft") await page.locator("[data-draft] button", { hasText: "Customize" }).first().click();
    if (tab === "playbook") await page.locator("#btn-pb-edit").click();
    if (tab === "summary") { await page.locator("#btn-summary").click(); await page.locator("#btn-issues").click(); }
    check(await noHorizontalScroll(page), `phone: no sideways scrolling on the ${tab} tab`);
  }
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

/* ------------------------------------------------------------------ */
console.log("W6. Opened straight from the folder (double-click index.html, no web server)");
{
  const { context, page } = await open(undefined, {}, pathToFileURL(join(ROOT, "index.html")).href);
  let s = await docState(page);
  check(s.count === Demo.lease.paragraphs.length, "file://: demo lease loads");
  await page.locator("#btn-run-review").click();
  check(await page.locator("#review-results .cc-card").count() === leaseTenant.findings.length + leaseTenant.missing.length, "file://: review runs with every finding");
  await card(page, titleOf("assign-sole-discretion")).locator("button", { hasText: "Accept" }).click();
  check((await docState(page)).dels[12].includes("sole and absolute discretion"), "file://: redlines apply");
  await page.locator(".sidebar-tabs .tab[data-tab=playbook]").click();
  check(await page.locator("#pb-items .cc-card").count() > 0, "file://: playbook runs");
  await page.locator(".sidebar-tabs .tab[data-tab=summary]").click();
  await page.locator("#btn-summary").click();
  check(await page.locator("#summary-output .cc-kv tr").count() >= 8, "file://: summary builds");
  const [dl] = await Promise.all([page.waitForEvent("download"), page.locator("#btn-docx").click()]);
  check(/\.docx$/.test(dl.suggestedFilename()), "file://: Word download works");
  await page.reload();
  await page.locator("#editor p").first().waitFor();
  check((await docState(page)).dels[12].includes("sole and absolute discretion"), "file://: work is kept after closing and reopening the page");
  await context.close();
}

const appErrors = errors.filter((e) => !/net::ERR_FAILED|ERR_NAME_NOT_RESOLVED|fonts\.g/.test(e));
check(appErrors.length === 0, "no JavaScript errors (" + appErrors.join(" | ") + ")");
check(offsite.length === 0, "no requests leave the page except Google Fonts (" + offsite.join(", ") + ")");
await browser.close();
server.close();
console.log(`\n${passed} passed, ${failed} failed. Screenshots: ${SHOTS}`);
process.exit(failed ? 1 : 0);
