// End-to-end tests for the Word task pane against the Office.js simulator (tests/office-mock.js).
// Usage: npm install && npm run test:ui     (set CHROMIUM_PATH to use a specific Chromium build)
// Screenshots are written to SHOTS_DIR (default tests/.shots).
import { chromium } from "playwright-core";
import { createServer } from "node:http";
import { readFileSync, mkdirSync, existsSync, statSync } from "node:fs";
import { join, extname, dirname, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const SHOTS = process.env.SHOTS_DIR || join(ROOT, "tests", ".shots");
mkdirSync(SHOTS, { recursive: true });
const Engine = require(join(ROOT, "engine", "clause-engine.js"));
const Demo = require(join(ROOT, "engine", "demo-docs.js"));
const MOCK = readFileSync(join(ROOT, "tests", "office-mock.js"), "utf8");

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".xml": "application/xml" };
const server = createServer((req, res) => {
  const path = normalize(decodeURIComponent(new URL(req.url, "http://x").pathname)).replace(/^(\.\.[/\\])+/, "");
  const file = join(ROOT, path);
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
function check(cond, label) {
  if (cond) { passed++; console.log("  ok   " + label); }
  else { failed++; console.log("  FAIL " + label); }
}

async function open(query = "") {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  await page.route(/^https:\/\/(officeapis\.public\.onecdn\.static\.microsoft|appsforoffice\.microsoft\.com)\//, (route) => route.fulfill({ status: 200, contentType: "text/javascript", body: MOCK }));
  await page.route(/fonts\.(googleapis|gstatic)\.com/, (route) => route.abort());
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error" && !/Failed to load resource/.test(m.text())) errors.push("console: " + m.text()); });
  await page.goto(`${BASE}/tests/harness.html${query}`);
  const pane = page.frameLocator("#pane");
  await pane.locator("#btn-run").waitFor();
  return { page, pane };
}

const docState = (page) => page.evaluate(() => {
  const d = window.MockWordHost.doc;
  const cur = (p) => p.runs.filter((r) => !r.del).map((r) => r.text).join("");
  return {
    count: d.paragraphs.length,
    current: d.paragraphs.map(cur),
    runs: d.paragraphs.map((p) => p.runs.map((r) => ({ text: r.text, ins: r.ins, del: r.del, fmt: r.fmt }))),
    tracked: d.paragraphs.map((p) => !!p.tracked),
    align: d.paragraphs.map((p) => p.align),
    bold: d.paragraphs.map((p) => p.bold),
    comments: d.comments.map((c) => ({ text: c.text, anchor: c.anchor, para: d.paragraphs.indexOf(c.para) })),
    selection: d.selection ? { para: d.paragraphs.indexOf(d.selection.para), text: d.selection.para.runs.map((r) => r.text).join("").slice(d.selection.start, d.selection.end) } : null,
    trackMode: d.trackMode
  };
});

const card = (pane, title) => pane.locator("#findings .card", { has: pane.locator("h4", { hasText: title }) }).first();
async function clickAndWait(pane, title, button, statusText) {
  const c = card(pane, title);
  await c.locator("button", { hasText: button }).click();
  if (statusText) await c.locator(".status", { hasText: statusText }).waitFor({ timeout: 5000 });
}
async function toastText(pane) {
  const t = pane.locator("#toast");
  await t.waitFor({ state: "visible", timeout: 5000 });
  return { text: await t.textContent(), cls: await t.getAttribute("class") };
}

/* ------------------------------------------------------------------ */
console.log("A. Microsoft 365 (WordApi 1.4): demo lease, tenant side");
{
  const { page, pane } = await open("?api=1.4");
  await pane.locator("[data-demo=lease]").click();
  await pane.locator("#review-body").waitFor();
  let s = await docState(page);
  const expected = Engine.review(Demo.lease.paragraphs, { perspective: "tenant" });
  check(s.count === Demo.lease.paragraphs.length, `demo inserted (${s.count} paragraphs)`);
  check(s.current.join("\n") === Demo.lease.paragraphs.join("\n"), "inserted text matches the demo exactly");
  check(s.align[0] === "Centered" && s.bold[0] === true, "title centered and bold");
  check(s.runs[2][0].fmt.bold === true && s.runs[2][1].fmt.bold === false, "section heading bold, body not bold");
  check(await pane.locator("#findings .card").count() === expected.findings.length, `all ${expected.findings.length} findings rendered`);
  check(await pane.locator("#missing .card").count() === expected.missing.length, `${expected.missing.length} missing clauses rendered`);
  check((await pane.locator("#doc-chip").textContent()).includes("Lease") && (await pane.locator("#doc-chip").textContent()).includes("Tenant"), "doc chip shows Lease · Tenant");
  await page.screenshot({ path: join(SHOTS, "word-01-review.png") });

  await clickAndWait(pane, "Assignment consent is in Landlord's sole discretion", "Go to");
  s = await docState(page);
  check(s.selection && s.selection.text === "sole and absolute discretion", "Go to selects the exact words");

  await clickAndWait(pane, "Assignment consent is in Landlord's sole discretion", "Apply redline", "Redlined");
  s = await docState(page);
  const p12 = s.runs[12];
  check(p12.some((r) => r.del && r.text === "sole and absolute discretion"), "original words kept as a tracked deletion");
  check(p12.some((r) => r.ins && r.text.startsWith("reasonable discretion (such consent")), "replacement inserted as a tracked insertion");
  check(s.trackMode === "Off", "user's track-changes setting restored afterwards");

  await clickAndWait(pane, "Landlord can relocate you", "Comment", "Commented");
  check(await card(pane, "Landlord can relocate you").locator("button", { hasText: "Add language" }).count() === 1, "after commenting, the redline action is still available");
  s = await docState(page);
  check(s.comments.length === 1 && s.comments[0].anchor === "Landlord may relocate" && s.comments[0].text.startsWith("Clause Craft"), "Word comment anchored on the risky words");

  await clickAndWait(pane, "Rent default with little or no notice", "Apply redline", "Redlined");
  await clickAndWait(pane, "Short cure period for non-monetary defaults", "Apply redline", "Redlined");
  s = await docState(page);
  check(/continues for ten \(10\) days after written notice from Landlord; or \(b\)/.test(s.current[21]) &&
        /such failure continues for thirty \(30\) days after written notice \(or such longer period/.test(s.current[21]),
        "two edits on the same phrase hit the right occurrences");

  await pane.locator("#btn-redline-all").click();
  const t = await toastText(pane);
  await page.waitForTimeout(300);
  check(/redlines added/.test((await pane.locator("#toast").textContent()) || t.text), "Redline all completed");
  await pane.locator("#btn-run").click();
  await page.waitForTimeout(300);
  const leftover = await pane.locator("#findings .card button", { hasText: /Apply redline|Add language/ }).count();
  check(leftover === 0, "after Redline all + re-run, no fixable issue remains (" + leftover + ")");
  const manual = await pane.locator("#findings .card h4").allTextContents();
  check(manual.length > 0 && manual.length <= 10, "only manual-review items remain: " + manual.length);
  s = await docState(page);
  check(s.current[2] === Demo.lease.paragraphs[2], "untouched paragraphs unchanged");
  check(/space of comparable size, quality/.test(s.current[19]), "an issue that was commented on is still redlined by Redline all");
  check(/no consent shall be required for cosmetic alterations/.test(s.current[11]), "alterations consent redlined");

  const casp = pane.locator("#missing .card", { has: pane.locator("h4", { hasText: "CASp" }) });
  await casp.locator("button", { hasText: "Insert at end" }).click();
  await page.waitForTimeout(300);
  s = await docState(page);
  check(s.current[s.count - 1].startsWith("CASp Disclosure.") && s.tracked[s.count - 1], "CASp clause inserted as a tracked paragraph");
  check(await pane.locator("#missing .card h4", { hasText: "CASp" }).count() === 0, "CASp no longer listed as missing");

  await pane.locator("#tab-clauses").click();
  check(await pane.locator("#clause-list li").count() > 20, "clause inventory listed");
  await pane.locator("#clause-list .clause-item", { hasText: "Assignment & subletting" }).click();
  s = await docState(page);
  check(s.selection && s.selection.para === 12, "clicking a clause jumps to it");

  await pane.locator("#tab-ask").click();
  await pane.locator("#ask-chips .chip", { hasText: "Who can terminate?" }).click();
  const ans = pane.locator(".msg.a").last();
  await ans.waitFor();
  const at = await ans.textContent();
  check(/Landlord/.test(at) && /Tenant/.test(at), "termination answer lists Landlord and (after redlines) Tenant rights");
  await ans.locator(".source").first().click();
  s = await docState(page);
  check(s.selection !== null, "citation jumps into the document");
  await pane.locator("#ask-input").fill("What is the rent?");
  await pane.locator("#ask-form button").click();
  check(/\$12,442\.50/.test(await pane.locator(".msg.a").last().textContent()), "rent question answered with the figure");

  await pane.locator("#tab-summary").click();
  await pane.locator("#btn-summary").click();
  check(await pane.locator(".kv tr").count() >= 8, "summary table built");
  const before = (await docState(page)).count;
  await pane.locator("#btn-summary-end").click();
  await page.waitForTimeout(200);
  s = await docState(page);
  check(s.count > before && s.current[before].startsWith("CLAUSE CRAFT SUMMARY") && !s.tracked[before], "summary inserted at the end as plain (untracked) text");

  await pane.locator("#tab-review").click();
  await page.locator(".canvas").evaluate((el) => { el.scrollTop = 520; });
  await page.screenshot({ path: join(SHOTS, "word-02-redlined.png") });
  await page.close();
}

/* ------------------------------------------------------------------ */
console.log("B. Older Word (WordApi 1.3): manual redline formatting, no comments");
{
  const { page, pane } = await open("?api=1.3");
  await pane.locator("[data-demo=lease]").click();
  await pane.locator("#review-body").waitFor();
  check(await pane.locator("#tracking").isDisabled(), "tracked-changes switch disabled");
  check(await pane.locator("#api-note").isVisible(), "explains the limitation");
  await clickAndWait(pane, "Assignment consent is in Landlord's sole discretion", "Apply redline", "Redlined");
  const s = await docState(page);
  const runs = s.runs[12];
  const struck = runs.find((r) => r.text === "sole and absolute discretion");
  const added = runs.find((r) => r.text.includes("reasonable discretion"));
  check(struck && struck.fmt.strikeThrough === true && struck.fmt.color === "#C0392B", "original shown in red strikethrough");
  check(added && added.fmt.strikeThrough === false && added.fmt.underline === "Double", "insertion shown double-underlined, not struck");
  await card(pane, "Landlord can relocate you").locator("button", { hasText: "Comment" }).click();
  check(/Comments need Microsoft 365, Word 2024 or Word on the web/.test((await toastText(pane)).text), "comment falls back with a clear message");
  await page.screenshot({ path: join(SHOTS, "word-03-older-word.png") });
  await page.close();
}

/* ------------------------------------------------------------------ */
console.log("C. Purchase agreement, buyer then seller");
{
  const { page, pane } = await open("?api=1.4");
  await pane.locator("[data-demo=psa]").click();
  await pane.locator("#review-body").waitFor();
  const chip = await pane.locator("#doc-chip").textContent();
  check(chip.includes("Purchase agreement") && chip.includes("Buyer"), "doc chip shows Purchase agreement · Buyer");
  check(await pane.locator("#findings .card").count() === Engine.review(Demo.psa.paragraphs, { perspective: "buyer" }).findings.length, "buyer findings rendered");
  await clickAndWait(pane, "Deposit goes hard immediately", "Add language", "Redlined");
  let s = await docState(page);
  check(s.runs[5].some((r) => r.ins && r.text.includes("fully refundable")), "deposit protection appended as tracked insertion");
  await pane.locator("#perspective").selectOption("seller");
  await page.waitForTimeout(300);
  const sellerExpected = Engine.review(s.current, { perspective: "seller" }).findings.length;
  check(await pane.locator("#findings .card").count() === sellerExpected, "switching to Seller re-runs the review");
  await pane.locator("#tab-summary").click();
  await pane.locator("#btn-summary").click();
  const dep = await pane.locator(".kv tr", { has: pane.locator("th", { hasText: "Deposit" }) }).locator("td").textContent();
  check(dep.includes("$100,000.00"), "summary deposit = $100,000.00");
  await page.close();
}

/* ------------------------------------------------------------------ */
console.log("D. Word builds where text/search exclude tracked deletions");
{
  const { page, pane } = await open("?api=1.4&search=current&ptext=current");
  await pane.locator("[data-demo=lease]").click();
  await pane.locator("#review-body").waitFor();
  await clickAndWait(pane, "Rent default with little or no notice", "Apply redline", "Redlined");
  await clickAndWait(pane, "Short cure period for non-monetary defaults", "Apply redline", "Redlined");
  const s = await docState(page);
  check(/for ten \(10\) days after written notice from Landlord; or \(b\)/.test(s.current[21]) && /for thirty \(30\) days after written notice \(or such/.test(s.current[21]), "correct occurrences edited");
  await page.close();
}

/* ------------------------------------------------------------------ */
console.log("E. Track changes already on / opened outside Word");
{
  const { page, pane } = await open("?api=1.4&track=TrackAll");
  await pane.locator("[data-demo=lease]").click();
  await pane.locator("#review-body").waitFor();
  await clickAndWait(pane, "Landlord can relocate you", "Add language", "Redlined");
  check((await docState(page)).trackMode === "TrackAll", "existing TrackAll mode left on");
  await page.close();
  const o = await open("?host=none");
  check(await o.pane.locator("#btn-run").isDisabled(), "outside Word: Run review disabled");
  check(/Open Clause Craft from Word/.test((await toastText(o.pane)).text), "outside Word: tells the user to open it from Word");
  await o.page.close();
}

/* ------------------------------------------------------------------ */
console.log("F. Task pane at narrow width");
{
  const { page, pane } = await open("?api=1.4");
  await pane.locator("[data-demo=lease]").click();
  await pane.locator("#review-body").waitFor();
  const frame = page.frames().find((f) => f.url().includes("taskpane.html"));
  const overflow = await frame.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
  check(overflow <= 0, "no horizontal scroll in a 380px task pane");
  await page.setViewportSize({ width: 1180, height: 900 });
  await page.close();
}

const appErrors = errors.filter((e) => !/net::ERR_FAILED|ERR_NAME_NOT_RESOLVED|fonts\.g/.test(e));
check(appErrors.length === 0, "no JavaScript errors (" + appErrors.join(" | ") + ")");
await browser.close();
server.close();
console.log(`\n${passed} passed, ${failed} failed. Screenshots: ${SHOTS}`);
process.exit(failed ? 1 : 0);
