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
  await pane.locator("#btn-setup-edit").click();
  await pane.locator("#btn-run").click();
  await page.waitForTimeout(300);
  const leftover = await pane.locator("#findings .card button", { hasText: /Apply redline|Add language/ }).count();
  check(leftover === 0, "after Redline all + re-run, no fixable issue remains (" + leftover + ")");
  const manual = await pane.locator("#findings .card h4").allTextContents();
  const expectManual = Engine.review((await docState(page)).current, { perspective: "tenant" }).findings.filter((f) => !f.action).length;
  check(manual.length > 0 && manual.length === expectManual, "only manual-review items remain: " + manual.length);
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

  await pane.locator("#tab-summary").click();
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
  await pane.locator("#btn-setup-edit").click();
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
console.log("G. Full workflow: setup, selection, playbook, drafting, saved clauses, explain, issues list");
{
  const { page, pane } = await open("?api=1.4");
  const frame = () => page.frames().find((f) => f.url().includes("taskpane.html"));
  const select = (index, start = 0, end = null) => page.evaluate(([i, a, b]) => {
    const d = window.MockWordHost.doc, p = d.paragraphs[i];
    const len = p.runs.map((r) => r.text).join("").length;
    d.selection = { para: p, start: a, end: b == null ? len : b };
  }, [index, start, end]);

  check(await pane.locator("#qs-review").isVisible() && await pane.locator("#qs-draft").isVisible() && await pane.locator("#qs-playbook").isVisible(), "getting-started shortcuts shown before the first review");
  await pane.locator("[data-demo=lease]").click();
  await pane.locator("#review-body").waitFor();
  check(await pane.locator("#setup-fields").isHidden() && /Tenant/.test(await pane.locator("#setup-summary-text").textContent()), "setup folds into a one-line summary after the review");
  check(/HARBORLINE ADVISORY GROUP, INC\./.test(await pane.locator("#perspective option[value=tenant]").textContent()), "party list shows the tenant's name from the document");
  const checks = await pane.locator("#checks li").allTextContents();
  check(checks.some((t) => /Term dates add up/.test(t)) && checks.some((t) => /Rent matches rate/.test(t)) && checks.some((t) => /Rent, Additional Rent have no definition/.test(t)), "document checks listed (dates, rent math, undefined terms)");

  // Review selection
  await select(12, 0, 30);
  await pane.locator("#btn-review-selection").click();
  await pane.locator("#scope-text").waitFor();
  const inPara12 = Engine.review(Demo.lease.paragraphs, { perspective: "tenant" }).findings.filter((f) => f.paragraphIndex === 12).length;
  check(await pane.locator("#findings .card").count() === inPara12, `Review selection shows only the ${inPara12} issues in the selected paragraph`);
  await pane.locator("#btn-scope-clear").click();
  check(await pane.locator("#findings .card").count() > inPara12, "Show all brings every issue back");

  // Playbook
  await pane.locator("#tab-playbook").click();
  const pb = Engine.playbook(Demo.lease.paragraphs, { perspective: "tenant" });
  check(await pane.locator("#pb-items .card").count() === pb.items.length, `playbook lists ${pb.items.length} standards`);
  check(await pane.locator("#count-playbook").textContent() === String(pb.stats.fail), "Playbook tab shows how many items are off standard");
  const lateCard = pane.locator("#pb-items .card", { has: pane.locator("h4", { hasText: "Late charge" }) });
  await lateCard.locator("button", { hasText: "Apply redline" }).click();
  await pane.locator("#pb-items .card.pb-pass", { has: pane.locator("h4", { hasText: "Late charge" }) }).waitFor({ timeout: 5000 });
  let s = await docState(page);
  check(s.runs[22].some((r) => r.del && r.text === "ten percent (10%)") && s.runs[22].some((r) => r.ins && r.text === "five percent (5%)"), "playbook redline: late charge 10% to your 5%, as a tracked change");
  await pane.locator("#btn-pb-edit").click();
  await pane.locator('[data-std="tn-late-charge"]').fill("4");
  await pane.locator("#btn-pb-save").click();
  const late4 = pane.locator("#pb-items .card.pb-fail", { has: pane.locator("h4", { hasText: "Late charge" }) });
  await late4.waitFor({ timeout: 5000 });
  check(/four percent \(4%\)/.test(await late4.textContent()), "your own number (4%) drives the check and the suggested redline");
  const stored = await frame().evaluate(() => JSON.parse(localStorage.getItem("clauseCraft.standards.v1") || "{}"));
  check(stored["tn-late-charge"] && stored["tn-late-charge"].value === 4, "standards are saved on this computer");
  await pane.locator("#btn-pb-edit").click();
  await pane.locator("#btn-pb-reset").click();
  await pane.locator("#pb-items .card.pb-pass", { has: pane.locator("h4", { hasText: "Late charge" }) }).waitFor({ timeout: 5000 });
  check(true, "reset to defaults");
  const renew = pane.locator("#pb-items .card", { has: pane.locator("h4", { hasText: "Option to extend" }) });
  await renew.locator("button", { hasText: "Insert clause" }).click();
  await page.waitForTimeout(300);
  s = await docState(page);
  check(s.current[s.count - 1].startsWith("Option to Extend.") && s.tracked[s.count - 1], "missing standard clause inserted as a tracked paragraph");

  // Draft with your numbers
  await pane.locator("#tab-draft").click();
  const draft = pane.locator('[data-draft="renewal_option"]');
  await draft.locator("button", { hasText: "Customize" }).click();
  await draft.locator('input[data-key="options"]').fill("2");
  await draft.locator('input[data-key="years"]').fill("3");
  check(/two \(2\) options to extend the Term for three \(3\) years each/.test(await draft.locator(".draft-preview").textContent()), "draft preview updates as you type");
  await draft.locator("button", { hasText: "Insert at cursor" }).click();
  await page.waitForTimeout(300);
  s = await docState(page);
  const drafted = s.current.findIndex((t) => /Tenant shall have two \(2\) options to extend the Term for three \(3\) years each/.test(t));
  check(drafted !== -1 && s.tracked[drafted], "drafted clause inserted at the cursor as a tracked change");

  // My clauses (the drafted clause went in after the selected paragraph, so find the estoppel paragraph by its text)
  await select((await docState(page)).current.indexOf(Demo.lease.paragraphs[18]));
  await pane.locator("#my-clause-title").fill("Estoppel - their form");
  await pane.locator("#btn-save-clause").click();
  await pane.locator("#my-clauses li", { hasText: "Estoppel - their form" }).waitFor();
  const saved = await frame().evaluate(() => JSON.parse(localStorage.getItem("clauseCraft.myClauses.v1") || "[]"));
  check(saved.length === 1 && saved[0].text === Demo.lease.paragraphs[18], "selected text saved to My clauses on this computer");
  const before = (await docState(page)).count;
  await pane.locator("#my-clauses li", { hasText: "Estoppel - their form" }).locator("button", { hasText: "Insert" }).click();
  await page.waitForTimeout(300);
  check((await docState(page)).count === before + 1, "saved clause inserted back into the document");

  // Explain selection
  await pane.locator("#tab-ask").click();
  await select(12, 0, 40);
  await pane.locator("#btn-explain").click();
  const ans = pane.locator(".msg.a").last();
  await ans.waitFor();
  check(/Assignment & subletting/.test(await ans.textContent()) && /Flagged here/.test(await ans.textContent()), "Explain describes the clause in plain English and lists what's flagged");

  // Issues list with deal notes
  await pane.locator("#tab-review").click();
  await pane.locator("#btn-setup-edit").click();
  await pane.locator("#deal-notes").fill("Tenant wants keys by Dec 1.");
  await pane.locator("#tab-summary").click();
  await pane.locator("#btn-issues").click();
  await pane.locator("#issues li").first().waitFor();
  const expectedIssues = Engine.issuesList((await docState(page)).current, { perspective: "tenant" }).items.length;
  check(await pane.locator("#issues li").count() === expectedIssues, `issues list built (${expectedIssues} items)`);
  await pane.locator("#btn-issues-end").click();
  await page.waitForTimeout(300);
  s = await docState(page);
  const head = s.current.indexOf("TENANT'S COMMENTS ON THE LEASE");
  check(head !== -1 && s.bold[head] && !s.tracked[head] && s.current[head + 2] === "Deal notes: Tenant wants keys by Dec 1.", "issues list inserted at the end with your deal notes");

  // Review type, jurisdiction and playbook choice
  await pane.locator("#tab-review").click();
  await pane.locator("#review-type").selectOption("quick");
  await page.waitForTimeout(300);
  const quick = Engine.review((await docState(page)).current, { perspective: "tenant", reviewType: "quick" });
  check(await pane.locator("#findings .card").count() === quick.findings.length && (await pane.locator("#findings .card .badge").allTextContents()).every((t) => t === "High"), "Quick review shows only high risks");
  check(await pane.locator("#setup-fields").isVisible(), "setup stays open while you change it (results update underneath)");
  await pane.locator("#review-type").selectOption("comprehensive");
  await page.waitForTimeout(200);
  await pane.locator("#jurisdiction").selectOption("other");
  await page.waitForTimeout(300);
  check(await pane.locator("#missing .card h4", { hasText: "CASp" }).count() === 0 && /Another U\.S\. state/.test(await pane.locator("#setup-summary-text").textContent()), "another state: California-only checks drop out");
  await pane.locator("#playbook-choice").selectOption("none");
  await page.waitForTimeout(300);
  await pane.locator("#tab-playbook").click();
  check(/Playbook is off/.test(await pane.locator("#pb-side").textContent()) && await pane.locator("#pb-items .card").count() === 0, "playbook can be switched off in the setup");

  // Named playbooks: save the editor as a new playbook, switch between playbooks, delete it
  await pane.locator("#btn-pb-edit").click();
  await pane.locator('[data-std="tn-late-charge"]').fill("6");
  await pane.locator("#pb-new-name").fill("Credit tenant deals");
  await pane.locator("#btn-pb-saveas").click();
  await pane.locator("#pb-side", { hasText: "Credit tenant deals" }).waitFor({ timeout: 5000 });
  let pbStore = await frame().evaluate(() => JSON.parse(localStorage.getItem("clauseCraft.playbooks.v1") || "[]"));
  check(pbStore.length === 1 && pbStore[0].name === "Credit tenant deals" && pbStore[0].standards["tn-late-charge"].value === 6, "Save as new playbook keeps your numbers under that name");
  check(await pane.locator("#playbook-choice").inputValue() === "pb:" + pbStore[0].id && (await pane.locator("#playbook-choice option").allTextContents()).join("|") === "My standards|Credit tenant deals|None", "the new playbook is offered in the setup and selected");
  check(/6%/.test(await pane.locator("#pb-items .card", { has: pane.locator("h4", { hasText: "Late charge" }) }).textContent()), "the checks use that playbook's numbers");
  await pane.locator("#tab-review").click();
  await pane.locator("#playbook-choice").selectOption("mine");
  await page.waitForTimeout(300);
  await pane.locator("#tab-playbook").click();
  check(/^My standards/.test(await pane.locator("#pb-side").textContent()), "switch back to My standards from the setup");
  await pane.locator("#tab-review").click();
  await pane.locator("#playbook-choice").selectOption("pb:" + pbStore[0].id);
  await page.waitForTimeout(300);
  await pane.locator("#tab-playbook").click();
  await pane.locator("#btn-pb-edit").click();
  await pane.locator("#btn-pb-delete").click();
  await page.waitForTimeout(300);
  pbStore = await frame().evaluate(() => JSON.parse(localStorage.getItem("clauseCraft.playbooks.v1") || "[]"));
  check(pbStore.length === 0 && (await pane.locator("#playbook-choice option").allTextContents()).join("|") === "My standards|None" && /^My standards/.test(await pane.locator("#pb-side").textContent()), "delete a playbook and fall back to My standards");

  // Parties: refresh from the document, add a name yourself, clear it
  await pane.locator("#tab-review").click();
  await pane.locator("#btn-parties-refresh").click();
  check(/Parties found: .*Landlord: PACIFIC MERIDIAN OWNER, LLC.*|Parties found: .*Tenant: HARBORLINE ADVISORY GROUP, INC\./.test((await toastText(pane)).text), "Refresh parties reads the names from the document");
  await pane.locator("#btn-party-add").click();
  await pane.locator("#party-add-role").selectOption("Landlord");
  await pane.locator("#party-add-name").fill("Pacific Meridian (owner)");
  await pane.locator("#btn-party-save").click();
  check(await pane.locator("#perspective option[value=landlord]").textContent() === "Landlord: Pacific Meridian (owner)" && await pane.locator("#party-add").isHidden(), "Add party names a side yourself");
  const savedSetup = await frame().evaluate(() => Office.context.document.settings.get("clauseCraft.setup"));
  check(savedSetup && savedSetup.parties && savedSetup.parties.Landlord === "Pacific Meridian (owner)", "the added name is kept with the document");
  await pane.locator("#btn-party-add").click();
  await pane.locator("#party-add-role").selectOption("Landlord");
  await pane.locator("#btn-party-save").click();
  check(await pane.locator("#perspective option[value=landlord]").textContent() === "Landlord: PACIFIC MERIDIAN OWNER, LLC", "clearing the name goes back to the one in the document");

  await pane.locator("#tab-review").click();
  await pane.locator("#btn-run").click();
  await page.waitForTimeout(300);
  check(await pane.locator("#setup-fields").isHidden() && await pane.locator("#setup-summary").isVisible(), "Start review folds the setup away again");
  await page.locator(".canvas").evaluate((el) => { el.scrollTop = 0; });
  await page.screenshot({ path: join(SHOTS, "word-04-setup.png") });
  await page.close();
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
  await page.evaluate(() => { document.querySelector(".main").style.gridTemplateColumns = "1fr 220px 320px"; });
  await page.waitForTimeout(200);
  for (const tab of ["review", "ask", "draft", "playbook", "summary"]) {
    await pane.locator("#tab-" + tab).click();
    const o = await frame.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
    check(o <= 0, "no horizontal scroll in a 320px task pane (" + tab + " tab)");
  }
  await page.setViewportSize({ width: 1180, height: 900 });
  await page.close();
}

const appErrors = errors.filter((e) => !/net::ERR_FAILED|ERR_NAME_NOT_RESOLVED|fonts\.g/.test(e));
check(appErrors.length === 0, "no JavaScript errors (" + appErrors.join(" | ") + ")");
await browser.close();
server.close();
console.log(`\n${passed} passed, ${failed} failed. Screenshots: ${SHOTS}`);
process.exit(failed ? 1 : 0);
