/* Clause Craft engine tests — run with: node --test tests/ */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const E = require("../engine/clause-engine.js");
const D = require("../engine/demo-docs.js");

const ids = (list) => list.map((f) => f.ruleId);
const cats = (list) => list.map((m) => m.category);

function applyAllFixes(paragraphs, perspective) {
  const r = E.review(paragraphs, { perspective });
  const out = paragraphs.slice();
  const applied = [];
  for (const f of r.findings) {
    if (!f.action) continue;
    const next = E.applyFixToText(f, out[f.paragraphIndex]);
    if (next == null) continue;
    out[f.paragraphIndex] = next;
    applied.push(f.ruleId + "@" + f.paragraphIndex);
  }
  return { before: r, after: E.review(out, { perspective }), applied, text: out };
}

test("number words parse", () => {
  assert.equal(E.wordsToNumber("sixty-two"), 62);
  assert.equal(E.wordsToNumber("one hundred fifty"), 150);
  assert.equal(E.wordsToNumber("twenty-one"), 21);
  assert.equal(E.wordsToNumber("two hundred"), 200);
  assert.ok(Number.isNaN(E.wordsToNumber("several")));
});

test("document type detection", () => {
  assert.equal(E.review(D.lease.paragraphs).docType, "lease");
  assert.equal(E.review(D.psa.paragraphs).docType, "psa");
  assert.equal(E.review("Call the broker on Tuesday about parking.").docType, "unknown");
  assert.equal(E.review(D.lease.paragraphs).perspective, "tenant");
  assert.equal(E.review(D.psa.paragraphs).perspective, "buyer");
});

test("lease, tenant side: key risks are flagged", () => {
  const r = E.review(D.lease.paragraphs, { perspective: "tenant" });
  for (const id of ["assign-sole-discretion", "assign-recapture", "assign-profit-100", "opex-no-cap", "opex-capital",
    "indemnity-own-negligence", "snda-missing", "relocation", "holdover-rate", "default-no-notice",
    "default-short-nonmonetary", "personal-guaranty", "landlord-lien", "entry-any-time", "late-charge-high",
    "utility-no-abatement", "casualty-no-tenant-termination", "number-mismatch", "ca-jury-waiver", "as-is-delivery",
    "alterations-sole-discretion"]) {
    assert.ok(ids(r.findings).includes(id), "missing finding " + id);
  }
  assert.ok(cats(r.missing).includes("casp"), "CASp disclosure should be missing (California lease)");
  assert.ok(cats(r.missing).includes("waiver_subrogation"));
  assert.ok(!cats(r.missing).includes("snda"), "SNDA is a finding, not missing (subordination clause exists)");
  assert.ok(r.stats.high >= 6);
  // severity order
  const rank = { high: 0, medium: 1, low: 2 };
  for (let i = 1; i < r.findings.length; i++) assert.ok(rank[r.findings[i - 1].severity] <= rank[r.findings[i].severity]);
});

test("lease, landlord side: tenant-only rules are not shown", () => {
  const r = E.review(D.lease.paragraphs, { perspective: "landlord" });
  assert.ok(!ids(r.findings).includes("assign-sole-discretion"));
  assert.ok(!ids(r.findings).includes("relocation"));
  assert.ok(ids(r.findings).includes("number-mismatch"), "drafting checks apply to both sides");
});

test("purchase agreement, buyer side", () => {
  const r = E.review(D.psa.paragraphs, { perspective: "buyer" });
  for (const id of ["psa-deposit-hard", "psa-dd-short", "psa-release-no-carveouts", "psa-buyer-remedy-deposit-only",
    "psa-buyer-no-assign", "psa-title-all-record", "psa-no-lien-cure", "psa-estoppel-not-condition", "psa-reps-short-survival", "number-mismatch"]) {
    assert.ok(ids(r.findings).includes(id), "missing finding " + id);
  }
  assert.ok(cats(r.missing).includes("operating_covenants"));
  assert.ok(cats(r.missing).includes("exchange_1031"));
});

test("purchase agreement, seller side: no false financing contingency", () => {
  const r = E.review(D.psa.paragraphs, { perspective: "seller" });
  assert.ok(ids(r.findings).includes("psa-reps-uncapped"));
  assert.ok(!ids(r.findings).includes("psa-financing-contingency"), "'not contingent upon financing' is not a contingency");
});

test("every finding points at real text that Word can search for", () => {
  for (const doc of D.all) {
    for (const persp of ["neutral", "tenant", "landlord", "buyer", "seller"]) {
      const r = E.review(doc.paragraphs, { perspective: persp });
      for (const f of r.findings) {
        const para = doc.paragraphs[f.paragraphIndex];
        assert.ok(para && para.trim(), f.id + " points at an empty paragraph");
        assert.ok(para.includes(f.matchText), f.id + " matchText not in paragraph");
        assert.ok(f.matchText.length > 0 && f.matchText.length <= 250, f.id + " matchText length " + f.matchText.length);
        if (f.action && f.action.type === "replace") assert.equal(f.action.find, f.matchText);
        assert.ok(f.ref && /^[§¶]/.test(f.ref));
      }
    }
  }
});

test("applying every suggested fix resolves those findings and adds no new ones", () => {
  for (const [doc, persp] of [[D.lease, "tenant"], [D.lease, "neutral"], [D.psa, "buyer"], [D.psa, "seller"]]) {
    const { before, after, applied } = applyAllFixes(doc.paragraphs, persp);
    assert.ok(applied.length > 0);
    const reflagged = after.findings.filter((f) => applied.includes(f.ruleId + "@" + f.paragraphIndex));
    assert.deepEqual(reflagged.map((f) => f.ruleId), [], doc.id + "/" + persp + " fixes did not clear");
    const fresh = after.findings.filter((f) => f.ruleId !== "placeholder" && !before.findings.some((o) => o.ruleId === f.ruleId && o.paragraphIndex === f.paragraphIndex));
    assert.deepEqual(fresh.map((f) => f.ruleId), [], doc.id + "/" + persp + " fixes introduced new findings");
  }
});

test("locate() finds the right occurrence after earlier edits in the same paragraph", () => {
  const r = E.review(D.lease.paragraphs, { perspective: "tenant" });
  const monetary = r.findings.find((f) => f.ruleId === "default-no-notice");
  const nonMonetary = r.findings.find((f) => f.ruleId === "default-short-nonmonetary");
  let text = D.lease.paragraphs[monetary.paragraphIndex];
  text = E.applyFixToText(monetary, text); // creates a second "ten (10) days after written notice"
  const loc = E.locate(nonMonetary, text);
  assert.ok(loc, "non-monetary target still present");
  assert.equal(loc.occurrence, 1, "must target the second occurrence");
  assert.equal(text.indexOf(loc.context), text.lastIndexOf(loc.context), "context must be unique");
  assert.ok(loc.context.length <= 250);
  const fixed = E.applyFixToText(nonMonetary, text);
  assert.match(fixed, /for ten \(10\) days after written notice from Landlord; or \(b\)/);
  assert.match(fixed, /continues for thirty \(30\) days after written notice/);
  assert.equal(E.locate(nonMonetary, fixed), null, "already fixed -> null");
});

test("curly quotes and Word paragraph characters are handled", () => {
  const paras = ["11. Assignment. Tenant shall not assign this Lease without Landlord’s consent, which Landlord may withhold in its sole and absolute discretion.", "Landlord and Tenant (“Landlord”) agree."];
  const r = E.review(paras, { perspective: "tenant", docType: "lease" });
  const f = r.findings.find((x) => x.ruleId === "assign-sole-discretion");
  assert.ok(f);
  assert.ok(paras[0].includes(f.matchText));
});

test("signature and initials blanks are not placeholders, but $[amount] is", () => {
  const r = E.review(D.psa.paragraphs, { perspective: "buyer" });
  assert.ok(!ids(r.findings).includes("placeholder"));
  const r2 = E.review(D.psa.paragraphs.concat(["Buyer may recover costs not to exceed $[amount]."]), { perspective: "buyer" });
  assert.ok(ids(r2.findings).includes("placeholder"));
});

test("Ask answers with cited sections", () => {
  const rent = E.ask("What is the rent?", D.lease.paragraphs);
  assert.ok(rent.found && rent.facts.includes("$12,442.50"));
  assert.match(rent.answer, /^Key terms:/);
  const term = E.ask("Who can terminate?", D.lease.paragraphs);
  assert.match(term.answer, /Landlord/);
  assert.match(term.answer, /Tenant has no termination right/);
  assert.equal(E.ask("Is there a renewal option?", D.lease.paragraphs).found, false);
  assert.match(E.ask("Who are the parties?", D.lease.paragraphs).answer, /HARBORLINE ADVISORY GROUP, INC\./);
  assert.ok(E.ask("Is the deposit refundable?", D.psa.paragraphs).facts.includes("$100,000.00"));
  assert.equal(E.ask("pizza delivery schedule", D.lease.paragraphs).found, false);
  assert.match(E.ask("Summarize risks", D.lease.paragraphs).answer, /^Top issues for the tenant side/);
});

test("summary extracts key terms", () => {
  const s = E.summarize(D.lease.paragraphs);
  const get = (label) => (s.fields.find((f) => f.label === label) || {}).value;
  assert.equal(get("Tenant"), "HARBORLINE ADVISORY GROUP, INC.");
  assert.equal(get("Term"), "sixty-two (62) months");
  assert.match(get("Base rent"), /^\$12,442\.50/);
  assert.equal(get("Security deposit"), "$25,000.00");
  const p = E.summarize(D.psa.paragraphs);
  const getp = (label) => (p.fields.find((f) => f.label === label) || {}).value;
  assert.equal(getp("Deposit"), "$100,000.00");
  assert.equal(getp("Purchase price"), "$4,200,000.00");
  assert.match(p.text, /Not legal advice/);
});

test("robust on empty, string and large input", () => {
  assert.equal(E.review([]).stats.paragraphs, 0);
  assert.equal(E.review("").docType, "unknown");
  assert.equal(E.review(D.lease.paragraphs.join("\r\n")).docType, "lease");
  const big = [];
  for (let i = 0; i < 60; i++) big.push(...D.lease.paragraphs);
  const t0 = Date.now();
  const r = E.review(big, { perspective: "tenant" });
  assert.ok(r.findings.length > 0);
  assert.ok(Date.now() - t0 < 4000, "2,000+ paragraph review took " + (Date.now() - t0) + "ms");
});

test("contract-style numbers", () => {
  assert.equal(E.numberToWords(62), "sixty-two");
  assert.equal(E.numberToWords(150), "one hundred fifty");
  assert.equal(E.numberToWords(1000), "one thousand");
  assert.equal(E.ordinalWords(36), "thirty-sixth (36th)");
  assert.equal(E.ordinalWords(12), "twelfth (12th)");
  assert.equal(E.ordinalWords(22), "twenty-second (22nd)");
  assert.equal(E.countPhrase(1, "months"), "one (1) month");
  assert.equal(E.countPhrase(10, "business days"), "ten (10) business days");
  assert.equal(E.percentPhrase(3), "three percent (3%)");
  assert.equal(E.percentPhrase(3.5), "3.5%");
  for (let n = 0; n < 400; n++) assert.equal(E.wordsToNumber(E.numberToWords(n)), n, "round trip " + n);
});

test("key numbers are read from the demo documents", () => {
  const m = E.metrics(D.lease.paragraphs);
  assert.equal(m.escalationPct.value, 4);
  assert.equal(m.lateChargePct.value, 10);
  assert.equal(m.holdoverPct.value, 200);
  assert.equal(m.monetaryCure.value, 3);
  assert.equal(m.monetaryCure.notice, false);
  assert.equal(m.nonMonetaryCureDays.value, 10);
  assert.equal(m.monthlyRent.value, 12442.5);
  assert.equal(m.rsf.value, 3950);
  assert.equal(m.termMonths.value, 62);
  assert.equal(m.commencement.value, "December 1, 2026");
  assert.equal(m.expiration.value, "January 31, 2032");
  assert.equal(m.estoppelDays.value, 5);
  const p = E.metrics(D.psa.paragraphs);
  assert.equal(p.ddDays.value, 21);
  assert.equal(p.repsSurvivalMonths.value, 3);
  assert.equal(p.purchasePrice.value, 4200000);
  assert.equal(p.earnestDeposit.value, 100000);
  assert.ok(Math.abs(p.depositPct.value - 2.381) < 0.01);
});

test("consistency checks: demo math is right, undefined terms are caught", () => {
  const r = E.review(D.lease.paragraphs, { perspective: "tenant" });
  const status = (id) => (r.checks.find((c) => c.id === id) || {}).status;
  assert.equal(status("term-dates"), "ok");
  assert.equal(status("rent-math"), "ok");
  assert.equal(status("party-names"), "ok");
  const undef = r.findings.filter((f) => f.ruleId === "check-undefined-term").map((f) => f.matchText).sort();
  assert.deepEqual(undef, ["Additional Rent", "Rent"]);
  const p = E.review(D.psa.paragraphs, { perspective: "buyer" });
  assert.deepEqual(p.findings.filter((f) => f.ruleId === "check-undefined-term").map((f) => f.matchText).sort(), ["Close of Escrow", "Escrow Holder", "Opening of Escrow"]);
  // a check finding can be located and quoted like any other
  const f = r.findings.find((x) => x.ruleId === "check-undefined-term");
  const loc = E.locate(f, D.lease.paragraphs[f.paragraphIndex]);
  assert.ok(loc && loc.target === f.matchText && D.lease.paragraphs[f.paragraphIndex].slice(loc.start, loc.start + loc.target.length) === f.matchText);
});

test("consistency checks catch planted drafting errors", () => {
  const paras = D.lease.paragraphs.slice();
  paras[3] = paras[3].replace("January 31, 2032", "January 31, 2031");
  paras[4] = paras[4].replace("$12,442.50", "$12,000.00");
  paras[36] = "TENANT: HARBORLINE ADVISORS, LLC";
  paras[31] = paras[31] + " Notices to Landlord shall also be copied to its lender (the \"Lender\") as provided in Section 45 and Exhibit C.";
  const r = E.review(paras, { perspective: "tenant" });
  const ids = r.findings.map((f) => f.ruleId);
  for (const id of ["check-term-dates", "check-rent-math", "check-party-name", "check-cross-ref", "check-exhibit", "check-defined-unused"]) assert.ok(ids.includes(id), "missing " + id);
  const term = r.findings.find((f) => f.ruleId === "check-term-dates");
  assert.equal(term.severity, "high");
  assert.match(term.why, /ends January 31, 2032/);
  assert.equal(term.matchText, "January 31, 2031");
  assert.match(r.findings.find((f) => f.ruleId === "check-rent-math").why, /\$12,442\.50/);
  assert.match(r.findings.find((f) => f.ruleId === "check-cross-ref").title, /Section 45/);
  // statutes are not treated as cross-references
  assert.ok(!r.findings.some((f) => f.ruleId === "check-cross-ref" && /1950|1932|1542/.test(f.matchText)));
});

test("playbook: tenant standards on the demo lease, with redlines to your numbers", () => {
  const pb = E.playbook(D.lease.paragraphs, { perspective: "tenant" });
  const item = (id) => pb.items.find((i) => i.standardId === id);
  assert.equal(item("tn-late-charge").status, "fail");
  assert.equal(item("tn-late-charge").action.find, "ten percent (10%)");
  assert.equal(item("tn-late-charge").action.replace, "five percent (5%)");
  assert.equal(item("tn-holdover").action.replace, "one hundred fifty percent (150%)");
  assert.equal(item("tn-monetary-cure").actualText, "3 days, no notice");
  assert.equal(item("tn-monetary-cure").action.find, "three (3) days, without notice");
  assert.equal(item("tn-deposit").status, "pass");
  assert.equal(item("tn-renewal").templateId, "renewal_option");
  assert.ok(pb.stats.fail >= 10);
  // your own numbers
  const mine = E.playbook(D.lease.paragraphs, { perspective: "tenant", standards: { "tn-late-charge": { value: 4 }, "tn-holdover": { off: true } } });
  assert.equal(mine.items.find((i) => i.standardId === "tn-late-charge").action.replace, "four percent (4%)");
  assert.ok(!mine.items.some((i) => i.standardId === "tn-holdover"));
  // applying every suggested redline makes those standards pass
  const text = D.lease.paragraphs.slice();
  for (const it of pb.items) {
    if (!it.action) continue;
    const next = E.applyFixToText(it, text[it.paragraphIndex]);
    if (next != null) text[it.paragraphIndex] = next;
  }
  const after = E.playbook(text, { perspective: "tenant" });
  const stillFailing = after.items.filter((i) => i.status === "fail" && pb.items.find((b) => b.standardId === i.standardId).action).map((i) => i.standardId);
  assert.deepEqual(stillFailing, []);
  assert.match(text[21], /continues for ten \(10\) days after written notice; or \(b\)/);
});

test("playbook: buyer and seller sides of the demo PSA", () => {
  const buyer = E.playbook(D.psa.paragraphs, { perspective: "buyer" });
  const dd = buyer.items.find((i) => i.standardId === "by-dd");
  assert.equal(dd.status, "fail");
  assert.equal(dd.action.replace, "thirty (30) days");
  assert.equal(buyer.items.find((i) => i.standardId === "by-refundable").status, "fail");
  assert.equal(buyer.items.find((i) => i.standardId === "by-deposit-pct").status, "pass");
  const seller = E.playbook(D.psa.paragraphs, { perspective: "seller" });
  assert.equal(seller.stats.fail, 0);
});

test("issues list for the other side", () => {
  const il = E.issuesList(D.lease.paragraphs, { perspective: "tenant" });
  assert.equal(il.heading, "Tenant's comments on the Lease");
  assert.ok(il.text.startsWith("TENANT'S COMMENTS ON THE LEASE"));
  assert.match(il.text, /\d+\. §11 Assignment and Subletting — Assignment consent is in Landlord's sole discretion\. Change “sole and absolute discretion” to “reasonable discretion/);
  assert.match(il.text, /Add a casp disclosure provision|Add a CASp disclosure provision/i);
  assert.deepEqual(il.items.map((i) => i.n), il.items.map((_, k) => k + 1));
});

test("an inserted summary or issues list is not read as contract text on the next review", () => {
  for (const [name, persp] of [["lease", "tenant"], ["psa", "buyer"]]) {
    const base = D[name].paragraphs;
    const before = E.review(base, { perspective: persp });
    const summary = E.summarize(base, { perspective: persp }).text.split("\n");
    const issues = E.issuesList(base, { perspective: persp, notes: "Close by Dec 1." }).text.split("\n");
    // Inserted at the end, and a summary inserted mid-document, as Word and the browser app do.
    const mid = Math.floor(base.length / 2);
    const variants = [
      base.concat([""], issues),
      base.concat([""], summary, [""], issues),
      base.slice(0, mid).concat(summary, base.slice(mid)),
    ];
    for (const doc of variants) {
      const after = E.review(doc, { perspective: persp });
      assert.deepEqual(cats(after.missing), cats(before.missing), name + ": missing clauses unchanged");
      assert.deepEqual(after.findings.map((f) => f.ruleId).sort(), before.findings.map((f) => f.ruleId).sort(), name + ": findings unchanged");
      const pb1 = E.playbook(base, { perspective: persp }).stats, pb2 = E.playbook(doc, { perspective: persp }).stats;
      assert.deepEqual(pb2, pb1, name + ": playbook unchanged");
    }
  }
  // Memo paragraphs don't count as document text...
  const base = D.lease.paragraphs;
  const n0 = E.review(base, { perspective: "tenant" }).stats.paragraphs;
  const withList = base.concat(E.issuesList(base, { perspective: "tenant" }).text.split("\n"));
  assert.equal(E.review(withList, { perspective: "tenant" }).stats.paragraphs, n0);
  // ...but a real contract section that happens to be titled "comments" is still read.
  const real = base.concat(["TENANT'S COMMENTS ON THE LEASE", "Tenant shall deliver comments within ten (10) days."]);
  assert.equal(E.review(real, { perspective: "tenant" }).stats.paragraphs, n0 + 2);
  // If someone deletes a summary's closing line, the contract after it is still read.
  const summary = E.summarize(base, { perspective: "tenant" }).text.split("\n").slice(0, -1);
  const mid = Math.floor(base.length / 2);
  const cut = base.slice(0, mid).concat(summary, base.slice(mid));
  assert.equal(E.review(cut, { perspective: "tenant" }).stats.paragraphs, n0);
  assert.deepEqual(cats(E.review(cut, { perspective: "tenant" }).missing), cats(E.review(base, { perspective: "tenant" }).missing));
});

test("drafting with your numbers", () => {
  const r = E.draftClause("renewal_option", { options: 2, years: 3, noticeMin: 6, noticeMax: 12, fmv: 100 });
  assert.match(r.text, /two \(2\) options to extend the Term for three \(3\) years each/);
  assert.match(r.text, /one hundred percent \(100%\) of the fair market rent/);
  assert.match(E.draftClause("early_termination", { month: 36 }).text, /thirty-sixth \(36th\) full calendar month/);
  assert.match(E.draftClause("ti_allowance", { perSf: 35 }, D.lease.paragraphs).text, /\$35\.00 per rentable square foot of the Premises \(\$138,250\.00 for 3,950 rentable square feet\)/);
  assert.match(E.draftClause("opex_cap", { pct: 4 }).text, /four percent \(4%\)/);
  assert.match(E.draftClause("rent_abatement", { months: 99 }).text, /twenty-four \(24\) full calendar months/, "values are clamped to the field range");
  assert.equal(E.draftClause("nope"), null);
});

test("explain a clause in plain English", () => {
  const x = E.explain(D.lease.paragraphs[12], D.lease.paragraphs, { perspective: "tenant" });
  assert.ok(x.found);
  assert.deepEqual(x.paragraphIndexes, [12]);
  assert.ok(x.categories.some((c) => c.id === "assignment"));
  assert.match(x.answer, /^§11 Assignment & subletting\./);
  assert.ok(x.findings.some((f) => f.ruleId === "assign-sole-discretion"));
  const part = E.explain("Landlord may elect to recapture the Premises", D.lease.paragraphs, {});
  assert.deepEqual(part.paragraphIndexes, [12]);
  assert.equal(E.explain("Call me tomorrow about lunch.", D.lease.paragraphs, {}).found, false);
  assert.equal(E.explain("   ", D.lease.paragraphs).found, false);
});

test("templates referenced by missing checks exist", () => {
  for (const type of Object.keys(E.EXPECTED)) {
    for (const e of E.EXPECTED[type]) if (e.template) assert.ok(E.TEMPLATES[e.template], "template " + e.template);
  }
});
