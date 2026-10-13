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

test("templates referenced by missing checks exist", () => {
  for (const type of Object.keys(E.EXPECTED)) {
    for (const e of E.EXPECTED[type]) if (e.template) assert.ok(E.TEMPLATES[e.template], "template " + e.template);
  }
});
