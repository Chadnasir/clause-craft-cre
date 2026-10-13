/**
 * Clause Craft — CRE contract review engine (rule-based, runs entirely on the device).
 *
 * Shared by the browser app (index.html + app.js) and the Word add-in (word/taskpane.html).
 * Input is a list of paragraphs (Word body paragraphs) or a plain-text string.
 * Nothing here makes network calls, stores data, or uses eval.
 *
 * Demo only — not legal advice. Findings are heuristics for a broker's first pass;
 * have counsel review before signing.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.ClauseCraftEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var VERSION = "1.0.0";

  /* ------------------------------------------------------------------ *
   * Text utilities
   * ------------------------------------------------------------------ */

  // Same-length normalisation so match offsets map 1:1 back to the original text.
  function norm(s) {
    return String(s == null ? "" : s)
      .replace(/[‘’ʼ′]/g, "'")
      .replace(/[“”″]/g, "\"")
      .replace(/[   ]/g, " ")
      .replace(/[‐‑‒–—]/g, "-")
      .replace(/[\u000B\u0007]/g, " ");
  }

  function trim(s) { return String(s == null ? "" : s).replace(/^\s+|\s+$/g, ""); }

  function clip(s, n) {
    s = trim(String(s == null ? "" : s).replace(/\s+/g, " "));
    return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, "") + "…" : s;
  }

  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

  function uniq(arr) {
    var out = [], seen = {};
    for (var i = 0; i < arr.length; i++) {
      var k = String(arr[i]);
      if (!seen[k]) { seen[k] = 1; out.push(arr[i]); }
    }
    return out;
  }

  function contains(arr, v) { return arr.indexOf(v) !== -1; }

  var NUM_WORDS = {
    zero: 0, one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9,
    ten: 10, eleven: 11, twelve: 12, thirteen: 13, fourteen: 14, fifteen: 15, sixteen: 16,
    seventeen: 17, eighteen: 18, nineteen: 19, twenty: 20, thirty: 30, forty: 40, fifty: 50,
    sixty: 60, seventy: 70, eighty: 80, ninety: 90
  };

  /** "sixty-two" -> 62, "one hundred fifty" -> 150. NaN when any token is unknown. */
  function wordsToNumber(phrase) {
    var tokens = String(phrase || "").toLowerCase().split(/[\s-]+/);
    var total = 0, current = 0, seen = false;
    for (var i = 0; i < tokens.length; i++) {
      var t = tokens[i];
      if (!t || t === "and") continue;
      if (Object.prototype.hasOwnProperty.call(NUM_WORDS, t)) { current += NUM_WORDS[t]; seen = true; }
      else if (t === "hundred") { current = (current || 1) * 100; seen = true; }
      else if (t === "thousand") { total += (current || 1) * 1000; current = 0; seen = true; }
      else return NaN;
    }
    return seen ? total + current : NaN;
  }

  var NUM_WORD_ALT = "zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|thousand";
  var WORD_NUMBER_RE = "(?:(?:" + NUM_WORD_ALT + ")(?:[\\s-]+(?:and[\\s-]+)?(?:" + NUM_WORD_ALT + "))*)";

  /** Parses "ten (10)", "10", "ten" into a number (digits win when present). */
  function parseCount(s) {
    var d = /\(\s*(\d+(?:\.\d+)?)\s*%?\s*\)/.exec(s) || /(\d+(?:\.\d+)?)/.exec(s);
    if (d) return parseFloat(d[1]);
    return wordsToNumber(s);
  }

  var ABBREV = /(?:^|[\s(])(?:inc|corp|co|ltd|no|nos|sec|secs|st|ave|blvd|rd|dr|mr|ms|mrs|u\.s|l\.p|l\.l\.c|e\.g|i\.e|vs|approx|cal|civ|proc|ft|sq|ex)$/i;

  /** Splits a paragraph into sentences with their offsets (keeps decimals, abbreviations and short headings intact). */
  function sentences(text) {
    var out = [], start = 0, n = text.length;
    for (var i = 0; i < n; i++) {
      var ch = text.charAt(i);
      if (ch !== "." && ch !== "!" && ch !== "?" && ch !== ";") continue;
      var next = text.charAt(i + 1);
      if (ch === "." && /\d/.test(text.charAt(i - 1)) && /\d/.test(next)) continue;
      if (next && !/[\s"\u201D')]/.test(next)) continue;
      if (ch === "." && ABBREV.test(text.slice(Math.max(0, i - 8), i))) continue;
      var end = i + 1;
      while (end < n && /["\u201D')]/.test(text.charAt(end))) end++;
      if (trim(text.slice(start, end)).length < 25) continue; // fold "1. Premises." into the next sentence
      out.push({ text: trim(text.slice(start, end)), start: start, end: end });
      start = end;
      i = end - 1;
    }
    if (trim(text.slice(start))) {
      if (out.length && trim(text.slice(start)).length < 25) { out[out.length - 1].text = trim(text.slice(out[out.length - 1].start)); out[out.length - 1].end = n; }
      else out.push({ text: trim(text.slice(start)), start: start, end: n });
    }
    return out.length ? out : [{ text: trim(text), start: 0, end: n }];
  }

  function sentenceAt(text, offset) {
    var s = sentences(text);
    for (var i = 0; i < s.length; i++) if (offset >= s[i].start && offset < s[i].end) return s[i].text;
    return s.length ? s[0].text : text;
  }

  /* ------------------------------------------------------------------ *
   * Paragraph model
   * ------------------------------------------------------------------ */

  var INLINE_HEADING = /^\s*(?:(?:article|section)\s+)?(\d+(?:\.\d+)*|[A-Z]|[IVXLC]+)[.)]?\s+([A-Z][A-Za-z0-9 ,&'\/()-]{2,70}?)[.:]\s+\S/;
  var SIGNATURE_LINE = /^\s*(by|name|its|title|date|signature|print name|buyer'?s initials|seller'?s initials|initials|landlord|tenant|buyer|seller|guarantor)\s*:|^\s*in witness whereof\b/i;

  function isStandaloneHeading(t) {
    var s = trim(t);
    if (!s || s.length > 90) return false;
    if (SIGNATURE_LINE.test(s)) return false;
    var words = s.split(/\s+/);
    if (words.length > 11) return false;
    var letters = s.replace(/[^A-Za-z]/g, "");
    if (letters.length < 3) return false;
    var upper = s.replace(/[^A-Z]/g, "").length / letters.length;
    var numbered = /^\s*(?:(?:article|section)\s+)?(?:\d+(?:\.\d+)*|[IVXLC]+)[.)]?\s+\S/i.test(s);
    if (upper > 0.7) return true;
    if (numbered && !/[.;,]$/.test(s) && words.length <= 8) return true;
    return /^(article|section)\s+\d+/i.test(s) && words.length <= 8;
  }

  /** Accepts an array of paragraph strings (Word) or one string (browser). */
  function toParagraphs(input) {
    var list;
    if (Object.prototype.toString.call(input) === "[object Array]") list = input;
    else list = String(input == null ? "" : input).split(/\r\n|\r|\n/);
    var out = [], currentHeading = "", currentNo = "";
    for (var i = 0; i < list.length; i++) {
      var raw = String(list[i] == null ? "" : list[i]);
      var text = norm(raw);
      var t = trim(text);
      var p = { index: i, raw: raw, text: text, empty: !t, heading: currentHeading, sectionNo: currentNo, inlineHeading: "", isHeading: false, isSignature: SIGNATURE_LINE.test(t) };
      if (p.isSignature) { currentHeading = ""; currentNo = ""; p.heading = ""; p.sectionNo = ""; }
      else if (t) {
        if (isStandaloneHeading(t)) {
          p.isHeading = true;
          currentHeading = t;
          p.heading = t;
          var hn = /^\s*(?:(?:article|section)\s+)?(\d+(?:\.\d+)*)[.)]?\s/i.exec(t);
          currentNo = hn ? hn[1] : "";
          p.sectionNo = currentNo;
        } else {
          var m = INLINE_HEADING.exec(text);
          if (m) {
            p.inlineHeading = m[2];
            currentHeading = m[2];
            p.heading = m[2];
            currentNo = /^\d/.test(m[1]) ? m[1] : "";
            p.sectionNo = currentNo;
          }
        }
      }
      out.push(p);
    }
    return out;
  }

  function refOf(p) { return p && p.sectionNo ? "\u00A7" + p.sectionNo : "\u00B6" + ((p ? p.index : 0) + 1); }

  function titleCase(s) {
    var letters = s.replace(/[^A-Za-z]/g, "");
    if (!letters || s.replace(/[^A-Z]/g, "").length / letters.length < 0.8) return s;
    return s.toLowerCase().replace(/(^|[\s(\/-])([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); })
      .replace(/\b(And|Or|Of|The|To|In|On|For|A|An)\b/g, function (w) { return w.toLowerCase(); })
      .replace(/^([a-z])/, function (c) { return c.toUpperCase(); })
      .replace(/\bAs-is\b/i, "AS-IS");
  }

  function sectionName(p) {
    return p && p.heading ? titleCase(p.heading.replace(/^\s*(?:(?:article|section)\s+)?\d+(?:\.\d+)*[.)]?\s*/i, "")) : "";
  }

  /* ------------------------------------------------------------------ *
   * Clause categories
   * ------------------------------------------------------------------ */

  // head: matched against the section heading; body: matched against paragraph text.
  var CATEGORIES = [
    { id: "premises", label: "Premises", docTypes: ["lease"], head: /^(the )?premises\b|demise/i, body: /\bpremises\b[\s\S]{0,160}\b(suite|rentable square feet|square feet)\b/i },
    { id: "term", label: "Term & commencement", docTypes: ["lease"], head: /^term\b|^lease term|commencement date|^commencement|^expiration/i, body: /\("(commencement date|expiration date|term)"\)|\bthe term of this lease (shall|will) (be|commence)\b/i },
    { id: "rent", label: "Base rent", docTypes: ["lease"], head: /\brent\b(?!\s*adjust)/i, body: /\bbase rent\b[\s\S]{0,120}\b(monthly|per month|installments|per rentable)/i },
    { id: "escalation", label: "Rent increases", docTypes: ["lease"], head: /adjust|escalat|increase/i, body: /\b(rent|base rent)\b[\s\S]{0,120}\b(increase|escalat|adjust)\w*\b|\bconsumer price index\b|\bCPI\b/i },
    { id: "operating_expenses", label: "Operating expenses / NNN", docTypes: ["lease"], head: /operating expenses|common area|cam\b|additional rent|taxes/i, body: /\boperating expenses\b|\bcommon area maintenance\b|\btriple net\b|\bNNN\b|\btenant's (pro rata )?share\b/i },
    { id: "security_deposit", label: "Security deposit", docTypes: ["lease"], head: /security deposit/i, body: /\bsecurity deposit\b/i },
    { id: "use", label: "Permitted use", docTypes: ["lease"], head: /^use\b|\buse\b of premises|permitted use/i, body: /\bpermitted use\b|\buse the premises\b/i },
    { id: "condition", label: "Delivery condition", docTypes: ["lease"], head: /condition|delivery|improvements/i, body: /\baccepts? the premises\b|\bas[- ]is condition\b/i },
    { id: "repairs", label: "Repairs & maintenance", docTypes: ["lease"], head: /repair|maintenance/i, body: /\b(maintain|repair|replace)\b[\s\S]{0,100}\b(hvac|roof|structural|plumbing|systems)\b/i },
    { id: "alterations", label: "Alterations", docTypes: ["lease"], head: /alteration/i, body: /\bmake any alterations\b/i },
    { id: "assignment", label: "Assignment & subletting", docTypes: ["lease", "psa"], head: /assign|sublet|transfer/i, body: /\bassign\w*\b[\s\S]{0,80}\b(sublet|sublease)\b|\bassignment and subletting\b|\b(buyer|seller|tenant) (shall not|may not|may) assign\b/i },
    { id: "insurance", label: "Insurance", docTypes: ["lease"], head: /insurance/i, body: /\bcommercial general liability\b|\badditional insureds?\b/i },
    { id: "indemnity", label: "Indemnity", docTypes: ["lease", "psa"], head: /indemn/i, body: /\bindemnif\w*\b|\bhold (\w+ )?harmless\b/i },
    { id: "waiver_subrogation", label: "Waiver of subrogation", docTypes: ["lease"], head: /subrogation/i, body: /\bsubrogation\b/i },
    { id: "limitation_liability", label: "Limitation of liability", docTypes: ["lease", "psa"], head: /limitation of liability|landlord's liability/i, body: /\blimited to landlord's interest\b|\bliability of landlord\b|\bin no event shall (either party|landlord|tenant|seller|buyer) be liable\b|\bwaive\w*\b[\s\S]{0,40}\bconsequential damages\b|\blandlord shall not be (liable|responsible) for any (damage|loss|injury)\b/i },
    { id: "utilities", label: "Utilities & services", docTypes: ["lease"], head: /utilit|services/i, body: /\binterruption\b[\s\S]{0,60}\b(utilities|services)\b/i },
    { id: "casualty", label: "Casualty", docTypes: ["lease", "psa"], head: /damage|destruction|casualty/i, body: /\b(fire or other casualty|damaged by (fire|casualty)|casualty)\b/i },
    { id: "condemnation", label: "Condemnation", docTypes: ["lease", "psa"], head: /condemnation|eminent domain/i, body: /\beminent domain\b|\b(is|are) (taken|condemned)\b|\bcondemnation\b[\s\S]{0,80}\b(terminate|award)\b/i },
    { id: "snda", label: "Subordination / SNDA", docTypes: ["lease"], head: /subordinat|non-?disturbance|attornment/i, body: /\bsubordinat\w*\b|\bnon-?disturbance\b|\battorn(?:ment|s)?\b/i },
    { id: "estoppel", label: "Estoppel certificates", docTypes: ["lease", "psa"], head: /estoppel/i, body: /\bestoppel\b/i },
    { id: "relocation", label: "Relocation", docTypes: ["lease"], head: /relocat/i, body: /\brelocate (tenant|the premises)\b|\bsubstitute premises\b/i },
    { id: "holdover", label: "Holdover", docTypes: ["lease"], head: /hold(ing)?\s*-?over/i, body: /\bhold(s|ing)?\s*-?over\b|\bremains in possession after\b/i },
    { id: "default", label: "Default & remedies", docTypes: ["lease", "psa"], head: /default|remed/i, body: /\bevent of default\b|\bdefaults under this (lease|agreement)\b/i },
    { id: "late_charge", label: "Late charges & interest", docTypes: ["lease"], head: /late charge|interest/i, body: /\blate charge\b|\blate fee\b/i },
    { id: "entry", label: "Landlord entry", docTypes: ["lease"], head: /entry|access|inspection by landlord/i, body: /\b(landlord|its agents) may enter\b|\bright to enter\b/i },
    { id: "landlord_lien", label: "Landlord's lien", docTypes: ["lease"], head: /lien/i, body: /\blandlord's lien\b|\bsecurity interest in (all of )?tenant's\b/i },
    { id: "parking", label: "Parking", docTypes: ["lease"], head: /parking/i, body: /\bparking (spaces|stalls|passes|ratio)\b/i },
    { id: "rules", label: "Rules & regulations", docTypes: ["lease"], head: /rules/i, body: /\brules and regulations\b/i },
    { id: "guaranty", label: "Guaranty", docTypes: ["lease"], head: /guarant/i, body: /\bguarant(y|ee|or)\b/i },
    { id: "renewal_option", label: "Renewal / extension option", docTypes: ["lease"], head: /option to (extend|renew)|renewal|extension/i, body: /\boption to (extend|renew)\b|\bextension option\b|\brenewal (option|term)\b/i },
    { id: "early_termination", label: "Early termination", docTypes: ["lease"], head: /early termination|termination option/i, body: /\b(right|option) to terminate this lease\b|\bearly termination\b/i },
    { id: "quiet_enjoyment", label: "Quiet enjoyment", docTypes: ["lease"], head: /quiet enjoyment/i, body: /\bquiet enjoyment\b|\bpeaceably and quietly\b/i },
    { id: "audit", label: "Audit rights", docTypes: ["lease"], head: /audit/i, body: /\baudit\b|\binspect landlord's (books|records)\b/i },
    { id: "hazardous", label: "Hazardous materials", docTypes: ["lease", "psa"], head: /hazardous|environmental/i, body: /\bhazardous (materials|substances)\b|\bphase i\b/i },
    { id: "casp", label: "CASp disclosure", docTypes: ["lease"], head: /casp|accessib/i, body: /\bCASp\b|\bcertified access specialist\b/i },
    { id: "signage", label: "Signage", docTypes: ["lease"], head: /signage|signs/i, body: /\bsignage\b/i },
    { id: "retail_ops", label: "Retail covenants", docTypes: ["lease"], head: /co-?tenancy|exclusive|continuous operation|radius/i, body: /\bco-?tenancy\b|\bexclusive use\b|\bcontinuous(ly)? operat\w*\b|\bgo dark\b|\bradius\b/i },
    { id: "attorneys_fees", label: "Attorneys' fees", docTypes: ["lease", "psa"], head: /attorneys/i, body: /\battorneys' fees\b/i },
    { id: "jury_waiver", label: "Jury waiver / disputes", docTypes: ["lease", "psa"], head: /jury|arbitration|dispute/i, body: /\btrial by jury\b|\bjury trial\b|\barbitration\b/i },
    { id: "brokers", label: "Brokers", docTypes: ["lease", "psa"], head: /broker/i, body: /\bbrokers?\b[\s\S]{0,80}\b(commission|dealt)\b/i },
    { id: "notices", label: "Notices", docTypes: ["lease", "psa"], head: /^notices?\b|\bnotices\b/i, body: /\bnotices?\b[\s\S]{0,60}\b(in writing|shall be delivered)\b/i },
    { id: "governing_law", label: "Governing law", docTypes: ["lease", "psa"], head: /governing law/i, body: /\bgoverned by\b[\s\S]{0,40}\blaws? of\b/i },
    // Purchase agreements
    { id: "purchase_price", label: "Purchase price", docTypes: ["psa"], head: /purchase price|consideration/i, body: /\("purchase price"\)|\bthe purchase price (for the property )?(is|shall be)\b/i },
    { id: "deposit", label: "Deposit / earnest money", docTypes: ["psa"], head: /deposit|earnest money/i, body: /\("deposit"\)|\bearnest money\b|\bshall deposit\b|\bdeposit\b[\s\S]{0,80}\b(non-?refundable|refundable|escrow holder)\b/i },
    { id: "due_diligence", label: "Due diligence period", docTypes: ["psa"], head: /due diligence|inspection|feasibility|contingenc/i, body: /\("(due diligence|inspection|feasibility|contingency) period"\)|\bshall have\b[\s\S]{0,60}\bto (inspect|review|investigate)\b/i },
    { id: "title", label: "Title & survey", docTypes: ["psa"], head: /title|survey/i, body: /\bpermitted exceptions\b|\bpreliminary (title )?report\b|\btitle (commitment|policy)\b|\bALTA survey\b/i },
    { id: "financing", label: "Financing", docTypes: ["psa"], head: /financing|loan/i, body: /\bfinancing contingency\b|\bobtaining financing\b|\bloan contingency\b/i },
    { id: "closing", label: "Closing", docTypes: ["psa"], head: /^closing|^close of escrow|\d\.?\s+clos(ing|e of escrow)/i, body: /\("closing date"\)|\b(close of escrow|closing) shall (occur|take place)\b/i },
    { id: "prorations", label: "Prorations & costs", docTypes: ["psa"], head: /proration|closing costs/i, body: /\bprorat\w*\b|\btransfer tax\b/i },
    { id: "reps", label: "Representations & warranties", docTypes: ["psa"], head: /representations|warrant/i, body: /\brepresents and warrants\b|\brepresentations and warranties\b/i },
    { id: "as_is", label: "AS-IS & release", docTypes: ["psa"], head: /as-?is|release/i, body: /\bAS[- ]IS\b|\bWHERE[- ]IS\b|\b[Ww]ith [Aa]ll [Ff]aults\b|\bas[- ]is,? where[- ]is\b|\bas[- ]is condition\b/ },
    { id: "exchange_1031", label: "1031 exchange", docTypes: ["psa"], head: /1031|exchange/i, body: /\b1031\b|\btax[- ]deferred exchange\b/i },
    { id: "operating_covenants", label: "Seller covenants before closing", docTypes: ["psa"], head: /covenants|operation of the property|pending closing/i, body: /\b(prior to|before|until) (the )?(close of escrow|closing)\b[\s\S]{0,120}\bseller shall (not )?(enter into|modify|amend|operate|maintain)\b/i }
  ];

  var CATEGORY_BY_ID = {};
  for (var ci = 0; ci < CATEGORIES.length; ci++) CATEGORY_BY_ID[CATEGORIES[ci].id] = CATEGORIES[ci];

  function categoryLabel(id) { return CATEGORY_BY_ID[id] ? CATEGORY_BY_ID[id].label : id; }

  function categoriesFor(p, docType) {
    var hits = [];
    if (p.empty || p.isSignature) return hits;
    for (var i = 0; i < CATEGORIES.length; i++) {
      var c = CATEGORIES[i];
      if (docType !== "unknown" && !contains(c.docTypes, docType)) continue;
      var headHit = p.heading && c.head.test(p.heading);
      var bodyHit = !p.isHeading && c.body.test(p.text);
      if (headHit || bodyHit) hits.push(c.id);
    }
    return hits;
  }

  /* ------------------------------------------------------------------ *
   * Document type and jurisdiction
   * ------------------------------------------------------------------ */

  function countMatches(re, text) {
    var m = text.match(re);
    return m ? m.length : 0;
  }

  function detectDocType(paragraphs) {
    var text = paragraphs.map ? paragraphs.map(function (p) { return typeof p === "string" ? norm(p) : p.text; }).join("\n") : norm(paragraphs);
    var lower = text.toLowerCase();
    var lease = countMatches(/\blandlord\b/g, lower) + countMatches(/\btenant\b/g, lower) +
      3 * countMatches(/\bpremises\b/g, lower) + 4 * countMatches(/\bbase rent\b/g, lower) +
      5 * countMatches(/\b(this|the) lease\b/g, lower);
    var psa = countMatches(/\b(buyer|purchaser)\b/g, lower) + countMatches(/\bseller\b/g, lower) +
      4 * countMatches(/\bpurchase price\b/g, lower) + 3 * countMatches(/\b(close of escrow|closing date)\b/g, lower) +
      3 * countMatches(/\bdue diligence\b/g, lower) + 3 * countMatches(/\b(earnest money|the deposit)\b/g, lower) +
      5 * countMatches(/\b(purchase and sale|purchase agreement|agreement of sale)\b/g, lower);
    if (lease < 6 && psa < 6) return "unknown";
    return lease >= psa ? "lease" : "psa";
  }

  function isCalifornia(text) {
    return /\bcalifornia\b|\bcal\.? civ(il)?\.? code\b/i.test(text);
  }

  var PERSPECTIVES = [
    { id: "tenant", label: "Tenant", docType: "lease" },
    { id: "landlord", label: "Landlord", docType: "lease" },
    { id: "buyer", label: "Buyer", docType: "psa" },
    { id: "seller", label: "Seller", docType: "psa" },
    { id: "neutral", label: "Both sides", docType: null }
  ];

  function defaultPerspective(docType) {
    return docType === "psa" ? "buyer" : docType === "lease" ? "tenant" : "neutral";
  }

  function perspectiveLabel(id) {
    for (var i = 0; i < PERSPECTIVES.length; i++) if (PERSPECTIVES[i].id === id) return PERSPECTIVES[i].label;
    return id;
  }

  /* ------------------------------------------------------------------ *
   * Insertable clause templates (starting points for negotiation)
   * ------------------------------------------------------------------ */

  var TEMPLATES = {
    snda: { title: "SNDA / non-disturbance", category: "snda", text: "Subordination and Non-Disturbance. Tenant's subordination of this Lease to any present or future mortgage, deed of trust or ground lease is conditioned on Tenant's receipt of a commercially reasonable subordination, non-disturbance and attornment agreement (\"SNDA\") from the holder, providing that so long as Tenant is not in default beyond applicable notice and cure periods, Tenant's possession of the Premises and its rights under this Lease shall not be disturbed. Landlord shall deliver an SNDA from each existing lender within thirty (30) days after the date of this Lease." },
    waiver_subrogation: { title: "Mutual waiver of subrogation", category: "waiver_subrogation", text: "Waiver of Subrogation. Landlord and Tenant each waive all rights of recovery against the other, and against the other's officers, employees, agents and contractors, for any loss or damage to its property to the extent such loss or damage is insured, or required by this Lease to be insured, under a property insurance policy. Each party shall cause its property insurance policies to include a waiver of subrogation consistent with this Section." },
    casualty: { title: "Casualty termination and abatement", category: "casualty", text: "Damage or Destruction. If the Premises or the Building are damaged by fire or other casualty, Landlord shall, within sixty (60) days, deliver to Tenant its contractor's estimate of the time required to restore. If the estimate exceeds one hundred eighty (180) days, or if restoration is not substantially complete within two hundred seventy (270) days after the casualty, Tenant may terminate this Lease by written notice to Landlord. Base Rent and Additional Rent shall abate in proportion to the portion of the Premises rendered untenantable until restoration is substantially complete." },
    quiet_enjoyment: { title: "Quiet enjoyment", category: "quiet_enjoyment", text: "Quiet Enjoyment. So long as Tenant is not in default beyond applicable notice and cure periods, Tenant shall peaceably and quietly hold and enjoy the Premises for the Term without hindrance by Landlord or anyone claiming by, through or under Landlord." },
    renewal_option: { title: "Option to extend", category: "renewal_option", text: "Option to Extend. Tenant shall have one (1) option to extend the Term for five (5) years by giving Landlord written notice not less than nine (9) nor more than twelve (12) months before the Expiration Date. Base Rent for the extension term shall be ninety-five percent (95%) of the fair market rent for comparable space in comparable buildings in the area, determined as of the commencement of the extension term." },
    audit: { title: "Operating expense audit", category: "audit", text: "Audit. Within one hundred eighty (180) days after receiving Landlord's annual statement of Operating Expenses, Tenant may audit Landlord's books and records relating to Operating Expenses for that year. If the audit shows that Tenant was overcharged, Landlord shall refund the overcharge within thirty (30) days, and if the overcharge exceeds five percent (5%), Landlord shall also reimburse Tenant's reasonable audit costs." },
    opex_cap: { title: "Cap on controllable expenses", category: "operating_expenses", text: "Notwithstanding anything to the contrary in this Lease, Tenant's Share of Controllable Operating Expenses (all Operating Expenses other than real property taxes, insurance premiums and utility costs) shall not increase by more than five percent (5%) per calendar year, on a non-cumulative basis, over the amount payable for the prior calendar year." },
    casp: { title: "CASp disclosure (Cal. Civ. Code §1938)", category: "casp", text: "CASp Disclosure. For purposes of California Civil Code Section 1938, Landlord states that the Premises have not undergone inspection by a Certified Access Specialist (CASp). A Certified Access Specialist (CASp) can inspect the subject premises and determine whether the subject premises comply with all of the applicable construction-related accessibility standards under state law. Although state law does not require a CASp inspection of the subject premises, the commercial property owner or lessor may not prohibit the lessee or tenant from obtaining a CASp inspection of the subject premises for the occupancy or potential occupancy of the lessee or tenant, if requested by the lessee or tenant. The parties shall mutually agree on the arrangements for the time and manner of the CASp inspection, the payment of the fee for the CASp inspection, and the cost of making any repairs necessary to correct violations of construction-related accessibility standards within the premises." },
    security_deposit: { title: "Security deposit", category: "security_deposit", text: "Security Deposit. Upon execution of this Lease, Tenant shall deposit with Landlord $[amount] as security for Tenant's performance. Landlord may apply the Security Deposit to cure any default beyond applicable notice and cure periods, and Tenant shall restore any amount so applied within ten (10) days after written demand. Landlord shall return any unapplied balance within thirty (30) days after the expiration of the Term." },
    holdover: { title: "Holdover", category: "holdover", text: "Holding Over. If Tenant remains in possession after the expiration or termination of the Term without Landlord's written consent, Tenant shall be a tenant at sufferance and shall pay Base Rent equal to one hundred fifty percent (150%) of the Base Rent in effect immediately before the holdover, prorated on a daily basis." },
    estoppel: { title: "Estoppel certificates", category: "estoppel", text: "Estoppel Certificates. Within ten (10) business days after written request, Tenant shall execute and deliver a certificate stating whether this Lease is in full force and effect, the dates through which Rent has been paid, and whether, to Tenant's knowledge, either party is in default." },
    insurance: { title: "Insurance", category: "insurance", text: "Insurance. Tenant shall maintain commercial general liability insurance with limits of not less than $[1,000,000] per occurrence and $[2,000,000] in the aggregate, naming Landlord as an additional insured, and property insurance covering Tenant's personal property and improvements at full replacement cost. Landlord shall maintain property insurance on the Building at full replacement cost and commercial general liability insurance on the common areas." },
    indemnity: { title: "Mutual indemnity", category: "indemnity", text: "Indemnity. Subject to the waiver of subrogation in this Lease, Tenant shall indemnify, defend and hold Landlord harmless from claims arising from the negligence or willful misconduct of Tenant or its agents, employees or contractors in or about the Premises, and Landlord shall indemnify, defend and hold Tenant harmless from claims arising from the negligence or willful misconduct of Landlord or its agents, employees or contractors in or about the Building." },
    default_notice: { title: "Notice and cure", category: "default", text: "Default. Neither party shall be in default unless it fails to cure (a) a monetary default within ten (10) days after written notice, or (b) any other default within thirty (30) days after written notice, or such longer period as is reasonably necessary provided the cure is commenced within such thirty (30) day period and diligently pursued." },
    assignment: { title: "Assignment and permitted transfers", category: "assignment", text: "Assignment and Subletting. Tenant may not assign this Lease or sublet the Premises without Landlord's consent, which shall not be unreasonably withheld, conditioned or delayed and shall be granted or denied within fifteen (15) days after request. Notwithstanding the foregoing, Tenant may, without Landlord's consent, assign this Lease or sublet all or part of the Premises to an affiliate, to an entity resulting from a merger or consolidation, or to a purchaser of all or substantially all of Tenant's assets or stock (each, a \"Permitted Transfer\"), upon written notice to Landlord." },
    attorneys_fees: { title: "Prevailing party attorneys' fees", category: "attorneys_fees", text: "Attorneys' Fees. In any action or proceeding arising out of this Agreement, the prevailing party shall be entitled to recover its reasonable attorneys' fees and costs from the non-prevailing party." },
    governing_law: { title: "Governing law", category: "governing_law", text: "Governing Law. This Agreement shall be governed by and construed in accordance with the laws of the State of California." },
    notices: { title: "Notices", category: "notices", text: "Notices. All notices under this Agreement shall be in writing and shall be delivered by personal delivery, nationally recognized overnight courier, or email with a confirming copy sent by courier, to the addresses set forth below the parties' signatures, and shall be effective upon receipt or refusal." },
    brokers: { title: "Brokers", category: "brokers", text: "Brokers. Each party represents that it has dealt with no broker in connection with this transaction other than [broker], whose commission shall be paid by [party] under a separate written agreement, and each party shall indemnify the other against claims by any other broker claiming through it." },
    due_diligence: { title: "Due diligence period", category: "due_diligence", text: "Due Diligence. Buyer shall have [forty-five (45)] days after the Opening of Escrow (the \"Due Diligence Period\") to review title, survey, leases, service contracts, environmental reports and all other matters relating to the Property, and Seller shall deliver all documents in its possession relating to the Property within five (5) business days after the Opening of Escrow. Buyer may terminate this Agreement for any reason by written notice before the end of the Due Diligence Period, in which case the Deposit shall be returned to Buyer." },
    title: { title: "Title review", category: "title", text: "Title. Within ten (10) days after the Opening of Escrow, Seller shall cause the title company to deliver a preliminary title report and copies of all exception documents. Buyer may object to any title matter before the end of the Due Diligence Period. Seller shall remove at or before Closing all deeds of trust, mechanics' liens and other monetary liens created or suffered by Seller." },
    deposit: { title: "Deposit", category: "deposit", text: "Deposit. Within three (3) business days after the Opening of Escrow, Buyer shall deposit $[amount] (the \"Deposit\") with Escrow Holder. The Deposit shall be fully refundable to Buyer until the expiration of the Due Diligence Period, and thereafter shall be refundable only upon Seller's default, the failure of a condition to Buyer's obligations, or as otherwise provided in this Agreement." },
    closing: { title: "Closing date", category: "closing", text: "Closing. The Close of Escrow shall occur on or before the date that is [thirty (30)] days after the expiration of the Due Diligence Period (the \"Closing Date\")." },
    prorations: { title: "Prorations", category: "prorations", text: "Prorations. Real property taxes and assessments, rents actually collected, operating expenses and other income and expenses of the Property shall be prorated as of 11:59 p.m. on the day before the Closing Date. Seller shall pay the documentary transfer tax and the premium for a standard coverage owner's title policy; Buyer shall pay the cost of any extended coverage and endorsements; escrow fees shall be shared equally." },
    reps: { title: "Seller representations", category: "reps", text: "Seller's Representations. Seller represents and warrants to Buyer that, to Seller's actual knowledge: (a) the rent roll and leases delivered to Buyer are true and complete; (b) there is no pending or threatened litigation or condemnation affecting the Property; (c) Seller has received no written notice of any uncured violation of law affecting the Property; and (d) Seller has full authority to enter into and perform this Agreement. Seller's representations shall survive the Closing for twelve (12) months." },
    operating_covenants: { title: "Seller covenants before closing", category: "operating_covenants", text: "Operation Before Closing. Until the Close of Escrow, Seller shall operate and maintain the Property in the ordinary course consistent with past practice and shall not, without Buyer's prior written consent, enter into, modify or terminate any lease or service contract that would bind Buyer after Closing, or grant any easement or encumbrance affecting the Property." },
    estoppel_psa: { title: "Tenant estoppel as closing condition", category: "estoppel", text: "Tenant Estoppels. As a condition to Buyer's obligation to close, Seller shall deliver, at least five (5) business days before the Closing Date, an estoppel certificate from each tenant of the Property dated not more than thirty (30) days before the Closing Date and disclosing no material defaults or adverse matters." },
    exchange_1031: { title: "1031 exchange cooperation", category: "exchange_1031", text: "Exchange Cooperation. Either party may structure this transaction as part of a tax-deferred exchange under Section 1031 of the Internal Revenue Code. The other party shall reasonably cooperate, at no cost, liability or delay to it, including by consenting to the assignment of this Agreement to a qualified intermediary." },
    as_is: { title: "AS-IS sale", category: "as_is", text: "AS-IS Sale. Except for Seller's express representations and warranties in this Agreement and in the closing documents, Buyer is acquiring the Property in its AS-IS, WHERE-IS condition, with all faults, and Seller makes no other representation or warranty regarding the Property." },
    liquidated_damages: { title: "Liquidated damages", category: "default", text: "Liquidated Damages. If the sale fails to close because of Buyer's default, Seller shall retain the Deposit as liquidated damages, which the parties agree is a reasonable estimate of Seller's damages under the circumstances existing on the date of this Agreement, and as Seller's sole remedy. Buyer's Initials: ______ Seller's Initials: ______" },
    financing: { title: "Financing contingency", category: "financing", text: "Financing. Buyer's obligations are conditioned on Buyer obtaining, within [thirty (30)] days after the Opening of Escrow, a written commitment for a loan of not less than [sixty-five percent (65%)] of the Purchase Price on commercially reasonable terms. If Buyer does not obtain the commitment, Buyer may terminate this Agreement by written notice and receive the return of the Deposit." }
  };

  /* ------------------------------------------------------------------ *
   * Risk rules
   * kind "phrase": regex tested on each paragraph in scope (one finding per paragraph).
   * kind "section": tested on the combined text of a category's paragraphs.
   * riskFor: whose risk it is. fix: replace the matched words, or append a sentence.
   * ------------------------------------------------------------------ */

  var RULES = [
    /* ---------- Lease: risks to the tenant ---------- */
    {
      id: "assign-sole-discretion", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "assignment", scope: "assignment",
      re: /\bsole\s+(?:and\s+absolute\s+)?discretion\b|\babsolute\s+discretion\b/i,
      title: "Assignment consent is in Landlord's sole discretion",
      why: "Landlord can refuse any assignment or sublease for any reason, including a sale of your business. The usual standard is consent not unreasonably withheld, conditioned or delayed.",
      fix: { type: "replace", text: "reasonable discretion (such consent not to be unreasonably withheld, conditioned or delayed)" }
    },
    {
      id: "assign-no-permitted-transfer", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "assignment",
      when: /\bassign\w*\b/i, unless: /\bpermitted transfer\b|\baffiliate\b|\bmerger\b|\bsubstantially all\b/i,
      title: "No permitted transfers to affiliates or a buyer of the business",
      why: "Without a permitted-transfer carve-out, a merger, reorganization or sale of the company needs Landlord's consent.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Tenant may assign this Lease or sublet all or part of the Premises without Landlord's consent to an affiliate, to an entity resulting from a merger or consolidation, or to a purchaser of all or substantially all of Tenant's assets or stock, upon written notice to Landlord." }
    },
    {
      id: "assign-recapture", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "assignment", scope: "assignment",
      re: /\brecapture\b/i,
      unless: /\bmore than fifty percent\b|\bwithdraw its request\b/i,
      title: "Landlord can recapture the space when you ask to sublease",
      why: "A recapture right lets Landlord take back the space just because you asked for consent. Limit it to transfers of most of the Premises for most of the remaining term, and let Tenant withdraw the request.",
      fix: { type: "append", text: "Landlord's recapture right shall apply only to a proposed assignment or a sublease of more than fifty percent (50%) of the Premises for substantially the remainder of the Term, and Tenant may withdraw its request within ten (10) days after Landlord's recapture notice, in which case the recapture notice shall be void." }
    },
    {
      id: "assign-profit-100", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "assignment", scope: "assignment",
      re: /\bone hundred percent\s*\(\s*100\s*%\s*\)|\b100\s*%(?=[\s\S]{0,120}\b(excess|profit|consideration)\b)/i,
      title: "Landlord takes 100% of sublease profit",
      why: "The common split is 50% of net profit after Tenant recovers its transfer costs (commissions, improvements, legal fees, free rent).",
      fix: { type: "replace", text: "fifty percent (50%), after Tenant first recovers its reasonable costs of the transfer (including brokerage commissions, improvement costs, legal fees and free rent)," }
    },
    {
      id: "opex-no-cap", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "operating_expenses",
      when: /\boperating expenses\b|\bcommon area\b|\btriple net\b|\bNNN\b/i, unless: /\bcap\b|\bshall not (increase|exceed)\b[\s\S]{0,60}\b(percent|%)|\bincrease by more than\b/i,
      title: "Operating expenses have no cap",
      why: "Tenant's share of expenses can rise without limit. A cap on controllable expenses (often about 5% a year, excluding taxes, insurance and utilities) limits that exposure.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Tenant's Share of Controllable Operating Expenses (all Operating Expenses other than real property taxes, insurance premiums and utility costs) shall not increase by more than five percent (5%) per calendar year, on a non-cumulative basis." }
    },
    {
      id: "opex-capital", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "operating_expenses", scope: "operating_expenses",
      re: /\bcapital (?:improvements?|expenditures?|repairs?|replacements?)\b/i, unless: /\bamortiz/i,
      title: "Capital costs pass through in full",
      why: "Tenant could pay for a new roof or major system in a single year. Pass-throughs are usually limited to capital items that save costs or are required by new laws, amortized over their useful life.",
      fix: { type: "append", text: "Capital expenditures shall be included in Operating Expenses only if they reduce Operating Expenses or are required by laws first enacted after the Commencement Date, and then only as amortized over their useful life under generally accepted accounting principles." }
    },
    {
      id: "opex-mgmt-fee", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "operating_expenses", scope: "operating_expenses",
      re: /\bmanagement fee\b[\s\S]{0,80}?\b(?:(?:four|five|six|seven|eight|nine|ten)\s+percent\s*\(\s*\d+(?:\.\d+)?\s*%\s*\)|(?:[4-9]|1\d)(?:\.\d+)?\s*%)/i,
      title: "Management fee looks high or is based on gross revenue",
      why: "Management fees are commonly about 3% of revenue, or calculated without counting the management fee itself. Confirm the base and the rate.",
      fix: null
    },
    {
      id: "opex-no-audit", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "operating_expenses",
      when: /\boperating expenses\b|\bcommon area\b/i, docUnless: /\baudit\b|\binspect landlord's (books|records)\b/i,
      title: "No right to audit operating expenses",
      why: "Without an audit right you cannot verify Landlord's annual expense reconciliation.",
      fix: { type: "append", text: "Tenant may audit Landlord's books and records relating to Operating Expenses within one hundred eighty (180) days after receiving Landlord's annual statement, and Landlord shall promptly refund any overcharge." }
    },
    {
      id: "indemnity-own-negligence", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "indemnity", scope: "indemnity",
      re: /\b(?:whether or not|regardless of whether|even if)\b[\s\S]{0,40}?\bcaused by\b[\s\S]{0,30}?\bnegligence of landlord\b|\bregardless of (?:landlord's )?negligence\b/i,
      title: "Tenant indemnifies Landlord even for Landlord's own negligence",
      why: "This puts Landlord's own negligence on Tenant (and Tenant's insurer). Carve out claims caused by Landlord's negligence or willful misconduct.",
      fix: { type: "replace", text: "except to the extent caused by the negligence or willful misconduct of Landlord or its agents, employees or contractors" }
    },
    {
      id: "indemnity-one-way", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "indemnity",
      when: /\btenant shall (indemnify|defend)\b/i, docUnless: /\blandlord shall (indemnify|defend)\b|\beach party shall indemnify\b|\bmutual\w*\b[\s\S]{0,20}\bindemn/i,
      title: "Indemnity runs one way",
      why: "Only Tenant indemnifies. A mutual indemnity covers claims caused by Landlord or its contractors in the common areas.",
      fix: { type: "append", text: "Landlord shall indemnify, defend and hold Tenant harmless from claims arising from the negligence or willful misconduct of Landlord or its agents, employees or contractors in or about the Building." }
    },
    {
      id: "landlord-no-liability-any-cause", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "limitation_liability",
      re: /\blandlord shall not be (?:liable|responsible) for any (?:damage|loss|injury)\b[\s\S]{0,80}?\b(from any cause(?: whatsoever)?)/i, group: 1,
      test: function (m, text) { return !/^\s*,?\s*except\b/i.test(text.slice(m.index + m[0].length, m.index + m[0].length + 14)); },
      title: "Landlord disclaims liability from any cause",
      why: "As written, Landlord is not responsible even when it damages your property through its own negligence.",
      fix: { type: "replace", text: "from any cause, except to the extent caused by the negligence or willful misconduct of Landlord or its agents, employees or contractors" }
    },
    {
      id: "relocation", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "relocation",
      re: /\blandlord may relocate\b|\brelocate tenant\b|\bsubstitute premises\b/i, unless: /\bonly once\b|\bcomparable\b[\s\S]{0,40}\b(size|quality)\b/i,
      title: "Landlord can relocate you",
      why: "Landlord can move you to other space with little protection. Limit it to comparable space, Landlord paying all moving costs, and a right for Tenant to terminate instead.",
      fix: { type: "append", text: "Any relocation shall be to space of comparable size, quality, view and floor height in the Building, may occur only once during the Term and not during the first twenty-four (24) or last twelve (12) months, requires at least ninety (90) days' notice, and Landlord shall pay all of Tenant's reasonable relocation costs; Tenant may instead terminate this Lease within thirty (30) days after receiving Landlord's relocation notice." }
    },
    {
      id: "holdover-rate", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "holdover", scope: "holdover",
      re: /\b(?:two hundred|three hundred)\s+percent\s*\(\s*[23]00\s*%\s*\)|\b[23]00\s*%|\bdouble the\b|\btwice the\b/i,
      title: "Holdover rent of 200% or more",
      why: "125–150% of the last month's rent is common, often with a grace period before consequential damages apply.",
      fix: { type: "replace", text: "one hundred twenty-five percent (125%) for the first sixty (60) days and one hundred fifty percent (150%) thereafter" }
    },
    {
      id: "holdover-consequential", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "holdover", scope: "holdover",
      re: /\bconsequential damages\b/i, unless: /\bafter\b[\s\S]{0,40}\b(days|notice)\b/i,
      title: "Holdover exposes you to consequential damages immediately",
      why: "A short holdover could make Tenant liable for Landlord's lost deal with a new tenant. Limit consequential damages to holdovers lasting more than 30–60 days after notice.",
      fix: { type: "append", text: "Tenant shall be liable for consequential damages only if the holdover continues for more than thirty (30) days after Landlord gives Tenant written notice that it has signed a lease with a new tenant for the Premises." }
    },
    {
      id: "personal-guaranty", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "guaranty",
      re: /\bpersonal(?:ly)?\s+guarant\w*\b|\bguarant(?:ee|y) (?:all|any) of tenant's obligations\b/i,
      unless: /\blimited to an amount equal to\b|\bshall terminate if\b/i,
      title: "Unlimited personal guaranty",
      why: "The principal is personally on the hook for the full lease. Negotiate a cap (for example 12 months' rent) and a burn-off after on-time payment.",
      fix: { type: "append", text: "The guarantor's liability shall be limited to an amount equal to twelve (12) months of Base Rent and shall terminate if Tenant has not been in monetary default beyond applicable cure periods during the first twenty-four (24) months of the Term." }
    },
    {
      id: "landlord-lien", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "landlord_lien",
      re: /\blandlord's lien\b|\bsecurity interest in (?:all of )?tenant's\b/i,
      title: "Landlord takes a lien on your furniture and equipment",
      why: "A contractual lien can block equipment financing or a business loan. Ask Landlord to waive it, or at least subordinate it to Tenant's lenders.",
      fix: null
    },
    {
      id: "snda-missing", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "snda",
      when: /\bsubordinat\w*\b/i, unless: /\bnon-?disturb\w*\b/i,
      title: "Subordination without non-disturbance",
      why: "If Landlord's lender forecloses, Tenant can lose the lease. Make subordination conditional on an SNDA.",
      fix: { type: "append", text: "Tenant's subordination is conditioned on Tenant's receipt of a commercially reasonable subordination, non-disturbance and attornment agreement providing that, so long as Tenant is not in default beyond applicable cure periods, Tenant's possession and rights under this Lease shall not be disturbed." }
    },
    {
      id: "attorney-in-fact", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "snda",
      re: /\battorney[- ]in[- ]fact\b/i,
      title: "Power of attorney for Landlord to sign for Tenant",
      why: "Landlord could sign subordination or estoppel documents in Tenant's name. Delete the power of attorney and rely on a notice-and-cure process.",
      fix: null
    },
    {
      id: "estoppel-short-turnaround", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "estoppel",
      re: /\bwithin\s+(?:(?:three|four|five|six|seven)\s*\(\s*[3-7]\s*\)|[3-7])\s+(?:business\s+)?days\b/i, scope: ["estoppel", "snda"],
      title: "Very short turnaround for estoppel or SNDA documents",
      why: "Ten business days is the common turnaround, and anything shorter is hard to meet with legal review.",
      fix: { type: "replace", text: "within ten (10) business days" }
    },
    {
      id: "default-no-notice", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "default", scope: "default",
      re: /\b(?:for|within)\s+(?:(?:one|two|three|four|five)\s*\(\s*[1-5]\s*\)|[1-5])\s+(?:business\s+)?days(?:\s+after (?:the )?(?:date )?(?:when )?due)?(?:,?\s+without (?:further )?notice)?/i,
      test: function (m, text) {
        // "for five (5) days after written notice" is acceptable; flag only short periods with no notice.
        var rest = text.slice(m.index + m[0].length, m.index + m[0].length + 40);
        return !/^\s*after\s+(?:written\s+)?notice/i.test(rest);
      },
      title: "Rent default with little or no notice",
      why: "An Event of Default can occur a few days after rent is late, without written notice. Ten days after written notice is common for monetary defaults.",
      fix: { type: "replace", text: "for ten (10) days after written notice from Landlord" }
    },
    {
      id: "default-short-nonmonetary", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "default", scope: "default",
      re: /\b(?:any other|other (?:obligations?|covenants?|terms?)|non-?monetary)[^;]{0,160}?\b((?:(?:five|ten|fifteen|twenty)\s*\(\s*(?:5|10|15|20)\s*\)|(?:5|10|15|20))\s+days after (?:written )?notice)\b/i, group: 1,
      title: "Short cure period for non-monetary defaults",
      why: "Thirty days, extended while Tenant diligently cures, is the usual period for non-monetary defaults.",
      fix: { type: "replace", text: "thirty (30) days after written notice (or such longer period as is reasonably necessary, provided Tenant commences the cure within such period and diligently pursues it)" }
    },
    {
      id: "late-charge-high", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "late_charge", scope: "late_charge",
      re: /\b(?:(?:six|seven|eight|nine|ten|twelve|fifteen|twenty)\s+percent\s*\(\s*\d+\s*%\s*\)|(?:[6-9]|[1-9]\d)\s*%)/i,
      title: "Late charge above 5%",
      why: "Late charges are commonly 5%, often with one grace notice per year before the charge applies.",
      fix: { type: "replace", text: "five percent (5%)" }
    },
    {
      id: "entry-any-time", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "entry",
      re: /\bat any time\b/i, scope: "entry",
      title: "Landlord may enter at any time",
      why: "Ask for at least 24 hours' notice except in emergencies, and for entry that doesn't unreasonably interfere with your business.",
      fix: { type: "replace", text: "upon at least twenty-four (24) hours' prior notice (except in an emergency) and in a manner that does not unreasonably interfere with Tenant's business" }
    },
    {
      id: "as-is-delivery", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "condition",
      re: /\baccepts? the premises\b[\s\S]{0,40}?\bas[- ]is\b|\bas[- ]is,? where[- ]is\b|\bas[- ]is condition\b/i,
      unless: /\blandlord warrants\b/i,
      title: "Premises delivered AS IS",
      why: "Tenant takes the risk of hidden defects in the roof, structure and building systems. Ask for a warranty that systems work at delivery, plus a repair period.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Landlord warrants that the roof, structure and building systems serving the Premises (including HVAC, electrical, plumbing and life-safety) will be in good working order on the Commencement Date, and Landlord shall, at its cost, repair any failure of such systems discovered within the first twelve (12) months of the Term." }
    },
    {
      id: "hvac-replacement", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "repairs", scope: "repairs",
      re: /\breplace the hvac\b|\breplace\b[\s\S]{0,40}\b(?:hvac|heating|air conditioning)\b/i,
      unless: /\blandlord shall replace\b/i,
      title: "Tenant must replace HVAC equipment",
      why: "Replacing a rooftop unit can cost more than several months of rent. Landlord usually replaces capital items, or Tenant pays only an amortized share.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Landlord shall replace any HVAC equipment that cannot be repaired at reasonable cost, and Tenant's share of the replacement cost shall be limited to the portion amortized over the remaining Term." }
    },
    {
      id: "structural-on-tenant", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "high", category: "repairs", scope: "repairs",
      re: /\btenant shall\b[\s\S]{0,120}?\b(?:roof|foundation|structural portions?|structure)\b/i, docUnless: /\blandlord shall\b[\s\S]{0,80}\b(roof|structural)\b/i,
      title: "Tenant responsible for roof or structure",
      why: "Structural and roof work is normally Landlord's cost, outside operating expenses.",
      fix: { type: "append", text: "Landlord shall, at its sole cost and not as an Operating Expense, maintain and repair the roof, foundation, exterior walls and structural portions of the Building." }
    },
    {
      id: "utility-no-abatement", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "utilities",
      re: /\bno such interruption shall entitle tenant to any abatement\b|\blandlord shall not be liable for any interruption\b/i,
      unless: /\brent shall abate\b/i,
      title: "No rent abatement when services fail",
      why: "If power or HVAC fails because of something within Landlord's control, rent should abate after a few business days.",
      fix: { type: "append", text: "Notwithstanding the foregoing, if an interruption caused by Landlord's negligence or within Landlord's reasonable control continues for more than three (3) consecutive business days and materially interferes with Tenant's use of the Premises, Rent shall abate from the first day of the interruption until service is restored." }
    },
    {
      id: "casualty-no-tenant-termination", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "casualty",
      when: /\bcasualty\b|\bdamage\w*\b/i, unless: /\btenant (may|shall have the right to) terminate\b|\beither party may terminate\b/i,
      title: "Only Landlord can terminate after a casualty",
      why: "If restoration drags on, Tenant has no exit. Add a Tenant termination right tied to the restoration estimate and an outside date.",
      fix: { type: "append", text: "If Landlord's contractor estimates that restoration will take more than one hundred eighty (180) days, or if restoration is not substantially complete within two hundred seventy (270) days after the casualty, Tenant may terminate this Lease by written notice to Landlord." }
    },
    {
      id: "casualty-no-abatement", kind: "section", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "casualty",
      when: /\bcasualty\b|\bdamage\w*\b/i, unless: /\babate\w*\b/i,
      title: "No rent abatement while the space is unusable",
      why: "Rent should abate for the part of the Premises you can't use until it is restored.",
      fix: { type: "append", text: "Rent shall abate in proportion to the portion of the Premises rendered untenantable from the date of the casualty until restoration is substantially complete." }
    },
    {
      id: "escalation-high", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "escalation", scope: "escalation",
      re: /\b(?:four|five|six|seven)\s+percent\s*\(\s*[4-7](?:\.\d+)?\s*%\s*\)|\b[4-7](?:\.\d+)?\s*%/i,
      title: "Annual increases above 3%",
      why: "Fixed bumps of about 3% are a frequent target. Check the rate against the market for this property type.",
      fix: { type: "replace", text: "three percent (3%)" }
    },
    {
      id: "cpi-uncapped", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "escalation",
      re: /\bconsumer price index\b|\bCPI\b/, unless: /(?:\bcap\b|\bnot (?:to )?exceed\b|\bmaximum\b|\bceiling\b)|\bshall exceed\b/i,
      title: "CPI increases with no ceiling",
      why: "Uncapped CPI adjustments pass inflation spikes straight to Tenant. Add a ceiling.",
      fix: { type: "append", text: "provided, however, that no annual increase shall exceed four percent (4%)." }
    },
    {
      id: "attorneys-fees-one-sided", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "attorneys_fees",
      re: /\btenant shall (?:pay|reimburse)\b[\s\S]{0,80}?\battorneys' fees\b/i, docUnless: /\bprevailing party\b/i,
      title: "Attorneys' fees run only to Landlord",
      why: "Make it a prevailing-party clause so Tenant recovers fees when it wins.",
      fix: { type: "append", text: "In any action or proceeding arising out of this Lease, the prevailing party shall be entitled to recover its reasonable attorneys' fees and costs from the non-prevailing party." }
    },
    {
      id: "rules-unilateral", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "rules",
      re: /\b(?:amend|modify|change)\w*\b[\s\S]{0,30}?\bfrom time to time\b/i, scope: "rules",
      unless: /\bapplied uniformly\b/i,
      title: "Landlord can change the rules at will",
      why: "Ask that new rules be reasonable, applied uniformly, and unable to override the Lease or materially increase Tenant's obligations.",
      fix: { type: "append", text: "Any new or amended rules shall be reasonable, applied uniformly and non-discriminatorily, and shall not materially increase Tenant's obligations or decrease Tenant's rights under this Lease." }
    },
    {
      id: "parking-floating-rate", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "parking",
      re: /\bthen[- ](?:prevailing|current)\b[\s\S]{0,20}?\brates?\b|\bmarket rates?\b/i, scope: "parking",
      title: "Parking at Landlord's floating rates",
      why: "Lock the parking ratio and rate (or free parking) for the initial term.",
      fix: null
    },
    {
      id: "use-narrow", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "use",
      re: /\b(?:and|for) no other (?:use|purpose)\b/i,
      title: "Narrow permitted use",
      why: "A narrow use clause can block a future sublease or assignment. Add \"and any other lawful use consistent with a first-class office building\".",
      fix: { type: "replace", text: "and for any other lawful use consistent with comparable buildings in the area" }
    },
    {
      id: "retail-continuous-operation", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "retail_ops",
      re: /\bcontinuous(?:ly)? operat\w*\b|\bgo dark\b|\bshall (?:remain )?open for business\b/i,
      title: "Continuous operation covenant",
      why: "Tenant can't close a weak store without defaulting. Ask for a go-dark right (Landlord's remedy limited to recapture).",
      fix: null
    },
    {
      id: "retail-radius", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "retail_ops",
      re: /\bradius\b/i,
      title: "Radius restriction",
      why: "Limits opening another location nearby. Narrow the radius and term, and exclude existing and acquired locations.",
      fix: null
    },

    /* ---------- Lease: risks to the landlord ---------- */
    {
      id: "ll-termination-no-fee", kind: "phrase", docTypes: ["lease"], riskFor: ["landlord"], severity: "medium", category: "early_termination",
      re: /\b(?:right|option) to terminate this lease\b|\bearly termination\b/i, unless: /\btermination fee\b|\bunamortized\b/i,
      title: "Tenant termination right without a fee",
      why: "Landlord would lose unamortized improvements and commissions. Require a termination fee.",
      fix: { type: "append", text: "As a condition to such termination, Tenant shall pay Landlord a termination fee equal to the unamortized tenant improvement allowance, leasing commissions and free rent (amortized at eight percent (8%) per annum) plus two (2) months of Base Rent." }
    },
    {
      id: "ll-opex-cap-uncontrollables", kind: "phrase", docTypes: ["lease"], riskFor: ["landlord"], severity: "medium", category: "operating_expenses", scope: "operating_expenses",
      re: /\bshall not increase by more than\b|\bcap\b/i, unless: /\bother than (real property )?taxes\b|\bexclud\w*\b[\s\S]{0,40}\btaxes\b|\bcontrollable\b/i,
      title: "Expense cap covers taxes and insurance",
      why: "Landlord can't control taxes, insurance or utilities. Limit the cap to controllable expenses.",
      fix: { type: "append", text: "The foregoing limitation shall apply only to Controllable Operating Expenses and shall not apply to real property taxes, insurance premiums, utility costs or costs of complying with laws." }
    },
    {
      id: "ll-retail-cotenancy", kind: "phrase", docTypes: ["lease"], riskFor: ["landlord"], severity: "high", category: "retail_ops",
      re: /\bco-?tenancy\b/i,
      title: "Co-tenancy requirement",
      why: "Tenant can pay reduced rent or terminate if anchors leave. Narrow the trigger, the remedy period and the rent reduction.",
      fix: null
    },
    {
      id: "ll-retail-exclusive", kind: "phrase", docTypes: ["lease"], riskFor: ["landlord"], severity: "medium", category: "retail_ops",
      re: /\bexclusive (?:right|use)\b/i,
      title: "Exclusive use granted",
      why: "Limit it to Tenant's primary use, carve out existing tenants and anchors, and make it end if Tenant stops operating.",
      fix: null
    },
    {
      id: "ll-audit-open-ended", kind: "phrase", docTypes: ["lease"], riskFor: ["landlord"], severity: "low", category: "audit",
      re: /\baudit\b/i, unless: /\bwithin\s+[\s\S]{0,30}\b(days|months)\b/i,
      title: "Open-ended audit right",
      why: "Limit audits to a window after each annual statement, require a CPA not paid on contingency, and keep results confidential.",
      fix: null
    },

    /* ---------- Purchase agreements: risks to the buyer ---------- */
    {
      id: "psa-deposit-hard", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "high", category: "deposit",
      re: /\bnon-?refundable\b/i, unless: /(?:\b(?:after|upon|following) the (?:expiration|end) of the (?:due diligence|inspection|feasibility|contingency) period\b)|\bfully refundable\b/i,
      title: "Deposit goes hard immediately",
      why: "The deposit is at risk before Buyer has finished due diligence. It is usually refundable until the due diligence period ends.",
      fix: { type: "append", text: "Notwithstanding the foregoing, the Deposit shall be fully refundable to Buyer if Buyer terminates this Agreement on or before the expiration of the Due Diligence Period, upon Seller's default, or upon the failure of any condition to Buyer's obligations." }
    },
    {
      id: "psa-dd-short", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "medium", category: "due_diligence", scope: "due_diligence",
      re: /\b(?:(?:seven|ten|fourteen|fifteen|twenty|twenty-one|twenty-five)\s*\(\s*(?:7|10|14|15|20|21|25)\s*\)|(?:7|10|14|15|20|21|25))\s+(?:calendar\s+|business\s+)?days\b/i,
      title: "Short due diligence period",
      why: "Under 30 days rarely leaves time for title, survey, a Phase I, lease review and lender work. Thirty to sixty days is common, longer for land.",
      fix: { type: "replace", text: "forty-five (45) days" }
    },
    {
      id: "psa-release-no-carveouts", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "high", category: "as_is",
      re: /\breleases? seller\b[\s\S]{0,60}?\b(?:any and all|all) claims\b|\bknown (?:or|and) unknown\b|\b1542\b/i, unless: /(?:\bexcept\b[\s\S]{0,80}\b(fraud|representations|breach)\b)|\bshall not apply to\b[\s\S]{0,40}\bfraud\b/i,
      title: "Broad release of Seller with no carve-outs",
      why: "Buyer gives up claims for known and unknown problems, including a California Civil Code §1542 waiver. Carve out fraud, breach of Seller's reps and covenants, and pre-closing third-party claims.",
      fix: { type: "append", text: "The foregoing release shall not apply to (i) Seller's fraud or intentional misrepresentation, (ii) any breach of Seller's express representations, warranties or covenants in this Agreement or the closing documents, or (iii) third-party claims arising from Seller's ownership or operation of the Property before the Closing." }
    },
    {
      id: "psa-reps-short-survival", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "medium", category: "reps",
      re: /\bsurviv\w*\b[\s\S]{0,80}?\b((?:(?:one|two|three|four|five|six)\s*\(\s*[1-6]\s*\)|[1-6])\s+months?)\b/i, group: 1,
      title: "Seller's reps survive only a few months",
      why: "Problems often surface after the first property-tax or reconciliation cycle. Nine to twelve months is a common survival period.",
      fix: { type: "replace", text: "twelve (12) months" }
    },
    {
      id: "psa-buyer-remedy-deposit-only", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "high", category: "default",
      re: /\bbuyer's sole and exclusive remedy\b|\bbuyer's (?:sole|exclusive) remedy\b/i, unless: /\bspecific performance\b/i,
      title: "Buyer's only remedy is a refund",
      why: "If Seller walks for a better offer, Buyer only gets its deposit back. Add specific performance and reimbursement of out-of-pocket costs.",
      fix: { type: "append", text: "Alternatively, Buyer may (a) seek specific performance, provided Buyer files suit within ninety (90) days after the scheduled Closing Date, or (b) terminate this Agreement and receive the return of the Deposit together with reimbursement of Buyer's documented out-of-pocket costs not to exceed $[amount]." }
    },
    {
      id: "psa-buyer-no-assign", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "medium", category: "assignment",
      re: /\bbuyer (?:shall not|may not) assign\b/i, unless: /\baffiliate\b|\bentity (?:controlled|formed)\b/i,
      title: "Buyer can't assign to its acquisition entity",
      why: "Buyers usually take title in a new LLC or need a 1031 intermediary. Allow assignment to affiliates without consent.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Buyer may assign this Agreement without Seller's consent to an entity controlled by or under common control with Buyer, or formed by Buyer to acquire the Property, upon written notice to Seller before the Closing." }
    },
    {
      id: "psa-title-all-record", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "medium", category: "title",
      re: /\ball matters of (?:public )?record\b/i,
      title: "Permitted exceptions include everything of record",
      why: "Buyer accepts unknown recorded matters. Limit permitted exceptions to items Buyer approves during title review.",
      fix: { type: "replace", text: "only those matters approved or deemed approved by Buyer during the Due Diligence Period" }
    },
    {
      id: "psa-no-lien-cure", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "high", category: "title",
      re: /\bseller shall have no obligation to cure\b|\bseller shall not be obligated to (?:cure|remove)\b/i,
      unless: /\bshall remove at or before\b/i,
      title: "Seller won't even remove its own loans and liens",
      why: "At minimum, Seller should pay off its deeds of trust and other monetary liens at closing.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Seller shall remove at or before the Closing all deeds of trust, mechanics' liens and other monetary liens created or suffered by Seller." }
    },
    {
      id: "psa-estoppel-not-condition", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "medium", category: "estoppel",
      re: /\b(?:commercially )?reasonable efforts\b|\bshall not be a condition\b/i, scope: "estoppel",
      unless: /\bshall be a condition to buyer's obligation\b/i,
      title: "Tenant estoppel is not a closing condition",
      why: "Without an estoppel, Buyer closes without confirming the tenant's rent, defaults and side deals.",
      fix: { type: "append", text: "Notwithstanding the foregoing, Buyer's receipt of a clean estoppel certificate from the tenant, dated not more than thirty (30) days before the Closing Date, shall be a condition to Buyer's obligation to close." }
    },
    {
      id: "psa-no-invasive", kind: "phrase", docTypes: ["psa"], riskFor: ["buyer"], severity: "low", category: "due_diligence",
      re: /\bno (?:invasive|intrusive|subsurface)\b|\bnot conduct any invasive\b/i,
      unless: /\bshall not unreasonably withhold\b/i,
      title: "No invasive testing allowed",
      why: "If the Phase I recommends a Phase II, Buyer needs consent to test. Pre-agree that consent won't be unreasonably withheld.",
      fix: { type: "append", text: "Seller shall not unreasonably withhold, condition or delay its consent to invasive testing recommended by Buyer's Phase I environmental consultant." }
    },

    /* ---------- Purchase agreements: risks to the seller ---------- */
    {
      id: "psa-dd-long", kind: "phrase", docTypes: ["psa"], riskFor: ["seller"], severity: "medium", category: "due_diligence", scope: "due_diligence",
      re: /\b(?:(?:ninety|one hundred twenty|one hundred eighty)\s*\(\s*(?:90|120|180)\s*\)|(?:90|120|180))\s+(?:calendar\s+|business\s+)?days\b/i,
      title: "Long due diligence period",
      why: "The property is off the market while Buyer holds a free option. Shorten it, or add deposits that go hard in stages.",
      fix: null
    },
    {
      id: "psa-deposit-refundable-thru-closing", kind: "phrase", docTypes: ["psa"], riskFor: ["seller"], severity: "high", category: "deposit",
      re: /\brefundable\b[\s\S]{0,60}?\b(?:until|through|prior to) (?:the )?(?:closing|close of escrow)\b/i,
      title: "Deposit stays refundable through closing",
      why: "Buyer can walk at any time. The deposit should go hard when due diligence ends.",
      fix: null
    },
    {
      id: "psa-reps-uncapped", kind: "section", docTypes: ["psa"], riskFor: ["seller"], severity: "high", category: "reps",
      when: /\bsurviv\w*\b/i, docUnless: /\bshall not exceed\b|\baggregate liability\b|\bcap\b/i,
      title: "Seller's post-closing liability is uncapped",
      why: "Survival without a cap leaves Seller's sale proceeds exposed. A cap of 1–3% of the price and a minimum claim threshold are common.",
      fix: { type: "append", text: "Seller's aggregate liability for breaches of its representations and warranties shall not exceed [two percent (2%)] of the Purchase Price, and Buyer may not bring any such claim unless its damages exceed $[25,000]." }
    },
    {
      id: "psa-financing-contingency", kind: "phrase", docTypes: ["psa"], riskFor: ["seller"], severity: "medium", category: "financing",
      re: /\bfinancing contingency\b|\bcontingent (?:up)?on buyer(?:'s)? (?:ability to )?obtain\w*\b/i, unless: /\bnot (?:be )?contingent\b|\bno financing contingency\b/i,
      title: "Financing contingency",
      why: "Buyer can cancel if its loan falls through. Tie the contingency to a short deadline after which the deposit goes hard.",
      fix: null
    },
    {
      id: "psa-specific-performance-open", kind: "phrase", docTypes: ["psa"], riskFor: ["seller"], severity: "medium", category: "default",
      re: /\bspecific performance\b/i, unless: /\bwithin\s+[\s\S]{0,30}\bdays\b/i,
      title: "Open-ended specific performance",
      why: "Require Buyer to file within 60–90 days of the scheduled closing and to have been ready to close.",
      fix: { type: "append", text: "Buyer must file any action for specific performance within sixty (60) days after the scheduled Closing Date and must have been ready, willing and able to close." }
    },

    /* ---------- Both sides: drafting checks ---------- */
    {
      id: "number-mismatch", kind: "phrase", docTypes: ["lease", "psa", "unknown"], riskFor: ["tenant", "landlord", "buyer", "seller"], severity: "high", category: "drafting",
      re: new RegExp("\\b(" + WORD_NUMBER_RE + ")\\s*(percent\\s*)?\\(\\s*(\\d[\\d,]*(?:\\.\\d+)?)\\s*%?\\s*\\)", "gi"),
      test: function (m) {
        var w = wordsToNumber(m[1]);
        var d = parseFloat(String(m[3]).replace(/,/g, ""));
        return !isNaN(w) && !isNaN(d) && w !== d;
      },
      title: "Words and numbers don't match",
      why: "The written number and the figure in parentheses disagree, so the term is ambiguous. Fix whichever is wrong.",
      fix: null, all: true
    },
    {
      id: "placeholder", kind: "phrase", docTypes: ["lease", "psa", "unknown"], riskFor: ["tenant", "landlord", "buyer", "seller"], severity: "medium", category: "drafting",
      re: /\[\s*_{2,}\s*\]|\[(?:insert|tbd|to be determined|amount|date|name|party|broker)[^\]]{0,40}\]|_{5,}|\bTBD\b|\$\s*\[/i, skipSignatures: true,
      title: "Blank or placeholder left in the document",
      why: "A blank or bracketed placeholder is still open. Fill it in or delete it before signing.",
      fix: null
    },
    {
      id: "ca-jury-waiver", kind: "phrase", docTypes: ["lease", "psa"], riskFor: ["tenant", "landlord", "buyer", "seller"], severity: "low", category: "jury_waiver", requiresCalifornia: true,
      re: /\bwaive\w*\b[\s\S]{0,30}?\btrial by jury\b|\bjury trial\b/i,
      title: "Jury waiver may be unenforceable in California",
      why: "California courts generally do not enforce pre-dispute jury waivers (Grafton Partners v. Superior Court, 2005). A judicial reference clause (Code Civ. Proc. §638) is the usual substitute.",
      fix: null
    },
    {
      id: "ca-civil-code-waivers", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "casualty",
      re: /\b(?:1932|1933\s*\(\s*4\s*\)|1941|1942|1950\.7)\b/,
      title: "Tenant waives California statutory rights",
      why: "Waivers of Civil Code §§1932, 1933(4), 1941, 1942 or 1950.7 are standard in landlord forms. Make sure the lease gives you the contractual rights you're waiving (casualty termination, repairs, deposit accounting).",
      fix: null
    }
  ];

  /* Clauses expected in each document type: missing ones are reported with a template to insert. */
  var EXPECTED = {
    lease: [
      { category: "snda", riskFor: ["tenant"], severity: "high", template: "snda", why: "Without non-disturbance protection, a foreclosure by Landlord's lender can wipe out the lease." },
      { category: "waiver_subrogation", riskFor: ["tenant", "landlord"], severity: "medium", template: "waiver_subrogation", why: "Without a mutual waiver, one party's insurer can sue the other after an insured loss." },
      { category: "casualty", riskFor: ["tenant", "landlord"], severity: "medium", template: "casualty", why: "Nothing says what happens to rent and the lease if the building burns." },
      { category: "quiet_enjoyment", riskFor: ["tenant"], severity: "low", template: "quiet_enjoyment", why: "An express covenant of quiet enjoyment protects possession against Landlord's interference." },
      { category: "renewal_option", riskFor: ["tenant"], severity: "low", template: "renewal_option", why: "No option to extend: Tenant has no right to stay after the term." },
      { category: "casp", riskFor: ["tenant", "landlord"], severity: "medium", template: "casp", requiresCalifornia: true, why: "California Civil Code §1938 requires commercial leases to state whether the premises have been CASp-inspected." },
      { category: "security_deposit", riskFor: ["landlord"], severity: "medium", template: "security_deposit", why: "No security deposit protects Landlord against default." },
      { category: "holdover", riskFor: ["landlord"], severity: "medium", template: "holdover", why: "Without a holdover clause, Tenant can stay over at the old rent." },
      { category: "estoppel", riskFor: ["landlord"], severity: "medium", template: "estoppel", why: "Landlord's lenders and buyers will need tenant estoppel certificates." },
      { category: "insurance", riskFor: ["tenant", "landlord"], severity: "high", template: "insurance", why: "No insurance requirements are stated." },
      { category: "indemnity", riskFor: ["tenant", "landlord"], severity: "medium", template: "indemnity", why: "Nothing allocates responsibility for third-party injury claims." },
      { category: "default", riskFor: ["tenant", "landlord"], severity: "medium", template: "default_notice", why: "No default and cure provisions." },
      { category: "assignment", riskFor: ["tenant", "landlord"], severity: "medium", template: "assignment", why: "Nothing addresses assignment or subletting." },
      { category: "attorneys_fees", riskFor: ["tenant", "landlord"], severity: "low", template: "attorneys_fees", why: "No attorneys' fees clause; each side bears its own costs in a dispute." },
      { category: "governing_law", riskFor: ["tenant", "landlord"], severity: "low", template: "governing_law", why: "No governing-law clause." },
      { category: "notices", riskFor: ["tenant", "landlord"], severity: "low", template: "notices", why: "No notice provision, so it's unclear how to give notice of default or exercise options." },
      { category: "brokers", riskFor: ["tenant", "landlord"], severity: "low", template: "brokers", why: "No broker representation or commission clause." }
    ],
    psa: [
      { category: "due_diligence", riskFor: ["buyer"], severity: "high", template: "due_diligence", why: "No due diligence period: Buyer has no free-look right to walk away." },
      { category: "title", riskFor: ["buyer"], severity: "high", template: "title", why: "No title review or objection process." },
      { category: "deposit", riskFor: ["seller"], severity: "medium", template: "deposit", why: "No deposit is required." },
      { category: "closing", riskFor: ["buyer", "seller"], severity: "high", template: "closing", why: "No closing date." },
      { category: "prorations", riskFor: ["buyer", "seller"], severity: "medium", template: "prorations", why: "No prorations or closing-cost allocation." },
      { category: "reps", riskFor: ["buyer"], severity: "high", template: "reps", why: "Seller makes no representations about the property." },
      { category: "operating_covenants", riskFor: ["buyer"], severity: "medium", template: "operating_covenants", when: "leased", why: "Nothing stops Seller from signing new leases or contracts before closing." },
      { category: "estoppel", riskFor: ["buyer"], severity: "medium", template: "estoppel_psa", when: "leased", why: "No tenant estoppel requirement for a leased property." },
      { category: "exchange_1031", riskFor: ["buyer", "seller"], severity: "low", template: "exchange_1031", why: "No 1031 exchange cooperation clause." },
      { category: "as_is", riskFor: ["seller"], severity: "high", template: "as_is", why: "No AS-IS clause: Seller may face implied-warranty claims." },
      { category: "default", riskFor: ["seller"], severity: "medium", template: "liquidated_damages", why: "No liquidated damages provision for Buyer's default." },
      { category: "casualty", riskFor: ["buyer"], severity: "medium", template: null, why: "Nothing addresses casualty or condemnation before closing." },
      { category: "brokers", riskFor: ["buyer", "seller"], severity: "low", template: "brokers", why: "No broker representation or commission clause." },
      { category: "governing_law", riskFor: ["buyer", "seller"], severity: "low", template: "governing_law", why: "No governing-law clause." },
      { category: "notices", riskFor: ["buyer", "seller"], severity: "low", template: "notices", why: "No notice provision." }
    ]
  };

  /* ------------------------------------------------------------------ *
   * Analysis
   * ------------------------------------------------------------------ */

  function analyze(input, opts) {
    opts = opts || {};
    var paragraphs = toParagraphs(input);
    var text = paragraphs.map(function (p) { return p.text; }).join("\n");
    var detected = detectDocType(paragraphs);
    var docType = opts.docType && opts.docType !== "auto" ? opts.docType : detected;
    var byCategory = {};
    for (var i = 0; i < paragraphs.length; i++) {
      var cats = categoriesFor(paragraphs[i], docType);
      paragraphs[i].categories = cats;
      for (var j = 0; j < cats.length; j++) {
        (byCategory[cats[j]] = byCategory[cats[j]] || []).push(i);
      }
    }
    return {
      paragraphs: paragraphs,
      text: text,
      docType: docType,
      detectedDocType: detected,
      california: isCalifornia(text),
      byCategory: byCategory
    };
  }

  function severityRank(s) { return s === "high" ? 0 : s === "medium" ? 1 : 2; }

  function ruleApplies(rule, doc, perspective) {
    if (!contains(rule.docTypes, doc.docType)) return false;
    if (rule.requiresCalifornia && !doc.california) return false;
    if (perspective !== "neutral" && !contains(rule.riskFor, perspective)) return false;
    if (rule.docUnless && rule.docUnless.test(doc.text)) return false;
    return true;
  }

  function scopeParagraphs(rule, doc) {
    if (!rule.scope) {
      var all = [];
      for (var i = 0; i < doc.paragraphs.length; i++) if (!doc.paragraphs[i].empty && !doc.paragraphs[i].isHeading) all.push(i);
      return all;
    }
    var scopes = typeof rule.scope === "string" ? [rule.scope] : rule.scope;
    var out = [];
    for (var s = 0; s < scopes.length; s++) out = out.concat(doc.byCategory[scopes[s]] || []);
    return uniq(out).sort(function (a, b) { return a - b; }).filter(function (idx) { return !doc.paragraphs[idx].isHeading; });
  }

  function makeAction(rule, p, start, matchText) {
    if (!rule.fix) return null;
    if (rule.fix.type === "replace") return { type: "replace", find: matchText, replace: rule.fix.text };
    return { type: "append", text: rule.fix.text };
  }

  function buildFinding(rule, p, start, end, doc) {
    var matchText = p.raw.slice(start, end);
    var tooLong = matchText.length > 240;
    var finding = {
      id: "f-" + rule.id + "-" + p.index + "-" + start,
      ruleId: rule.id,
      category: rule.category,
      categoryLabel: rule.category === "drafting" ? "Drafting" : categoryLabel(rule.category),
      severity: rule.severity,
      riskFor: rule.riskFor.slice(0),
      title: rule.title,
      why: rule.why,
      paragraphIndex: p.index,
      ref: refOf(p),
      section: sectionName(p),
      excerpt: clip(sentenceAt(p.raw, start), 260),
      matchText: tooLong ? matchText.slice(0, 240) : matchText,
      action: tooLong && rule.fix && rule.fix.type === "replace" ? null : makeAction(rule, p, start, matchText),
      status: "open"
    };
    if (rule.id === "number-mismatch") {
      finding.title = "Words and numbers don't match: “" + clip(matchText, 60) + "”";
    }
    finding.comment = "Clause Craft — " + finding.title + ". " + finding.why + (finding.action && finding.action.type === "replace" ? " Suggested: “" + finding.action.replace + "”." : finding.action ? " Suggested addition: “" + finding.action.text + "”" : "");
    return finding;
  }

  function runPhraseRule(rule, doc, out) {
    var idxs = scopeParagraphs(rule, doc);
    var count = 0, limit = rule.all ? 12 : 3;
    for (var k = 0; k < idxs.length && count < limit; k++) {
      var p = doc.paragraphs[idxs[k]];
      if (p.empty || (rule.skipSignatures && p.isSignature)) continue;
      if (rule.unless && rule.unless.test(p.text)) continue;
      var flags = rule.re.flags || ((rule.re.ignoreCase ? "i" : "") + (rule.re.multiline ? "m" : ""));
      var re = new RegExp(rule.re.source, flags.indexOf("g") === -1 ? flags + "g" : flags);
      var m;
      while ((m = re.exec(p.text)) !== null) {
        if (m[0] === "") { re.lastIndex++; continue; }
        if (!rule.test || rule.test(m, p.text)) {
          var st = m.index, en = m.index + m[0].length;
          if (rule.group && m[rule.group]) { st = m.index + m[0].lastIndexOf(m[rule.group]); en = st + m[rule.group].length; }
          out.push(buildFinding(rule, p, st, en, doc));
          count++;
          if (!rule.all) break;
        }
        if (count >= limit) break;
      }
    }
  }

  function runSectionRule(rule, doc, out) {
    var idxs = (doc.byCategory[rule.category] || []).filter(function (i) { return !doc.paragraphs[i].isHeading; });
    if (!idxs.length) return;
    var sectionText = idxs.map(function (i) { return doc.paragraphs[i].text; }).join("\n");
    if (rule.when && !rule.when.test(sectionText)) return;
    if (rule.unless && rule.unless.test(sectionText)) return;
    // Anchor on the first paragraph that satisfies "when" (or the first paragraph of the section).
    var anchor = doc.paragraphs[idxs[0]];
    var m = null;
    for (var k = 0; k < idxs.length; k++) {
      var p = doc.paragraphs[idxs[k]];
      var re = rule.when ? new RegExp(rule.when.source, rule.when.ignoreCase ? "i" : "") : null;
      var hit = re ? re.exec(p.text) : null;
      if (hit || !re) { anchor = p; m = hit; break; }
    }
    var start = m ? m.index : 0;
    var end = m ? m.index + m[0].length : Math.min(anchor.raw.length, 40);
    out.push(buildFinding(rule, anchor, start, end, doc));
  }

  function hasLeasedProperty(doc) {
    return /\b(the |existing |tenant )?leases?\b|\brent roll\b|\btenant\b/i.test(doc.text);
  }

  function findMissing(doc, perspective) {
    var list = EXPECTED[doc.docType] || [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var e = list[i];
      if (perspective !== "neutral" && !contains(e.riskFor, perspective)) continue;
      if (e.requiresCalifornia && !doc.california) continue;
      if (e.when === "operating_expenses" && !(doc.byCategory.operating_expenses || []).length) continue;
      if (e.when === "leased" && !hasLeasedProperty(doc)) continue;
      var present = (doc.byCategory[e.category] || []).some(function (idx) { return !doc.paragraphs[idx].isHeading; });
      if (present) continue;
      var t = e.template ? TEMPLATES[e.template] : null;
      out.push({
        id: "m-" + e.category + (e.template ? "-" + e.template : ""),
        category: e.category,
        label: categoryLabel(e.category),
        severity: e.severity,
        riskFor: e.riskFor.slice(0),
        title: "Missing: " + categoryLabel(e.category),
        why: e.why,
        templateId: e.template,
        templateTitle: t ? t.title : null,
        templateText: t ? t.text : null,
        status: "open"
      });
    }
    return out;
  }

  function clauseInventory(doc) {
    var out = [];
    for (var i = 0; i < CATEGORIES.length; i++) {
      var c = CATEGORIES[i];
      var idxs = (doc.byCategory[c.id] || []).filter(function (x) { return !doc.paragraphs[x].isHeading; });
      if (!idxs.length) continue;
      var first = doc.paragraphs[idxs[0]];
      out.push({ id: c.id, label: c.label, paragraphs: idxs, first: idxs[0], ref: refOf(first), count: idxs.length, excerpt: clip(first.raw, 180) });
    }
    out.sort(function (a, b) { return a.first - b.first; });
    return out;
  }

  function countOccurrences(hay, needle, before) {
    var n = 0, i = hay.indexOf(needle);
    while (i !== -1 && i < before) { n++; i = hay.indexOf(needle, i + 1); }
    return n;
  }

  /** Smallest window around [start,end) (<= 250 chars) that occurs only once in text. */
  function uniqueContext(text, start, end) {
    var a = start, b = end, guard = 0;
    while (guard++ < 400) {
      var ctx = text.slice(a, b);
      if (text.indexOf(ctx) === text.lastIndexOf(ctx)) return { context: ctx, offset: start - a };
      if (b - a >= 250 || (a === 0 && b === text.length)) break;
      if (b < text.length) b++;
      if (text.indexOf(text.slice(a, b)) === text.lastIndexOf(text.slice(a, b))) continue;
      if (a > 0) a--;
    }
    return { context: text.slice(start, end), offset: 0 };
  }

  /**
   * Re-locate a finding inside the paragraph's current text (after other edits).
   * Returns { target, context, offset, occurrence, start } or null when the issue is no longer present.
   */
  function locate(finding, currentText) {
    var raw = String(currentText == null ? "" : currentText);
    var text = norm(raw);
    var rule = null;
    for (var i = 0; i < RULES.length; i++) if (RULES[i].id === finding.ruleId) { rule = RULES[i]; break; }
    if (!rule) return null;
    if (rule.kind === "section" || (finding.action && finding.action.type === "append")) {
      if (finding.action && finding.action.type === "append" && norm(finding.action.text) && text.indexOf(norm(finding.action.text).slice(0, 60)) !== -1) return null;
      var idx = finding.matchText ? raw.indexOf(finding.matchText) : -1;
      return { target: finding.matchText, context: finding.matchText, offset: 0, occurrence: 0, start: Math.max(idx, 0) };
    }
    if (rule.unless && rule.unless.test(text)) return null;
    var flags = rule.re.flags || ((rule.re.ignoreCase ? "i" : "") + (rule.re.multiline ? "m" : ""));
    var re = new RegExp(rule.re.source, flags.indexOf("g") === -1 ? flags + "g" : flags), m, best = null;
    while ((m = re.exec(text)) !== null) {
      if (m[0] === "") { re.lastIndex++; continue; }
      if (rule.test && !rule.test(m, text)) continue;
      var st = m.index, en = m.index + m[0].length;
      if (rule.group && m[rule.group]) { st = m.index + m[0].lastIndexOf(m[rule.group]); en = st + m[rule.group].length; }
      var cand = { start: st, end: en, target: raw.slice(st, en) };
      if (!best) best = cand;
      if (cand.target === finding.matchText) { best = cand; break; }
    }
    if (!best) return null;
    var ctx = uniqueContext(raw, best.start, best.end);
    return {
      target: best.target,
      context: ctx.context,
      offset: ctx.offset,
      occurrence: countOccurrences(raw, best.target, best.start),
      start: best.start
    };
  }

  /** Apply a finding's fix to a plain string (used by the browser app and tests). */
  function applyFixToText(finding, currentText) {
    if (!finding.action) return null;
    if (finding.action.type === "append") {
      var loc = locate(finding, currentText);
      if (!loc) return null;
      return String(currentText).replace(/\s+$/, "") + " " + finding.action.text;
    }
    var l = locate(finding, currentText);
    if (!l) return null;
    return currentText.slice(0, l.start) + finding.action.replace + currentText.slice(l.start + l.target.length);
  }

  /**
   * Review a document.
   * @param input  array of paragraph strings, or one string
   * @param opts   { perspective: "tenant"|"landlord"|"buyer"|"seller"|"neutral"|"auto", docType: "lease"|"psa"|"auto" }
   */
  function review(input, opts) {
    opts = opts || {};
    var doc = analyze(input, opts);
    var perspective = opts.perspective && opts.perspective !== "auto" ? opts.perspective : defaultPerspective(doc.docType);
    var findings = [];
    for (var r = 0; r < RULES.length; r++) {
      var rule = RULES[r];
      if (!ruleApplies(rule, doc, perspective)) continue;
      if (rule.kind === "section") runSectionRule(rule, doc, findings);
      else runPhraseRule(rule, doc, findings);
    }
    // One finding per rule+paragraph+offset; order by severity, then document position.
    var seen = {};
    findings = findings.filter(function (f) { if (seen[f.id]) return false; seen[f.id] = 1; return true; });
    findings.sort(function (a, b) {
      return severityRank(a.severity) - severityRank(b.severity) || a.paragraphIndex - b.paragraphIndex;
    });
    var missing = doc.docType === "unknown" ? [] : findMissing(doc, perspective);
    missing.sort(function (a, b) { return severityRank(a.severity) - severityRank(b.severity); });
    var clauses = clauseInventory(doc);
    var stats = { high: 0, medium: 0, low: 0, missing: missing.length, clauses: clauses.length, paragraphs: 0 };
    for (var f = 0; f < findings.length; f++) stats[findings[f].severity]++;
    for (var q = 0; q < doc.paragraphs.length; q++) if (!doc.paragraphs[q].empty) stats.paragraphs++;
    return {
      engine: VERSION,
      docType: doc.docType,
      detectedDocType: doc.detectedDocType,
      perspective: perspective,
      perspectiveLabel: perspectiveLabel(perspective),
      california: doc.california,
      findings: findings,
      missing: missing,
      clauses: clauses,
      stats: stats
    };
  }

  /* ------------------------------------------------------------------ *
   * Ask the document (keyword retrieval with cited paragraphs)
   * ------------------------------------------------------------------ */

  var MONEY_RE = /\$\s?\d[\d,]*(?:\.\d{2})?(?:\s*(?:per|\/)\s*(?:rentable\s+|usable\s+)?(?:square foot|sq\.?\s*ft\.?|sf|rsf|month|year|annum)(?:\s+per\s+(?:month|year|annum))?)?/gi;
  var PERCENT_RE = new RegExp("\\b(?:" + WORD_NUMBER_RE + "\\s+percent\\s*\\(\\s*\\d+(?:\\.\\d+)?\\s*%\\s*\\)|\\d+(?:\\.\\d+)?\\s*%)", "gi");
  var DURATION_RE = new RegExp("\\b(?:" + WORD_NUMBER_RE + "\\s*\\(\\s*\\d+\\s*\\)|\\d+)\\s+(?:calendar\\s+|business\\s+)?(?:days?|months?|years?)\\b", "gi");
  var DATE_RE = /\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b|\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/g;
  var STATE_RE = /\blaws of the state of ([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)|\blaws of ([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b/;

  var INTENTS = [
    { id: "risks", q: /\b(risks?|red flags?|issues?|problems?|concerns?|watch out|summari[sz]e)\b/i, special: "risks" },
    { id: "nnn", q: /\b(nnn|triple net|cam|operating expenses?|opex|common area|pass[- ]?throughs?|taxes)\b/i, keys: ["operating expenses", "triple net", "nnn", "common area", "tenant's share", "pro rata", "taxes", "insurance costs", "management fee"], cats: ["operating_expenses"], extract: ["percent", "money"] },
    { id: "deposit", q: /\b(deposit|earnest money)\b/i, keys: ["security deposit", "deposit", "earnest money"], cats: ["security_deposit", "deposit"], extract: ["money", "duration"] },
    { id: "terminate", label: "termination right", q: /\b(terminat\w*|cancel\w*|get out|early out|walk away|exit)\b/i, keys: ["terminate", "termination", "cancel"], extract: ["duration"], special: "terminate" },
    { id: "governing_law", label: "governing-law clause", q: /\b(governing law|which state|jurisdiction|venue|laws? of)\b/i, keys: ["governed by", "laws of", "governing law", "venue"], cats: ["governing_law"], extract: ["state"] },
    { id: "insurance", label: "insurance requirement", q: /\b(insurance|insured|coverage)\b/i, keys: ["insurance", "commercial general liability", "per occurrence", "additional insured"], cats: ["insurance"], extract: ["money"] },
    { id: "assignment", q: /\b(assign\w*|sublet\w*|sublease|transfer)\b/i, keys: ["assign", "sublet", "sublease", "transfer", "consent"], cats: ["assignment"], extract: [] },
    { id: "renewal", label: "renewal or extension option", q: /\b(renew\w*|extend|extension|options?)\b/i, keys: ["option to extend", "renewal", "extension", "option"], cats: ["renewal_option"], extract: ["duration"] },
    { id: "escalation", q: /\b(escalat\w*|increases?|bumps?|annual adjustment|cpi)\b/i, keys: ["increase", "escalat", "anniversary", "percent", "consumer price index"], cats: ["escalation"], extract: ["percent"] },
    { id: "holdover", label: "holdover provision", q: /\bhold\s*-?over\b/i, keys: ["holdover", "hold over", "holding over", "remains in possession"], cats: ["holdover"], extract: ["percent"] },
    { id: "default", q: /\b(default|cure|late)\b/i, keys: ["default", "cure", "late charge", "notice"], cats: ["default", "late_charge"], extract: ["duration", "percent"] },
    { id: "repairs", q: /\b(repairs?|maintain|maintenance|hvac|roof|structural)\b/i, keys: ["repair", "maintain", "hvac", "roof", "structural", "replace"], cats: ["repairs"], extract: [] },
    { id: "parking", label: "parking provision", q: /\bparking\b/i, keys: ["parking", "spaces", "stalls"], cats: ["parking"], extract: ["money"] },
    { id: "guaranty", label: "guaranty", q: /\bguarant\w*\b/i, keys: ["guaranty", "guarantor", "guarantee"], cats: ["guaranty"], extract: [] },
    { id: "due_diligence", label: "due diligence period", q: /\b(due diligence|inspection period|feasibility|contingenc\w*)\b/i, keys: ["due diligence", "inspection", "feasibility", "contingency"], cats: ["due_diligence"], extract: ["duration"] },
    { id: "closing", q: /\b(clos(e|ing)|escrow)\b/i, keys: ["close of escrow", "closing date", "closing", "escrow"], cats: ["closing"], extract: ["duration", "date"] },
    { id: "price", q: /\b(purchase price|price)\b/i, keys: ["purchase price"], cats: ["purchase_price"], extract: ["money"] },
    { id: "title", q: /\btitle\b/i, keys: ["title", "permitted exceptions", "survey"], cats: ["title"], extract: [] },
    { id: "rent", q: /\b(rent|payments?|how much|monthly)\b/i, keys: ["base rent", "monthly installments", "per rentable square foot", "per month"], cats: ["rent"], extract: ["money"] },
    { id: "term", q: /\b(term|how long|expir\w*|commenc\w*|start date|end date)\b/i, keys: ["term", "commencement date", "expiration date", "months", "years"], cats: ["term"], extract: ["duration", "date"] },
    { id: "use", q: /\b(use|permitted use|allowed to do)\b/i, keys: ["use the premises", "permitted use", "solely for", "general office"], cats: ["use"], extract: [] },
    { id: "parties", q: /\b(who are the parties|parties|who is the (landlord|tenant|buyer|seller))\b/i, special: "parties" }
  ];

  var STOP = ("a an and are as at be by can do does for from has have how i if in is it its me my of on or our shall so " +
    "that the their there this to under was we what when where which who whom why will with would you your about any").split(" ");

  function tokens(s) {
    return norm(s).toLowerCase().replace(/[^a-z0-9$%'\s-]/g, " ").split(/\s+/).filter(function (t) {
      return t && t.length > 1 && STOP.indexOf(t) === -1;
    });
  }

  function extractFacts(text, kinds) {
    var facts = [];
    function grab(re) { var m = text.match(re); return m ? uniq(m.map(trim)) : []; }
    for (var i = 0; i < kinds.length; i++) {
      var k = kinds[i];
      if (k === "money") facts = facts.concat(grab(MONEY_RE));
      else if (k === "percent") facts = facts.concat(grab(PERCENT_RE));
      else if (k === "duration") facts = facts.concat(grab(DURATION_RE));
      else if (k === "date") facts = facts.concat(grab(DATE_RE));
      else if (k === "state") { var s = STATE_RE.exec(text); if (s) facts.push(s[1] || s[2]); }
    }
    return uniq(facts).slice(0, 5);
  }

  function scoreParagraph(p, intent, qTokens) {
    if (p.empty || p.isHeading || p.isSignature) return 0;
    var t = p.text.toLowerCase();
    var score = 0, i;
    if (intent && intent.keys) {
      for (i = 0; i < intent.keys.length; i++) if (t.indexOf(intent.keys[i]) !== -1) score += 3;
      if (intent.cats) for (i = 0; i < intent.cats.length; i++) if (contains(p.categories || [], intent.cats[i])) score += 6;
    }
    for (i = 0; i < qTokens.length; i++) if (t.indexOf(qTokens[i]) !== -1) score += 1;
    return score;
  }

  function isHeadingSentence(s) { return /^\s*(?:(?:article|section)\s+)?\d+(?:\.\d+)*\.?\s+[A-Z][^.]{0,70}\.\s*$/.test(s); }

  function keySentences(text, intent, qTokens) {
    return sentences(text).filter(function (s) {
      if (isHeadingSentence(s.text)) return false;
      var t = s.text.toLowerCase(), j;
      if (intent && intent.keys) for (j = 0; j < intent.keys.length; j++) if (t.indexOf(intent.keys[j]) !== -1) return true;
      for (j = 0; j < qTokens.length; j++) if (t.indexOf(qTokens[j]) !== -1) return true;
      return false;
    });
  }

  var TERMINATION_RE = /\b(Landlord|Tenant|Buyer|Seller|either party|each party)\b[^.;]{0,80}?\b(?:may|shall have the right to|has the right to|can)\b[^.;]{0,60}?\bterminat\w*\b|\b(?:recapture)\b/gi;

  function answerTermination(doc) {
    var who = [], sources = [];
    for (var i = 0; i < doc.paragraphs.length; i++) {
      var p = doc.paragraphs[i];
      if (p.empty || p.isHeading || p.isSignature) continue;
      var re = new RegExp(TERMINATION_RE.source, "gi"), m;
      while ((m = re.exec(p.text)) !== null) {
        var party = m[1] ? cap(m[1].toLowerCase()) : "Landlord";
        var where = p.heading ? cap(sectionName(p).toLowerCase()) : "this provision";
        who.push(party + " \u2014 " + (m[1] ? where : where + " (recapture)") + " (" + refOf(p) + ")");
        sources.push({ paragraphIndex: p.index, ref: refOf(p), excerpt: clip(sentenceAt(p.raw, m.index), 220), score: 10 });
      }
    }
    who = uniq(who);
    if (!who.length) return { found: false, intent: "terminate", answer: "No party has an express right to terminate early in this document (other than any default remedies).", sources: [], facts: [] };
    var parties = uniq(who.map(function (w) { return w.split(" \u2014 ")[0]; }));
    var note = "";
    if (doc.docType === "lease" && !contains(parties, "Tenant") && !contains(parties, "Either party")) note = "\nTenant has no termination right in this draft.";
    if (doc.docType === "psa" && !contains(parties, "Buyer") && !contains(parties, "Either party")) note = "\nBuyer has no express termination right in this draft.";
    return { found: true, intent: "terminate", answer: "Termination rights:\n" + who.join("\n") + note, sources: uniqSources(sources), facts: [] };
  }

  function uniqSources(list) {
    var seen = {}, out = [];
    for (var i = 0; i < list.length; i++) { if (seen[list[i].paragraphIndex]) continue; seen[list[i].paragraphIndex] = 1; out.push(list[i]); }
    return out.slice(0, 4);
  }

  function bestSentence(text, intent, qTokens) {
    var ss = sentences(text).filter(function (s) { return !isHeadingSentence(s.text); });
    if (!ss.length) ss = sentences(text);
    var best = ss[0], bestScore = -1;
    for (var i = 0; i < ss.length; i++) {
      var s = ss[i].text.toLowerCase(), sc = 0, j;
      if (intent && intent.keys) for (j = 0; j < intent.keys.length; j++) if (s.indexOf(intent.keys[j]) !== -1) sc += 2;
      for (j = 0; j < qTokens.length; j++) if (s.indexOf(qTokens[j]) !== -1) sc += 1;
      if (sc > bestScore) { bestScore = sc; best = ss[i]; }
    }
    return best ? best.text : text;
  }

  function ask(question, input, opts) {
    opts = opts || {};
    var q = trim(question);
    if (!q) return { found: false, answer: "Type a question about this document.", sources: [], facts: [] };
    var doc = analyze(input, opts);
    var intent = null;
    for (var i = 0; i < INTENTS.length; i++) if (INTENTS[i].q.test(q)) { intent = INTENTS[i]; break; }
    var qTokens = tokens(q);

    if (intent && intent.special === "risks") {
      var rv = review(input, opts);
      var top = rv.findings.slice(0, 5);
      if (!top.length && !rv.missing.length) return { found: true, intent: "risks", answer: "No issues flagged for the " + rv.perspectiveLabel.toLowerCase() + " side. Clause Craft checks common CRE clauses only; have counsel review.", sources: [], facts: [] };
      var lines = top.map(function (f) { return cap(f.severity) + ": " + f.title + " (" + f.ref + ")"; });
      if (rv.missing.length) lines.push("Missing: " + rv.missing.slice(0, 4).map(function (m) { return m.label; }).join(", ") + (rv.missing.length > 4 ? "…" : ""));
      return {
        found: true, intent: "risks",
        answer: "Top issues for the " + rv.perspectiveLabel.toLowerCase() + " side:\n" + lines.join("\n"),
        sources: top.map(function (f) { return { paragraphIndex: f.paragraphIndex, ref: f.ref, excerpt: f.excerpt, score: 10 }; }),
        facts: []
      };
    }

    if (intent && intent.special === "terminate") return answerTermination(doc);

    if (intent && intent.special === "parties") {
      var parties = extractParties(doc);
      if (!parties.length) return { found: false, intent: "parties", answer: "No defined parties found (look for “(\"Landlord\")”-style definitions).", sources: [], facts: [] };
      return {
        found: true, intent: "parties",
        answer: parties.map(function (p) { return p.role + ": " + p.name; }).join("\n"),
        sources: uniq(parties.map(function (p) { return p.paragraphIndex; })).map(function (idx) { return { paragraphIndex: idx, ref: refOf(doc.paragraphs[idx]), excerpt: clip(doc.paragraphs[idx].raw, 220), score: 10 }; }),
        facts: []
      };
    }

    var scored = [];
    for (var k = 0; k < doc.paragraphs.length; k++) {
      var sc = scoreParagraph(doc.paragraphs[k], intent, qTokens);
      if (sc > 0) scored.push({ p: doc.paragraphs[k], score: sc });
    }
    scored.sort(function (a, b) { return b.score - a.score || a.p.index - b.p.index; });
    var minScore = intent ? 3 : 2;
    var hits = scored.filter(function (s) { return s.score >= minScore; }).slice(0, 3);
    var catHit = false;
    if (intent && intent.cats) for (var ci2 = 0; ci2 < intent.cats.length; ci2++) if ((doc.byCategory[intent.cats[ci2]] || []).length) catHit = true;
    if (intent && intent.cats && !catHit && intent.label) {
      return { found: false, intent: intent.id, answer: "No " + intent.label + " found in this document. If you need one, the Review tab lists missing clauses you can insert.", sources: hits.slice(0, 2).map(function (h) { return { paragraphIndex: h.p.index, excerpt: clip(h.p.raw, 220), score: h.score }; }), facts: [] };
    }
    if (!hits.length) {
      return { found: false, intent: intent ? intent.id : null, answer: "Couldn't find language on that in this document. Try different words, or check the Clauses tab for what was detected.", sources: [], facts: [] };
    }
    var top1 = hits[0].p;
    var sentence = bestSentence(top1.raw, intent, qTokens);
    var kinds = intent ? intent.extract || [] : [];
    var factText = keySentences(top1.text, intent, qTokens).map(function (s) { return s.text; }).join(" ");
    var facts = extractFacts(factText || sentence, kinds);
    var answer = refOf(top1) + ": “" + clip(sentence, 320) + "”";
    if (facts.length) answer = "Key terms: " + facts.join("; ") + "\n" + answer;
    return {
      found: true,
      intent: intent ? intent.id : null,
      answer: answer,
      facts: facts,
      sources: hits.map(function (h) { return { paragraphIndex: h.p.index, ref: refOf(h.p), excerpt: clip(h.p.raw, 220), score: h.score }; })
    };
  }

  /* ------------------------------------------------------------------ *
   * Summary
   * ------------------------------------------------------------------ */

  var ROLE_WORDS = /\(\s*(?:the\s+|hereinafter\s+)?"?(Landlord|Tenant|Lessor|Lessee|Buyer|Purchaser|Seller|Guarantor)"?\s*\)/g;
  var ENTITY_RE = /([A-Z][A-Za-z0-9&.'-]*(?:\s+[A-Z0-9][A-Za-z0-9&.'-]*){0,7},?\s+(?:LLC|L\.L\.C\.|Inc\.?|INC\.?|Incorporated|Corp\.?|Corporation|L\.?P\.?|LLP|Ltd\.?|Trust|Company))/g;

  function extractParties(doc) {
    var out = [], seenRole = {};
    for (var i = 0; i < doc.paragraphs.length && out.length < 6; i++) {
      var p = doc.paragraphs[i];
      if (p.empty) continue;
      var re = new RegExp(ROLE_WORDS.source, "g"), m;
      while ((m = re.exec(p.raw)) !== null) {
        var role = m[1] === "Lessor" ? "Landlord" : m[1] === "Lessee" ? "Tenant" : m[1] === "Purchaser" ? "Buyer" : m[1];
        if (seenRole[role]) continue;
        var before = norm(p.raw.slice(Math.max(0, m.index - 220), m.index));
        var cut = Math.max(before.lastIndexOf(" between "), before.lastIndexOf(" and "), before.lastIndexOf(" by "));
        var seg = cut >= 0 ? before.slice(cut).replace(/^\s*(between|and|by)\s+/i, "") : before;
        var em = null, mm, ere = new RegExp(ENTITY_RE.source, "g");
        while ((mm = ere.exec(seg)) !== null) em = mm[1];
        var name = em || trim(seg.split(/,\s+(?:a|an)\s+/i)[0]);
        name = trim(name.replace(/^[,;\s]+|[,;\s]+$/g, ""));
        if (name && name.length <= 120) {
          seenRole[role] = 1;
          out.push({ role: role, name: name, paragraphIndex: p.index });
        }
      }
    }
    return out;
  }

  function tidy(v) {
    v = trim(String(v || "").replace(/\s*\((?:the\s+)?"[^"]{1,40}"\)/g, "").replace(/\s+/g, " "));
    if (/^\(.*\)$/.test(v)) v = trim(v.slice(1, -1));
    var open = (v.match(/\(/g) || []).length, close = (v.match(/\)/g) || []).length;
    while (close < open) { v += ")"; close++; }
    return v;
  }

  /** First money amount after (or just before) a keyword inside the given categories. */
  function moneyNear(doc, cats, keyRe) {
    for (var c = 0; c < cats.length; c++) {
      var idxs = doc.byCategory[cats[c]] || [];
      for (var i = 0; i < idxs.length; i++) {
        var p = doc.paragraphs[idxs[i]];
        var re = new RegExp(keyRe.source, "gi"), m;
        while ((m = re.exec(p.text)) !== null) {
          var after = p.text.slice(m.index, m.index + 160).match(/\$\s?\d[\d,]*(?:\.\d{2})?/);
          if (after) return { value: after[0], paragraphIndex: p.index };
        }
      }
    }
    return null;
  }

  function firstIn(doc, cats, re, extraRe) {
    for (var c = 0; c < cats.length; c++) {
      var idxs = doc.byCategory[cats[c]] || [];
      for (var i = 0; i < idxs.length; i++) {
        var p = doc.paragraphs[idxs[i]];
        if (p.isHeading) continue;
        if (extraRe && !extraRe.test(p.text)) continue;
        var m = p.text.match(re);
        if (m) return { value: trim(m[0]), paragraphIndex: p.index };
      }
    }
    return null;
  }

  function summarize(input, opts) {
    opts = opts || {};
    var doc = analyze(input, opts);
    var rv = review(input, opts);
    var fields = [];
    function add(label, hit) { if (hit && hit.value) fields.push({ label: label, value: hit.value, paragraphIndex: hit.paragraphIndex }); }

    var parties = extractParties(doc);
    for (var i = 0; i < parties.length; i++) fields.push({ label: parties[i].role, value: parties[i].name, paragraphIndex: parties[i].paragraphIndex });

    if (doc.docType === "lease" || doc.docType === "unknown") {
      add("Premises", firstIn(doc, ["premises"], /\bSuite\s+[\w-]+\b[\s\S]{0,140}?\b\d{5}\b|\bSuite\s+[\w-]+/i));
      add("Size", firstIn(doc, ["premises"], /\b\d[\d,]*\s+(?:rentable|usable)\s+square\s+feet\b/i));
      add("Term", firstIn(doc, ["term"], new RegExp("\\b(?:" + WORD_NUMBER_RE + "\\s*\\(\\s*\\d+\\s*\\)|\\d+)\\s+(?:months|years)\\b", "i")));
      add("Commencement", firstIn(doc, ["term"], /\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b/));
      add("Base rent", firstIn(doc, ["rent"], /\$\s?\d[\d,]*(?:\.\d{2})?[\s\S]{0,60}?\bper (?:month|rentable square foot per month)\b|\$\s?\d[\d,]*(?:\.\d{2})?/i));
      add("Increases", firstIn(doc, ["escalation"], PERCENT_RE));
      add("Expenses", firstIn(doc, ["operating_expenses"], /\btriple net\b[^.]*|\bNNN\b|\bmodified gross\b|\bfull service\b|\btenant's share\s*\(\s*\d+(?:\.\d+)?\s*%\s*\)/i));
      add("Security deposit", firstIn(doc, ["security_deposit"], /\$\s?\d[\d,]*(?:\.\d{2})?/));
      add("Options", firstIn(doc, ["renewal_option", "early_termination"], /\boption to (?:extend|renew)\b[^.]*|\bright to terminate\b[^.]*/i));
    }
    if (doc.docType === "psa") {
      add("Property", firstIn(doc, ["purchase_price", "premises"], /\b\d{2,6}\s+[A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){0,3}\s+(?:Road|Rd\.?|Street|St\.?|Avenue|Ave\.?|Boulevard|Blvd\.?|Drive|Dr\.?|Way|Parkway|Pkwy\.?|Lane|Ln\.?|Court|Ct\.?)\b[^"(]*?\b\d{5}\b/));
      if (!fields.some(function (f) { return f.label === "Property"; })) {
        for (var pp = 0; pp < doc.paragraphs.length; pp++) {
          var am = doc.paragraphs[pp].text.match(/\b\d{2,6}\s+[A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){0,3}\s+(?:Road|Street|Avenue|Boulevard|Drive|Way|Parkway|Lane|Court)\b,?\s+[A-Z][A-Za-z]+(?:\s+[A-Z][a-z]+)?,\s+California\s+\d{5}\b/);
          if (am) { fields.push({ label: "Property", value: am[0], paragraphIndex: pp }); break; }
        }
      }
      add("Purchase price", firstIn(doc, ["purchase_price"], /\(\s*\$\s?\d[\d,]*(?:\.\d{2})?\s*\)|\$\s?\d[\d,]*(?:\.\d{2})?/));
      add("Deposit", moneyNear(doc, ["deposit"], /\b(?:shall deposit|earnest money|deposit of)\b/i) || moneyNear(doc, ["deposit"], /\bdeposit\b/i));
      add("Due diligence", firstIn(doc, ["due_diligence"], DURATION_RE));
      add("Closing", firstIn(doc, ["closing"], DURATION_RE) || firstIn(doc, ["closing"], DATE_RE));
      add("Financing", firstIn(doc, ["financing"], /\bnot contingent\b[^.]*|\bcontingent\b[^.]*/i));
      add("Survival", firstIn(doc, ["reps"], /\bsurviv\w*\b[^.]*/i));
    }
    add("Governing law", firstIn(doc, ["governing_law"], /\blaws of (?:the State of )?[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?/));

    fields.forEach(function (f) { f.value = clip(tidy(f.value), 160); });

    var risks = rv.findings.filter(function (f) { return f.severity !== "low"; }).slice(0, 6).map(function (f) {
      return { severity: f.severity, title: f.title, paragraphIndex: f.paragraphIndex, ref: f.ref };
    });
    var missing = rv.missing.map(function (m) { return m.label; });
    var typeLabel = doc.docType === "lease" ? "Lease" : doc.docType === "psa" ? "Purchase agreement" : "Document";

    var lines = ["CLAUSE CRAFT SUMMARY — " + typeLabel + " (reviewed for the " + rv.perspectiveLabel.toLowerCase() + " side)"];
    fields.forEach(function (f) { lines.push(f.label + ": " + f.value); });
    if (risks.length) {
      lines.push("Key risks:");
      risks.forEach(function (r) { lines.push("• " + cap(r.severity) + " — " + r.title + " (" + r.ref + ")"); });
    }
    if (missing.length) lines.push("Missing: " + missing.join(", "));
    lines.push("Generated by Clause Craft (rule-based). Not legal advice — have counsel review.");

    return {
      docType: doc.docType,
      typeLabel: typeLabel,
      perspective: rv.perspective,
      perspectiveLabel: rv.perspectiveLabel,
      fields: fields,
      risks: risks,
      missing: missing,
      stats: rv.stats,
      text: lines.join("\n")
    };
  }

  return {
    VERSION: VERSION,
    CATEGORIES: CATEGORIES,
    RULES: RULES,
    TEMPLATES: TEMPLATES,
    PERSPECTIVES: PERSPECTIVES,
    EXPECTED: EXPECTED,
    toParagraphs: toParagraphs,
    detectDocType: detectDocType,
    defaultPerspective: defaultPerspective,
    perspectiveLabel: perspectiveLabel,
    categoryLabel: categoryLabel,
    review: review,
    locate: locate,
    applyFixToText: applyFixToText,
    ask: ask,
    summarize: summarize,
    wordsToNumber: wordsToNumber,
    norm: norm
  };
});
