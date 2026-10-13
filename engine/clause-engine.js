/**
 * Clause Craft - CRE contract review engine (rule-based, runs entirely on the device).
 *
 * Shared by the browser app (index.html + app.js) and the Word add-in (word/taskpane.html).
 * Input is a list of paragraphs (Word body paragraphs) or a plain-text string.
 * Nothing here makes network calls, stores data, or uses eval.
 *
 * Demo only - not legal advice. Findings are heuristics for a broker's first pass;
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
      .replace(/[\u2018\u2019\u02BC\u2032]/g, "'")
      .replace(/[\u201C\u201D\u2033]/g, "\"")
      .replace(/[\u00A0\u2007\u202F]/g, " ")
      .replace(/[\u2010\u2011\u2012\u2013\u2014]/g, "-")
      .replace(/[\u000B\u0007]/g, " ");
  }

  function trim(s) { return String(s == null ? "" : s).replace(/^\s+|\s+$/g, ""); }

  function clip(s, n) {
    s = trim(String(s == null ? "" : s).replace(/\s+/g, " "));
    return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, "") + "\u2026" : s;
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
    var texts = [];
    for (var k = 0; k < list.length; k++) texts.push(norm(String(list[k] == null ? "" : list[k])));
    var memo = memoMask(texts);
    var out = [], currentHeading = "", currentNo = "";
    for (var i = 0; i < list.length; i++) {
      var raw = String(list[i] == null ? "" : list[i]);
      var text = texts[i];
      if (memo[i]) {
        // Clause Craft's own notes: kept in place (indexes stay aligned with Word) but read as blank.
        out.push({ index: i, raw: raw, text: "", memoText: text, memo: true, empty: true, heading: currentHeading, sectionNo: currentNo, inlineHeading: "", isHeading: false, isSignature: false });
        continue;
      }
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

  // Clause Craft's own notes (a summary or an issues list inserted into the document) are not contract
  // language: they are skipped so they don't satisfy or trigger checks on the next review. A block runs
  // from its title to its closing line; if someone deleted the closing line, it stops at the first line
  // that isn't shaped like one of its own lines, so the contract after it is still read.
  var MEMO_KINDS = [
    {
      start: /^\s*CLAUSE CRAFT SUMMARY\b/,
      second: null,
      end: /^\s*Generated by Clause Craft\b/,
      line: /^\s*(?:$|Key risks:|Missing:|[A-Z][A-Za-z ]{1,30}:\s+\S|\u2022\s*(?:High|Medium|Low)\b)/
    },
    {
      start: /^\s*[A-Z' ]{0,40}COMMENTS ON THE (?:LEASE|PURCHASE AGREEMENT|AGREEMENT)\s*$/,
      second: /^\s*Prepared with Clause Craft\b/,
      end: /^\s*End of comments\.?\s*$/,
      line: /^\s*(?:$|Prepared with Clause Craft\b|Deal notes:|\d+\.\s+(?:\u00A7|\u00B6|New\b))/
    }
  ];

  function memoMask(texts) {
    var mask = [];
    for (var i = 0; i < texts.length; i++) {
      for (var k = 0; k < MEMO_KINDS.length; k++) {
        var kind = MEMO_KINDS[k];
        if (!kind.start.test(texts[i]) || (kind.second && !kind.second.test(texts[i + 1] || ""))) continue;
        mask[i] = true;
        var j = i + 1;
        for (; j < texts.length; j++) {
          if (kind.end.test(texts[j])) { mask[j] = true; break; }
          if (!kind.line.test(texts[j])) { j--; break; }
          mask[j] = true;
        }
        i = Math.min(j, texts.length - 1);
        break;
      }
    }
    return mask;
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
    casp: { title: "CASp disclosure (Cal. Civ. Code \u00A71938)", category: "casp", text: "CASp Disclosure. For purposes of California Civil Code Section 1938, Landlord states that the Premises have not undergone inspection by a Certified Access Specialist (CASp). A Certified Access Specialist (CASp) can inspect the subject premises and determine whether the subject premises comply with all of the applicable construction-related accessibility standards under state law. Although state law does not require a CASp inspection of the subject premises, the commercial property owner or lessor may not prohibit the lessee or tenant from obtaining a CASp inspection of the subject premises for the occupancy or potential occupancy of the lessee or tenant, if requested by the lessee or tenant. The parties shall mutually agree on the arrangements for the time and manner of the CASp inspection, the payment of the fee for the CASp inspection, and the cost of making any repairs necessary to correct violations of construction-related accessibility standards within the premises." },
    security_deposit: { title: "Security deposit", category: "security_deposit", text: "Security Deposit. Upon execution of this Lease, Tenant shall deposit with Landlord $[amount] as security for Tenant's performance. Landlord may apply the Security Deposit to cure any default beyond applicable notice and cure periods, and Tenant shall restore any amount so applied within ten (10) days after written demand. Landlord shall return any unapplied balance within thirty (30) days after the expiration of the Term." },
    holdover: { title: "Holdover", category: "holdover", text: "Holding Over. If Tenant remains in possession after the expiration or termination of the Term without Landlord's written consent, Tenant shall be a tenant at sufferance and shall pay Base Rent equal to one hundred fifty percent (150%) of the Base Rent in effect immediately before the holdover, prorated on a daily basis." },
    estoppel: { title: "Estoppel certificates", category: "estoppel", text: "Estoppel Certificates. Within ten (10) business days after written request, Tenant shall execute and deliver a certificate stating whether this Lease is in full force and effect, the dates through which Rent has been paid, and whether, to Tenant's knowledge, either party is in default." },
    insurance: { title: "Insurance", category: "insurance", text: "Insurance. Tenant shall maintain commercial general liability insurance with limits of not less than $[1,000,000] per occurrence and $[2,000,000] in the aggregate, naming Landlord as an additional insured, and property insurance covering Tenant's personal property and improvements at full replacement cost. Landlord shall maintain property insurance on the Building at full replacement cost and commercial general liability insurance on the common areas." },
    indemnity: { title: "Mutual indemnity", category: "indemnity", text: "Indemnity. Subject to the waiver of subrogation in this Lease, Tenant shall indemnify, defend and hold Landlord harmless from claims arising from the negligence or willful misconduct of Tenant or its agents, employees or contractors in or about the Premises, and Landlord shall indemnify, defend and hold Tenant harmless from claims arising from the negligence or willful misconduct of Landlord or its agents, employees or contractors in or about the Building." },
    default_notice: { title: "Notice and cure", category: "default", text: "Default. Neither party shall be in default unless it fails to cure (a) a monetary default within ten (10) days after written notice, or (b) any other default within thirty (30) days after written notice, or such longer period as is reasonably necessary provided the cure is commenced within such thirty (30) day period and diligently pursued." },
    assignment: { title: "Assignment and permitted transfers", category: "assignment", text: "Assignment and Subletting. Tenant may not assign this Lease or sublet the Premises without Landlord's consent, which shall not be unreasonably withheld, conditioned or delayed and shall be granted or denied within fifteen (15) days after request. Notwithstanding the foregoing, Tenant may, without Landlord's consent, assign this Lease or sublet all or part of the Premises to an affiliate, to an entity resulting from a merger or consolidation, or to a purchaser of all or substantially all of Tenant's assets or stock (each, a \"Permitted Transfer\"), upon written notice to Landlord. No Permitted Transfer shall require Landlord's consent or trigger any recapture or profit-sharing right." },
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
      id: "alterations-sole-discretion", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "alterations", scope: "alterations",
      re: /\bsole\s+(?:and\s+absolute\s+)?discretion\b|\babsolute\s+discretion\b/i,
      title: "Alterations need Landlord's consent in its sole discretion",
      why: "Even minor work such as paint, carpet or cabling could be refused. Ask for reasonable consent and no consent for cosmetic changes.",
      fix: { type: "replace", text: "reasonable discretion (such consent not to be unreasonably withheld, conditioned or delayed, and no consent shall be required for cosmetic alterations such as painting, carpeting and cabling)" }
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
      why: "125\u2013150% of the last month's rent is common, often with a grace period before consequential damages apply.",
      fix: { type: "replace", text: "one hundred twenty-five percent (125%) for the first sixty (60) days and one hundred fifty percent (150%) thereafter" }
    },
    {
      id: "holdover-consequential", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "medium", category: "holdover", scope: "holdover",
      re: /\bconsequential damages\b/i, unless: /\bafter\b[\s\S]{0,40}\b(days|notice)\b/i,
      title: "Holdover exposes you to consequential damages immediately",
      why: "A short holdover could make Tenant liable for Landlord's lost deal with a new tenant. Limit consequential damages to holdovers lasting more than 30\u201360 days after notice.",
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
      why: "Buyer gives up claims for known and unknown problems, including a California Civil Code \u00A71542 waiver. Carve out fraud, breach of Seller's reps and covenants, and pre-closing third-party claims.",
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
      why: "Survival without a cap leaves Seller's sale proceeds exposed. A cap of 1\u20133% of the price and a minimum claim threshold are common.",
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
      why: "Require Buyer to file within 60\u201390 days of the scheduled closing and to have been ready to close.",
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
      why: "California courts generally do not enforce pre-dispute jury waivers (Grafton Partners v. Superior Court, 2005). A judicial reference clause (Code Civ. Proc. \u00A7638) is the usual substitute.",
      fix: null
    },
    {
      id: "ca-civil-code-waivers", kind: "phrase", docTypes: ["lease"], riskFor: ["tenant"], severity: "low", category: "casualty",
      re: /\b(?:1932|1933\s*\(\s*4\s*\)|1941|1942|1950\.7)\b/,
      title: "Tenant waives California statutory rights",
      why: "Waivers of Civil Code \u00A7\u00A71932, 1933(4), 1941, 1942 or 1950.7 are standard in landlord forms. Make sure the lease gives you the contractual rights you're waiving (casualty termination, repairs, deposit accounting).",
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
      { category: "casp", riskFor: ["tenant", "landlord"], severity: "medium", template: "casp", requiresCalifornia: true, why: "California Civil Code \u00A71938 requires commercial leases to state whether the premises have been CASp-inspected." },
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
      // jurisdiction: "ca" forces the California checks on, "other" turns them off, anything else detects from the text
      california: opts.jurisdiction === "ca" ? true : opts.jurisdiction === "other" ? false : isCalifornia(text),
      byCategory: byCategory
    };
  }

  var REVIEW_TYPES = [
    { id: "comprehensive", label: "Comprehensive review" },
    { id: "quick", label: "Quick review (high risks only)" },
    { id: "drafting", label: "Drafting check (numbers, dates, terms)" }
  ];

  function isDraftingFinding(f) { return f.category === "drafting" || f.kind === "check"; }

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
      finding.title = "Words and numbers don't match: \u201C" + clip(matchText, 60) + "\u201D";
    }
    finding.comment = "Clause Craft \u2014 " + finding.title + ". " + finding.why + (finding.action && finding.action.type === "replace" ? " Suggested: \u201C" + finding.action.replace + "\u201D." : finding.action ? " Suggested addition: \u201C" + finding.action.text + "\u201D" : "");
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
  /** Locate for findings that are not backed by a RULES regex (checks, playbook items): plain text search. */
  function locateText(finding, raw) {
    var target = finding.matchText || "";
    if (finding.action && finding.action.type === "append") {
      if (norm(finding.action.text) && norm(raw).indexOf(norm(finding.action.text).slice(0, 60)) !== -1) return null;
      var ai = target ? raw.indexOf(target) : -1;
      return { target: target, context: target, offset: 0, occurrence: 0, start: Math.max(ai, 0) };
    }
    if (!target) return null;
    var hits = [], at = raw.indexOf(target);
    while (at !== -1) { hits.push(at); at = raw.indexOf(target, at + 1); }
    if (!hits.length) return null;
    var start = -1;
    // Prefer the words around the target as they were at review time (other edits may have added copies of it).
    if (finding.anchor && hits.length > 1) {
      var b = finding.anchor.before || "", a = finding.anchor.after || "";
      var tries = [[b, a], [b, ""], ["", a], [b.slice(-20), ""], ["", a.slice(0, 20)], [b.slice(-20), a.slice(0, 20)]];
      for (var t = 0; t < tries.length && start === -1; t++) {
        if (!tries[t][0] && !tries[t][1]) continue;
        var probe = tries[t][0] + target + tries[t][1];
        var pi = raw.indexOf(probe);
        if (pi !== -1 && raw.indexOf(probe, pi + 1) === -1) start = pi + tries[t][0].length;
      }
    }
    if (start === -1) start = hits[Math.min(finding.occurrence || 0, hits.length - 1)];
    var ctx = uniqueContext(raw, start, start + target.length);
    return { target: target, context: ctx.context, offset: ctx.offset, occurrence: countOccurrences(raw, target, start), start: start };
  }

  function locate(finding, currentText) {
    var raw = String(currentText == null ? "" : currentText);
    var text = norm(raw);
    var rule = null;
    for (var i = 0; i < RULES.length; i++) if (RULES[i].id === finding.ruleId) { rule = RULES[i]; break; }
    if (!rule) return locateText(finding, raw);
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
    // Consistency checks (dates, rent math, defined terms, cross-references, party names) apply to every side.
    var checked = runChecks(doc);
    findings = findings.concat(checked.findings);
    // One finding per rule+paragraph+offset; order by severity, then document position.
    var seen = {};
    findings = findings.filter(function (f) { if (seen[f.id]) return false; seen[f.id] = 1; return true; });
    findings.sort(function (a, b) {
      return severityRank(a.severity) - severityRank(b.severity) || a.paragraphIndex - b.paragraphIndex;
    });
    var missing = doc.docType === "unknown" ? [] : findMissing(doc, perspective);
    missing.sort(function (a, b) { return severityRank(a.severity) - severityRank(b.severity); });
    var reviewType = opts.reviewType === "quick" || opts.reviewType === "drafting" ? opts.reviewType : "comprehensive";
    if (reviewType === "quick") {
      findings = findings.filter(function (f) { return f.severity === "high"; });
      missing = missing.filter(function (m) { return m.severity === "high"; });
    } else if (reviewType === "drafting") {
      findings = findings.filter(isDraftingFinding);
      missing = [];
    }
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
      reviewType: reviewType,
      california: doc.california,
      findings: findings,
      missing: missing,
      clauses: clauses,
      checks: checked.items,
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
      if (rv.missing.length) lines.push("Missing: " + rv.missing.slice(0, 4).map(function (m) { return m.label; }).join(", ") + (rv.missing.length > 4 ? "\u2026" : ""));
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
      if (!parties.length) return { found: false, intent: "parties", answer: "No defined parties found (look for \u201C(\"Landlord\")\u201D-style definitions).", sources: [], facts: [] };
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
    var answer = refOf(top1) + ": \u201C" + clip(sentence, 320) + "\u201D";
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

    var lines = ["CLAUSE CRAFT SUMMARY \u2014 " + typeLabel + " (reviewed for the " + rv.perspectiveLabel.toLowerCase() + " side)"];
    fields.forEach(function (f) { lines.push(f.label + ": " + f.value); });
    if (risks.length) {
      lines.push("Key risks:");
      risks.forEach(function (r) { lines.push("\u2022 " + cap(r.severity) + " \u2014 " + r.title + " (" + r.ref + ")"); });
    }
    if (missing.length) lines.push("Missing: " + missing.join(", "));
    lines.push("Generated by Clause Craft (rule-based). Not legal advice \u2014 have counsel review.");

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

  /* ------------------------------------------------------------------ *
   * Numbers, money and dates in contract style
   * ------------------------------------------------------------------ */

  var ONES = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve",
    "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"];
  var TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];

  /** 62 -> "sixty-two", 150 -> "one hundred fifty" (whole numbers below one million). */
  function numberToWords(n) {
    n = Math.floor(Math.abs(Number(n) || 0));
    if (n < 20) return ONES[n];
    if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? "-" + ONES[n % 10] : "");
    if (n < 1000) return ONES[Math.floor(n / 100)] + " hundred" + (n % 100 ? " " + numberToWords(n % 100) : "");
    if (n < 1000000) return numberToWords(Math.floor(n / 1000)) + " thousand" + (n % 1000 ? " " + numberToWords(n % 1000) : "");
    return String(n);
  }

  var ORDINAL_WORD = { one: "first", two: "second", three: "third", five: "fifth", eight: "eighth", nine: "ninth", twelve: "twelfth" };

  /** 36 -> "thirty-sixth (36th)". */
  function ordinalWords(n) {
    n = Math.max(1, Math.floor(Math.abs(Number(n) || 1)));
    var w = numberToWords(n);
    var last = /([a-z]+)$/.exec(w)[1];
    var ord = ORDINAL_WORD[last] || (/y$/.test(last) ? last.replace(/y$/, "ieth") : last + "th");
    var suffix = (n % 100 >= 11 && n % 100 <= 13) ? "th" : ({ 1: "st", 2: "nd", 3: "rd" })[n % 10] || "th";
    return w.slice(0, w.length - last.length) + ord + " (" + n + suffix + ")";
  }

  function trimNum(v) { return String(Math.round(Number(v) * 100) / 100); }

  /** 10, "days" -> "ten (10) days"; 1, "months" -> "one (1) month"; fractions stay as digits. */
  function countPhrase(n, unit) {
    n = Number(n);
    var u = n === 1 ? String(unit).replace(/s$/, "") : String(unit);
    if (n !== Math.floor(n) || n < 0) return trimNum(n) + " " + u;
    return numberToWords(n) + " (" + trimNum(n) + ") " + u;
  }

  /** 5 -> "five percent (5%)"; fractions -> "3.5%". */
  function percentPhrase(n, words) {
    n = Number(n);
    if (words === false || n !== Math.floor(n) || n < 0) return trimNum(n) + "%";
    return numberToWords(n) + " percent (" + trimNum(n) + "%)";
  }

  function commas(n) { return String(Math.round(Number(n))).replace(/\B(?=(\d{3})+(?!\d))/g, ","); }

  function moneyText(v) {
    var parts = (Math.round(Number(v) * 100) / 100).toFixed(2).split(".");
    return "$" + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "." + parts[1];
  }

  function moneyValue(s) {
    var m = /\$\s?(\d[\d,]*(?:\.\d+)?)/.exec(s || "");
    return m ? parseFloat(m[1].replace(/,/g, "")) : NaN;
  }

  /** True when a number is written "ten (10)" or "four percent (4%)" style. */
  function wordsStyle(s) { return new RegExp("^\\s*" + WORD_NUMBER_RE + "\\b", "i").test(s || "") && /\(\s*\d/.test(s || ""); }

  var MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var DATE_TEXT = "(?:January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{1,2},\\s+\\d{4}";

  function parseDate(s) {
    var m = new RegExp("(" + MONTHS.join("|") + ")\\s+(\\d{1,2}),\\s+(\\d{4})").exec(s || "");
    return m ? new Date(Date.UTC(+m[3], MONTHS.indexOf(m[1]), +m[2])) : null;
  }
  function fmtDate(d) { return MONTHS[d.getUTCMonth()] + " " + d.getUTCDate() + ", " + d.getUTCFullYear(); }
  function addMonths(d, n) {
    var y = d.getUTCFullYear(), mo = d.getUTCMonth() + n;
    var last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
    return new Date(Date.UTC(y, mo, Math.min(d.getUTCDate(), last)));
  }
  function addDays(d, n) { return new Date(d.getTime() + n * 86400000); }
  function endOfMonth(d) { return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)); }
  function sameDay(a, b) { return !!(a && b) && a.getTime() === b.getTime(); }

  function escapeRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  /* ------------------------------------------------------------------ *
   * Key numbers (used by the consistency checks and the playbook)
   * ------------------------------------------------------------------ */

  var COUNT_SRC = "(?:" + WORD_NUMBER_RE + "\\s*\\(\\s*\\d+(?:\\.\\d+)?\\s*\\)|\\d+(?:\\.\\d+)?)";
  var PCT_SRC = "(?:" + WORD_NUMBER_RE + "\\s+percent\\s*\\(\\s*\\d+(?:\\.\\d+)?\\s*%\\s*\\)|\\d+(?:\\.\\d+)?\\s*%)";
  var MONEY_SRC = "\\$\\s?\\d[\\d,]*(?:\\.\\d{2})?";

  /** First hit of re in the paragraphs of the given categories (document order, headings skipped). */
  function hitIn(doc, cats, re, group, accept) {
    var seen = {}, idxs = [], c, i;
    for (c = 0; c < cats.length; c++) {
      var list = doc.byCategory[cats[c]] || [];
      for (i = 0; i < list.length; i++) if (!seen[list[i]]) { seen[list[i]] = 1; idxs.push(list[i]); }
    }
    idxs.sort(function (a, b) { return a - b; });
    for (var k = 0; k < idxs.length; k++) {
      var p = doc.paragraphs[idxs[k]];
      if (p.isHeading || p.empty) continue;
      var r = new RegExp(re.source, re.ignoreCase ? "gi" : "g"), m;
      while ((m = r.exec(p.text)) !== null) {
        if (m[0] === "") { r.lastIndex++; continue; }
        var g = group && m[group] ? m[group] : m[0];
        var start = m.index + m[0].indexOf(g);
        if (accept && !accept(m, p, start, g)) continue;
        return { text: p.raw.slice(start, start + g.length), norm: g, paragraphIndex: p.index, start: start, end: start + g.length, match: m, p: p };
      }
    }
    return null;
  }

  function withValue(hit, fn) {
    if (!hit) return null;
    var v = fn(hit.norm, hit);
    if (v == null || isNaN(v)) return null;
    hit.value = v;
    return hit;
  }

  function metricsOf(doc) {
    var M = {};
    var PCT = new RegExp(PCT_SRC, "i");
    if (doc.docType === "lease") {
      M.escalationPct = withValue(hitIn(doc, ["escalation"], PCT, 0, function (m, p) { return /increase|escalat|adjust/i.test(p.text); }), parseCount);
      M.lateChargePct = withValue(hitIn(doc, ["late_charge"], PCT), parseCount);
      M.holdoverPct = withValue(hitIn(doc, ["holdover"], PCT), parseCount);
      var mon = hitIn(doc, ["default"], new RegExp("\\bfails?\\s+to\\s+pay\\b[^;]{0,140}?\\b(?:continues\\s+for|for|within)\\s+(" + COUNT_SRC + "\\s+(?:business\\s+|calendar\\s+)?days?)", "i"), 1);
      if (mon) {
        mon.value = parseCount(mon.norm);
        var after = mon.p.text.slice(mon.end, mon.end + 60);
        mon.notice = /^\s*(?:after|following)\s+(?:the\s+)?(?:written\s+)?notice/i.test(after);
        var tail = /^,?\s*without\s+(?:further\s+)?notice/i.exec(after);
        mon.fullEnd = mon.end + (tail ? tail[0].length : 0);
        M.monetaryCure = mon;
      }
      M.nonMonetaryCureDays = withValue(hitIn(doc, ["default"], new RegExp("\\b(?:any other|other (?:obligations?|covenants?|terms?)|non-?monetary)[^;]{0,160}?\\b(" + COUNT_SRC + "\\s+(?:business\\s+)?days?)", "i"), 1), parseCount);
      M.securityDeposit = withValue(hitIn(doc, ["security_deposit"], new RegExp(MONEY_SRC)), moneyValue);
      M.monthlyRent = withValue(hitIn(doc, ["rent"], new RegExp("\\bmonthly installments of\\s+(" + MONEY_SRC + ")|(" + MONEY_SRC + ")\\s+per month\\b", "i"), 1), moneyValue);
      var psf = hitIn(doc, ["rent"], new RegExp("(" + MONEY_SRC + ")\\s+per\\s+(?:rentable\\s+|usable\\s+)?square\\s+f(?:oo|ee)t(?:\\s+per\\s+(month|year|annum))?", "i"), 1);
      if (psf) { psf.value = moneyValue(psf.norm); psf.period = psf.match[2] ? (/month/i.test(psf.match[2]) ? "month" : "year") : null; M.rentPerSf = psf; }
      M.rsf = withValue(hitIn(doc, ["premises"], /(\d[\d,]*)\s+(?:rentable|usable)\s+square\s+feet/i, 1), function (s) { return parseFloat(s.replace(/,/g, "")); });
      var term = hitIn(doc, ["term"], new RegExp("\\b(?:shall be|of|for)\\s+(" + COUNT_SRC + "\\s+(months?|years?))", "i"), 1);
      if (term) { term.value = parseCount(term.norm) * (/year/i.test(term.match[2]) ? 12 : 1); M.termMonths = term; }
      M.commencement = withValue(hitIn(doc, ["term"], new RegExp("commenc\\w*\\s+on\\s+(" + DATE_TEXT + ")|(" + DATE_TEXT + ")\\s*\\(the\\s+\"Commencement Date\"\\)", "i"), 1), function (s, h) { return parseDate(h.match[1] || h.match[2]); });
      if (M.commencement && !M.commencement.match[1]) M.commencement = withValue(hitIn(doc, ["term"], new RegExp("(" + DATE_TEXT + ")\\s*\\(the\\s+\"Commencement Date\"\\)", "i"), 1), parseDate);
      M.expiration = withValue(hitIn(doc, ["term"], new RegExp("expir\\w*\\s+on\\s+(" + DATE_TEXT + ")", "i"), 1), parseDate) ||
        withValue(hitIn(doc, ["term"], new RegExp("(" + DATE_TEXT + ")\\s*\\(the\\s+\"Expiration Date\"\\)", "i"), 1), parseDate);
      M.estoppelDays = withValue(hitIn(doc, ["estoppel"], new RegExp("\\bwithin\\s+(" + COUNT_SRC + "\\s+(?:business\\s+)?days?)", "i"), 1), parseCount);
      M.opexCapPct = withValue(hitIn(doc, ["operating_expenses"], new RegExp("(?:shall\\s+not|not\\s+to)\\s+(?:increase|exceed)\\b[^.;]{0,60}?(" + PCT_SRC + ")", "i"), 1), parseCount);
      if (M.securityDeposit && M.monthlyRent && M.monthlyRent.value > 0) {
        M.depositMonths = { value: M.securityDeposit.value / M.monthlyRent.value, text: M.securityDeposit.text, norm: M.securityDeposit.norm,
          paragraphIndex: M.securityDeposit.paragraphIndex, start: M.securityDeposit.start, end: M.securityDeposit.end, p: M.securityDeposit.p };
      }
    }
    if (doc.docType === "psa") {
      M.ddDays = withValue(hitIn(doc, ["due_diligence"], new RegExp("\\b(?:period\\s+of|have|within)\\s+(" + COUNT_SRC + "\\s+(?:business\\s+)?days?)", "i"), 1), parseCount);
      M.closingDays = withValue(hitIn(doc, ["closing"], new RegExp("\\b(" + COUNT_SRC + "\\s+(?:business\\s+)?days?)\\s+after\\b", "i"), 1), parseCount);
      var rep = hitIn(doc, ["reps"], new RegExp("\\bsurviv\\w*\\b[^.;]{0,80}?\\b(" + COUNT_SRC + "\\s+(months?|years?))", "i"), 1);
      if (rep) { rep.value = parseCount(rep.norm) * (/year/i.test(rep.match[2]) ? 12 : 1); M.repsSurvivalMonths = rep; }
      M.purchasePrice = withValue(hitIn(doc, ["purchase_price"], new RegExp("(" + MONEY_SRC + ")"), 1), moneyValue);
      M.earnestDeposit = withValue(hitIn(doc, ["deposit"], new RegExp("\\b(?:shall deposit|deposit of|earnest money(?: deposit)? of)\\b[^.;]{0,80}?(" + MONEY_SRC + ")", "i"), 1), moneyValue);
      if (M.purchasePrice && M.earnestDeposit && M.purchasePrice.value > 0) {
        M.depositPct = { value: M.earnestDeposit.value / M.purchasePrice.value * 100, text: M.earnestDeposit.text, norm: M.earnestDeposit.norm,
          paragraphIndex: M.earnestDeposit.paragraphIndex, start: M.earnestDeposit.start, end: M.earnestDeposit.end, p: M.earnestDeposit.p };
      }
    }
    return M;
  }

  /* ------------------------------------------------------------------ *
   * Consistency checks: dates, rent math, defined terms, cross-references, party names
   * ------------------------------------------------------------------ */

  var UNDEFINED_CANDIDATES = {
    lease: ["Rent", "Additional Rent", "Base Rent", "Operating Expenses", "Tenant's Share", "Common Areas", "Project",
      "Commencement Date", "Expiration Date", "Security Deposit", "Permitted Use", "Event of Default", "Effective Date", "Premises", "Building"],
    psa: ["Close of Escrow", "Closing Date", "Escrow Holder", "Opening of Escrow", "Title Company", "Due Diligence Period",
      "Permitted Exceptions", "Effective Date", "Purchase Price", "Deposit", "Property", "Title Policy"]
  };
  var DETERMINERS = /^(the|any|all|such|each|no|said|this|that|every|other|and|or|of|to|in|for|as|by|with|on|at|from|its|their|his|her)$/i;
  var DEFINITION_RE = /\(\s*(?:the\s+|each,?\s+(?:a|an)\s+|collectively,?\s+(?:the\s+)?|hereinafter\s+(?:the\s+|referred\s+to\s+as\s+(?:the\s+)?)?|individually,?\s+(?:a|an)\s+|a\s+|an\s+)?"([^"]{2,60})"\s*\)|"([A-Z][^"]{1,60})"\s+(?:means|shall mean|has the meaning|refers to)\b/g;

  function definitionsOf(doc) {
    var defs = [];
    for (var i = 0; i < doc.paragraphs.length; i++) {
      var p = doc.paragraphs[i];
      if (p.empty) continue;
      var re = new RegExp(DEFINITION_RE.source, "g"), m;
      while ((m = re.exec(p.text)) !== null) {
        var term = trim(m[1] || m[2]);
        var qs = m.index + m[0].indexOf("\"" + term);
        defs.push({ term: term, paragraphIndex: i, start: qs, end: qs + term.length + 2 });
      }
    }
    return defs;
  }

  function termUses(doc, term) {
    var out = [], re = new RegExp("(^|[^A-Za-z0-9])(" + escapeRe(term) + ")(?![A-Za-z0-9])", "g");
    for (var i = 0; i < doc.paragraphs.length; i++) {
      var p = doc.paragraphs[i];
      if (p.empty || p.isHeading) continue;
      var headEnd = 0;
      if (p.inlineHeading) { var hi = p.text.indexOf(p.inlineHeading); headEnd = hi === -1 ? 0 : hi + p.inlineHeading.length; }
      var m;
      re.lastIndex = 0;
      while ((m = re.exec(p.text)) !== null) {
        var start = m.index + m[1].length;
        re.lastIndex = start + term.length;
        if (start < headEnd) continue; // section caption ("4. Rent Adjustments."), not a use
        if (p.text.charAt(start - 1) === "\"") continue; // inside its own definition
        var before = p.text.slice(Math.max(0, start - 30), start);
        var prev = /([A-Z][A-Za-z'-]*)\s+$/.exec(before);
        // Part of a longer capitalized term ("Base Rent", "Security Deposit", "Rent Commencement Date"), not a use of this one.
        if (prev && !DETERMINERS.test(prev[1]) && !/'s$/.test(prev[1])) continue;
        if (/^\s+(?!I\b)[A-Z][a-z]/.test(p.text.slice(start + term.length, start + term.length + 4))) continue;
        out.push({ paragraphIndex: i, start: start, end: start + term.length });
      }
    }
    return out;
  }

  function isDefinedInText(doc, term, defs) {
    for (var d = 0; d < defs.length; d++) if (defs[d].term === term) return true;
    var lower = term.toLowerCase();
    var esc = escapeRe(term);
    var means = new RegExp("\\b" + esc + "\\b,?\\s+(?:which\\s+)?(?:means|shall mean|includes?|shall include|include|is defined|are defined|has the meaning)\\b|\\b" + esc + "\\s*\\(\\s*\\d|\\b(?:is|are|be|constitutes?|constitute)\\s+(?:an?\\s+|the\\s+)?" + esc + "\\b");
    for (var i = 0; i < doc.paragraphs.length; i++) {
      var p = doc.paragraphs[i];
      if (p.empty) continue;
      var head = (p.inlineHeading || (p.isHeading ? sectionName(p) : "") || "").toLowerCase().replace(/[.:]$/, "");
      if (head && (head === lower || head === lower + "s")) return true;
      if (means.test(p.text)) return true;
    }
    return false;
  }

  function normName(s) {
    return norm(s).toUpperCase().replace(/\bL\.\s?L\.\s?C\.?/g, "LLC").replace(/\bLIMITED LIABILITY COMPANY\b/g, "LLC")
      .replace(/\bINCORPORATED\b/g, "INC").replace(/\bCORPORATION\b/g, "CORP").replace(/\bL\.\s?P\.?/g, "LP")
      .replace(/[^A-Z0-9& ]+/g, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
  }

  function checkRiskFor(doc) {
    return doc.docType === "lease" ? ["tenant", "landlord"] : doc.docType === "psa" ? ["buyer", "seller"] : ["tenant", "landlord", "buyer", "seller"];
  }

  function checkFinding(doc, o) {
    var p = doc.paragraphs[o.paragraphIndex];
    var rule = { id: o.id, category: "drafting", severity: o.severity, riskFor: checkRiskFor(doc), title: o.title, why: o.why, fix: null };
    var f = buildFinding(rule, p, o.start, o.end, doc);
    f.kind = "check";
    f.request = o.request || null;
    f.occurrence = countOccurrences(p.raw, f.matchText, o.start);
    f.anchor = { before: p.raw.slice(Math.max(0, o.start - 40), o.start), after: p.raw.slice(o.end, o.end + 40) };
    return f;
  }

  function runChecks(doc) {
    var items = [], findings = [];
    if (doc.docType === "unknown" && !doc.paragraphs.some(function (p) { return !p.empty; })) return { items: items, findings: findings };
    var M = metricsOf(doc);

    // 1. Term dates add up.
    if (doc.docType === "lease" && M.termMonths && M.commencement && M.expiration) {
      var c0 = M.commencement.value, months = M.termMonths.value, exp = M.expiration.value;
      var e1 = addDays(addMonths(c0, months), -1), e2 = endOfMonth(addMonths(c0, months)), e3 = endOfMonth(e1);
      if (sameDay(exp, e1) || sameDay(exp, e2) || sameDay(exp, e3)) {
        items.push({ id: "term-dates", label: "Term dates add up", status: "ok", detail: months + " months from " + fmtDate(c0) + " ends " + fmtDate(exp) + "." });
      } else {
        items.push({ id: "term-dates", label: "Term dates add up", status: "issue", detail: months + " months from " + fmtDate(c0) + " ends " + fmtDate(e1) + ", not " + fmtDate(exp) + "." });
        findings.push(checkFinding(doc, { id: "check-term-dates", severity: "high", paragraphIndex: M.expiration.paragraphIndex, start: M.expiration.start, end: M.expiration.end,
          title: "Expiration date doesn't match the term",
          why: "A " + months + "-month term commencing " + fmtDate(c0) + " ends " + fmtDate(e1) + ", but the lease says " + fmtDate(exp) + ". Fix whichever is wrong before signing.",
          request: "Conform the Expiration Date to the " + months + "-month Term (" + fmtDate(e1) + "), or correct the Term." }));
      }
    }

    // 2. Rent matches rate x square feet.
    if (doc.docType === "lease" && M.monthlyRent && M.rentPerSf && M.rsf && M.rentPerSf.period) {
      var calc = M.rentPerSf.value * M.rsf.value / (M.rentPerSf.period === "year" ? 12 : 1);
      var diff = Math.abs(calc - M.monthlyRent.value);
      var rateText = moneyText(M.rentPerSf.value) + "/SF/" + (M.rentPerSf.period === "year" ? "yr" : "mo");
      if (diff <= Math.max(1, M.monthlyRent.value * 0.002)) {
        items.push({ id: "rent-math", label: "Rent matches rate \u00D7 square feet", status: "ok", detail: rateText + " \u00D7 " + commas(M.rsf.value) + " RSF = " + moneyText(calc) + " per month." });
      } else {
        items.push({ id: "rent-math", label: "Rent matches rate \u00D7 square feet", status: "issue", detail: rateText + " \u00D7 " + commas(M.rsf.value) + " RSF = " + moneyText(calc) + ", not " + moneyText(M.monthlyRent.value) + "." });
        findings.push(checkFinding(doc, { id: "check-rent-math", severity: "high", paragraphIndex: M.monthlyRent.paragraphIndex, start: M.monthlyRent.start, end: M.monthlyRent.end,
          title: "Monthly rent doesn't match the rate and square footage",
          why: rateText + " \u00D7 " + commas(M.rsf.value) + " rentable square feet = " + moneyText(calc) + " per month, but the lease says " + moneyText(M.monthlyRent.value) + ".",
          request: "Reconcile the monthly Base Rent with the rate per square foot and the Premises area." }));
      }
    }

    // 3. Party names match the signature blocks.
    var parties = extractParties(doc), sigChecked = 0, sigBad = 0;
    for (var s = 0; s < doc.paragraphs.length; s++) {
      var sp = doc.paragraphs[s];
      var sm = /^(\s*(LANDLORD|TENANT|BUYER|SELLER|LESSOR|LESSEE|PURCHASER|GUARANTOR)\s*:\s*)(\S[^\r\n]*?)\s*$/i.exec(sp.text);
      if (!sm) continue;
      var role = cap(sm[2].toLowerCase());
      role = role === "Lessor" ? "Landlord" : role === "Lessee" ? "Tenant" : role === "Purchaser" ? "Buyer" : role;
      var party = null;
      for (var q = 0; q < parties.length; q++) if (parties[q].role === role) party = parties[q];
      if (!party || /^_+$|^by\b/i.test(sm[3])) continue;
      sigChecked++;
      if (normName(party.name) !== normName(sm[3])) {
        sigBad++;
        findings.push(checkFinding(doc, { id: "check-party-name", severity: "medium", paragraphIndex: s, start: sm[1].length, end: sm[1].length + sm[3].length,
          title: role + "'s name in the signature block doesn't match the preamble",
          why: "The preamble names " + party.name + ", but the signature block says " + trim(sm[3]) + ". The wrong entity could end up bound, or no one at all.",
          request: "Conform the " + role + "'s name in the signature block to the preamble (" + party.name + ")." }));
      }
    }
    if (sigChecked) items.push({ id: "party-names", label: "Party names match the signature blocks", status: sigBad ? "issue" : "ok", detail: sigBad ? sigBad + " mismatch" + (sigBad > 1 ? "es" : "") + "." : sigChecked + " signature block" + (sigChecked > 1 ? "s" : "") + " checked." });

    // 4. Defined terms: defined but never used.
    var defs = definitionsOf(doc), unused = [];
    var ROLE_TERMS = /^(Landlord|Tenant|Buyer|Seller|Lessor|Lessee|Purchaser|Guarantor|Agreement|Lease)$/;
    var seenDef = {};
    for (var d = 0; d < defs.length; d++) {
      var def = defs[d];
      if (seenDef[def.term] || ROLE_TERMS.test(def.term) || !/^[A-Z0-9]/.test(def.term)) continue;
      seenDef[def.term] = 1;
      if (!termUses(doc, def.term).length) unused.push(def);
    }
    unused.slice(0, 5).forEach(function (def) {
      findings.push(checkFinding(doc, { id: "check-defined-unused", severity: "low", paragraphIndex: def.paragraphIndex, start: def.start, end: def.end,
        title: "\u201C" + def.term + "\u201D is defined but never used",
        why: "A defined term that's never used is usually a leftover from another draft, or a sign the wrong term is used elsewhere.",
        request: "Delete the unused definition of \u201C" + def.term + "\u201D, or use it where intended." }));
    });
    if (defs.length) items.push({ id: "defined-used", label: "Defined terms are used", status: unused.length ? "issue" : "ok", detail: unused.length ? unused.map(function (x) { return x.term; }).join(", ") + (unused.length > 1 ? " are" : " is") + " never used." : defs.length + " defined term" + (defs.length > 1 ? "s" : "") + " checked." });

    // 5. Capitalized terms used but never defined.
    var cands = UNDEFINED_CANDIDATES[doc.docType] || [], undef = [];
    for (var u = 0; u < cands.length; u++) {
      var uses = termUses(doc, cands[u]);
      if (!uses.length || isDefinedInText(doc, cands[u], defs)) continue;
      undef.push({ term: cands[u], use: uses[0] });
    }
    undef.slice(0, 6).forEach(function (x) {
      findings.push(checkFinding(doc, { id: "check-undefined-term", severity: "low", paragraphIndex: x.use.paragraphIndex, start: x.use.start, end: x.use.end,
        title: "\u201C" + x.term + "\u201D is used but never defined",
        why: "Capitalized terms read as defined terms. Without a definition, what \u201C" + x.term + "\u201D covers is open to argument.",
        request: "Define \u201C" + x.term + "\u201D, or replace it with a defined term." }));
    });
    if (cands.length && doc.docType !== "unknown") items.push({ id: "terms-defined", label: "Capitalized terms are defined", status: undef.length ? "issue" : "ok", detail: undef.length ? undef.map(function (x) { return x.term; }).join(", ") + (undef.length > 1 ? " have" : " has") + " no definition." : "No undefined standard terms found." });

    // 6. Cross-references point to real sections; referenced exhibits are attached.
    var sectionNos = {}, nSections = 0;
    doc.paragraphs.forEach(function (p) { if (p.sectionNo) { var major = p.sectionNo.split(".")[0]; if (!sectionNos[major]) { sectionNos[major] = 1; nSections++; } } });
    var badRefs = 0, refCount = 0, badEx = 0, exCount = 0, exSeen = {};
    for (var r = 0; r < doc.paragraphs.length && nSections >= 3; r++) {
      var rp = doc.paragraphs[r];
      if (rp.empty || rp.isHeading) continue;
      var xre = /\b(Section|Paragraph|Article)\s+(\d+)((?:\.\d+)*(?:\([a-z0-9]+\))*)/g, xm;
      while ((xm = xre.exec(rp.text)) !== null) {
        var pre = rp.text.slice(Math.max(0, xm.index - 40), xm.index);
        var post = rp.text.slice(xm.index + xm[0].length, xm.index + xm[0].length + 60);
        if (/\b(?:Code|U\.?\s?S\.?\s?C\.?|Regulations?|Act|Rule)\s*$/i.test(pre) || /^\s*(?:\([^)]*\)\s*)*of\s+(?:the\s+)?(?:Internal Revenue|[A-Z][a-z]+\s+Code|Code|United States|California|Securities)/.test(post)) continue;
        refCount++;
        if (!sectionNos[xm[2]]) {
          badRefs++;
          if (badRefs <= 5) findings.push(checkFinding(doc, { id: "check-cross-ref", severity: "medium", paragraphIndex: r, start: xm.index, end: xm.index + xm[0].length,
            title: xm[1] + " " + xm[2] + xm[3] + " is referenced but doesn't exist",
            why: "This document has no " + xm[1] + " " + xm[2] + ". A broken cross-reference can leave an obligation or exception pointing at nothing.",
            request: "Correct the reference to " + xm[1] + " " + xm[2] + xm[3] + "." }));
        }
      }
    }
    if (refCount) items.push({ id: "cross-refs", label: "Section references point to real sections", status: badRefs ? "issue" : "ok", detail: badRefs ? badRefs + " of " + refCount + " references are broken." : refCount + " reference" + (refCount > 1 ? "s" : "") + " checked." });
    for (var e = 0; e < doc.paragraphs.length; e++) {
      var ep = doc.paragraphs[e];
      if (ep.empty) continue;
      var ere = /\b(Exhibit|Schedule|Addendum)\s+"?([A-Z]|\d+)\b/g, em2;
      while ((em2 = ere.exec(ep.text)) !== null) {
        var key = em2[1].toLowerCase() + " " + em2[2];
        if (exSeen[key]) continue;
        exSeen[key] = 1;
        exCount++;
        var attached = new RegExp("^\\s*" + em2[1] + "\\s+\"?" + em2[2] + "\\b", "i");
        var found = doc.paragraphs.some(function (pp, idx) { return idx !== e && attached.test(pp.text) && (pp.isHeading || trim(pp.text).length < 80); });
        if (!found) {
          badEx++;
          findings.push(checkFinding(doc, { id: "check-exhibit", severity: "low", paragraphIndex: e, start: em2.index, end: em2.index + em2[0].length,
            title: em2[1] + " " + em2[2] + " is referenced but not attached",
            why: "The document refers to " + em2[1] + " " + em2[2] + ", but it isn't in this file. Make sure it's attached before signing.",
            request: "Attach " + em2[1] + " " + em2[2] + "." }));
        }
      }
    }
    if (exCount) items.push({ id: "exhibits", label: "Referenced exhibits are attached", status: badEx ? "issue" : "ok", detail: badEx ? badEx + " not attached." : exCount + " checked." });

    return { items: items, findings: findings, metrics: M };
  }

  /** Key numbers in a document (for display and tests). */
  function metrics(input, opts) {
    var M = metricsOf(analyze(input, opts || {})), out = {};
    for (var k in M) if (Object.prototype.hasOwnProperty.call(M, k) && M[k]) {
      out[k] = { value: M[k].value instanceof Date ? fmtDate(M[k].value) : M[k].value, text: M[k].text, paragraphIndex: M[k].paragraphIndex };
      if (k === "monetaryCure") out[k].notice = M[k].notice;
    }
    return out;
  }

  /* ------------------------------------------------------------------ *
   * Playbook: your standards, checked against the document
   * ------------------------------------------------------------------ */

  var STANDARDS = [
    // Lease: tenant side
    { id: "tn-escalation", side: "tenant", docType: "lease", type: "max", metric: "escalationPct", value: 3, unit: "%", label: "Annual rent increase" },
    { id: "tn-late-charge", side: "tenant", docType: "lease", type: "max", metric: "lateChargePct", value: 5, unit: "%", label: "Late charge" },
    { id: "tn-holdover", side: "tenant", docType: "lease", type: "max", metric: "holdoverPct", value: 150, unit: "%", label: "Holdover rent" },
    { id: "tn-monetary-cure", side: "tenant", docType: "lease", type: "min", metric: "monetaryCure", value: 10, unit: "days", label: "Rent default: cure period after written notice", severity: "high" },
    { id: "tn-nonmonetary-cure", side: "tenant", docType: "lease", type: "min", metric: "nonMonetaryCureDays", value: 30, unit: "days", label: "Other defaults: cure period" },
    { id: "tn-estoppel", side: "tenant", docType: "lease", type: "min", metric: "estoppelDays", value: 10, unit: "business days", label: "Estoppel turnaround" },
    { id: "tn-deposit", side: "tenant", docType: "lease", type: "max", metric: "depositMonths", value: 2, unit: "months of rent", label: "Security deposit", noFix: true },
    { id: "tn-opex-cap", side: "tenant", docType: "lease", type: "max", metric: "opexCapPct", value: 5, unit: "%", label: "Cap on controllable operating expenses", required: true, ruleIds: ["opex-no-cap"], category: "operating_expenses", severity: "high" },
    { id: "tn-assignment", side: "tenant", docType: "lease", type: "absent", ruleIds: ["assign-sole-discretion"], category: "assignment", label: "Assignment consent not unreasonably withheld", severity: "high", template: "assignment" },
    { id: "tn-indemnity", side: "tenant", docType: "lease", type: "absent", ruleIds: ["indemnity-own-negligence"], category: "indemnity", label: "No indemnity for Landlord's own negligence", severity: "high", okIfNoClause: true },
    { id: "tn-relocation", side: "tenant", docType: "lease", type: "absent", ruleIds: ["relocation"], category: "relocation", label: "No open-ended relocation right", okIfNoClause: true },
    { id: "tn-guaranty", side: "tenant", docType: "lease", type: "absent", ruleIds: ["personal-guaranty"], category: "guaranty", label: "Personal guaranty capped (or none)", okIfNoClause: true },
    { id: "tn-snda", side: "tenant", docType: "lease", type: "absent", ruleIds: ["snda-missing"], category: "snda", label: "Non-disturbance (SNDA) protection", severity: "high", template: "snda" },
    { id: "tn-renewal", side: "tenant", docType: "lease", type: "present", category: "renewal_option", template: "renewal_option", label: "Option to extend" },
    { id: "tn-casp", side: "tenant", docType: "lease", type: "present", category: "casp", template: "casp", label: "CASp disclosure (California)", california: true },
    // Lease: landlord side
    { id: "ll-escalation", side: "landlord", docType: "lease", type: "min", metric: "escalationPct", value: 3, unit: "%", label: "Annual rent increase" },
    { id: "ll-late-charge", side: "landlord", docType: "lease", type: "min", metric: "lateChargePct", value: 5, unit: "%", label: "Late charge" },
    { id: "ll-holdover", side: "landlord", docType: "lease", type: "min", metric: "holdoverPct", value: 150, unit: "%", label: "Holdover rent" },
    { id: "ll-deposit", side: "landlord", docType: "lease", type: "min", metric: "depositMonths", value: 1, unit: "months of rent", label: "Security deposit", noFix: true },
    { id: "ll-estoppel", side: "landlord", docType: "lease", type: "max", metric: "estoppelDays", value: 10, unit: "business days", label: "Estoppel turnaround" },
    { id: "ll-holdover-clause", side: "landlord", docType: "lease", type: "present", category: "holdover", template: "holdover", label: "Holdover clause" },
    { id: "ll-subrogation", side: "landlord", docType: "lease", type: "present", category: "waiver_subrogation", template: "waiver_subrogation", label: "Mutual waiver of subrogation" },
    { id: "ll-insurance", side: "landlord", docType: "lease", type: "present", category: "insurance", template: "insurance", label: "Tenant insurance requirements" },
    { id: "ll-casp", side: "landlord", docType: "lease", type: "present", category: "casp", template: "casp", label: "CASp disclosure (California)", california: true },
    // Purchase agreement: buyer side
    { id: "by-dd", side: "buyer", docType: "psa", type: "min", metric: "ddDays", value: 30, unit: "days", label: "Due diligence period", severity: "high" },
    { id: "by-closing", side: "buyer", docType: "psa", type: "min", metric: "closingDays", value: 30, unit: "days", label: "Days from end of due diligence to closing" },
    { id: "by-reps", side: "buyer", docType: "psa", type: "min", metric: "repsSurvivalMonths", value: 12, unit: "months", label: "Survival of Seller's representations" },
    { id: "by-deposit-pct", side: "buyer", docType: "psa", type: "max", metric: "depositPct", value: 3, unit: "% of price", label: "Deposit", noFix: true },
    { id: "by-refundable", side: "buyer", docType: "psa", type: "absent", ruleIds: ["psa-deposit-hard"], category: "deposit", label: "Deposit refundable during due diligence", severity: "high", template: "deposit" },
    { id: "by-title", side: "buyer", docType: "psa", type: "absent", ruleIds: ["psa-no-lien-cure", "psa-title-all-record"], category: "title", label: "Seller clears monetary liens; Buyer can object to title", severity: "high", template: "title" },
    { id: "by-estoppel", side: "buyer", docType: "psa", type: "absent", ruleIds: ["psa-estoppel-not-condition"], category: "estoppel", label: "Tenant estoppel is a closing condition", template: "estoppel_psa", leasedOnly: true },
    { id: "by-remedy", side: "buyer", docType: "psa", type: "absent", ruleIds: ["psa-buyer-remedy-deposit-only"], category: "default", label: "Specific performance if Seller defaults", severity: "high" },
    { id: "by-assign", side: "buyer", docType: "psa", type: "absent", ruleIds: ["psa-buyer-no-assign"], category: "assignment", label: "Buyer may assign to an affiliate", okIfNoClause: true },
    { id: "by-release", side: "buyer", docType: "psa", type: "absent", ruleIds: ["psa-release-no-carveouts"], category: "as_is", label: "AS-IS release carves out fraud and Seller's reps", okIfNoClause: true },
    // Purchase agreement: seller side
    { id: "sl-dd", side: "seller", docType: "psa", type: "max", metric: "ddDays", value: 30, unit: "days", label: "Due diligence period" },
    { id: "sl-reps", side: "seller", docType: "psa", type: "max", metric: "repsSurvivalMonths", value: 9, unit: "months", label: "Survival of Seller's representations" },
    { id: "sl-deposit-pct", side: "seller", docType: "psa", type: "min", metric: "depositPct", value: 1, unit: "% of price", label: "Deposit", noFix: true },
    { id: "sl-as-is", side: "seller", docType: "psa", type: "present", category: "as_is", template: "as_is", label: "AS-IS sale" },
    { id: "sl-liquidated", side: "seller", docType: "psa", type: "present", test: /\bliquidated damages\b/i, template: "liquidated_damages", label: "Liquidated damages for Buyer's default" },
    { id: "sl-financing", side: "seller", docType: "psa", type: "absent", ruleIds: ["psa-financing-contingency"], category: "financing", label: "No financing contingency", okIfNoClause: true },
    { id: "sl-deposit-hard", side: "seller", docType: "psa", type: "absent", ruleIds: ["psa-deposit-refundable-thru-closing"], category: "deposit", label: "Deposit goes hard after due diligence", okIfNoClause: true }
  ];

  function unitFor(unit, v) {
    if (unit.charAt(0) === "%") return unit;
    var u = Number(v) === 1 ? unit.replace(/^(\w+?)s\b/, "$1") : unit; // 1 "months of rent" -> "month of rent"
    return " " + u;
  }

  function standardText(std, value) {
    var v = value == null ? std.value : value;
    if (std.type === "present" || std.type === "absent") return "Required";
    return trimNum(v) + unitFor(std.unit, v) + (std.type === "max" ? " or less" : " or more");
  }

  function actualText(std, hit) {
    var v = std.metric === "depositMonths" || std.metric === "depositPct" ? (Math.round(hit.value * 10) / 10) : hit.value;
    var unit = std.unit;
    if (/days$/.test(unit)) unit = /business/i.test(hit.norm) ? "business days" : "days"; // say what the document says
    var s = trimNum(v) + unitFor(unit, v);
    if (std.metric === "monetaryCure" && !hit.notice) s += ", no notice";
    return s;
  }

  /** Replacement wording for a number in the document, written the way the document writes it. */
  function numberFix(std, hit, value) {
    var words = wordsStyle(hit.norm);
    if (std.unit.charAt(0) === "%") return percentPhrase(value, words);
    var unit = std.unit === "business days" ? (/business/i.test(hit.norm) ? "business days" : "business days") : (/years?/i.test(hit.norm) && std.unit === "months" ? "months" : std.unit);
    return words ? countPhrase(value, unit) : trimNum(value) + " " + (Number(value) === 1 ? unit.replace(/s$/, "") : unit);
  }

  function pbItem(std, doc, extra) {
    var item = {
      id: "pb-" + std.id, standardId: std.id, kind: "playbook", label: std.label, title: std.label,
      standardText: standardText(std, extra.value), actualText: extra.actualText || "", status: extra.status,
      severity: std.severity || "medium", ruleId: "pb-" + std.id, riskFor: [std.side],
      category: std.category || (std.metric ? "drafting" : null), categoryLabel: std.category ? categoryLabel(std.category) : "Playbook",
      paragraphIndex: -1, ref: "", section: "", excerpt: "", matchText: "", action: null,
      templateId: null, templateTitle: null, templateText: null, why: extra.why || ""
    };
    if (extra.finding) {
      var f = extra.finding;
      item.ruleId = f.ruleId; item.paragraphIndex = f.paragraphIndex; item.ref = f.ref; item.section = f.section;
      item.excerpt = f.excerpt; item.matchText = f.matchText; item.action = extra.action || f.action; item.why = item.why || f.why;
    } else if (extra.hit) {
      var h = extra.hit, p = doc.paragraphs[h.paragraphIndex];
      item.paragraphIndex = h.paragraphIndex; item.ref = refOf(p); item.section = sectionName(p);
      item.excerpt = clip(sentenceAt(p.raw, h.start), 260);
      item.matchText = p.raw.slice(h.start, extra.end || h.end);
      item.occurrence = countOccurrences(p.raw, item.matchText, h.start);
      item.anchor = { before: p.raw.slice(Math.max(0, h.start - 40), h.start), after: p.raw.slice(extra.end || h.end, (extra.end || h.end) + 40) };
      if (extra.replace) item.action = { type: "replace", find: item.matchText, replace: extra.replace };
    }
    if (extra.templateId && TEMPLATES[extra.templateId]) {
      item.templateId = extra.templateId; item.templateTitle = TEMPLATES[extra.templateId].title; item.templateText = TEMPLATES[extra.templateId].text;
    }
    item.title = std.label + ": " + (item.actualText || (item.status === "pass" ? "meets your standard" : "not found")) + " (standard: " + item.standardText + ")";
    item.comment = "Clause Craft playbook \u2014 " + item.title + "." + (item.why ? " " + item.why : "") +
      (item.action && item.action.type === "replace" ? " Suggested: \u201C" + item.action.replace + "\u201D." : item.action ? " Suggested addition: \u201C" + item.action.text + "\u201D" : "");
    return item;
  }

  /**
   * Check the document against a playbook of standards.
   * opts.standards: { "<standard id>": { value: <number>, off: <bool> } } overrides the defaults.
   */
  function playbook(input, opts) {
    opts = opts || {};
    var doc = analyze(input, opts);
    var perspective = opts.perspective && opts.perspective !== "auto" ? opts.perspective : defaultPerspective(doc.docType);
    if (perspective === "neutral") perspective = defaultPerspective(doc.docType);
    var overrides = opts.standards || {};
    var M = metricsOf(doc);
    var all = review(input, { perspective: "neutral", docType: doc.docType, jurisdiction: opts.jurisdiction }).findings;
    var items = [];
    for (var i = 0; i < STANDARDS.length; i++) {
      var std = STANDARDS[i];
      if (std.docType !== doc.docType || std.side !== perspective) continue;
      if (std.california && !doc.california) continue;
      if (std.leasedOnly && !hasLeasedProperty(doc)) continue;
      var ov = overrides[std.id] || {};
      if (ov.off) continue;
      var value = ov.value != null && !isNaN(Number(ov.value)) ? Number(ov.value) : std.value;
      var item;
      if (std.type === "max" || std.type === "min") {
        var hit = M[std.metric];
        if (!hit) {
          var rf = std.ruleIds ? all.filter(function (f) { return std.ruleIds.indexOf(f.ruleId) !== -1; })[0] : null;
          if (std.required && rf && rf.action) {
            var text = rf.action.text;
            if (std.unit === "%") text = text.replace(new RegExp(PCT_SRC, "i"), percentPhrase(value));
            item = pbItem(std, doc, { status: "fail", value: value, actualText: "None", finding: rf, action: { type: rf.action.type, text: text, find: rf.action.find, replace: rf.action.replace },
              why: "No cap found. Your standard is " + standardText(std, value) + "." });
          } else {
            item = pbItem(std, doc, { status: std.required ? "fail" : "not-found", value: value, actualText: "Not found", why: std.required ? "Your playbook requires this, and it isn't in the document." : "Clause Craft couldn't find this term in the document." });
          }
        } else {
          var ok = std.type === "max" ? hit.value <= value + 1e-9 : hit.value >= value - 1e-9;
          if (std.metric === "depositMonths" || std.metric === "depositPct") ok = std.type === "max" ? Math.round(hit.value * 10) / 10 <= value : Math.round(hit.value * 10) / 10 >= value;
          if (std.metric === "monetaryCure" && !hit.notice) ok = false;
          var extra = { status: ok ? "pass" : "fail", value: value, actualText: actualText(std, hit), hit: hit };
          if (!ok && !std.noFix) {
            if (std.metric === "monetaryCure") { extra.end = hit.fullEnd; extra.replace = numberFix(std, hit, Math.max(value, hit.value)) + " after written notice"; }
            else extra.replace = numberFix(std, hit, value);
          }
          extra.why = ok ? "" : "The document says " + extra.actualText + "; your standard is " + standardText(std, value) + ".";
          item = pbItem(std, doc, extra);
        }
      } else if (std.type === "absent") {
        var bad = all.filter(function (f) { return std.ruleIds.indexOf(f.ruleId) !== -1; })[0];
        var hasClause = std.category ? (doc.byCategory[std.category] || []).some(function (x) { return !doc.paragraphs[x].isHeading; }) : true;
        if (bad) item = pbItem(std, doc, { status: "fail", actualText: bad.title, finding: bad });
        else if (!hasClause && !std.okIfNoClause) item = pbItem(std, doc, { status: "fail", actualText: "No clause", templateId: std.template, why: "There is no " + categoryLabel(std.category).toLowerCase() + " clause in this document." });
        else item = pbItem(std, doc, { status: "pass", actualText: hasClause ? "Meets standard" : "No such clause" });
      } else {
        var present = std.test ? std.test.test(doc.text) : (doc.byCategory[std.category] || []).some(function (x) { return !doc.paragraphs[x].isHeading; });
        item = pbItem(std, doc, present ? { status: "pass", actualText: "Present" } : { status: "fail", actualText: "Missing", templateId: std.template, why: "Your playbook requires this clause, and the document doesn't have one." });
      }
      items.push(item);
    }
    var stats = { pass: 0, fail: 0, notFound: 0 };
    items.forEach(function (it) { if (it.status === "pass") stats.pass++; else if (it.status === "fail") stats.fail++; else stats.notFound++; });
    items.sort(function (a, b) {
      var ra = a.status === "fail" ? 0 : a.status === "not-found" ? 1 : 2, rb = b.status === "fail" ? 0 : b.status === "not-found" ? 1 : 2;
      return ra - rb || severityRank(a.severity) - severityRank(b.severity);
    });
    return { docType: doc.docType, perspective: perspective, perspectiveLabel: perspectiveLabel(perspective), items: items, stats: stats };
  }

  /* ------------------------------------------------------------------ *
   * Issues list (comments to send to the other side)
   * ------------------------------------------------------------------ */

  function requestFor(f) {
    if (f.request) return f.request;
    if (f.action && f.action.type === "replace") return "Change \u201C" + clip(f.action.find, 160) + "\u201D to \u201C" + f.action.replace + "\u201D.";
    if (f.action) return "Add: \u201C" + f.action.text + "\u201D";
    return f.why;
  }

  function issuesList(input, opts) {
    opts = opts || {};
    var rv = review(input, opts);
    var docName = rv.docType === "lease" ? "Lease" : rv.docType === "psa" ? "Purchase Agreement" : "Agreement";
    var heading = (rv.perspective === "neutral" ? "Comments" : rv.perspectiveLabel + "'s comments") + " on the " + docName;
    var items = [];
    rv.findings.slice().sort(function (a, b) { return a.paragraphIndex - b.paragraphIndex || severityRank(a.severity) - severityRank(b.severity); }).forEach(function (f) {
      items.push({ ref: f.ref, section: f.section || f.categoryLabel, title: f.title, request: requestFor(f), why: f.why, severity: f.severity, paragraphIndex: f.paragraphIndex });
    });
    rv.missing.forEach(function (m) {
      items.push({ ref: "New", section: m.label, title: "Add a " + m.label.toLowerCase() + " provision", request: m.templateText ? "Add: \u201C" + m.templateText + "\u201D" : "Add a " + m.label.toLowerCase() + " provision.", why: m.why, severity: m.severity, paragraphIndex: -1 });
    });
    var lines = [heading.toUpperCase(), "Prepared with Clause Craft (rule-based first pass). Not legal advice; subject to review by counsel.", ""];
    var notes = trim(opts.notes || "").replace(/\s+/g, " ");
    if (notes) lines.splice(2, 0, "Deal notes: " + notes);
    items.forEach(function (it, i) {
      it.n = i + 1;
      lines.push(it.n + ". " + it.ref + (it.section ? " " + it.section : "") + " \u2014 " + it.title + ". " + it.request);
    });
    lines.push("End of comments.");
    return { heading: heading, docType: rv.docType, perspective: rv.perspective, notes: notes, items: items, text: lines.join("\n") };
  }

  /** The parties named in the document, e.g. [{ role: "Tenant", name: "HARBORLINE ADVISORY GROUP, INC." }]. */
  function parties(input) {
    var doc = analyze(input, {});
    return extractParties(doc).map(function (p) { return { role: p.role, name: p.name, paragraphIndex: p.paragraphIndex }; });
  }

  /* ------------------------------------------------------------------ *
   * Drafting with your numbers (fill-in clauses)
   * ------------------------------------------------------------------ */

  var DRAFTS = {
    renewal_option: {
      title: "Option to extend", docType: "lease", category: "renewal_option",
      fields: [
        { key: "options", label: "Number of options", value: 1, min: 1, max: 5 },
        { key: "years", label: "Years per option", value: 5, min: 1, max: 25 },
        { key: "noticeMin", label: "Notice: at least (months before expiration)", value: 9, min: 1, max: 36 },
        { key: "noticeMax", label: "Notice: no more than (months before expiration)", value: 12, min: 1, max: 48 },
        { key: "fmv", label: "Extension rent (% of fair market rent)", value: 95, min: 50, max: 150 }
      ],
      build: function (v) {
        return "Option to Extend. Tenant shall have " + countPhrase(v.options, "options") + " to extend the Term for " + countPhrase(v.years, "years") + (v.options > 1 ? " each" : "") +
          ", exercisable by written notice to Landlord given not less than " + countPhrase(v.noticeMin, "months") + " nor more than " + countPhrase(Math.max(v.noticeMax, v.noticeMin), "months") +
          " before the then-scheduled Expiration Date. Base Rent for " + (v.options > 1 ? "each extension term" : "the extension term") + " shall be " + percentPhrase(v.fmv) +
          " of the fair market rent for comparable space in comparable buildings in the area, determined as of the commencement of the extension term.";
      }
    },
    early_termination: {
      title: "Early termination right", docType: "lease", category: "early_termination",
      fields: [
        { key: "month", label: "Effective at the end of month", value: 36, min: 1, max: 240 },
        { key: "notice", label: "Notice (months)", value: 9, min: 1, max: 24 },
        { key: "fee", label: "Fee (months of Base Rent)", value: 3, min: 0, max: 24 }
      ],
      build: function (v) {
        return "Early Termination. Tenant may terminate this Lease effective as of the last day of the " + ordinalWords(v.month) + " full calendar month of the Term by giving Landlord at least " +
          countPhrase(v.notice, "months") + "' prior written notice and paying Landlord, with the notice, a termination fee equal to " + countPhrase(v.fee, "months") +
          " of Base Rent at the rate then in effect plus the unamortized cost of the tenant improvements paid for by Landlord and the leasing commissions for this Lease, amortized on a straight-line basis over the initial Term.";
      }
    },
    rent_abatement: {
      title: "Free rent (abatement)", docType: "lease", category: "rent",
      fields: [{ key: "months", label: "Months of abated Base Rent", value: 3, min: 1, max: 24 }],
      build: function (v) {
        return "Abated Rent. Base Rent for the first " + countPhrase(v.months, "full calendar months") + " of the Term shall be abated, provided that Tenant is not then in default beyond applicable notice and cure periods. Tenant shall pay its share of operating expenses, taxes and insurance during the abatement period.";
      }
    },
    ti_allowance: {
      title: "Tenant improvement allowance", docType: "lease", category: "condition",
      fields: [
        { key: "perSf", label: "Allowance ($ per rentable SF)", value: 40, min: 0, max: 500, step: 0.5 },
        { key: "rsf", label: "Rentable square feet", value: 0, min: 0, max: 10000000 },
        { key: "months", label: "Use it within (months after commencement)", value: 12, min: 1, max: 60 }
      ],
      build: function (v) {
        var total = v.perSf * v.rsf;
        return "Tenant Improvement Allowance. Landlord shall provide Tenant an allowance of up to " + moneyText(v.perSf) + " per rentable square foot of the Premises" +
          (v.rsf > 0 ? " (" + moneyText(total) + " for " + commas(v.rsf) + " rentable square feet)" : "") +
          " for the design, permitting and construction of Tenant's initial improvements, payable within thirty (30) days after Tenant delivers paid invoices and unconditional lien releases. Any portion not used within " +
          countPhrase(v.months, "months") + " after the Commencement Date may be applied against Base Rent next coming due.";
      }
    },
    opex_cap: {
      title: "Cap on controllable expenses", docType: "lease", category: "operating_expenses",
      fields: [{ key: "pct", label: "Annual cap (%)", value: 5, min: 0, max: 25, step: 0.5 }],
      build: function (v) { return TEMPLATES.opex_cap.text.replace("five percent (5%)", percentPhrase(v.pct)); }
    },
    guaranty_cap: {
      title: "Guaranty cap and burn-off", docType: "lease", category: "guaranty",
      fields: [
        { key: "months", label: "Cap (months of Base Rent)", value: 12, min: 1, max: 60 },
        { key: "burnOff", label: "Burns off after (months of on-time payment)", value: 24, min: 1, max: 120 }
      ],
      build: function (v) {
        return "Guaranty Limitation. The guarantor's liability shall be limited to an amount equal to " + countPhrase(v.months, "months") + " of Base Rent and shall terminate if Tenant has not been in monetary default beyond applicable cure periods during the first " + countPhrase(v.burnOff, "months") + " of the Term.";
      }
    },
    exclusive_use: {
      title: "Exclusive use (retail)", docType: "lease", category: "retail_ops",
      fields: [{ key: "use", label: "Protected use", type: "text", value: "the sale of coffee and espresso drinks for on- or off-site consumption" }],
      build: function (v) {
        return "Exclusive Use. So long as Tenant is open and operating in the Premises and is not in default beyond applicable notice and cure periods, Landlord shall not lease, or consent to the use of, any other space in the shopping center for " + trim(v.use || "Tenant's primary use") +
          ". If this covenant is breached, Base Rent shall be reduced by fifty percent (50%) until the breach ends, and Tenant may terminate this Lease if the breach continues for more than one hundred eighty (180) days.";
      }
    },
    rofo: {
      title: "Right of first offer on space", docType: "lease", category: "premises",
      fields: [
        { key: "space", label: "Which space", type: "text", value: "any space on the same floor as the Premises" },
        { key: "days", label: "Days to accept (business days)", value: 10, min: 1, max: 60 }
      ],
      build: function (v) {
        return "Right of First Offer. Before leasing " + trim(v.space || "any space adjacent to the Premises") + " to a third party during the Term, Landlord shall first offer it to Tenant in writing on the terms Landlord is prepared to accept. Tenant shall have " +
          countPhrase(v.days, "business days") + " to accept the offer by written notice, failing which Landlord may lease the space to a third party on terms not materially more favorable to the tenant than those offered to Tenant.";
      }
    },
    due_diligence: {
      title: "Due diligence period", docType: "psa", category: "due_diligence",
      fields: [
        { key: "days", label: "Days after opening of escrow", value: 45, min: 1, max: 180 },
        { key: "docsDays", label: "Seller delivers documents within (business days)", value: 5, min: 1, max: 30 }
      ],
      build: function (v) {
        return "Due Diligence. Buyer shall have " + countPhrase(v.days, "days") + " after the Opening of Escrow (the \"Due Diligence Period\") to review title, survey, leases, service contracts, environmental reports and all other matters relating to the Property, and Seller shall deliver all documents in its possession relating to the Property within " +
          countPhrase(v.docsDays, "business days") + " after the Opening of Escrow. Buyer may terminate this Agreement for any reason by written notice before the end of the Due Diligence Period, in which case the Deposit shall be returned to Buyer.";
      }
    },
    dd_extension: {
      title: "Due diligence extension", docType: "psa", category: "due_diligence",
      fields: [
        { key: "days", label: "Extension (days)", value: 30, min: 1, max: 180 },
        { key: "amount", label: "Additional deposit ($)", value: 25000, min: 0, max: 100000000 }
      ],
      build: function (v) {
        return "Extension. Buyer may extend the Due Diligence Period once by up to " + countPhrase(v.days, "days") + " by written notice to Seller before the Due Diligence Period expires and by depositing an additional " + moneyText(v.amount) +
          " with Escrow Holder, which shall be added to the Deposit and applied to the Purchase Price at the Close of Escrow.";
      }
    },
    closing: {
      title: "Closing date", docType: "psa", category: "closing",
      fields: [{ key: "days", label: "Days after due diligence ends", value: 30, min: 1, max: 365 }],
      build: function (v) {
        return "Closing. The Close of Escrow shall occur on or before the date that is " + countPhrase(v.days, "days") + " after the expiration of the Due Diligence Period (the \"Closing Date\").";
      }
    },
    reps_survival: {
      title: "Survival and cap on Seller's reps", docType: "psa", category: "reps",
      fields: [
        { key: "months", label: "Survival (months)", value: 12, min: 1, max: 60 },
        { key: "capPct", label: "Cap (% of purchase price)", value: 3, min: 0, max: 100, step: 0.5 }
      ],
      build: function (v) {
        return "Survival; Limitation. Seller's representations and warranties shall survive the Close of Escrow for " + countPhrase(v.months, "months") +
          ". Seller's aggregate liability for breaches of its representations and warranties discovered after the Close of Escrow shall not exceed " + percentPhrase(v.capPct) +
          " of the Purchase Price, and no claim may be made unless Buyer gives Seller written notice of it before the end of the survival period.";
      }
    }
  };

  /** Field defaults for a drafted clause, filled from the document where possible (e.g. square footage). */
  function draftDefaults(id, input) {
    var d = DRAFTS[id];
    if (!d) return null;
    var out = {};
    d.fields.forEach(function (f) { out[f.key] = f.value; });
    if (input && id === "ti_allowance") {
      var M = metricsOf(analyze(input, {}));
      if (M.rsf) out.rsf = M.rsf.value;
    }
    return out;
  }

  /** Builds a clause from a draft template and your values (numbers are clamped to each field's range). */
  function draftClause(id, values, input) {
    var d = DRAFTS[id];
    if (!d) return null;
    var v = draftDefaults(id, input);
    values = values || {};
    d.fields.forEach(function (f) {
      if (values[f.key] == null || values[f.key] === "") return;
      if (f.type === "text") { v[f.key] = String(values[f.key]); return; }
      var n = Number(values[f.key]);
      if (isNaN(n)) return;
      if (f.min != null) n = Math.max(f.min, n);
      if (f.max != null) n = Math.min(f.max, n);
      v[f.key] = f.step ? Math.round(n / f.step) * f.step : Math.round(n);
    });
    return { id: id, title: d.title, category: d.category, values: v, text: d.build(v) };
  }

  /* ------------------------------------------------------------------ *
   * Explain a clause in plain English
   * ------------------------------------------------------------------ */

  var EXPLAIN = {
    premises: "Defines the space being leased: the suite, the building and the square footage that rent and expense shares are based on.",
    term: "Sets when the lease starts and ends. The dates and the number of months should agree.",
    rent: "The base rent, when it's due and how it's paid. Free-rent (abatement) months usually appear here too.",
    escalation: "How base rent goes up over the term, by fixed percentage bumps or CPI. Small differences compound over a long lease.",
    operating_expenses: "The tenant's share of the building's operating costs, taxes and insurance. Look for a cap on controllable costs, exclusions for capital items, and an audit right.",
    security_deposit: "Money the landlord holds against a default. Watch the amount (in months of rent), when it can be applied, and when it must be returned.",
    use: "What business the tenant may run in the space. A narrow use clause can block a sublease or a change in the business.",
    condition: "The condition the space is delivered in. AS IS means the tenant takes the risk of existing defects unless the landlord gives a warranty.",
    repairs: "Who maintains and replaces what: roof, structure, HVAC and building systems. Capital replacements are the expensive part.",
    alterations: "Whether the tenant can change the space and what consent it needs. Cosmetic work usually shouldn't need consent.",
    assignment: "Whether the tenant can assign the lease or sublet. Consent \u201Cnot unreasonably withheld\u201D and free transfers to affiliates or a buyer of the business are the usual asks.",
    insurance: "The insurance each side must carry, the limits, and who must be named as additional insured.",
    indemnity: "Who covers whom for third-party claims. Each side should generally answer for its own negligence.",
    waiver_subrogation: "Each party's property insurer can't sue the other side after an insured loss. Usually mutual.",
    limitation_liability: "Limits on what a party can recover. Landlords often limit their liability to their interest in the building.",
    utilities: "Who provides utilities and services, and whether rent abates if they're interrupted.",
    casualty: "What happens after a fire or other damage: restoration, rent abatement and termination rights.",
    condemnation: "What happens if the government takes all or part of the property: termination rights and who gets the award.",
    snda: "Puts the lease behind the landlord's loan. A non-disturbance agreement (SNDA) keeps the tenant in place if the lender forecloses.",
    estoppel: "A signed statement confirming the lease terms and that no one is in default, usually for the landlord's lender or buyer. Watch the turnaround time.",
    relocation: "Lets the landlord move the tenant to other space. Limit it to comparable space at the landlord's cost, or delete it.",
    holdover: "Rent and liability if the tenant stays after the lease ends. 125\u2013150% of the last rent is common; consequential damages are the big exposure.",
    "default": "What counts as a default, the notice and cure periods, and the remedies. Short cure periods without notice are the main risk.",
    late_charge: "The fee for late rent, sometimes with interest. 5% with a grace period is common.",
    entry: "When the landlord can come into the space. Advance notice except in emergencies is the usual ask.",
    landlord_lien: "A security interest for the landlord in the tenant's furniture and equipment, which can get in the way of equipment financing.",
    parking: "The number of spaces, reserved or unreserved, and what they cost now and later.",
    rules: "Building rules the tenant must follow. Changes should be reasonable, applied uniformly and consistent with the lease.",
    guaranty: "A personal or corporate guarantee of the tenant's obligations. A cap and a burn-off limit the exposure.",
    renewal_option: "The tenant's right to extend the lease, the notice window, and how the new rent is set.",
    early_termination: "A right to end the lease early, usually with advance notice and a termination fee.",
    quiet_enjoyment: "The landlord's promise not to interfere with the tenant's possession while the tenant isn't in default.",
    audit: "The tenant's right to check the landlord's operating expense records and recover overcharges.",
    hazardous: "Responsibility for hazardous materials and environmental conditions.",
    casp: "California requires commercial leases to state whether the space has been inspected by a Certified Access Specialist (Civil Code \u00A71938).",
    signage: "Where and how the tenant can put up signs.",
    retail_ops: "Retail covenants: co-tenancy, exclusive use, operating hours and radius restrictions.",
    attorneys_fees: "Who pays legal fees in a dispute. A prevailing-party clause is the balanced version.",
    jury_waiver: "How disputes are decided. In California, pre-dispute jury waivers generally aren't enforced; judicial reference is the usual substitute.",
    brokers: "Which brokers are involved and who pays their commission.",
    notices: "How formal notices must be given: method, addresses and when they take effect.",
    governing_law: "Which state's law applies.",
    purchase_price: "The price and how it's paid at closing.",
    deposit: "The earnest money: the amount, when it's due, and when it becomes non-refundable.",
    due_diligence: "The buyer's inspection period and its right to walk away during it.",
    title: "How title is reviewed, which exceptions the buyer must accept, and what the seller must clear (monetary liens).",
    financing: "Whether the buyer's obligation depends on getting a loan.",
    closing: "When and how escrow closes.",
    prorations: "How taxes, rents and expenses are split as of closing, and who pays closing costs.",
    reps: "The seller's statements of fact about the property, how long they survive closing, and any cap on liability.",
    as_is: "The buyer takes the property in its current condition and releases the seller, except for listed carve-outs.",
    exchange_1031: "Lets either party do a 1031 tax-deferred exchange with the other's cooperation.",
    operating_covenants: "What the seller may and may not do with the property (new leases, contracts) before closing."
  };

  /**
   * Explain selected text: which kind of clause it is, in plain English, and what Clause Craft flags in it.
   * input: the whole document (for context); opts: { perspective }.
   */
  function explain(selection, input, opts) {
    opts = opts || {};
    var sel = trim(norm(selection));
    if (!sel) return { found: false, answer: "Select some contract text first, or put the cursor in a paragraph.", categories: [], findings: [], paragraphIndexes: [] };
    var docInput = input || [selection];
    var doc = analyze(docInput, opts);
    var lines = sel.split(/\r\n|\r|\n/).map(trim).filter(function (l) { return l.length >= 8; });
    if (!lines.length) lines = [sel];
    var idxs = [];
    for (var i = 0; i < doc.paragraphs.length; i++) {
      var t = trim(doc.paragraphs[i].text);
      if (!t) continue;
      for (var j = 0; j < lines.length; j++) {
        if (t.indexOf(lines[j].slice(0, 80)) !== -1 || (t.length >= 8 && lines[j].indexOf(t) !== -1)) { idxs.push(i); break; }
      }
    }
    var cats = [];
    if (idxs.length) idxs.forEach(function (x) { cats = cats.concat(doc.paragraphs[x].categories || []); });
    else toParagraphs(lines).forEach(function (p) { cats = cats.concat(categoriesFor(p, doc.docType)); });
    cats = uniq(cats);
    var findings;
    if (idxs.length) findings = review(docInput, opts).findings.filter(function (f) { return idxs.indexOf(f.paragraphIndex) !== -1; });
    else findings = review(lines, { perspective: opts.perspective, docType: doc.docType }).findings;
    var categories = cats.map(function (c) { return { id: c, label: categoryLabel(c), plain: EXPLAIN[c] || "" }; });
    var where = idxs.length ? refOf(doc.paragraphs[idxs[0]]) + " " : "";
    var parts = [];
    if (categories.length) {
      parts.push(where + categories.map(function (c) { return c.label; }).join(" / ") + ".");
      categories.forEach(function (c) { if (c.plain) parts.push(c.plain); });
    } else {
      parts.push(where + "Clause Craft doesn't recognize this as one of the CRE clauses it checks.");
    }
    if (findings.length) parts.push("Flagged here: " + findings.slice(0, 4).map(function (f) { return f.title; }).join("; ") + (findings.length > 4 ? "; and " + (findings.length - 4) + " more" : "") + ".");
    else if (categories.length) parts.push("Nothing flagged in this text for the " + perspectiveLabel(opts.perspective && opts.perspective !== "auto" ? opts.perspective : defaultPerspective(doc.docType)).toLowerCase() + " side.");
    return { found: !!categories.length, categories: categories, findings: findings, paragraphIndexes: idxs, answer: parts.join("\n") };
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
    REVIEW_TYPES: REVIEW_TYPES,
    parties: parties,
    STANDARDS: STANDARDS,
    playbook: playbook,
    issuesList: issuesList,
    DRAFTS: DRAFTS,
    draftDefaults: draftDefaults,
    draftClause: draftClause,
    EXPLAIN: EXPLAIN,
    explain: explain,
    metrics: metrics,
    numberToWords: numberToWords,
    ordinalWords: ordinalWords,
    countPhrase: countPhrase,
    percentPhrase: percentPhrase,
    wordsToNumber: wordsToNumber,
    norm: norm
  };
});
