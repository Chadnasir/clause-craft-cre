/*
 * Clause Craft - browser app (index.html).
 * A Word-like editor with a review sidebar, powered by the same engine as the Word add-in
 * (engine/clause-engine.js). Everything runs in the page: no uploads, no API keys, works offline
 * when opened straight from the folder (file://).
 * Document text is only ever inserted with textContent / DOM nodes, never innerHTML.
 */
(function () {
  "use strict";

  var Engine = window.ClauseCraftEngine;
  var Demo = window.ClauseCraftDemoDocs;
  var STORE_KEY = "clauseCraft.doc.v1";
  var SETUP_KEY = "clauseCraft.webSetup.v1";
  var STANDARDS_KEY = "clauseCraft.standards.v1";
  var CLAUSES_KEY = "clauseCraft.myClauses.v1";
  var PLAYBOOKS_KEY = "clauseCraft.playbooks.v1";
  var editor = document.getElementById("editor");

  var LIBRARY = {
    lease: ["snda", "waiver_subrogation", "casualty", "quiet_enjoyment", "renewal_option", "audit", "opex_cap", "casp",
      "assignment", "indemnity", "default_notice", "holdover", "estoppel", "security_deposit", "insurance",
      "attorneys_fees", "governing_law", "notices", "brokers"],
    psa: ["due_diligence", "deposit", "title", "reps", "operating_covenants", "estoppel_psa", "prorations", "closing",
      "exchange_1031", "as_is", "liquidated_damages", "financing", "attorneys_fees", "governing_law", "notices", "brokers"]
  };

  var CHIPS = {
    lease: ["Who can terminate?", "What's the NNN?", "Summarize risks", "What is the rent?", "Governing law?", "Insurance requirements?"],
    psa: ["Who can terminate?", "What is the purchase price?", "How long is due diligence?", "When is closing?", "Is the deposit refundable?", "Summarize risks"]
  };

  /* Per-computer storage (never sent anywhere). Reads and writes can fail in private windows. */
  function loadStore(key, fallback) {
    try { var v = window.localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; }
  }
  function saveStore(key, value) {
    try { window.localStorage.setItem(key, JSON.stringify(value)); return true; } catch (e) { return false; }
  }

  // base / pbBase: paragraph texts and count when the review (or playbook) ran, so a redline never lands
  // on the wrong paragraph if the document is edited between the review and the click.
  var state = {
    perspective: "auto", reviewType: "comprehensive", jurisdiction: "auto", playbookChoice: "mine", notes: "",
    setupOpen: true,
    result: null, playbook: null, base: null, pbBase: null,
    status: {}, filter: { high: true, medium: true, low: true, missing: true }, scope: null,
    library: "lease", pbEditing: false,
    standards: loadStore(STANDARDS_KEY, {}),
    myClauses: loadStore(CLAUSES_KEY, []),
    playbooks: loadStore(PLAYBOOKS_KEY, []),   // named playbooks: [{ id, name, standards }]
    customParties: {},                         // role -> name you added, kept with this document
    summary: null, issues: null, lastRange: null
  };

  function reviewOpts() { return { perspective: state.perspective, reviewType: state.reviewType, jurisdiction: state.jurisdiction }; }
  function pbOpts() { return { perspective: state.perspective, jurisdiction: state.jurisdiction, standards: activeStandards() }; }

  /* Playbooks: "My standards" (state.standards) plus any named playbooks saved in this browser. */
  function activePlaybook() {
    var c = state.playbookChoice || "";
    if (c.indexOf("pb:") !== 0) return null;
    for (var i = 0; i < state.playbooks.length; i++) if ("pb:" + state.playbooks[i].id === c) return state.playbooks[i];
    return null;
  }
  function activeStandards() { var pb = activePlaybook(); return pb ? pb.standards : state.standards; }
  function playbookName() { var pb = activePlaybook(); return pb ? pb.name : "My standards"; }
  function saveActiveStandards() { return activePlaybook() ? saveStore(PLAYBOOKS_KEY, state.playbooks) : saveStore(STANDARDS_KEY, state.standards); }

  /* ---------------- helpers ---------------- */

  function $(id) { return document.getElementById(id); }
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (!Object.prototype.hasOwnProperty.call(attrs, k) || attrs[k] == null || attrs[k] === false) continue;
      if (k === "text") el.textContent = attrs[k];
      else if (k === "class") el.className = attrs[k];
      else if (k.indexOf("on") === 0 && typeof attrs[k] === "function") el.addEventListener(k.slice(2), attrs[k]);
      else el.setAttribute(k, attrs[k] === true ? "" : attrs[k]);
    }
    for (var i = 2; i < arguments.length; i++) {
      var c = arguments[i];
      if (c == null || c === false) continue;
      [].concat(c).forEach(function (x) { if (x != null && x !== false) el.appendChild(typeof x === "string" ? document.createTextNode(x) : x); });
    }
    return el;
  }
  function clear(el) { while (el && el.firstChild) el.removeChild(el.firstChild); }
  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }
  function trim(s) { return String(s == null ? "" : s).replace(/^\s+|\s+$/g, ""); }
  function keyOf(f) { return f.ruleId + "|" + (f.matchText || "") + "|" + (f.excerpt || "").slice(0, 60); }
  function plural(n, word) { return n + " " + word + (n === 1 ? "" : "s"); }
  function sevLabel(s) { return s === "medium" ? "Med" : cap(s); }
  // Keep the document's selection when a sidebar button is clicked with the mouse.
  function keepSelection(e) { e.preventDefault(); }

  function toast(msg, kind) {
    var t = $("cc-toast");
    if (!t) { t = h("div", { id: "cc-toast", class: "cc-toast", role: "status", "aria-live": "polite" }); document.body.appendChild(t); }
    t.textContent = msg;
    t.className = "cc-toast is-visible" + (kind === "err" ? " is-err" : kind === "ok" ? " is-ok" : "");
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.className = "cc-toast"; }, kind === "err" ? 4200 : 2600);
  }

  function copyText(text, done) {
    function fallback() {
      var ok = false;
      try {
        var ta = h("textarea", { style: "position:fixed;left:-9999px;top:0", "aria-hidden": "true" });
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        ok = document.execCommand("copy");
        document.body.removeChild(ta);
      } catch (e) { ok = false; }
      if (ok) toast(done, "ok"); else toast("Copy isn't available in this browser. Select the text and press Ctrl+C.", "err");
    }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () { toast(done, "ok"); }, fallback);
        return;
      }
    } catch (e) { /* use the fallback */ }
    fallback();
  }

  function download(name, blob) {
    var a = h("a", { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
  }

  /* ---------------- document model ---------------- */

  function paragraphs() {
    var ps = Array.prototype.slice.call(editor.children).filter(function (n) { return /^(P|H1|H2|H3|DIV|LI)$/.test(n.tagName); });
    if (!ps.length && editor.textContent.trim()) {
      // Normalise loose text (e.g. pasted) into paragraphs.
      var text = editor.innerText;
      setDocument(text.split(/\r?\n/), { keepSetup: true });
      return paragraphs();
    }
    return ps;
  }

  function skipNode(c) { return c.nodeType === 1 && (c.tagName === "DEL" || c.classList.contains("cc-note")); }

  /** Paragraph text as it currently reads: deleted (struck) redline text is excluded. */
  function currentText(p) {
    var out = "";
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) out += c.nodeValue;
        else if (c.nodeType === 1 && !skipNode(c)) walk(c);
      }
    })(p);
    return out;
  }

  function texts() { return paragraphs().map(currentText); }

  function docTypeOf(lines) {
    var t = Engine.detectDocType(lines);
    return t === "psa" ? "psa" : t === "lease" ? "lease" : null;
  }

  function setDocument(lines, opts) {
    opts = opts || {};
    clear(editor);
    lines.forEach(function (line, i) {
      var text = String(line);
      var p = h("p");
      var md = /^#{1,3}\s+(.*)$/.exec(text);
      var m = /^(\d+\.\s+[^.]{2,60}\.)\s+([\s\S]+)$/.exec(text);
      if (md) { p.className = "doc-heading"; p.textContent = md[1]; }
      else if (m) { p.appendChild(h("strong", { text: m[1] })); p.appendChild(document.createTextNode(" " + m[2])); }
      else if (/^[^a-z]{3,90}$/.test(text) && !/^By:/.test(text)) { p.className = i === 0 ? "doc-title" : "doc-heading"; p.textContent = text; }
      else if (!text.trim()) { p.appendChild(h("br")); }
      else p.textContent = text;
      editor.appendChild(p);
    });
    state.result = null;
    state.playbook = null;
    state.status = {};
    state.base = null;
    state.pbBase = null;
    state.scope = null;
    state.summary = null;
    state.issues = null;
    state.lastRange = null;
    if (!opts.keepSetup) { state.perspective = "auto"; state.notes = ""; state.customParties = {}; state.setupOpen = true; saveSetup(); syncSetupControls(); }
    var dt = docTypeOf(lines);
    if (dt) state.library = dt;
    clear($("summary-output"));
    clear($("issues"));
    $("issues-actions").hidden = true;
    updatePartyLabels();
    renderAll();
    save();
  }

  function save() {
    try { localStorage.setItem(STORE_KEY, editor.innerHTML); } catch (e) { /* private mode */ }
  }

  function restore() {
    var html = null;
    try { html = localStorage.getItem(STORE_KEY); } catch (e) { html = null; }
    if (!html) return false;
    // Re-create from our own saved markup through a detached document, keeping only safe elements.
    var tmp = document.implementation.createHTMLDocument("x").createElement("div");
    tmp.innerHTML = html;
    var allowed = { P: 1, DIV: 1, H1: 1, H2: 1, H3: 1, STRONG: 1, B: 1, EM: 1, I: 1, S: 1, STRIKE: 1, U: 1, DEL: 1, INS: 1, BR: 1, SPAN: 1 };
    (function sanitize(n) {
      Array.prototype.slice.call(n.childNodes).forEach(function (c) {
        if (c.nodeType === 1) {
          if (!allowed[c.tagName]) { n.replaceChild(document.createTextNode(c.textContent), c); return; }
          Array.prototype.slice.call(c.attributes).forEach(function (a) { if (a.name !== "class") c.removeAttribute(a.name); });
          sanitize(c);
        } else if (c.nodeType !== 3) n.removeChild(c);
      });
    })(tmp);
    Array.prototype.forEach.call(tmp.querySelectorAll(".cc-flash"), function (el) { el.classList.remove("cc-flash"); });
    clear(editor);
    while (tmp.firstChild) editor.appendChild(tmp.firstChild);
    return !!editor.textContent.trim();
  }

  /* Map an offset in currentText(p) to a DOM position (skipping <del>). */
  function domPoint(p, offset) {
    var count = 0, found = null;
    (function walk(n) {
      for (var c = n.firstChild; c && !found; c = c.nextSibling) {
        if (c.nodeType === 3) {
          var len = c.nodeValue.length;
          if (offset <= count + len) { found = { node: c, offset: offset - count }; return; }
          count += len;
        } else if (c.nodeType === 1 && !skipNode(c)) walk(c);
      }
    })(p);
    return found;
  }

  function flash(el) {
    el.classList.remove("cc-flash");
    void el.offsetWidth;
    el.classList.add("cc-flash");
    clearTimeout(el._ccFlash);
    el._ccFlash = setTimeout(function () { el.classList.remove("cc-flash"); if (!el.className) el.removeAttribute("class"); }, 2000);
  }

  /** Same paragraph as at review time? Redlines change the middle or end, not the opening words. */
  function samePara(cur, orig) {
    if (orig == null) return false;
    if (cur === orig) return true;
    var n = Math.min(24, orig.length);
    return n > 0 && cur.slice(0, n) === orig.slice(0, n);
  }

  /**
   * Find the paragraph a finding refers to, as the document reads now: the same index if it is still
   * the same paragraph, otherwise a paragraph whose text is unchanged since the review. Never guesses;
   * { none, changed } means the issue is gone (changed=false) or the document moved under it (changed=true).
   */
  function locateFinding(f, base) {
    base = base || state.base;
    var ps = paragraphs();
    if (!base) return { none: true, changed: true };
    var orig = base.snapshot[f.paragraphIndex];
    var changed = ps.length !== base.count;
    var tried = {};
    function attempt(i) {
      tried[i] = true;
      var text = currentText(ps[i]);
      var loc = Engine.locate(f, text);
      return loc ? { p: ps[i], index: i, text: text, loc: loc } : null;
    }
    if (!changed && ps[f.paragraphIndex]) {
      if (samePara(currentText(ps[f.paragraphIndex]), orig)) {
        var same = attempt(f.paragraphIndex);
        if (same) return same;
      } else changed = true;
    }
    if (orig != null && orig.trim()) {
      for (var i = 0; i < ps.length; i++) {
        if (tried[i] || currentText(ps[i]) !== orig) continue;
        var moved = attempt(i);
        if (moved) return moved;
      }
    }
    return { none: true, changed: changed };
  }

  function select(p, start, end) {
    var a = domPoint(p, start), b = domPoint(p, end);
    if (!a || !b) return;
    var r = document.createRange();
    r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset);
    var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
  }

  function goTo(index, f, base) {
    var ps = paragraphs();
    var hit = f && f.matchText ? locateFinding(f, base) : null;
    if (hit && hit.none && hit.changed) { toast("The document changed since the review. Run the review again.", "err"); return; }
    var p = hit && !hit.none ? hit.p : ps[index];
    if (!p) { toast("That paragraph no longer exists. Run the review again.", "err"); return; }
    p.scrollIntoView({ behavior: "smooth", block: "center" });
    flash(p);
    if (hit && !hit.none) select(p, hit.loc.start, hit.loc.start + hit.loc.target.length);
  }

  /** Apply a finding as an in-document redline: <del>old</del><ins>new</ins>, or an appended <ins>. */
  function accept(f, base) {
    base = base || state.base;
    if (!f.action) return "missing";
    var hit = locateFinding(f, base);
    if (hit.none) return hit.changed ? "moved" : "resolved";
    var p = hit.p, loc = hit.loc;
    if (f.action.type === "append") {
      p.appendChild(document.createTextNode(" "));
      p.appendChild(h("ins", { class: "cc-ins", text: f.action.text }));
    } else {
      var a = domPoint(p, loc.start), b = domPoint(p, loc.start + loc.target.length);
      if (!a || !b) return "missing";
      var r = document.createRange();
      r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset);
      var del = h("del", { class: "cc-del" });
      del.appendChild(r.extractContents());
      r.insertNode(del);
      del.parentNode.insertBefore(h("ins", { class: "cc-ins", text: f.action.replace }), del.nextSibling);
    }
    // Remember the paragraph as Clause Craft left it, so later fixes in it still find their place.
    base.snapshot[f.paragraphIndex] = currentText(p);
    return "redlined";
  }

  /* ---------------- selection in the document ---------------- */

  function inEditor(node) { return !!node && node !== editor && editor.contains(node); }

  /** The live selection if it is in the document, else the last one the user made there. */
  function editorRange() {
    var sel = window.getSelection();
    if (sel && sel.rangeCount) {
      var r = sel.getRangeAt(0);
      if (editor.contains(r.startContainer) && editor.contains(r.endContainer)) return r;
    }
    var lr = state.lastRange;
    if (lr && editor.contains(lr.startContainer) && editor.contains(lr.endContainer)) return lr;
    return null;
  }

  /** The paragraph element holding a DOM point (handles points that sit directly on the editor). */
  function paraAt(node, offset, isEnd) {
    var ps = paragraphs();
    if (node === editor) {
      var idx = isEnd ? offset - 1 : offset;
      var child = editor.childNodes[Math.max(0, Math.min(idx, editor.childNodes.length - 1))];
      return ps.indexOf(child) !== -1 ? child : ps[ps.length - 1] || null;
    }
    while (node && node.parentNode !== editor) node = node.parentNode;
    return node && ps.indexOf(node) !== -1 ? node : null;
  }

  function visibleLength(n) {
    if (n.nodeType === 3) return n.nodeValue.length;
    if (n.nodeType !== 1 || skipNode(n)) return 0;
    var len = 0;
    for (var c = n.firstChild; c; c = c.nextSibling) len += visibleLength(c);
    return len;
  }

  /** Number of visible characters in paragraph p before the DOM point (node, offset). */
  function textOffset(p, node, offset) {
    if (node === p || node === editor) {
      if (node === editor) return offset > Array.prototype.indexOf.call(editor.childNodes, p) ? visibleLength(p) : 0;
      var n0 = 0;
      for (var i = 0; i < offset && i < p.childNodes.length; i++) n0 += visibleLength(p.childNodes[i]);
      return n0;
    }
    var count = 0, done = false;
    (function walk(n) {
      for (var c = n.firstChild; c && !done; c = c.nextSibling) {
        if (c === node) {
          if (c.nodeType === 3) count += skipNode(c.parentNode) ? 0 : offset;
          else for (var k = 0; k < offset && k < c.childNodes.length; k++) count += visibleLength(c.childNodes[k]);
          done = true;
          return;
        }
        if (c.nodeType === 3) count += c.nodeValue.length;
        else if (c.nodeType === 1 && !skipNode(c)) walk(c);
        else if (c.nodeType === 1 && c.contains(node)) { done = true; return; }
      }
    })(p);
    return count;
  }

  /** What is selected in the document: paragraph range, their texts, and the selected text itself. */
  function selectionInfo() {
    var r = editorRange();
    if (!r) return null;
    var a = paraAt(r.startContainer, r.startOffset, false), b = paraAt(r.endContainer, r.endOffset, true);
    if (!a || !b) return null;
    var ps = paragraphs();
    var from = ps.indexOf(a), to = ps.indexOf(b);
    if (from > to) { var x = from; from = to; to = x; }
    var lines = [], paraTexts = [];
    for (var i = from; i <= to; i++) {
      var t = currentText(ps[i]);
      paraTexts.push(t);
      var s = i === from ? textOffset(ps[i], r.startContainer, r.startOffset) : 0;
      var e = i === to ? textOffset(ps[i], r.endContainer, r.endOffset) : t.length;
      lines.push(t.slice(s, Math.max(s, e)));
    }
    return { from: from, to: to, texts: paraTexts, text: r.collapsed ? "" : lines.join("\n"), collapsed: r.collapsed };
  }

  function matchSpan(body, sel) {
    if (!sel.length) return null;
    for (var i = 0; i + sel.length <= body.length; i++) {
      var ok = true;
      for (var j = 0; j < sel.length && ok; j++) if (body[i + j] !== sel[j]) ok = false;
      if (ok) return { from: i, to: i + sel.length - 1 };
    }
    return null;
  }

  /**
   * Insert lines as new paragraphs after the paragraph the cursor is in ("cursor") or at the end ("end").
   * tracked: shown as an insertion (blue double underline), like a tracked change in Word.
   * boldFirst: first line in bold (summaries and issues lists).
   */
  function insertLines(lines, where, opts) {
    opts = opts || {};
    var r = where === "cursor" ? editorRange() : null;
    var anchor = r ? paraAt(r.endContainer, r.endOffset, true) : null;
    var ps = paragraphs();
    var atEnd = !anchor || anchor === ps[ps.length - 1];
    var before = ps.length;
    var created = [];
    lines.forEach(function (line, i) {
      var text = String(line);
      var p = h("p");
      if (!text) p.appendChild(h("br"));
      else if (opts.tracked) p.appendChild(h("ins", { class: "cc-ins", text: text }));
      else if (opts.boldFirst && i === 0) p.appendChild(h("strong", { text: text }));
      else p.textContent = text;
      if (anchor) { editor.insertBefore(p, anchor.nextSibling); anchor = p; }
      else editor.appendChild(p);
      created.push(p);
    });
    // Added at the end: earlier paragraphs keep their places, so the review still lines up.
    [state.base, state.pbBase].forEach(function (b) { if (atEnd && b && b.count === before) b.count += created.length; });
    save();
    if (created.length) { created[0].scrollIntoView({ behavior: "smooth", block: "center" }); flash(created[0]); }
    return { hadCursor: where !== "cursor" || !!anchor, first: created[0] || null };
  }

  function insertedToast(title, res, where) {
    if (where === "cursor" && !res.hadCursor) return "\u201C" + title + "\u201D added at the end (click in the document first to place it)";
    return "\u201C" + title + "\u201D inserted" + (where === "end" ? " at the end" : "");
  }

  /* ---------------- review ---------------- */

  /** keepSetupOpen: re-run after a setup change without folding the setup away while it is being edited. */
  function runReview(keepSetupOpen) {
    var t = texts();
    if (!t.some(function (x) { return x.trim(); })) { toast("Load the demo lease or upload a .docx, .txt or .md contract first.", "err"); return; }
    state.result = Engine.review(t, reviewOpts());
    state.playbook = state.playbookChoice === "none" ? null : Engine.playbook(t, pbOpts());
    state.base = { snapshot: t.slice(), count: t.length };
    state.pbBase = { snapshot: t.slice(), count: t.length };
    if (keepSetupOpen !== true) state.setupOpen = false;
    state.scope = null;
    state.summary = null;
    state.issues = null;
    clear($("summary-output"));
    clear($("issues"));
    $("issues-actions").hidden = true;
    if (state.result.docType === "lease" || state.result.docType === "psa") state.library = state.result.docType;
    updatePartyLabels(t);
    renderAll();
    var s = state.result.stats;
    toast("Review done: " + (s.high + s.medium + s.low) + " issues, " + s.missing + " missing", "ok");
  }

  function refreshPlaybook() {
    var t = texts();
    state.playbook = state.playbookChoice === "none" ? null : Engine.playbook(t, pbOpts());
    state.pbBase = { snapshot: t.slice(), count: t.length };
    renderPlaybook();
    updateCounts();
  }

  function renderAll() {
    $("review-empty").hidden = !!state.result;
    $("review-toolbar").hidden = !state.result;
    renderSetup();
    renderStats();
    renderResults();
    renderScope();
    renderChecks();
    renderPlaybook();
    renderClauses();
    renderLibraryFilter();
    renderLibrary();
    renderDrafts();
    renderMyClauses();
    renderChips();
    updateCounts();
  }

  function updateCounts() {
    var r = state.result;
    $("count-review").textContent = r ? String(r.stats.high + r.stats.medium + r.stats.low + r.stats.missing) : "";
    $("count-playbook").textContent = state.playbook && state.playbook.stats.fail ? String(state.playbook.stats.fail) : "";
  }

  function visibleFindings() {
    if (!state.result) return [];
    var sc = state.scope;
    return state.result.findings.filter(function (f) {
      if (!state.filter[f.severity]) return false;
      return !sc || (f.paragraphIndex >= sc.from && f.paragraphIndex <= sc.to);
    });
  }

  function renderStats() {
    var box = $("review-stats");
    clear(box);
    if (!state.result) { box.classList.add("hidden"); return; }
    box.classList.remove("hidden");
    var s = state.result.stats;
    box.appendChild(h("p", { class: "doc-type", text: (state.result.docType === "lease" ? "Lease" : state.result.docType === "psa" ? "Purchase agreement" : "Document") + " \u00B7 reviewed for the " + state.result.perspectiveLabel.toLowerCase() + " side" }));
    var row = h("div", { class: "stat-row" });
    [["high", "High", s.high], ["medium", "Medium", s.medium], ["low", "Low", s.low], ["missing", "Missing", s.missing]].forEach(function (x) {
      row.appendChild(h("button", { type: "button", class: "stat stat-" + x[0] + (state.filter[x[0]] ? "" : " is-off"), "aria-pressed": state.filter[x[0]] ? "true" : "false",
        title: (state.filter[x[0]] ? "Hide " : "Show ") + x[1].toLowerCase() + " items",
        onclick: function () { state.filter[x[0]] = !state.filter[x[0]]; renderStats(); renderResults(); } }, h("b", { text: String(x[2]) }), h("span", { text: x[1] })));
    });
    box.appendChild(row);
  }

  function quote(f) {
    var q = h("blockquote", { class: "cc-quote" });
    var i = f.matchText ? f.excerpt.indexOf(f.matchText) : -1;
    if (i === -1) { q.textContent = "\u201C" + f.excerpt + "\u201D"; return q; }
    q.appendChild(document.createTextNode("\u201C" + f.excerpt.slice(0, i)));
    q.appendChild(h("mark", { text: f.matchText }));
    q.appendChild(document.createTextNode(f.excerpt.slice(i + f.matchText.length) + "\u201D"));
    return q;
  }

  var STATUS_TEXT = { redlined: "\u2713 Redlined", rejected: "Rejected", resolved: "\u2713 Already resolved", inserted: "\u2713 Inserted" };

  function statusSpan(st) {
    return h("span", { class: "cc-status" + (st === "rejected" ? " is-warn" : ""), text: STATUS_TEXT[st] || st });
  }

  function onAccept(f) {
    var out = accept(f, state.base);
    if (out === "missing") { toast("Couldn't find that text. Run the review again.", "err"); return; }
    if (out === "moved") { toast("The document changed since the review. Run the review again.", "err"); return; }
    state.status[keyOf(f)] = out;
    save(); renderResults(); renderPlaybook();
    toast(out === "resolved" ? "Already resolved" : "Redline applied", "ok");
  }

  function findingCard(f) {
    var st = state.status[keyOf(f)];
    var actions = h("div", { class: "cc-actions" },
      h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { goTo(f.paragraphIndex, f); } }, "Go to"));
    if (!st) {
      if (f.action) actions.appendChild(h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: function () { onAccept(f); } }, "Accept"));
      actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { state.status[keyOf(f)] = "rejected"; renderResults(); renderPlaybook(); } }, "Reject"));
    } else actions.appendChild(statusSpan(st));
    var suggest = null;
    if (f.action && f.action.type === "replace") suggest = h("div", { class: "cc-suggest" }, h("strong", { text: "Suggested redline" }), h("del", { text: f.action.find }), " \u2192 ", h("ins", { text: f.action.replace }));
    else if (f.action) suggest = h("div", { class: "cc-suggest" }, h("strong", { text: "Suggested addition" }), h("ins", { text: f.action.text }));
    else if (f.request) suggest = h("div", { class: "cc-suggest is-request" }, h("strong", { text: "What to ask for" }), f.request);
    var riskFor = state.result.perspective === "neutral" ? h("span", { class: "cc-riskfor", text: "Risk for " + f.riskFor.map(cap).join(" / ") }) : null;
    return h("article", { class: "cc-card sev-" + f.severity + (st ? " is-done" : ""), "data-kind": "finding", "data-rule": f.ruleId },
      h("div", { class: "cc-card-head" },
        h("span", { class: "cc-badge " + f.severity, text: sevLabel(f.severity) }),
        h("div", null, h("h3", { text: f.title }), h("p", { class: "cc-meta" },
          h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(f.paragraphIndex, f); } }, f.ref),
          " \u00B7 " + (f.section || f.categoryLabel), riskFor))),
      h("p", { class: "cc-why", text: f.why }),
      quote(f), suggest, actions);
  }

  function insertMissing(m, where) {
    var t = Engine.TEMPLATES[m.templateId];
    var res = insertLines([m.templateText || (t && t.text)], where, { tracked: true });
    state.status["m|" + m.id] = "inserted";
    renderResults();
    toast(insertedToast(m.templateTitle || m.label, res, where), "ok");
  }

  function missingCard(m) {
    var st = state.status["m|" + m.id];
    var actions = h("div", { class: "cc-actions" });
    if (!st && m.templateText) {
      actions.appendChild(h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: function () { insertMissing(m, "end"); } }, "Insert clause"));
      actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onmousedown: keepSelection, onclick: function () { insertMissing(m, "cursor"); } }, "At cursor"));
    }
    if (!st) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { state.status["m|" + m.id] = "rejected"; renderResults(); } }, "Reject"));
    else actions.appendChild(statusSpan(st));
    var riskFor = state.result.perspective === "neutral" ? h("span", { class: "cc-riskfor", text: "Risk for " + m.riskFor.map(cap).join(" / ") }) : null;
    return h("article", { class: "cc-card sev-" + m.severity + (st ? " is-done" : ""), "data-kind": "missing", "data-category": m.category },
      h("div", { class: "cc-card-head" }, h("span", { class: "cc-badge " + m.severity, text: sevLabel(m.severity) }),
        h("div", null, h("h3", { text: m.title }), riskFor ? h("p", { class: "cc-meta" }, riskFor) : null)),
      h("p", { class: "cc-why", text: m.why }),
      m.templateText ? h("details", { class: "cc-template" }, h("summary", { text: "Preview: " + m.templateTitle }), h("p", { text: m.templateText })) : null,
      actions);
  }

  function renderResults() {
    var box = $("review-results");
    clear(box);
    if (!state.result) return;
    var list = visibleFindings();
    if (!list.length) box.appendChild(h("p", { class: "cc-empty", text: state.scope ? "Nothing flagged in your selection." : state.result.findings.length ? "All issues are hidden by the filters." : "No issues flagged for the " + state.result.perspectiveLabel.toLowerCase() + " side. Clause Craft checks common CRE clauses only." }));
    list.forEach(function (f) { box.appendChild(findingCard(f)); });
    if (state.filter.missing && !state.scope && state.result.missing.length) {
      box.appendChild(h("h3", { class: "cc-section", text: "Missing clauses" }));
      state.result.missing.forEach(function (m) { box.appendChild(missingCard(m)); });
    }
  }

  function renderChecks() {
    var ul = $("checks");
    clear(ul);
    var items = state.result && !state.scope ? (state.result.checks || []) : [];
    $("checks-title").hidden = !items.length;
    items.forEach(function (c) {
      ul.appendChild(h("li", { class: "check " + (c.status === "ok" ? "is-ok" : "is-issue") },
        h("span", { class: "check-icon", "aria-hidden": "true", text: c.status === "ok" ? "\u2713" : "!" }),
        h("span", { class: "check-body" }, h("strong", { text: c.label }), h("span", { class: "check-detail", text: c.detail }))));
    });
  }

  function renderScope() {
    var on = !!state.scope;
    $("scope-text").hidden = !on;
    $("btn-scope-clear").hidden = !on;
    if (on) $("scope-text").textContent = "Showing your selection (" + plural(state.scope.to - state.scope.from + 1, "paragraph") + ")";
  }

  function onReviewSelection() {
    if (!state.result) runReview();
    if (!state.result) return;
    var info = selectionInfo();
    if (!info) { toast("Select the paragraphs you want reviewed in the document first.", "err"); return; }
    var span = paragraphs().length === state.base.count ? { from: info.from, to: info.to } : matchSpan(state.base.snapshot, info.texts);
    if (!span) { toast("The document changed since the review. Run the review again, then select once more.", "err"); return; }
    state.scope = span;
    renderResults(); renderScope(); renderChecks();
    var n = visibleFindings().length;
    toast(n ? plural(n, "issue") + " in your selection" : "Nothing flagged in your selection", "ok");
  }

  function onAcceptAll() {
    if (!state.result) return;
    var open = visibleFindings().filter(function (f) { return f.action && !state.status[keyOf(f)]; });
    if (!open.length) { toast("No open suggested redlines"); return; }
    var done = 0, stale = 0;
    open.forEach(function (f) {
      var out = accept(f, state.base);
      if (out === "redlined" || out === "resolved") { state.status[keyOf(f)] = out; if (out === "redlined") done++; }
      else stale++;
    });
    save(); renderResults(); renderPlaybook();
    toast(plural(done, "redline") + " applied" + (stale ? "; " + stale + " need a fresh review" : ""), stale ? "err" : "ok");
  }

  function renderClauses() {
    var box = $("clause-list");
    clear(box);
    var r = state.result;
    $("clauses-hint").textContent = r ? (r.clauses.length + " clause types found. Click one to jump to it.") : "Run a review to list the clauses Clause Craft recognizes. Click one to jump to it.";
    if (!r) return;
    r.clauses.forEach(function (c) {
      var flagged = r.findings.filter(function (f) { return c.paragraphs.indexOf(f.paragraphIndex) !== -1; });
      var worst = flagged.some(function (f) { return f.severity === "high"; }) ? "high" : flagged.some(function (f) { return f.severity === "medium"; }) ? "medium" : flagged.length ? "low" : null;
      box.appendChild(h("button", { type: "button", class: "cc-clause", onclick: function () { goTo(c.first); } },
        h("span", { class: "cc-clause-ref", text: c.ref }),
        h("span", { class: "cc-clause-body" }, h("strong", { text: c.label }), h("span", { text: c.excerpt })),
        worst ? h("span", { class: "cc-badge " + worst, text: plural(flagged.length, "issue") }) : null));
    });
  }

  /* ---------------- setup ---------------- */

  function loadSetup() {
    var s = loadStore(SETUP_KEY, null);
    if (s && typeof s === "object") {
      if (s.perspective) state.perspective = s.perspective;
      if (s.reviewType) state.reviewType = s.reviewType;
      if (s.jurisdiction) state.jurisdiction = s.jurisdiction;
      if (s.playbookChoice) state.playbookChoice = s.playbookChoice;
      if (typeof s.notes === "string") state.notes = s.notes;
      if (s.parties && typeof s.parties === "object") state.customParties = s.parties;
    }
    syncSetupControls();
  }

  function saveSetup() {
    saveStore(SETUP_KEY, { perspective: state.perspective, reviewType: state.reviewType, jurisdiction: state.jurisdiction, playbookChoice: state.playbookChoice, notes: state.notes, parties: state.customParties });
  }

  function syncSetupControls() {
    $("perspective").value = state.perspective;
    $("review-type").value = state.reviewType;
    $("jurisdiction").value = state.jurisdiction;
    renderPlaybookOptions();
    $("deal-notes").value = state.notes;
  }

  function optionText(id, value) {
    var sel = $(id);
    for (var i = 0; i < sel.options.length; i++) if (sel.options[i].value === value) return sel.options[i].text;
    return value;
  }

  function renderSetup() {
    var open = state.setupOpen || !state.result;
    $("setup-fields").hidden = !open;
    $("setup-summary").hidden = open;
    if (state.result) {
      var r = state.result;
      var party = r.perspectiveLabel + (state.perspective === "auto" ? " (auto)" : "");
      var where = state.jurisdiction === "auto" ? (r.california ? "California (detected)" : "Not California") : optionText("jurisdiction", state.jurisdiction);
      $("setup-summary-text").textContent = [party, optionText("review-type", state.reviewType), where, state.playbookChoice === "none" ? "No playbook" : playbookName()].join(" \u00B7 ");
    }
  }

  /** Shows the party names found in the document next to each side ("Tenant: HARBORLINE ADVISORY GROUP, INC."). */
  function updatePartyLabels(t) {
    var found = [];
    try { found = Engine.parties(t || texts()); } catch (e) { found = []; }
    var labels = { tenant: "Tenant", landlord: "Landlord", buyer: "Buyer", seller: "Seller" };
    var sel = $("perspective");
    for (var i = 0; i < sel.options.length; i++) {
      var o = sel.options[i], base = labels[o.value];
      if (!base) continue;
      var hit = null;
      for (var k = 0; k < found.length; k++) if (found[k].role === base) hit = found[k];
      var name = state.customParties[base] || (hit ? hit.name : "");
      o.text = name ? base + ": " + name : base;
    }
    return found;
  }

  /** Re-reads the party names from the document (after edits, or a different document). */
  function onRefreshParties() {
    var found = updatePartyLabels().map(function (p) { return p.role + ": " + p.name; });
    toast(found.length ? "Parties found: " + found.join("; ") : "No party names found in this document. Use Add party to name them.", found.length ? "ok" : "err");
  }

  function toggleAddParty(open) {
    $("party-add").hidden = !open;
    $("btn-party-add").setAttribute("aria-expanded", open ? "true" : "false");
    if (open) $("party-add-name").focus();
  }

  /** Names a side yourself (kept with this document), e.g. when the document calls it Lessee or the name isn't found. */
  function onSaveParty() {
    var role = $("party-add-role").value;
    var name = trim(String($("party-add-name").value || "").replace(/\s+/g, " ")).slice(0, 80);
    if (name) state.customParties[role] = name;
    else delete state.customParties[role];
    saveSetup();
    updatePartyLabels();
    $("party-add-name").value = "";
    toggleAddParty(false);
    toast(name ? "Added " + role + ": " + name : role + " name cleared", "ok");
  }

  function onSetupChange(key, value) {
    state[key] = value;
    saveSetup();
    if (state.result) runReview(true);
  }

  /* ---------------- playbook ---------------- */

  function pbSide() {
    if (state.playbook) return { side: state.playbook.perspective, docType: state.playbook.docType };
    var docType = state.library === "psa" ? "psa" : "lease";
    var p = state.perspective;
    var ok = (docType === "lease" && (p === "tenant" || p === "landlord")) || (docType === "psa" && (p === "buyer" || p === "seller"));
    return { side: ok ? p : (docType === "psa" ? "buyer" : "tenant"), docType: docType };
  }

  function applyPlaybook(it) {
    var out = accept(it, state.pbBase);
    if (out === "missing") { toast("Couldn't find that text. Run the review again.", "err"); return; }
    if (out === "moved") { toast("The document changed since the review. Run the review again.", "err"); return; }
    state.status[keyOf(it)] = out;
    save();
    refreshPlaybook();
    renderResults();
    toast(out === "resolved" ? "Already resolved" : "Redline applied to your standard", "ok");
  }

  function insertPlaybookClause(it) {
    var t = Engine.TEMPLATES[it.templateId];
    if (!t) return;
    var res = insertLines([t.text], "end", { tracked: true });
    state.status[keyOf(it)] = "inserted";
    refreshPlaybook();
    toast(insertedToast(t.title, res, "end"), "ok");
  }

  function pbCard(it) {
    var st = state.status[keyOf(it)];
    var kind = it.status === "pass" ? "pass" : it.status === "fail" ? "fail" : "unknown";
    var actions = h("div", { class: "cc-actions" });
    if (it.paragraphIndex >= 0) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { goTo(it.paragraphIndex, it, state.pbBase); } }, "Go to"));
    if (kind === "fail" && !st && it.action) actions.appendChild(h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: function () { applyPlaybook(it); } }, it.action.type === "replace" ? "Apply redline" : "Add language"));
    if (kind === "fail" && !st && !it.action && it.templateId) actions.appendChild(h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: function () { insertPlaybookClause(it); } }, "Insert clause"));
    if (st) actions.appendChild(statusSpan(st));
    var suggest = null;
    if (kind === "fail" && it.action && it.action.type === "replace") suggest = h("div", { class: "cc-suggest" }, h("strong", { text: "Redline to your standard" }), h("del", { text: it.action.find }), " \u2192 ", h("ins", { text: it.action.replace }));
    else if (kind === "fail" && it.action) suggest = h("div", { class: "cc-suggest" }, h("strong", { text: "Suggested addition" }), h("ins", { text: it.action.text }));
    return h("article", { class: "cc-card pb-card pb-" + kind + (st ? " is-done" : ""), "data-kind": "playbook", "data-standard": it.standardId },
      h("div", { class: "cc-card-head" },
        h("span", { class: "cc-badge pb-badge " + kind, text: kind === "pass" ? "OK" : kind === "fail" ? "Off" : "N/A" }),
        h("div", null,
          h("h3", { text: it.label }),
          h("p", { class: "cc-meta" },
            it.ref ? h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(it.paragraphIndex, it, state.pbBase); } }, it.ref) : null,
            (it.ref ? " \u00B7 " : "") + "Document: " + it.actualText + " \u00B7 Yours: " + it.standardText))),
      kind !== "pass" && it.why ? h("p", { class: "cc-why", text: it.why }) : null,
      suggest,
      actions.childNodes.length ? actions : null);
  }

  function renderPlaybook() {
    var box = $("pb-items"), statsBox = $("pb-stats");
    clear(box); clear(statsBox);
    var side = pbSide();
    $("pb-side").textContent = state.playbook ? (playbookName() + " \u00B7 " + cap(side.side) + " side \u00B7 " + (side.docType === "psa" ? "Purchase agreement" : "Lease")) :
      state.result && state.playbookChoice === "none" ? "Playbook is off. Choose \u201CMy standards\u201D in the review setup." : "Run a review to check the document.";
    renderPbEditor();
    if (!state.playbook) { statsBox.hidden = true; return; }
    var s = state.playbook.stats;
    statsBox.hidden = false;
    [["fail", "Off", s.fail], ["pass", "OK", s.pass], ["unknown", "N/A", s.notFound]].forEach(function (row) {
      statsBox.appendChild(h("div", { class: "stat pb-" + row[0] }, h("b", { text: String(row[2]) }), h("span", { text: row[1] })));
    });
    if (!state.playbook.items.length) box.appendChild(h("p", { class: "cc-empty", text: "No standards apply to this document type and side." }));
    state.playbook.items.forEach(function (it) { box.appendChild(pbCard(it)); });
  }

  function renderPbEditor() {
    var box = $("pb-editor");
    clear(box);
    $("btn-pb-edit").setAttribute("aria-expanded", state.pbEditing ? "true" : "false");
    $("btn-pb-edit").textContent = state.pbEditing ? "Close" : "Edit standards";
    box.hidden = !state.pbEditing;
    if (!state.pbEditing) return;
    var side = pbSide();
    box.appendChild(h("p", { class: "pb-editor-title", text: playbookName() + ": " + side.side + " standards for " + (side.docType === "psa" ? "purchase agreements" : "leases") }));
    Engine.STANDARDS.forEach(function (std) {
      if (std.side !== side.side || std.docType !== side.docType) return;
      var ov = activeStandards()[std.id] || {};
      var on = h("input", { type: "checkbox", "data-std-on": std.id, "aria-label": "Check " + std.label });
      on.checked = !ov.off;
      var row = h("div", { class: "pb-row" }, h("label", { class: "pb-row-label" }, on, h("span", { text: std.label })));
      if (std.type === "max" || std.type === "min") {
        var inp = h("input", { type: "number", step: "any", min: "0", "data-std": std.id, "aria-label": std.label + " standard" });
        inp.value = String(ov.value != null ? ov.value : std.value);
        row.appendChild(h("span", { class: "pb-row-value" }, h("span", { class: "pb-op", text: std.type === "max" ? "\u2264" : "\u2265" }), inp, h("span", { class: "pb-unit", text: std.unit })));
      }
      box.appendChild(row);
    });
    box.appendChild(h("div", { class: "btn-row" },
      h("button", { type: "button", class: "btn btn-primary btn-sm", id: "btn-pb-save", onclick: savePbEditor }, "Save standards"),
      h("button", { type: "button", class: "btn btn-ghost btn-sm", id: "btn-pb-reset", onclick: resetPbEditor }, "Reset to defaults"),
      activePlaybook() ? h("button", { type: "button", class: "btn btn-ghost btn-sm", id: "btn-pb-delete", onclick: deletePlaybook }, "Delete this playbook") : null));
    box.appendChild(h("div", { class: "pb-saveas" },
      h("input", { type: "text", id: "pb-new-name", placeholder: "New playbook name", "aria-label": "New playbook name" }),
      h("button", { type: "button", class: "btn btn-ghost btn-sm", id: "btn-pb-saveas", onclick: saveAsNewPlaybook }, "Save as new playbook")));
  }

  /** The editor's values for the side being edited, as overrides of the built-in defaults. */
  function editorOverrides(side, target) {
    Engine.STANDARDS.forEach(function (std) {
      if (std.side !== side.side || std.docType !== side.docType) return;
      var on = document.querySelector('[data-std-on="' + std.id + '"]');
      var inp = document.querySelector('[data-std="' + std.id + '"]');
      var ov = {};
      if (on && !on.checked) ov.off = true;
      if (inp && inp.value !== "" && !isNaN(Number(inp.value)) && Number(inp.value) !== std.value) ov.value = Number(inp.value);
      if (ov.off || ov.value != null) target[std.id] = ov;
      else delete target[std.id];
    });
    return target;
  }

  function afterPlaybookChange() {
    state.pbEditing = false;
    saveSetup();
    renderPlaybookOptions();
    renderSetup();
    if (state.result && state.playbookChoice !== "none") refreshPlaybook(); else renderPlaybook();
  }

  function saveAsNewPlaybook() {
    var name = trim($("pb-new-name").value).slice(0, 60);
    if (!name) { toast("Give the new playbook a name first.", "err"); $("pb-new-name").focus(); return; }
    var copy = editorOverrides(pbSide(), JSON.parse(JSON.stringify(activeStandards())));
    var pb = { id: String(Date.now()), name: name, standards: copy };
    state.playbooks.push(pb);
    var stored = saveStore(PLAYBOOKS_KEY, state.playbooks);
    state.playbookChoice = "pb:" + pb.id;
    afterPlaybookChange();
    toast(stored ? "Saved as \u201C" + name + "\u201D and switched to it" : "Playbook made for this session; this browser blocked saving it", stored ? "ok" : "err");
  }

  function deletePlaybook() {
    var pb = activePlaybook();
    if (!pb) return;
    state.playbooks = state.playbooks.filter(function (x) { return x.id !== pb.id; });
    saveStore(PLAYBOOKS_KEY, state.playbooks);
    state.playbookChoice = "mine";
    afterPlaybookChange();
    toast("\u201C" + pb.name + "\u201D deleted. Back to My standards.", "ok");
  }

  function renderPlaybookOptions() {
    var sel = $("playbook-choice");
    if (state.playbookChoice.indexOf("pb:") === 0 && !activePlaybook()) state.playbookChoice = "mine";
    clear(sel);
    sel.appendChild(h("option", { value: "mine", text: "My standards" }));
    state.playbooks.forEach(function (pb) { sel.appendChild(h("option", { value: "pb:" + pb.id, text: pb.name })); });
    sel.appendChild(h("option", { value: "none", text: "None" }));
    sel.value = state.playbookChoice;
  }

  function savePbEditor() {
    editorOverrides(pbSide(), activeStandards());
    var stored = saveActiveStandards();
    state.pbEditing = false;
    if (state.result && state.playbookChoice !== "none") refreshPlaybook(); else renderPlaybook();
    toast(stored ? "Standards saved" : "Standards applied, but this browser blocked saving them", stored ? "ok" : "err");
  }

  function resetPbEditor() {
    var side = pbSide();
    var target = activeStandards();
    Engine.STANDARDS.forEach(function (std) { if (std.side === side.side && std.docType === side.docType) delete target[std.id]; });
    saveActiveStandards();
    state.pbEditing = false;
    if (state.result && state.playbookChoice !== "none") refreshPlaybook(); else renderPlaybook();
    toast("Standards reset to defaults", "ok");
  }

  /* ---------------- drafting and saved clauses ---------------- */

  function renderLibraryFilter() {
    Array.prototype.forEach.call(document.querySelectorAll("[data-lib]"), function (b) {
      b.classList.toggle("is-active", b.getAttribute("data-lib") === state.library);
    });
  }

  function renderDrafts() {
    var ul = $("drafts");
    clear(ul);
    var t = texts();
    Object.keys(Engine.DRAFTS).forEach(function (id) {
      var d = Engine.DRAFTS[id];
      if (d.docType !== state.library) return;
      var defaults = Engine.draftDefaults(id, t);
      var inputs = {};
      var preview = h("p", { class: "draft-preview" });
      var form = h("div", { class: "draft-form", hidden: true });
      function values() { var v = {}; for (var k in inputs) v[k] = inputs[k].value; return v; }
      function update() { preview.textContent = Engine.draftClause(id, values(), texts()).text; }
      d.fields.forEach(function (f) {
        var inp = h("input", { type: f.type === "text" ? "text" : "number", "data-key": f.key, "aria-label": f.label });
        if (f.type !== "text") { inp.min = f.min; inp.max = f.max; inp.step = f.step || 1; }
        inp.value = defaults[f.key];
        inp.addEventListener("input", update);
        inputs[f.key] = inp;
        form.appendChild(h("label", { class: "draft-field" }, h("span", { text: f.label }), inp));
      });
      form.appendChild(preview);
      form.appendChild(h("div", { class: "btn-row" },
        h("button", { type: "button", class: "btn btn-primary btn-sm", onmousedown: keepSelection, onclick: function () {
          var c = Engine.draftClause(id, values(), texts());
          if (!c) return;
          var res = insertLines([c.text], "cursor", { tracked: true });
          toast(insertedToast(c.title, res, "cursor"), "ok");
        } }, "Insert at cursor")));
      var toggle = h("button", { type: "button", class: "btn btn-ghost btn-sm", "aria-expanded": "false" }, "Customize");
      toggle.addEventListener("click", function () {
        var open = form.hidden;
        form.hidden = !open;
        toggle.setAttribute("aria-expanded", open ? "true" : "false");
        toggle.textContent = open ? "Close" : "Customize";
        if (open) update();
      });
      ul.appendChild(h("li", { class: "draft", "data-draft": id }, h("div", { class: "lib-head" }, h("span", { class: "lib-title", text: d.title }), toggle), form));
    });
  }

  function renderLibrary() {
    var ul = $("library");
    clear(ul);
    LIBRARY[state.library].forEach(function (id) {
      var t = Engine.TEMPLATES[id];
      if (!t) return;
      ul.appendChild(h("li", { "data-template": id },
        h("div", { class: "lib-head" },
          h("span", { class: "lib-title", text: t.title }),
          h("button", { type: "button", class: "btn btn-ghost btn-sm", onmousedown: keepSelection, onclick: function () {
            var res = insertLines([t.text], "cursor", { tracked: true });
            toast(insertedToast(t.title, res, "cursor"), "ok");
          } }, "Insert")),
        h("details", { class: "cc-template" }, h("summary", { text: "Preview" }), h("p", { text: t.text }))));
    });
  }

  function clauseTitle(text) {
    var m = /^\s*(?:(?:section|article)\s+)?(?:\d+(?:\.\d+)*\.?\s+)?([A-Z][A-Za-z ,&'\/()-]{2,60}?)[.:]\s/.exec(text);
    if (m) return m[1];
    var words = text.split(/\s+/).slice(0, 7).join(" ");
    return words.length < text.length ? words + "\u2026" : words;
  }

  function renderMyClauses() {
    var ul = $("my-clauses");
    clear(ul);
    if (!state.myClauses.length) { ul.appendChild(h("li", { class: "empty-note", text: "No saved clauses yet." })); return; }
    state.myClauses.forEach(function (c, i) {
      ul.appendChild(h("li", { class: "my-clause" },
        h("div", { class: "lib-head" },
          h("span", { class: "lib-title", text: c.title }),
          h("span", { class: "lib-actions" },
            h("button", { type: "button", class: "btn btn-ghost btn-sm", onmousedown: keepSelection, onclick: function () {
              var res = insertLines(c.text.split("\n"), "cursor", { tracked: true });
              toast(insertedToast(c.title, res, "cursor"), "ok");
            } }, "Insert"),
            h("button", { type: "button", class: "btn btn-ghost btn-sm", "aria-label": "Delete " + c.title, onclick: function () {
              state.myClauses.splice(i, 1);
              saveStore(CLAUSES_KEY, state.myClauses);
              renderMyClauses();
            } }, "Delete"))),
        h("details", { class: "cc-template" }, h("summary", { text: "Preview" }), h("p", { text: c.text }))));
    });
  }

  function onSaveClause() {
    var info = selectionInfo();
    var text = info ? trim(String(info.text).replace(/\r\n?/g, "\n")) : "";
    if (text.length < 15) { toast("Select the language you want to save in the document first.", "err"); return; }
    var title = trim($("my-clause-title").value) || clauseTitle(text);
    state.myClauses.unshift({ id: "c" + Date.now(), title: title.slice(0, 80), text: text.slice(0, 20000) });
    var stored = saveStore(CLAUSES_KEY, state.myClauses);
    $("my-clause-title").value = "";
    renderMyClauses();
    toast(stored ? "Saved to My clauses" : "Saved for this session; this browser blocked permanent storage", stored ? "ok" : "err");
  }

  /* ---------------- ask, explain, summary, issues list ---------------- */

  function renderChips() {
    var box = $("ask-chips");
    clear(box);
    var type = state.result ? (state.result.docType === "psa" ? "psa" : "lease") : state.library;
    CHIPS[type].forEach(function (q) { box.appendChild(h("button", { type: "button", class: "chip", onclick: function () { ask(q); } }, q)); });
  }

  function sourcesRow(items) {
    var src = h("div", { class: "cc-sources" });
    items.forEach(function (s) { src.appendChild(h("button", { type: "button", class: "cc-source", title: s.title || "", onclick: function () { goTo(s.index); } }, s.label)); });
    return src;
  }

  function ask(q) {
    q = trim(q);
    if (!q) return;
    var chat = $("ask-chat");
    chat.appendChild(h("div", { class: "cc-msg q", text: q }));
    var a = Engine.ask(q, texts(), { perspective: state.perspective, jurisdiction: state.jurisdiction });
    var msg = h("div", { class: "cc-msg a" + (a.found ? "" : " not-found"), text: a.answer });
    if (a.sources && a.sources.length) msg.appendChild(sourcesRow(a.sources.map(function (s) { return { index: s.paragraphIndex, label: s.ref || "\u00B6" + (s.paragraphIndex + 1), title: s.excerpt }; })));
    chat.appendChild(msg);
    if (msg.scrollIntoView) msg.scrollIntoView({ block: "nearest" });
    $("ask-input").value = "";
  }

  function onExplain() {
    var info = selectionInfo();
    var text = info ? (trim(info.text) ? info.text : info.texts.join("\n")) : "";
    if (!trim(text)) { toast("Select some contract text first, or click into a paragraph.", "err"); return; }
    var x = Engine.explain(text, texts(), reviewOpts());
    var chat = $("ask-chat");
    var quoted = trim(String(text).replace(/\s+/g, " "));
    chat.appendChild(h("div", { class: "cc-msg q", text: "Explain: \u201C" + (quoted.length > 90 ? quoted.slice(0, 89) + "\u2026" : quoted) + "\u201D" }));
    var msg = h("div", { class: "cc-msg a" + (x.found ? "" : " not-found"), text: x.answer });
    if (x.paragraphIndexes.length) msg.appendChild(sourcesRow(x.paragraphIndexes.slice(0, 3).map(function (i) { return { index: i, label: "Go to \u00B6" + (i + 1) }; })));
    chat.appendChild(msg);
    if (msg.scrollIntoView) msg.scrollIntoView({ block: "nearest" });
  }

  function summary() {
    var out = $("summary-output");
    clear(out);
    var t = texts();
    if (!t.some(function (x) { return x.trim(); })) { out.appendChild(h("p", { class: "cc-empty", text: "Load or upload a contract first." })); return; }
    var s = Engine.summarize(t, { perspective: state.perspective, jurisdiction: state.jurisdiction });
    state.summary = s;
    out.appendChild(h("p", { class: "doc-type", text: s.typeLabel + " \u00B7 reviewed for the " + s.perspectiveLabel.toLowerCase() + " side" }));
    var table = h("table", { class: "cc-kv" });
    s.fields.forEach(function (f) {
      table.appendChild(h("tr", null, h("th", { text: f.label }), h("td", null, f.value, " ", h("button", { type: "button", class: "cc-ref", title: "Go to source", onclick: function () { goTo(f.paragraphIndex); } }, "\u2192"))));
    });
    if (s.fields.length) out.appendChild(table);
    if (s.risks.length) {
      out.appendChild(h("h3", { class: "cc-section", text: "Key risks" }));
      var ul = h("ul", { class: "cc-risks" });
      s.risks.forEach(function (r) { ul.appendChild(h("li", null, h("span", { class: "cc-badge " + r.severity, text: sevLabel(r.severity) }), " " + r.title + " ", h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(r.paragraphIndex); } }, r.ref))); });
      out.appendChild(ul);
    }
    if (s.missing.length) out.appendChild(h("p", { class: "cc-missing", text: "Missing: " + s.missing.join(", ") }));
    out.appendChild(h("div", { class: "btn-row" },
      h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { copyText(s.text, "Summary copied"); } }, "Copy summary"),
      h("button", { type: "button", class: "btn btn-ghost btn-sm", onmousedown: keepSelection, onclick: function () {
        var res = insertLines(s.text.split("\n"), "cursor", { boldFirst: true });
        toast(insertedToast("Summary", res, "cursor"), "ok");
      } }, "Insert at cursor"),
      h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () {
        insertLines(s.text.split("\n"), "end", { boldFirst: true });
        toast("Summary added at the end", "ok");
      } }, "Insert at end")));
  }

  function onIssues() {
    var t = texts();
    if (!t.some(function (x) { return x.trim(); })) { toast("Load or upload a contract first.", "err"); return; }
    var o = reviewOpts();
    o.notes = state.notes;
    var il = Engine.issuesList(t, o);
    state.issues = il;
    var ol = $("issues");
    clear(ol);
    il.items.forEach(function (it) {
      ol.appendChild(h("li", { class: "issue" },
        h("div", { class: "issue-head" }, h("span", { class: "cc-badge " + it.severity, text: sevLabel(it.severity) }), h("strong", { text: it.title })),
        h("p", { class: "cc-meta" },
          it.paragraphIndex >= 0 ? h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(it.paragraphIndex); } }, it.ref) : it.ref,
          it.section ? " \u00B7 " + it.section : ""),
        h("p", { class: "issue-request", text: it.request })));
    });
    if (!il.items.length) ol.appendChild(h("li", { class: "empty-note", text: "Nothing to raise for this side." }));
    $("issues-actions").hidden = !il.items.length;
  }

  /* ---------------- export: HTML and Word (.docx with tracked changes) ---------------- */

  function docTitle() {
    var first = paragraphs()[0];
    return (first && trim(currentText(first))) || "Contract";
  }

  function fileStem() {
    return docTitle().replace(/[^\w\- ]+/g, "").trim().replace(/\s+/g, "-").toLowerCase().slice(0, 60) || "contract";
  }

  function exportHtml() {
    var title = docTitle();
    var doc = document.implementation.createHTMLDocument(title + " - Clause Craft redline");
    var style = doc.createElement("style");
    style.textContent = "body{font-family:Georgia,'Times New Roman',serif;max-width:760px;margin:40px auto;line-height:1.6;color:#111;padding:0 20px}" +
      "del{color:#B42318;text-decoration:line-through}ins{color:#1D4ED8;text-decoration:underline double}.doc-title{text-align:center;font-weight:700}" +
      ".doc-heading{font-weight:700}footer{margin-top:40px;font:12px system-ui;color:#666}";
    doc.head.appendChild(style);
    paragraphs().forEach(function (p) { doc.body.appendChild(doc.importNode(p, true)); });
    var foot = doc.createElement("footer");
    foot.textContent = "Redline prepared with Clause Craft (rule-based). Demo only \u2014 not legal advice. Have counsel review before signing.";
    doc.body.appendChild(foot);
    download(fileStem() + "-redline.html", new Blob(["<!DOCTYPE html>\n" + doc.documentElement.outerHTML], { type: "text/html" }));
  }

  var W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  function xmlEsc(s) {
    return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  /** Runs of one paragraph: text with bold/italic/underline/strike and whether it is a Clause Craft insertion or deletion. */
  function paraParts(p) {
    var parts = [];
    (function walk(n, fmt) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) { if (c.nodeValue) parts.push({ text: c.nodeValue.replace(/\u00A0/g, " "), fmt: fmt }); continue; }
        if (c.nodeType !== 1 || c.classList.contains("cc-note")) continue;
        var t = c.tagName, f = { b: fmt.b, i: fmt.i, u: fmt.u, s: fmt.s, kind: fmt.kind };
        if (t === "BR") { parts.push({ br: true, fmt: fmt }); continue; }
        if (t === "STRONG" || t === "B") f.b = true;
        if (t === "EM" || t === "I") f.i = true;
        if (t === "U") f.u = true;
        if (t === "S" || t === "STRIKE") f.s = true;
        if (t === "DEL") f.kind = "del";
        if (t === "INS" && f.kind !== "del") f.kind = "ins";
        walk(c, f);
      }
    })(p, {});
    return parts;
  }

  function docxDocumentXml() {
    var ps = paragraphs();
    var stamp = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
    var rev = 0;
    function mark() { rev++; return ' w:id="' + rev + '" w:author="Clause Craft" w:date="' + stamp + '"'; }
    var body = ps.map(function (p, idx) {
      var title = p.classList.contains("doc-title"), heading = p.classList.contains("doc-heading");
      var parts = paraParts(p);
      // A trailing <br> in an otherwise empty paragraph is just the browser's placeholder.
      if (parts.length === 1 && parts[0].br) parts = [];
      var visible = parts.filter(function (x) { return x.br || trim(x.text); });
      var wholeInsert = visible.length > 0 && visible.every(function (x) { return x.fmt.kind === "ins"; });
      var pPr = "<w:pPr>" + (heading ? "<w:keepNext/>" : "") +
        (title ? '<w:spacing w:after="360"/><w:jc w:val="center"/>' : heading ? '<w:spacing w:before="240" w:after="120"/>' : '<w:jc w:val="both"/>') +
        (wholeInsert ? "<w:rPr><w:ins" + mark() + "/></w:rPr>" : "") + "</w:pPr>";
      var out = "", open = null;
      parts.forEach(function (x) {
        var kind = x.fmt.kind || null;
        if (kind !== open) {
          if (open) out += open === "ins" ? "</w:ins>" : "</w:del>";
          if (kind) out += kind === "ins" ? "<w:ins" + mark() + ">" : "<w:del" + mark() + ">";
          open = kind;
        }
        var b = x.fmt.b || title || heading;
        var rPr = (b ? "<w:b/>" : "") + (x.fmt.i ? "<w:i/>" : "") + (x.fmt.s ? "<w:strike/>" : "") + (x.fmt.u ? '<w:u w:val="single"/>' : "");
        out += "<w:r>" + (rPr ? "<w:rPr>" + rPr + "</w:rPr>" : "") +
          (x.br ? "<w:br/>" : kind === "del" ? '<w:delText xml:space="preserve">' + xmlEsc(x.text) + "</w:delText>" : '<w:t xml:space="preserve">' + xmlEsc(x.text) + "</w:t>") + "</w:r>";
      });
      if (open) out += open === "ins" ? "</w:ins>" : "</w:del>";
      return "<w:p>" + pPr + out + "</w:p>";
    }).join("");
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="' + W + '"><w:body>' + body +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      "</w:body></w:document>";
  }

  var CRC_TABLE = (function () {
    var t = [];
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /** A minimal zip (stored, no compression), which is all a .docx needs. */
  function zip(files) {
    var enc = new TextEncoder(), chunks = [], central = [], offset = 0;
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
      var head = new DataView(new ArrayBuffer(30));
      head.setUint32(0, 0x04034b50, true); head.setUint16(4, 20, true); head.setUint16(6, 0, true); head.setUint16(8, 0, true);
      head.setUint16(10, dosTime, true); head.setUint16(12, dosDate, true); head.setUint32(14, crc, true);
      head.setUint32(18, data.length, true); head.setUint32(22, data.length, true); head.setUint16(26, name.length, true); head.setUint16(28, 0, true);
      var cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0, true); cd.setUint16(10, 0, true);
      cd.setUint16(12, dosTime, true); cd.setUint16(14, dosDate, true); cd.setUint32(16, crc, true);
      cd.setUint32(20, data.length, true); cd.setUint32(24, data.length, true); cd.setUint16(28, name.length, true);
      cd.setUint16(30, 0, true); cd.setUint16(32, 0, true); cd.setUint16(34, 0, true); cd.setUint16(36, 0, true); cd.setUint32(38, 0, true);
      cd.setUint32(42, offset, true);
      chunks.push(new Uint8Array(head.buffer), name, data);
      central.push(new Uint8Array(cd.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var cdSize = central.reduce(function (n, c) { return n + c.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(4, 0, true); end.setUint16(6, 0, true);
    end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
    return new Blob(chunks.concat(central, [new Uint8Array(end.buffer)]), { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  function exportDocx() {
    if (!paragraphs().length) { toast("Nothing to download yet.", "err"); return; }
    var RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
    var OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    var files = [
      { name: "[Content_Types].xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>' },
      { name: "_rels/.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + RELS + '">' +
        '<Relationship Id="rId1" Type="' + OFFICE + '/officeDocument" Target="word/document.xml"/></Relationships>' },
      { name: "word/_rels/document.xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + RELS + '">' +
        '<Relationship Id="rId1" Type="' + OFFICE + '/styles" Target="styles.xml"/><Relationship Id="rId2" Type="' + OFFICE + '/settings" Target="settings.xml"/></Relationships>' },
      { name: "word/document.xml", data: docxDocumentXml() },
      { name: "word/styles.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + W + '"><w:docDefaults>' +
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>' },
      // Track changes stays on, so edits made in Word are tracked too.
      { name: "word/settings.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="' + W + '"><w:trackRevisions/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>' }
    ];
    download(fileStem() + "-redline.docx", zip(files));
    toast("Word file downloaded. Your redlines open as tracked changes.", "ok");
  }

  /* ---------------- upload (.docx / .txt / .md) ---------------- */

  /* .docx reader: a .docx is a zip; pull out word/document.xml with the browser's own
     DecompressionStream and keep each paragraph's visible text (tracked deletions skipped). */
  function docxXml(buf) {
    var u8 = new Uint8Array(buf), dv = new DataView(buf), dec = new TextDecoder();
    var eocd = -1;
    for (var i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) return Promise.reject(new Error("That file isn't a valid .docx."));
    var entries = dv.getUint16(eocd + 10, true), off = dv.getUint32(eocd + 16, true);
    for (var k = 0; k < entries && off + 46 <= u8.length; k++) {
      if (dv.getUint32(off, true) !== 0x02014b50) break;
      var method = dv.getUint16(off + 10, true), size = dv.getUint32(off + 20, true);
      var nameLen = dv.getUint16(off + 28, true), extraLen = dv.getUint16(off + 30, true), commentLen = dv.getUint16(off + 32, true);
      var local = dv.getUint32(off + 42, true);
      if (dec.decode(u8.subarray(off + 46, off + 46 + nameLen)) === "word/document.xml") {
        var start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
        var data = u8.subarray(start, start + size);
        if (method === 0) return Promise.resolve(dec.decode(data));
        if (method !== 8) return Promise.reject(new Error("That .docx uses a compression this page can't read."));
        if (typeof DecompressionStream === "undefined") return Promise.reject(new Error("This browser can't open .docx files. Save the contract as .txt and upload that."));
        return new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text();
      }
      off += 46 + nameLen + extraLen + commentLen;
    }
    return Promise.reject(new Error("That file isn't a Word document."));
  }

  function docxParagraphs(xml) {
    var doc = new DOMParser().parseFromString(xml, "application/xml");
    var body = doc.getElementsByTagNameNS(W, "body")[0];
    if (!body) throw new Error("That file isn't a Word document.");
    var SKIP = { del: 1, moveFrom: 1, delText: 1, instrText: 1, pPr: 1, rPr: 1, p: 1 };
    function text(p) {
      var out = "";
      (function walk(n) {
        for (var c = n.firstChild; c; c = c.nextSibling) {
          if (c.nodeType !== 1 || c.localName === "Fallback") continue;
          if (c.namespaceURI === W) {
            if (SKIP[c.localName]) continue;
            if (c.localName === "t") { out += c.textContent; continue; }
            if (c.localName === "tab" || c.localName === "br" || c.localName === "cr") { out += " "; continue; }
            if (c.localName === "noBreakHyphen") { out += "-"; continue; }
          }
          walk(c);
        }
      })(p);
      return out.trim();
    }
    var lines = [];
    (function collect(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || c.localName === "Fallback") continue;
        if (c.namespaceURI === W && c.localName === "p") lines.push(text(c));
        if (c.namespaceURI === W && (c.localName === "del" || c.localName === "moveFrom" || c.localName === "sectPr")) continue;
        collect(c); // tables, content controls, text boxes
      }
    })(body);
    while (lines.length && !lines[lines.length - 1]) lines.pop();
    return lines;
  }

  function upload(file) {
    if (!file) return;
    var name = file.name || "file";
    if (/\.docx$/i.test(name) || file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
      if (file.size > 25 * 1024 * 1024) { toast("That .docx is larger than 25 MB.", "err"); return; }
      file.arrayBuffer().then(docxXml).then(docxParagraphs).then(function (lines) {
        if (!lines.some(function (x) { return x.trim(); })) throw new Error("No text found in " + name + ".");
        setDocument(lines);
        toast("Loaded " + name + ". Run Review when ready.", "ok");
      }).catch(function (e) { toast(e && e.message ? e.message : "Couldn't open " + name + ".", "err"); });
      return;
    }
    if (/\.(doc|pdf|rtf|pages)$/i.test(name)) { toast("Save it as .docx (or .txt) first, then upload.", "err"); return; }
    if (file.size > 2 * 1024 * 1024) { toast("File is larger than 2 MB.", "err"); return; }
    var reader = new FileReader();
    reader.onload = function () {
      setDocument(String(reader.result || "").replace(/\r\n?/g, "\n").split("\n"));
      toast("Loaded " + name + ". Run Review when ready.", "ok");
    };
    reader.readAsText(file);
  }

  /* ---------------- chrome ---------------- */

  function switchTab(name) {
    Array.prototype.forEach.call(document.querySelectorAll(".sidebar-tabs .tab"), function (t) {
      var on = t.getAttribute("data-tab") === name;
      t.classList.toggle("active", on);
      t.setAttribute("aria-selected", on ? "true" : "false");
    });
    Array.prototype.forEach.call(document.querySelectorAll(".tab-panel"), function (p) {
      var on = p.id === "panel-" + name;
      p.classList.toggle("active", on);
      p.hidden = !on;
    });
  }

  function isDemoText(t) {
    return ["lease", "psa"].some(function (id) {
      var d = Demo[id].paragraphs;
      return t.length >= d.length && d.every(function (x, i) { return t[i] === x; });
    });
  }

  function loadDemo(id) {
    var t = texts();
    if (t.some(function (x) { return x.trim(); }) && !isDemoText(t) && !window.confirm("Replace the current document with the demo " + (id === "psa" ? "purchase agreement" : "lease") + "?")) return;
    setDocument(Demo[id].paragraphs);
    runReview();
  }

  function bind() {
    $("btn-load-demo").addEventListener("click", function () { loadDemo("lease"); });
    Array.prototype.forEach.call(document.querySelectorAll("[data-demo]"), function (b) {
      b.addEventListener("click", function () { loadDemo(b.getAttribute("data-demo")); });
    });
    $("file-input").addEventListener("change", function (e) { upload(e.target.files[0]); e.target.value = ""; });
    $("file-label").addEventListener("keydown", function (e) {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("file-input").click(); }
    });
    var pane = document.querySelector(".editor-pane");
    pane.addEventListener("dragover", function (e) {
      if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], "Files") !== -1) { e.preventDefault(); pane.classList.add("is-dropping"); }
    });
    pane.addEventListener("dragleave", function (e) { if (!pane.contains(e.relatedTarget)) pane.classList.remove("is-dropping"); });
    pane.addEventListener("drop", function (e) {
      pane.classList.remove("is-dropping");
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) { e.preventDefault(); upload(e.dataTransfer.files[0]); }
    });
    $("btn-docx").addEventListener("click", exportDocx);
    $("btn-export").addEventListener("click", exportHtml);
    $("btn-print").addEventListener("click", function () { window.print(); });
    $("btn-run-review").addEventListener("click", function () { switchTab("review"); runReview(); });
    $("review-setup").addEventListener("submit", function (e) { e.preventDefault(); runReview(); });
    $("btn-run-review-side").addEventListener("mousedown", keepSelection);
    $("btn-setup-edit").addEventListener("click", function () { state.setupOpen = true; renderSetup(); $("review-type").focus(); });
    $("perspective").addEventListener("change", function (e) { onSetupChange("perspective", e.target.value); });
    $("review-type").addEventListener("change", function (e) { onSetupChange("reviewType", e.target.value); });
    $("jurisdiction").addEventListener("change", function (e) { onSetupChange("jurisdiction", e.target.value); });
    $("playbook-choice").addEventListener("change", function (e) { onSetupChange("playbookChoice", e.target.value); });
    $("btn-parties-refresh").addEventListener("click", onRefreshParties);
    $("btn-party-add").addEventListener("click", function () { toggleAddParty($("party-add").hidden); });
    $("btn-party-save").addEventListener("click", onSaveParty);
    $("btn-party-cancel").addEventListener("click", function () { toggleAddParty(false); });
    $("party-add-name").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); onSaveParty(); } });
    var notesTimer = null;
    $("deal-notes").addEventListener("input", function (e) {
      state.notes = e.target.value;
      clearTimeout(notesTimer);
      notesTimer = setTimeout(saveSetup, 400);
    });
    $("qs-review").addEventListener("click", function () { runReview(); });
    $("qs-draft").addEventListener("click", function () { switchTab("draft"); });
    $("qs-playbook").addEventListener("click", function () { if (!state.result) runReview(); switchTab("playbook"); });
    ["btn-review-selection", "btn-explain", "btn-save-clause"].forEach(function (id) { $(id).addEventListener("mousedown", keepSelection); });
    $("btn-review-selection").addEventListener("click", onReviewSelection);
    $("btn-scope-clear").addEventListener("click", function () { state.scope = null; renderResults(); renderScope(); renderChecks(); });
    $("btn-accept-all").addEventListener("click", onAcceptAll);
    $("btn-pb-edit").addEventListener("click", function () { state.pbEditing = !state.pbEditing; renderPbEditor(); });
    Array.prototype.forEach.call(document.querySelectorAll("[data-lib]"), function (b) {
      b.addEventListener("click", function () { state.library = b.getAttribute("data-lib"); renderLibraryFilter(); renderLibrary(); renderDrafts(); });
    });
    $("btn-save-clause").addEventListener("click", onSaveClause);
    $("btn-explain").addEventListener("click", onExplain);
    $("btn-summary").addEventListener("click", summary);
    $("btn-issues").addEventListener("click", onIssues);
    $("btn-issues-end").addEventListener("click", function () {
      if (!state.issues) return;
      insertLines(state.issues.text.split("\n"), "end", { boldFirst: true });
      toast("Issues list added at the end", "ok");
    });
    $("btn-issues-copy").addEventListener("click", function () { if (state.issues) copyText(state.issues.text, "Issues list copied"); });
    $("ask-form").addEventListener("submit", function (e) { e.preventDefault(); ask($("ask-input").value); });
    Array.prototype.forEach.call(document.querySelectorAll(".sidebar-tabs .tab"), function (t) { t.addEventListener("click", function () { switchTab(t.getAttribute("data-tab")); }); });
    Array.prototype.forEach.call(document.querySelectorAll("#floating-toolbar [data-cmd]"), function (b) {
      b.addEventListener("mousedown", function (e) { e.preventDefault(); });
      b.addEventListener("click", function () { document.execCommand(b.getAttribute("data-cmd"), false, null); save(); });
    });
    document.addEventListener("selectionchange", function () {
      var sel = window.getSelection();
      if (!sel || !sel.rangeCount) return;
      var r = sel.getRangeAt(0);
      if (inEditor(r.startContainer) || (r.startContainer === editor && editor.childNodes.length)) state.lastRange = r.cloneRange();
    });
    var timer = null;
    editor.addEventListener("input", function () { clearTimeout(timer); timer = setTimeout(save, 400); });
    editor.addEventListener("paste", function (e) {
      // Paste as plain text, one paragraph per line.
      var text = (e.clipboardData || window.clipboardData).getData("text");
      if (!text) return;
      e.preventDefault();
      if (!editor.textContent.trim()) { setDocument(text.replace(/\r\n?/g, "\n").split("\n")); return; }
      document.execCommand("insertText", false, text);
    });
  }

  function init() {
    if (!Engine || !Demo) { document.body.insertBefore(h("p", { class: "cc-fatal", text: "Clause Craft engine failed to load (engine/clause-engine.js). Keep the engine folder next to index.html." }), document.body.firstChild); return; }
    $("version").textContent = "v" + Engine.VERSION;
    loadSetup();
    bind();
    if (!restore()) setDocument(Demo.lease.paragraphs);
    else {
      var dt = docTypeOf(texts());
      if (dt) state.library = dt;
      updatePartyLabels();
    }
    renderAll();
    switchTab("review");
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
