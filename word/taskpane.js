/*
 * Clause Craft - Word task pane (Office.js).
 * Reads the open document, runs the shared rule engine (../engine/clause-engine.js) locally,
 * and writes redlines (tracked changes), comments and clauses back into Word.
 * No network calls, no API keys, no eval. Text from the document is only ever set with textContent.
 */
(function () {
  "use strict";

  var Engine = window.ClauseCraftEngine;
  var Demo = window.ClauseCraftDemoDocs;
  var RED = "#C0392B";
  var BLUE = "#1D4ED8";
  var SETTINGS_KEY = "clauseCraft.perspective";
  var STANDARDS_KEY = "clauseCraft.standards.v1";
  var CLAUSES_KEY = "clauseCraft.myClauses.v1";

  /* Per-computer storage for the playbook standards and saved clauses (never sent anywhere). */
  function loadStore(key, fallback) {
    try { var v = window.localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; }
  }
  function saveStore(key, value) {
    try { window.localStorage.setItem(key, JSON.stringify(value)); return true; } catch (e) { return false; }
  }

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

  var state = {
    inWord: false,
    api13: false,
    api14: false,
    tracking: true,
    perspective: "auto",
    result: null,
    texts: null,
    status: {},          // finding key -> { redlined, commented, dismissed, resolved, error, inserted }
    filter: { high: true, medium: true, low: true, missing: true },
    library: "lease",
    summary: null,
    playbook: null,
    standards: loadStore(STANDARDS_KEY, {}),
    myClauses: loadStore(CLAUSES_KEY, []),
    scope: null,          // { from, to } when showing only the issues in the selected paragraphs
    pbEditing: false,
    issues: null,
    reviewType: "comprehensive",
    jurisdiction: "auto",
    playbookChoice: "mine",
    notes: "",
    setupOpen: true,
    parties: []
  };

  function reviewOpts() { return { perspective: state.perspective, reviewType: state.reviewType, jurisdiction: state.jurisdiction }; }
  function pbOpts() { return { perspective: state.perspective, jurisdiction: state.jurisdiction, standards: state.standards }; }

  /* ---------------------------------------------------------------- *
   * DOM helpers
   * ---------------------------------------------------------------- */

  function $(id) { return document.getElementById(id); }

  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) {
      for (var k in attrs) {
        if (!Object.prototype.hasOwnProperty.call(attrs, k) || attrs[k] == null || attrs[k] === false) continue;
        if (k === "text") el.textContent = attrs[k];
        else if (k === "class") el.className = attrs[k];
        else if (k.indexOf("on") === 0 && typeof attrs[k] === "function") el.addEventListener(k.slice(2), attrs[k]);
        else el.setAttribute(k, attrs[k] === true ? "" : attrs[k]);
      }
    }
    for (var i = 2; i < arguments.length; i++) {
      var c = arguments[i];
      if (c == null || c === false) continue;
      if (Array.isArray(c)) c.forEach(function (x) { if (x) el.appendChild(typeof x === "string" ? document.createTextNode(x) : x); });
      else el.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    }
    return el;
  }

  function clear(el) { while (el && el.firstChild) el.removeChild(el.firstChild); }

  function toast(message, kind) {
    var t = $("toast");
    t.textContent = message;
    t.className = "toast" + (kind === "ok" ? " is-ok" : kind === "err" ? " is-err" : "");
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, kind === "err" ? 5200 : 2400);
  }

  function busy(btn, on) {
    if (!btn) return;
    btn.classList.toggle("is-busy", !!on);
    btn.setAttribute("aria-busy", on ? "true" : "false");
  }

  function errorMessage(e) {
    if (!e) return "Unknown error";
    var msg = e.message || String(e);
    if (e.debugInfo && e.debugInfo.message && e.debugInfo.message !== msg) msg += " (" + e.debugInfo.message + ")";
    return msg;
  }

  function keyOf(f) { return f.ruleId + "|" + (f.matchText || "") + "|" + (f.excerpt || "").slice(0, 60); }

  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

  /* ---------------------------------------------------------------- *
   * Word access
   * ---------------------------------------------------------------- */

  function searchSafe(s) {
    // Word's search treats "^" as a special-character prefix even without wildcards.
    return String(s).replace(/\^/g, "^^").slice(0, 255);
  }

  /** Loads every body paragraph and its current text (tracked deletions excluded when the API allows). */
  async function loadTexts(context) {
    var ps = context.document.body.paragraphs;
    ps.load("items/text");
    await context.sync();
    var items = ps.items;
    var texts = items.map(function (p) { return p.text || ""; });
    if (state.api14 && items.length && typeof items[0].getReviewedText === "function") {
      try {
        var results = items.map(function (p) { return p.getReviewedText("Current"); });
        await context.sync();
        texts = results.map(function (r, i) { return typeof r.value === "string" ? r.value : texts[i]; });
      } catch (e) {
        // Older builds: fall back to paragraph.text
      }
    }
    return { items: items, texts: texts };
  }

  function readParagraphs() {
    return Word.run(async function (context) {
      var loaded = await loadTexts(context);
      return loaded.texts;
    });
  }

  /**
   * Finds the paragraph that still holds the finding. Returns null when it has already been fixed.
   * Indexes can shift when paragraphs are added above, so nearby paragraphs are checked too.
   */
  function findTarget(f, texts) {
    if (f.action && f.action.type === "append") {
      var probe = Engine.norm(f.action.text).slice(0, 60);
      for (var a = 0; a < texts.length; a++) if (Engine.norm(texts[a]).indexOf(probe) !== -1) return null;
    }
    function has(i) { return i >= 0 && i < texts.length && !!f.matchText && texts[i].indexOf(f.matchText) !== -1; }
    function at(i) { var loc = Engine.locate(f, texts[i]); return loc ? { index: i, loc: loc } : null; }
    if (has(f.paragraphIndex)) return at(f.paragraphIndex); // still there (or fixed in place -> null)
    for (var d = 1; d < texts.length; d++) {
      if (has(f.paragraphIndex - d)) { var x = at(f.paragraphIndex - d); if (x) return x; }
      if (has(f.paragraphIndex + d)) { var y = at(f.paragraphIndex + d); if (y) return y; }
      if (f.paragraphIndex - d < 0 && f.paragraphIndex + d >= texts.length) break;
    }
    return null;
  }

  function occurrencesBefore(hay, needle, before) {
    var n = 0, i = hay.indexOf(needle);
    while (i !== -1 && i < before) { n++; i = hay.indexOf(needle, i + 1); }
    return n;
  }

  /** Resolves the Word range for a located target inside a paragraph. */
  async function findRange(context, para, loc) {
    if (!loc || !loc.target) return null;
    if (loc.context && loc.context.length <= 255) {
      var cRes = para.search(searchSafe(loc.context), { matchCase: true });
      cRes.load("items");
      await context.sync();
      if (cRes.items.length) {
        var base = cRes.items[0];
        if (loc.context === loc.target) return base;
        var tRes = base.search(searchSafe(loc.target), { matchCase: true });
        tRes.load("items");
        await context.sync();
        if (tRes.items.length) {
          var k = occurrencesBefore(loc.context, loc.target, loc.offset);
          return tRes.items[Math.min(k, tRes.items.length - 1)];
        }
      }
    }
    var res = para.search(searchSafe(loc.target), { matchCase: true });
    res.load("items");
    await context.sync();
    if (!res.items.length) return null;
    return res.items[Math.min(loc.occurrence || 0, res.items.length - 1)];
  }

  async function withTracking(context, fn) {
    var track = state.api14 && state.tracking;
    var prev = null;
    if (track) {
      context.document.load("changeTrackingMode");
      await context.sync();
      prev = context.document.changeTrackingMode;
      if (prev !== "TrackAll") context.document.changeTrackingMode = "TrackAll";
    }
    var out = await fn(track);
    await context.sync();
    if (track && prev && prev !== "TrackAll") {
      context.document.changeTrackingMode = prev;
      await context.sync();
    }
    return out;
  }

  function markInserted(range) {
    range.font.strikeThrough = false;
    range.font.underline = "Double";
    range.font.color = BLUE;
  }

  function goToFinding(f) {
    return Word.run(async function (context) {
      var loaded = await loadTexts(context);
      var hit = findTarget(f, loaded.texts);
      var idx = hit ? hit.index : f.paragraphIndex;
      if (idx < 0 || idx >= loaded.items.length) throw new Error("That text is no longer in the document. Run the review again.");
      var para = loaded.items[idx];
      var range = hit ? await findRange(context, para, hit.loc) : null;
      (range || para).select();
      await context.sync();
    }).catch(function (e) { toast(errorMessage(e), "err"); });
  }

  function goToParagraph(index) {
    return Word.run(async function (context) {
      var ps = context.document.body.paragraphs;
      ps.load("items");
      await context.sync();
      if (index < 0 || index >= ps.items.length) throw new Error("That paragraph no longer exists. Run the review again.");
      ps.items[index].select();
      await context.sync();
    }).catch(function (e) { toast(errorMessage(e), "err"); });
  }

  /** Applies one finding's suggested edit. Returns "redlined" or "resolved". */
  function applyFinding(f) {
    return Word.run(async function (context) {
      var loaded = await loadTexts(context);
      var hit = findTarget(f, loaded.texts);
      if (!hit) return "resolved";
      var para = loaded.items[hit.index];
      var range = null;
      if (f.action.type === "replace") {
        range = await findRange(context, para, hit.loc);
        if (!range) throw new Error("Couldn't find \u201C" + f.matchText.slice(0, 60) + "\u201D in Word. Select it and edit by hand.");
      }
      await withTracking(context, async function (track) {
        if (f.action.type === "append") {
          var added = para.insertText(" " + f.action.text, "End");
          if (!track) markInserted(added);
        } else if (track) {
          range.insertText(f.action.replace, "Replace");
        } else {
          range.font.strikeThrough = true;
          range.font.color = RED;
          var ins = range.insertText(" " + f.action.replace, "After");
          markInserted(ins);
        }
      });
      return "redlined";
    });
  }

  function commentFinding(f, context, loaded) {
    var run = async function (ctx, data) {
      var hit = findTarget(f, data.texts);
      var idx = hit ? hit.index : f.paragraphIndex;
      if (idx < 0 || idx >= data.items.length) return "resolved";
      var para = data.items[idx];
      var range = hit ? await findRange(ctx, para, hit.loc) : null;
      (range || para.getRange("Whole")).insertComment(f.comment);
      return "commented";
    };
    if (context) return run(context, loaded);
    return Word.run(async function (ctx) {
      var data = await loadTexts(ctx);
      var out = await run(ctx, data);
      await ctx.sync();
      return out;
    });
  }

  function insertParagraphs(lines, where, opts) {
    opts = opts || {};
    return Word.run(async function (context) {
      var body = context.document.body;
      // Contract language (templates) is inserted as a tracked change; notes such as summaries are not.
      var wrap = opts.tracked ? withTracking : async function (ctx, fn) { var r = await fn(true); await ctx.sync(); return r; };
      await wrap(context, async function (track) {
        var anchor = null;
        if (where === "cursor") {
          var sel = context.document.getSelection();
          anchor = sel;
        }
        for (var i = 0; i < lines.length; i++) {
          var p = anchor ? anchor.insertParagraph(lines[i], "After") : body.insertParagraph(lines[i], "End");
          if (opts.boldFirst && i === 0) p.font.bold = true;
          else if (opts.boldFirst) p.font.bold = false;
          if (!track && opts.markInserted) markInserted(p.getRange("Whole"));
          anchor = where === "cursor" ? p : null;
        }
      });
    });
  }

  /* ---------------------------------------------------------------- *
   * Review
   * ---------------------------------------------------------------- */

  /** keepSetupOpen: re-run after a setup change without folding the setup away while it is being edited. */
  async function runReview(keepSetupOpen) {
    var btn = $("btn-run");
    busy(btn, true);
    try {
      var texts = await readParagraphs();
      state.texts = texts;
      var nonEmpty = texts.filter(function (t) { return t && t.trim(); }).length;
      if (!nonEmpty) {
        state.result = null;
        renderAll();
        toast("This document is empty. Open a contract, or insert a demo.", "err");
        return;
      }
      state.result = Engine.review(texts, reviewOpts());
      state.playbook = state.playbookChoice === "none" ? null : Engine.playbook(texts, pbOpts());
      if (keepSetupOpen !== true) state.setupOpen = false;
      updatePartyLabels(texts);
      state.scope = null;
      state.summary = null;
      clear($("summary"));
      $("summary-actions").hidden = true;
      state.issues = null;
      clear($("issues"));
      $("issues-actions").hidden = true;
      if (state.result.docType === "lease" || state.result.docType === "psa") state.library = state.result.docType;
      renderAll();
      var s = state.result.stats;
      toast("Review done: " + (s.high + s.medium + s.low) + " issues, " + s.missing + " missing", "ok");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally {
      busy(btn, false);
    }
  }

  async function refreshAfterEdit() {
    try {
      state.texts = await readParagraphs();
      state.result = Engine.review(state.texts, reviewOpts());
      state.playbook = state.playbookChoice === "none" ? null : Engine.playbook(state.texts, pbOpts());
      renderAll();
    } catch (e) { /* keep the current view */ }
  }

  async function refreshPlaybook() {
    try {
      state.texts = await readParagraphs();
      state.playbook = state.playbookChoice === "none" ? null : Engine.playbook(state.texts, pbOpts());
      renderPlaybook();
      updateCounts();
    } catch (e) { /* keep the current view */ }
  }

  function flags(key) { return state.status[key] || (state.status[key] = {}); }
  function st(f) { return state.status[keyOf(f)] || {}; }
  function isClosed(s) { return !!(s.redlined || s.resolved || s.dismissed); }

  function setStatus(f, flag) {
    var s = flags(keyOf(f));
    s[flag] = true;
    if (flag === "redlined" || flag === "resolved") s.error = false;
    renderFindings();
    renderPlaybook();
  }

  async function onApply(f, btn, fromPlaybook) {
    busy(btn, true);
    try {
      var outcome = await applyFinding(f);
      setStatus(f, outcome === "resolved" ? "resolved" : "redlined");
      toast(outcome === "resolved" ? "Already resolved in the document" : (state.api14 && state.tracking ? "Redline added as a tracked change" : "Redline added"), "ok");
      if (fromPlaybook) await refreshPlaybook();
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onComment(f, btn) {
    if (!state.api14) {
      copyText(f.comment);
      toast("Comments need Microsoft 365, Word 2024 or Word on the web. Comment text copied instead.", "err");
      return;
    }
    busy(btn, true);
    try {
      var outcome = await commentFinding(f);
      setStatus(f, outcome === "resolved" ? "resolved" : "commented");
      toast("Comment added", "ok");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  /* ---------------- Review selection ---------------- */

  function matchSpan(body, sel) {
    var want = sel.filter(function (t) { return t != null; });
    if (!want.length) return null;
    for (var i = 0; i + want.length <= body.length; i++) {
      var ok = true;
      for (var j = 0; j < want.length && ok; j++) if (body[i + j] !== want[j]) ok = false;
      if (ok) return { from: i, to: i + want.length - 1 };
    }
    return null;
  }

  function readSelection() {
    return Word.run(async function (context) {
      var range = context.document.getSelection();
      range.load("text");
      var sps = range.paragraphs;
      sps.load("items/text");
      var body = context.document.body.paragraphs;
      body.load("items/text");
      await context.sync();
      return {
        text: range.text || "",
        paragraphs: sps.items.map(function (p) { return p.text || ""; }),
        body: body.items.map(function (p) { return p.text || ""; })
      };
    });
  }

  async function onReviewSelection() {
    var btn = $("btn-review-selection");
    busy(btn, true);
    try {
      if (!state.result) await runReview();
      if (!state.result) return;
      var sel = await readSelection();
      var span = matchSpan(sel.body, sel.paragraphs);
      if (!span) { toast("Couldn't match your selection to the document. Run the review again, then try once more.", "err"); return; }
      state.scope = span;
      renderFindings();
      renderScope();
      renderChecks();
      var n = visibleFindings().length;
      toast(n ? n + " issue" + (n === 1 ? "" : "s") + " in your selection" : "Nothing flagged in your selection", "ok");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  function renderScope() {
    var on = !!state.scope;
    $("scope-text").hidden = !on;
    $("btn-scope-clear").hidden = !on;
    if (on) {
      var n = state.scope.to - state.scope.from + 1;
      $("scope-text").textContent = "Showing issues in your selection (" + n + " paragraph" + (n === 1 ? "" : "s") + ")";
    }
  }

  /* ---------------- Playbook ---------------- */

  function pbSide() {
    if (state.playbook) return { side: state.playbook.perspective, docType: state.playbook.docType };
    var docType = state.library === "psa" ? "psa" : "lease";
    var p = state.perspective;
    var ok = (docType === "lease" && (p === "tenant" || p === "landlord")) || (docType === "psa" && (p === "buyer" || p === "seller"));
    return { side: ok ? p : (docType === "psa" ? "buyer" : "tenant"), docType: docType };
  }

  function pbCard(it) {
    var s = st(it);
    var closed = isClosed(s) || !!s.inserted;
    var kind = it.status === "pass" ? "pass" : it.status === "fail" ? "fail" : "unknown";
    var actions = h("div", { class: "card-actions" });
    if (it.paragraphIndex >= 0) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { goToFinding(it); } }, "Go to"));
    if (kind === "fail" && it.action && !closed) actions.appendChild(h("button", { type: "button", class: "btn btn-secondary btn-sm", onclick: function (ev) { onApply(it, ev.currentTarget, true); } }, it.action.type === "replace" ? "Apply redline" : "Add language"));
    if (kind === "fail" && !it.action && it.templateId && !closed) actions.appendChild(h("button", { type: "button", class: "btn btn-secondary btn-sm", onclick: function (ev) { onInsertTemplate(it.templateId, "end", ev.currentTarget, null, it); } }, "Insert clause"));
    if (kind === "fail" && it.paragraphIndex >= 0 && !s.commented && !closed) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function (ev) { onComment(it, ev.currentTarget); } }, "Comment"));
    var label = statusLabels(s);
    if (label) actions.appendChild(h("span", { class: "status", text: label }));
    var suggest = null;
    if (kind === "fail" && it.action && it.action.type === "replace") suggest = h("div", { class: "suggest" }, h("strong", { text: "Redline to your standard" }), h("span", { class: "del", text: it.action.find }), " \u2192 ", h("span", { class: "ins", text: it.action.replace }));
    else if (kind === "fail" && it.action) suggest = h("div", { class: "suggest" }, h("strong", { text: "Suggested addition" }), h("span", { class: "ins", text: it.action.text }));
    return h("article", { class: "card pb-card pb-" + kind + (closed ? " is-done" : "") },
      h("div", { class: "card-head" },
        h("span", { class: "badge pb-badge " + kind, text: kind === "pass" ? "OK" : kind === "fail" ? "Off" : "N/A" }),
        h("div", null,
          h("h4", { text: it.label }),
          h("p", { class: "card-meta" },
            it.ref ? h("button", { type: "button", class: "ref", title: "Go to this text", onclick: function () { goToFinding(it); } }, it.ref) : null,
            (it.ref ? " \u00B7 " : "") + "Document: " + it.actualText + " \u00B7 Yours: " + it.standardText))),
      kind !== "pass" && it.why ? h("p", { class: "why", text: it.why }) : null,
      suggest,
      actions.childNodes.length ? actions : null);
  }

  function renderPlaybook() {
    var box = $("pb-items"), statsBox = $("pb-stats");
    clear(box); clear(statsBox);
    var side = pbSide();
    $("pb-side").textContent = state.playbook ? (cap(side.side) + " standards \u00B7 " + (side.docType === "psa" ? "Purchase agreement" : "Lease")) :
      state.result && state.playbookChoice === "none" ? "Playbook is off. Choose \u201CMy standards\u201D in the review setup." : "Run a review to check the document.";
    renderPbEditor();
    if (!state.playbook) { statsBox.hidden = true; return; }
    var s = state.playbook.stats;
    statsBox.hidden = false;
    [["fail", "Off", s.fail], ["pass", "OK", s.pass], ["unknown", "N/A", s.notFound]].forEach(function (row) {
      statsBox.appendChild(h("div", { class: "stat pb-" + row[0] }, h("b", { text: String(row[2]) }), h("span", { text: row[1] })));
    });
    if (!state.playbook.items.length) box.appendChild(h("p", { class: "hint", text: "No standards apply to this document type and side." }));
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
    box.appendChild(h("p", { class: "pb-editor-title", text: cap(side.side) + " standards (" + (side.docType === "psa" ? "purchase agreements" : "leases") + ")" }));
    Engine.STANDARDS.forEach(function (std) {
      if (std.side !== side.side || std.docType !== side.docType) return;
      var ov = state.standards[std.id] || {};
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
      h("button", { type: "button", class: "btn btn-secondary btn-sm", id: "btn-pb-save", onclick: savePbEditor }, "Save standards"),
      h("button", { type: "button", class: "btn btn-ghost btn-sm", id: "btn-pb-reset", onclick: resetPbEditor }, "Reset to defaults")));
  }

  async function savePbEditor() {
    var side = pbSide();
    Engine.STANDARDS.forEach(function (std) {
      if (std.side !== side.side || std.docType !== side.docType) return;
      var on = document.querySelector('[data-std-on="' + std.id + '"]');
      var inp = document.querySelector('[data-std="' + std.id + '"]');
      var ov = {};
      if (on && !on.checked) ov.off = true;
      if (inp && inp.value !== "" && !isNaN(Number(inp.value)) && Number(inp.value) !== std.value) ov.value = Number(inp.value);
      if (ov.off || ov.value != null) state.standards[std.id] = ov;
      else delete state.standards[std.id];
    });
    var stored = saveStore(STANDARDS_KEY, state.standards);
    state.pbEditing = false;
    if (state.result) await refreshPlaybook();
    else renderPlaybook();
    toast(stored ? "Standards saved" : "Standards applied, but this computer blocked saving them", stored ? "ok" : "err");
  }

  async function resetPbEditor() {
    var side = pbSide();
    Engine.STANDARDS.forEach(function (std) { if (std.side === side.side && std.docType === side.docType) delete state.standards[std.id]; });
    saveStore(STANDARDS_KEY, state.standards);
    state.pbEditing = false;
    if (state.result) await refreshPlaybook();
    else renderPlaybook();
    toast("Standards reset to defaults", "ok");
  }

  /* ---------------- Drafting and saved clauses ---------------- */

  function renderDrafts() {
    var ul = $("drafts");
    clear(ul);
    Object.keys(Engine.DRAFTS).forEach(function (id) {
      var d = Engine.DRAFTS[id];
      if (d.docType !== state.library) return;
      var defaults = Engine.draftDefaults(id, state.texts);
      var inputs = {};
      var preview = h("p", { class: "draft-preview" });
      var form = h("div", { class: "draft-form", hidden: true });
      function values() { var v = {}; for (var k in inputs) v[k] = inputs[k].value; return v; }
      function update() { preview.textContent = Engine.draftClause(id, values(), state.texts).text; }
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
        h("button", { type: "button", class: "btn btn-secondary btn-sm", onclick: function (ev) { onInsertDraft(id, values(), ev.currentTarget); } }, "Insert at cursor")));
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

  async function onInsertDraft(id, values, btn) {
    var d = Engine.draftClause(id, values, state.texts);
    if (!d) return;
    busy(btn, true);
    try {
      await insertParagraphs([d.text], "cursor", { tracked: true, markInserted: true });
      toast("\u201C" + d.title + "\u201D inserted" + (state.api14 && state.tracking ? " as a tracked change" : ""), "ok");
      if (state.result) await refreshAfterEdit();
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
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
            h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function (ev) { onInsertMyClause(c, ev.currentTarget); } }, "Insert"),
            h("button", { type: "button", class: "btn btn-ghost btn-sm", "aria-label": "Delete " + c.title, onclick: function () {
              state.myClauses.splice(i, 1);
              saveStore(CLAUSES_KEY, state.myClauses);
              renderMyClauses();
            } }, "Delete"))),
        h("details", { class: "template" }, h("summary", { text: "Preview" }), h("p", { text: c.text }))));
    });
  }

  async function onSaveClause() {
    var btn = $("btn-save-clause");
    busy(btn, true);
    try {
      var sel = await readSelection();
      var text = (sel.text || "").replace(/\r\n?/g, "\n").replace(/^\s+|\s+$/g, "");
      if (text.length < 15) { toast("Select the language you want to save first.", "err"); return; }
      var title = ($("my-clause-title").value || "").replace(/^\s+|\s+$/g, "") || clauseTitle(text);
      state.myClauses.unshift({ id: "c" + Date.now(), title: title.slice(0, 80), text: text.slice(0, 20000) });
      var stored = saveStore(CLAUSES_KEY, state.myClauses);
      $("my-clause-title").value = "";
      renderMyClauses();
      toast(stored ? "Saved to My clauses" : "Saved for this session; this computer blocked permanent storage", stored ? "ok" : "err");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onInsertMyClause(c, btn) {
    busy(btn, true);
    try {
      await insertParagraphs(c.text.split("\n"), "cursor", { tracked: true, markInserted: true });
      toast("\u201C" + c.title + "\u201D inserted" + (state.api14 && state.tracking ? " as a tracked change" : ""), "ok");
      if (state.result) await refreshAfterEdit();
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  /* ---------------- Explain + issues list ---------------- */

  async function onExplain() {
    var btn = $("btn-explain");
    busy(btn, true);
    try {
      var sel = await readSelection();
      var text = (sel.text || "").replace(/^\s+|\s+$/g, "") ? sel.text : sel.paragraphs.join("\n");
      if (!String(text).replace(/\s+/g, "")) { toast("Select some contract text first, or click into a paragraph.", "err"); return; }
      state.texts = await readParagraphs();
      var x = Engine.explain(text, state.texts, reviewOpts());
      var chat = $("chat");
      var quoted = String(text).replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
      chat.appendChild(h("div", { class: "msg q", text: "Explain: \u201C" + (quoted.length > 90 ? quoted.slice(0, 89) + "\u2026" : quoted) + "\u201D" }));
      var msg = h("div", { class: "msg a" + (x.found ? "" : " not-found") });
      msg.appendChild(document.createTextNode(x.answer));
      if (x.paragraphIndexes.length) {
        var src = h("div", { class: "sources" });
        x.paragraphIndexes.slice(0, 3).forEach(function (i) { src.appendChild(h("button", { type: "button", class: "source", onclick: function () { goToParagraph(i); } }, "Go to")); });
        msg.appendChild(src);
      }
      chat.appendChild(msg);
      msg.scrollIntoView({ block: "nearest" });
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onIssues() {
    var btn = $("btn-issues");
    busy(btn, true);
    try {
      state.texts = await readParagraphs();
      var o = reviewOpts();
      o.notes = state.notes;
      var il = Engine.issuesList(state.texts, o);
      state.issues = il;
      var ol = $("issues");
      clear(ol);
      il.items.forEach(function (it) {
        ol.appendChild(h("li", { class: "issue" },
          h("div", { class: "issue-head" },
            h("span", { class: "badge " + it.severity, text: it.severity === "medium" ? "Med" : cap(it.severity) }),
            h("strong", { text: it.title })),
          h("p", { class: "card-meta" },
            it.paragraphIndex >= 0 ? h("button", { type: "button", class: "ref", onclick: function () { goToParagraph(it.paragraphIndex); } }, it.ref) : it.ref,
            it.section ? " \u00B7 " + it.section : ""),
          h("p", { class: "issue-request", text: it.request })));
      });
      if (!il.items.length) ol.appendChild(h("li", { class: "empty-note", text: "Nothing to raise for this side." }));
      $("issues-actions").hidden = !il.items.length;
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onIssuesEnd(btn) {
    if (!state.issues) return;
    busy(btn, true);
    try {
      await insertParagraphs(state.issues.text.split("\n"), "end", { boldFirst: true });
      toast("Issues list inserted at the end", "ok");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onCommentAll() {
    if (!state.result) return;
    if (!state.api14) { toast("Comments need Microsoft 365, Word 2024 or Word on the web.", "err"); return; }
    var open = visibleFindings().filter(function (f) { var s = st(f); return !s.commented && !s.dismissed && !s.resolved; });
    if (!open.length) { toast("No open issues to comment on"); return; }
    var btn = $("btn-comment-all");
    busy(btn, true);
    try {
      var done = 0;
      await Word.run(async function (context) {
        var loaded = await loadTexts(context);
        for (var i = 0; i < open.length; i++) {
          var out = await commentFinding(open[i], context, loaded);
          flags(keyOf(open[i]))[out === "resolved" ? "resolved" : "commented"] = true;
          if (out === "commented") done++;
        }
        await context.sync();
      });
      renderFindings();
      toast(done + " comments added", "ok");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onRedlineAll() {
    if (!state.result) return;
    var open = visibleFindings().filter(function (f) { return f.action && !isClosed(st(f)); });
    if (!open.length) { toast("No open suggested edits"); return; }
    var btn = $("btn-redline-all");
    busy(btn, true);
    var done = 0, failed = 0;
    for (var i = 0; i < open.length; i++) {
      toast("Redlining " + (i + 1) + " of " + open.length + "\u2026");
      try {
        var out = await applyFinding(open[i]);
        flags(keyOf(open[i]))[out === "resolved" ? "resolved" : "redlined"] = true;
        if (out === "redlined") done++;
      } catch (e) {
        failed++;
        flags(keyOf(open[i])).error = true;
      }
    }
    busy(btn, false);
    renderFindings();
    toast(done + " redlines added" + (failed ? ", " + failed + " need manual edits" : ""), failed ? "err" : "ok");
  }

  async function onInsertTemplate(templateId, where, btn, missingItem, playbookItem) {
    var t = Engine.TEMPLATES[templateId];
    if (!t) return;
    busy(btn, true);
    try {
      await insertParagraphs([t.text], where, { tracked: true, markInserted: true });
      if (missingItem) flags("missing|" + missingItem.id).inserted = true;
      if (playbookItem) flags(keyOf(playbookItem)).inserted = true;
      toast("\u201C" + t.title + "\u201D inserted" + (state.api14 && state.tracking ? " as a tracked change" : ""), "ok");
      if (state.result) await refreshAfterEdit();
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onInsertDemo(id) {
    try {
      var texts = await readParagraphs();
      if (texts.some(function (t) { return t && t.trim(); })) {
        toast("Demo text goes into an empty document. Open a new blank document first.", "err");
        return;
      }
      var doc = Demo[id];
      await Word.run(async function (context) {
        var body = context.document.body;
        var ps = body.paragraphs;
        ps.load("items");
        await context.sync();
        for (var i = 0; i < doc.paragraphs.length; i++) {
          var text = doc.paragraphs[i];
          var m = /^(\d+\.\s+[^.]{2,60}\.)\s+([\s\S]+)$/.exec(text);
          var p;
          if (i === 0 && ps.items.length) { p = ps.items[0]; p.insertText(m ? m[1] : text, "Replace"); }
          else p = body.insertParagraph(m ? m[1] : text, "End");
          var heading = !m && /^[^a-z]{3,90}$/.test(text) && !/^By:/.test(text);
          p.font.bold = !!(m || heading);
          if (i === 0) p.alignment = "Centered";
          else p.alignment = "Justified";
          if (m) {
            var rest = p.insertText(" " + m[2], "End");
            rest.font.bold = false;
          }
        }
        await context.sync();
      });
      toast("Demo inserted. Running review\u2026", "ok");
      await runReview();
    } catch (e) {
      toast(errorMessage(e), "err");
    }
  }

  /* ---------------------------------------------------------------- *
   * Rendering
   * ---------------------------------------------------------------- */

  function visibleFindings() {
    if (!state.result) return [];
    var sc = state.scope;
    return state.result.findings.filter(function (f) {
      if (!state.filter[f.severity]) return false;
      return !sc || (f.paragraphIndex >= sc.from && f.paragraphIndex <= sc.to);
    });
  }

  function updateCounts() {
    var r = state.result;
    $("count-review").textContent = r ? String(r.stats.high + r.stats.medium + r.stats.low + r.stats.missing) : "";
    $("count-playbook").textContent = state.playbook && state.playbook.stats.fail ? String(state.playbook.stats.fail) : "";
  }

  function renderAll() {
    var r = state.result;
    var chip = $("doc-chip");
    if (r) {
      var type = r.docType === "lease" ? "Lease" : r.docType === "psa" ? "Purchase agreement" : "Document";
      chip.textContent = type + " \u00B7 " + r.perspectiveLabel;
      chip.classList.add("is-set");
    } else {
      chip.textContent = "No review yet";
      chip.classList.remove("is-set");
    }
    $("review-empty").hidden = !!r;
    $("review-body").hidden = !r;
    updateCounts();
    renderSetup();
    renderStats();
    renderFindings();
    renderScope();
    renderChecks();
    renderPlaybook();
    renderClauses();
    renderLibrary();
    renderDrafts();
    renderMyClauses();
    renderChips();
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

  function renderStats() {
    var box = $("stats");
    clear(box);
    if (!state.result) return;
    var s = state.result.stats;
    [["high", "High", s.high], ["medium", "Medium", s.medium], ["low", "Low", s.low], ["missing", "Missing", s.missing]].forEach(function (row) {
      var on = state.filter[row[0]];
      box.appendChild(h("button", {
        type: "button", class: "stat " + row[0] + (on ? "" : " is-off"), "aria-pressed": on ? "true" : "false",
        title: (on ? "Hide " : "Show ") + row[1].toLowerCase() + " items",
        onclick: function () { state.filter[row[0]] = !state.filter[row[0]]; renderStats(); renderFindings(); }
      }, h("b", { text: String(row[2]) }), h("span", { text: row[1] })));
    });
  }

  function quoteWithMark(excerpt, match) {
    var q = h("blockquote", { class: "quote" });
    var i = match ? excerpt.indexOf(match) : -1;
    if (i === -1) { q.textContent = "\u201C" + excerpt + "\u201D"; return q; }
    q.appendChild(document.createTextNode("\u201C" + excerpt.slice(0, i)));
    q.appendChild(h("mark", { text: match }));
    q.appendChild(document.createTextNode(excerpt.slice(i + match.length) + "\u201D"));
    return q;
  }

  function statusLabels(s) {
    var out = [];
    if (s.redlined) out.push("\u2713 Redlined");
    if (s.inserted) out.push("\u2713 Inserted");
    if (s.resolved) out.push("\u2713 Already resolved");
    if (s.commented) out.push("\u2713 Commented");
    if (s.dismissed) out.push("Dismissed");
    if (s.error && !s.redlined) out.push("Needs manual edit");
    return out.join(" \u00B7 ");
  }

  function findingCard(f) {
    var s = st(f);
    var closed = isClosed(s);
    var riskFor = state.result.perspective === "neutral" ? h("span", { class: "risk-for", text: "Risk for " + f.riskFor.map(cap).join(" / ") }) : null;
    var actions = h("div", { class: "card-actions" });
    actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { goToFinding(f); } }, "Go to"));
    if (f.action && !closed) actions.appendChild(h("button", { type: "button", class: "btn btn-secondary btn-sm", onclick: function (ev) { onApply(f, ev.currentTarget); } }, f.action.type === "replace" ? "Apply redline" : "Add language"));
    if (!s.commented && !s.dismissed && !s.resolved) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function (ev) { onComment(f, ev.currentTarget); } }, "Comment"));
    if (!closed) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { setStatus(f, "dismissed"); } }, "Dismiss"));
    var label = statusLabels(s);
    if (label) actions.appendChild(h("span", { class: "status" + ((s.error && !s.redlined) || s.dismissed ? " is-warn" : ""), text: label }));

    var suggest = null;
    if (f.action && f.action.type === "replace") {
      suggest = h("div", { class: "suggest" }, h("strong", { text: "Suggested redline" }), h("span", { class: "del", text: f.action.find }), " \u2192 ", h("span", { class: "ins", text: f.action.replace }));
    } else if (f.action) {
      suggest = h("div", { class: "suggest" }, h("strong", { text: "Suggested addition" }), h("span", { class: "ins", text: f.action.text }));
    }

    return h("article", { class: "card sev-" + f.severity + (closed ? " is-done" : "") },
      h("div", { class: "card-head" },
        h("span", { class: "badge " + f.severity, text: f.severity === "medium" ? "Med" : cap(f.severity) }),
        h("div", null,
          h("h4", { text: f.title }),
          h("p", { class: "card-meta" },
            h("button", { type: "button", class: "ref", title: "Go to this text", onclick: function () { goToFinding(f); } }, f.ref),
            f.section ? " \u00B7 " + f.section : (" \u00B7 " + f.categoryLabel),
            riskFor))),
      h("p", { class: "why", text: f.why }),
      quoteWithMark(f.excerpt, f.matchText),
      suggest,
      actions);
  }

  function missingCard(m) {
    var ms = state.status["missing|" + m.id] || {};
    var mdone = !!(ms.inserted || ms.dismissed);
    var riskFor = state.result.perspective === "neutral" ? h("span", { class: "risk-for", text: "Risk for " + m.riskFor.map(cap).join(" / ") }) : null;
    var actions = h("div", { class: "card-actions" });
    if (m.templateId && !mdone) {
      actions.appendChild(h("button", { type: "button", class: "btn btn-secondary btn-sm", onclick: function (ev) { onInsertTemplate(m.templateId, "cursor", ev.currentTarget, m); } }, "Insert at cursor"));
      actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function (ev) { onInsertTemplate(m.templateId, "end", ev.currentTarget, m); } }, "Insert at end"));
    }
    if (!mdone) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { flags("missing|" + m.id).dismissed = true; renderFindings(); } }, "Dismiss"));
    if (mdone) actions.appendChild(h("span", { class: "status" + (ms.dismissed ? " is-warn" : ""), text: statusLabels(ms) }));
    return h("article", { class: "card sev-" + m.severity + (mdone ? " is-done" : "") },
      h("div", { class: "card-head" },
        h("span", { class: "badge " + m.severity, text: m.severity === "medium" ? "Med" : cap(m.severity) }),
        h("div", null, h("h4", { text: m.title }), h("p", { class: "card-meta" }, m.label, riskFor))),
      h("p", { class: "why", text: m.why }),
      m.templateText ? h("details", { class: "template" }, h("summary", { text: "Preview: " + m.templateTitle }), h("p", { text: m.templateText })) : null,
      actions);
  }

  function renderFindings() {
    var box = $("findings"), missBox = $("missing");
    clear(box); clear(missBox);
    if (!state.result) return;
    var list = visibleFindings();
    if (!list.length) {
      box.appendChild(h("p", { class: "hint", text: state.result.findings.length ? "All issues are hidden by the filters above." : "No issues flagged for the " + state.result.perspectiveLabel.toLowerCase() + " side. Clause Craft checks common CRE clauses only." }));
    }
    list.forEach(function (f) { box.appendChild(findingCard(f)); });
    var missing = state.filter.missing && !state.scope ? state.result.missing : [];
    $("missing-title").hidden = !missing.length;
    missing.forEach(function (m) { missBox.appendChild(missingCard(m)); });
  }

  function renderClauses() {
    var list = $("clause-list");
    clear(list);
    var r = state.result;
    $("clauses-hint").textContent = r ? (r.clauses.length + " clause types found. Click one to jump to it.") : "Run a review to list the clauses Clause Craft recognizes. Click one to jump to it.";
    if (!r) return;
    r.clauses.forEach(function (c) {
      var worst = null;
      r.findings.forEach(function (f) {
        if (c.paragraphs.indexOf(f.paragraphIndex) === -1) return;
        if (!worst || (f.severity === "high") || (f.severity === "medium" && worst === "low")) worst = f.severity;
      });
      list.appendChild(h("li", null, h("button", { type: "button", class: "clause-item", onclick: function () { goToParagraph(c.first); } },
        h("span", { class: "ref", text: c.ref }),
        h("span", { class: "label", text: c.label }),
        worst ? h("span", { class: "flag " + worst, text: worst === "medium" ? "Med" : cap(worst) }) : null)));
    });
  }

  function renderLibrary() {
    var ul = $("library");
    clear(ul);
    Array.prototype.forEach.call(document.querySelectorAll("[data-lib]"), function (b) {
      b.classList.toggle("is-active", b.getAttribute("data-lib") === state.library);
    });
    LIBRARY[state.library].forEach(function (id) {
      var t = Engine.TEMPLATES[id];
      if (!t) return;
      ul.appendChild(h("li", null,
        h("div", { class: "lib-head" },
          h("span", { class: "lib-title", text: t.title }),
          h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function (ev) { onInsertTemplate(id, "cursor", ev.currentTarget); } }, "Insert")),
        h("details", { class: "template" }, h("summary", { text: "Preview" }), h("p", { text: t.text }))));
    });
  }

  function renderChips() {
    var box = $("ask-chips");
    clear(box);
    var type = state.result && state.result.docType === "psa" ? "psa" : "lease";
    CHIPS[type].forEach(function (q) {
      box.appendChild(h("button", { type: "button", class: "chip", onclick: function () { onAsk(q); } }, q));
    });
  }

  /* ---------------------------------------------------------------- *
   * Ask + Summary
   * ---------------------------------------------------------------- */

  async function ensureTexts() {
    if (!state.texts) state.texts = await readParagraphs();
    return state.texts;
  }

  async function onAsk(question) {
    var q = String(question || "").trim();
    if (!q) return;
    var chat = $("chat");
    chat.appendChild(h("div", { class: "msg q", text: q }));
    $("ask-input").value = "";
    try {
      state.texts = await readParagraphs();
      var a = Engine.ask(q, state.texts, { perspective: state.perspective, jurisdiction: state.jurisdiction });
      var msg = h("div", { class: "msg a" + (a.found ? "" : " not-found") });
      msg.appendChild(document.createTextNode(a.answer));
      if (a.sources && a.sources.length) {
        var src = h("div", { class: "sources" });
        a.sources.forEach(function (s) {
          src.appendChild(h("button", { type: "button", class: "source", title: s.excerpt, onclick: function () { goToParagraph(s.paragraphIndex); } }, s.ref || ("\u00B6" + (s.paragraphIndex + 1))));
        });
        msg.appendChild(src);
      }
      chat.appendChild(msg);
    } catch (e) {
      chat.appendChild(h("div", { class: "msg a not-found", text: errorMessage(e) }));
    }
    chat.lastChild.scrollIntoView({ block: "nearest" });
  }

  async function onSummary() {
    var btn = $("btn-summary");
    busy(btn, true);
    try {
      state.texts = await readParagraphs();
      var s = Engine.summarize(state.texts, { perspective: state.perspective, jurisdiction: state.jurisdiction });
      state.summary = s;
      var box = $("summary");
      clear(box);
      var table = h("table", { class: "kv" });
      s.fields.forEach(function (f) {
        table.appendChild(h("tr", null, h("th", { text: f.label }), h("td", null, f.value,
          h("button", { type: "button", class: "ref", title: "Go to source", onclick: function () { goToParagraph(f.paragraphIndex); } }, "\u2192"))));
      });
      box.appendChild(h("p", { class: "card-meta", text: s.typeLabel + " \u00B7 reviewed for the " + s.perspectiveLabel.toLowerCase() + " side" }));
      if (s.fields.length) box.appendChild(table);
      if (s.risks.length) {
        box.appendChild(h("h3", { class: "section-title", text: "Key risks" }));
        var ul = h("ul", { class: "risk-list" });
        s.risks.forEach(function (r) {
          ul.appendChild(h("li", null, h("span", { class: "badge " + r.severity, text: r.severity === "medium" ? "Med" : cap(r.severity) }),
            h("span", null, r.title + " ", h("button", { type: "button", class: "ref", onclick: function () { goToParagraph(r.paragraphIndex); } }, r.ref))));
        });
        box.appendChild(ul);
      }
      if (s.missing.length) box.appendChild(h("p", { class: "missing-line", text: "Missing: " + s.missing.join(", ") }));
      $("summary-actions").hidden = false;
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  async function onInsertSummary(where, btn) {
    if (!state.summary) return;
    busy(btn, true);
    try {
      await insertParagraphs(state.summary.text.split("\n"), where, { boldFirst: true });
      toast("Summary inserted", "ok");
    } catch (e) {
      toast(errorMessage(e), "err");
    } finally { busy(btn, false); }
  }

  function copyText(text) {
    var done = false;
    try {
      var ta = h("textarea", { style: "position:fixed;left:-9999px;top:0" });
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      done = document.execCommand("copy");
      document.body.removeChild(ta);
    } catch (e) { done = false; }
    if (!done && navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(function () {});
      done = true;
    }
    return done;
  }

  /* ---------------------------------------------------------------- *
   * Setup
   * ---------------------------------------------------------------- */

  function switchTab(name) {
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
      var on = t.getAttribute("data-tab") === name;
      t.classList.toggle("is-active", on);
      t.setAttribute("aria-selected", on ? "true" : "false");
    });
    Array.prototype.forEach.call(document.querySelectorAll(".panel"), function (p) {
      var on = p.id === "panel-" + name;
      p.hidden = !on;
      p.classList.toggle("is-active", on);
    });
  }

  var SETUP_KEY = "clauseCraft.setup";

  /** Review setup is saved with the document (Word document settings), so it travels with the file. */
  function loadSettings() {
    try {
      var s = Office.context.document.settings;
      var v = s && s.get(SETTINGS_KEY);
      if (v) state.perspective = v;
      var setup = s && s.get(SETUP_KEY);
      if (setup && typeof setup === "object") {
        if (setup.reviewType) state.reviewType = setup.reviewType;
        if (setup.jurisdiction) state.jurisdiction = setup.jurisdiction;
        if (setup.playbookChoice) state.playbookChoice = setup.playbookChoice;
        if (typeof setup.notes === "string") state.notes = setup.notes;
      }
    } catch (e) { /* settings unavailable */ }
    $("perspective").value = state.perspective;
    $("review-type").value = state.reviewType;
    $("jurisdiction").value = state.jurisdiction;
    $("playbook-choice").value = state.playbookChoice;
    $("deal-notes").value = state.notes;
  }

  function saveSettings() {
    try {
      var s = Office.context.document.settings;
      if (!s) return;
      s.set(SETTINGS_KEY, state.perspective);
      s.set(SETUP_KEY, { reviewType: state.reviewType, jurisdiction: state.jurisdiction, playbookChoice: state.playbookChoice, notes: state.notes });
      s.saveAsync(function () {});
    } catch (e) { /* ignore */ }
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
      $("setup-summary-text").textContent = [party, optionText("review-type", state.reviewType), where, state.playbookChoice === "none" ? "No playbook" : "My standards"].join(" \u00B7 ");
    }
  }

  /** Shows the party names found in the document next to each side ("Tenant: HARBORLINE ADVISORY GROUP, INC."). */
  function updatePartyLabels(texts) {
    try { state.parties = Engine.parties(texts || []); } catch (e) { state.parties = []; }
    var labels = { tenant: "Tenant", landlord: "Landlord", buyer: "Buyer", seller: "Seller" };
    var sel = $("perspective");
    for (var i = 0; i < sel.options.length; i++) {
      var o = sel.options[i], base = labels[o.value];
      if (!base) continue;
      var hit = null;
      for (var k = 0; k < state.parties.length; k++) if (state.parties[k].role === base) hit = state.parties[k];
      o.text = hit ? base + ": " + hit.name : base;
    }
  }

  function onSetupChange(key, value) {
    state[key] = value;
    saveSettings();
    if (state.result) runReview(true);
  }

  function bind() {
    $("review-setup").addEventListener("submit", function (ev) { ev.preventDefault(); runReview(); });
    $("btn-setup-edit").addEventListener("click", function () { state.setupOpen = true; renderSetup(); $("review-type").focus(); });
    $("perspective").addEventListener("change", function (ev) { onSetupChange("perspective", ev.target.value); });
    $("review-type").addEventListener("change", function (ev) { onSetupChange("reviewType", ev.target.value); });
    $("jurisdiction").addEventListener("change", function (ev) { onSetupChange("jurisdiction", ev.target.value); });
    $("playbook-choice").addEventListener("change", function (ev) { onSetupChange("playbookChoice", ev.target.value); });
    var notesTimer = null;
    $("deal-notes").addEventListener("input", function (ev) {
      state.notes = ev.target.value;
      clearTimeout(notesTimer);
      notesTimer = setTimeout(saveSettings, 500);
    });
    $("qs-review").addEventListener("click", function () { runReview(); });
    $("qs-draft").addEventListener("click", function () { switchTab("draft"); });
    $("qs-playbook").addEventListener("click", function () { switchTab("playbook"); if (!state.result) runReview(); });
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
      t.addEventListener("click", function () { switchTab(t.getAttribute("data-tab")); });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-demo]"), function (b) {
      b.addEventListener("click", function () { onInsertDemo(b.getAttribute("data-demo")); });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-lib]"), function (b) {
      b.addEventListener("click", function () { state.library = b.getAttribute("data-lib"); renderLibrary(); renderDrafts(); });
    });
    $("btn-review-selection").addEventListener("click", onReviewSelection);
    $("btn-scope-clear").addEventListener("click", function () { state.scope = null; renderFindings(); renderScope(); renderChecks(); });
    $("btn-pb-edit").addEventListener("click", function () { state.pbEditing = !state.pbEditing; renderPbEditor(); });
    $("btn-save-clause").addEventListener("click", onSaveClause);
    $("btn-explain").addEventListener("click", onExplain);
    $("btn-issues").addEventListener("click", onIssues);
    $("btn-issues-end").addEventListener("click", function (ev) { onIssuesEnd(ev.currentTarget); });
    $("btn-issues-copy").addEventListener("click", function () {
      if (state.issues && copyText(state.issues.text)) toast("Issues list copied", "ok");
    });
    $("tracking").addEventListener("change", function (ev) { state.tracking = ev.target.checked; });
    $("btn-comment-all").addEventListener("click", onCommentAll);
    $("btn-redline-all").addEventListener("click", onRedlineAll);
    $("ask-form").addEventListener("submit", function (ev) { ev.preventDefault(); onAsk($("ask-input").value); });
    $("btn-summary").addEventListener("click", onSummary);
    $("btn-summary-insert").addEventListener("click", function (ev) { onInsertSummary("cursor", ev.currentTarget); });
    $("btn-summary-end").addEventListener("click", function (ev) { onInsertSummary("end", ev.currentTarget); });
    $("btn-summary-copy").addEventListener("click", function () {
      if (state.summary && copyText(state.summary.text)) toast("Summary copied", "ok");
    });
  }

  function applyCapabilities() {
    var note = $("api-note");
    var wrap = $("tracking-wrap");
    var box = $("tracking");
    if (!state.api14) {
      state.tracking = false;
      box.checked = false;
      box.disabled = true;
      wrap.classList.add("is-disabled");
      note.hidden = false;
      note.textContent = "This version of Word can't create tracked changes or comments from an add-in. Redlines are shown as red strikethrough with blue double-underlined insertions instead.";
      $("btn-comment-all").disabled = true;
    }
  }

  Office.onReady(function (info) {
    $("version").textContent = "v" + Engine.VERSION;
    bind();
    renderAll();
    if (!info || info.host !== Office.HostType.Word || typeof Word === "undefined") {
      state.inWord = false;
      $("btn-run").disabled = true;
      Array.prototype.forEach.call(document.querySelectorAll("[data-demo]"), function (b) { b.disabled = true; });
      toast("Open Clause Craft from Word (Home \u203A Review Contract).", "err");
      return;
    }
    state.inWord = true;
    var req = Office.context.requirements;
    state.api13 = req.isSetSupported("WordApi", "1.3");
    state.api14 = req.isSetSupported("WordApi", "1.4");
    loadSettings();
    applyCapabilities();
    renderSetup();
    readParagraphs().then(function (texts) { state.texts = texts; updatePartyLabels(texts); renderDrafts(); }).catch(function () {});
  });
})();
