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
    summary: null
  };

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

  async function runReview() {
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
      state.result = Engine.review(texts, { perspective: state.perspective });
      state.summary = null;
      clear($("summary"));
      $("summary-actions").hidden = true;
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
      state.result = Engine.review(state.texts, { perspective: state.perspective });
      renderAll();
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
  }

  async function onApply(f, btn) {
    busy(btn, true);
    try {
      var outcome = await applyFinding(f);
      setStatus(f, outcome === "resolved" ? "resolved" : "redlined");
      toast(outcome === "resolved" ? "Already resolved in the document" : (state.api14 && state.tracking ? "Redline added as a tracked change" : "Redline added"), "ok");
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

  async function onInsertTemplate(templateId, where, btn, missingItem) {
    var t = Engine.TEMPLATES[templateId];
    if (!t) return;
    busy(btn, true);
    try {
      await insertParagraphs([t.text], where, { tracked: true, markInserted: true });
      if (missingItem) flags("missing|" + missingItem.id).inserted = true;
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
    return state.result.findings.filter(function (f) { return state.filter[f.severity]; });
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
    $("count-review").textContent = r ? String(r.stats.high + r.stats.medium + r.stats.low + r.stats.missing) : "";
    $("count-clauses").textContent = r ? String(r.clauses.length) : "";
    renderStats();
    renderFindings();
    renderClauses();
    renderLibrary();
    renderChips();
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
    var missing = state.filter.missing ? state.result.missing : [];
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
      var a = Engine.ask(q, state.texts, { perspective: state.perspective });
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
      var s = Engine.summarize(state.texts, { perspective: state.perspective });
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

  function loadSettings() {
    try {
      var s = Office.context.document.settings;
      var v = s && s.get(SETTINGS_KEY);
      if (v) state.perspective = v;
    } catch (e) { /* settings unavailable */ }
    $("perspective").value = state.perspective;
  }

  function saveSettings() {
    try {
      var s = Office.context.document.settings;
      if (!s) return;
      s.set(SETTINGS_KEY, state.perspective);
      s.saveAsync(function () {});
    } catch (e) { /* ignore */ }
  }

  function bind() {
    $("btn-run").addEventListener("click", runReview);
    $("perspective").addEventListener("change", function (ev) {
      state.perspective = ev.target.value;
      saveSettings();
      if (state.result) runReview();
    });
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
      t.addEventListener("click", function () { switchTab(t.getAttribute("data-tab")); });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-demo]"), function (b) {
      b.addEventListener("click", function () { onInsertDemo(b.getAttribute("data-demo")); });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-lib]"), function (b) {
      b.addEventListener("click", function () { state.library = b.getAttribute("data-lib"); renderLibrary(); });
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
  });
})();
