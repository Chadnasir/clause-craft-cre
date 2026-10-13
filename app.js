/*
 * Clause Craft - browser app (index.html).
 * A Word-like editor with a review sidebar, powered by the same engine as the Word add-in
 * (engine/clause-engine.js). Everything runs in the page: no uploads, no API keys.
 * Document text is only ever inserted with textContent / DOM nodes, never innerHTML.
 */
(function () {
  "use strict";

  var Engine = window.ClauseCraftEngine;
  var Demo = window.ClauseCraftDemoDocs;
  var STORE_KEY = "clauseCraft.doc.v1";
  var editor = document.getElementById("editor");

  // snapshot/count: paragraph texts and count at review time, so a redline never lands on the wrong
  // paragraph if the document is edited between the review and the click.
  var state = { perspective: "auto", result: null, status: {}, snapshot: null, count: 0, filter: { high: true, medium: true, low: true, missing: true } };

  /* ---------------- helpers ---------------- */

  function $(id) { return document.getElementById(id); }
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (attrs[k] == null || attrs[k] === false) continue;
      if (k === "text") el.textContent = attrs[k];
      else if (k === "class") el.className = attrs[k];
      else if (k.indexOf("on") === 0 && typeof attrs[k] === "function") el.addEventListener(k.slice(2), attrs[k]);
      else el.setAttribute(k, attrs[k]);
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
  function keyOf(f) { return f.ruleId + "|" + f.matchText + "|" + (f.excerpt || "").slice(0, 60); }

  function toast(msg, kind) {
    var t = $("cc-toast");
    if (!t) { t = h("div", { id: "cc-toast", class: "cc-toast", role: "status", "aria-live": "polite" }); document.body.appendChild(t); }
    t.textContent = msg;
    t.className = "cc-toast is-visible" + (kind === "err" ? " is-err" : kind === "ok" ? " is-ok" : "");
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.className = "cc-toast"; }, 2600);
  }

  /* ---------------- document model ---------------- */

  function paragraphs() {
    var ps = Array.prototype.slice.call(editor.children).filter(function (n) { return /^(P|H1|H2|H3|DIV|LI)$/.test(n.tagName); });
    if (!ps.length && editor.textContent.trim()) {
      // Normalise loose text (e.g. pasted) into paragraphs.
      var text = editor.innerText;
      setDocument(text.split(/\r?\n/));
      return paragraphs();
    }
    return ps;
  }

  /** Paragraph text as it currently reads: deleted (struck) redline text is excluded. */
  function currentText(p) {
    var out = "";
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) out += c.nodeValue;
        else if (c.nodeType === 1 && c.tagName !== "DEL" && !c.classList.contains("cc-note")) walk(c);
      }
    })(p);
    return out;
  }

  function texts() { return paragraphs().map(currentText); }

  function setDocument(lines) {
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
    state.status = {};
    state.snapshot = null;
    state.count = 0;
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
        } else if (c.nodeType === 1 && c.tagName !== "DEL" && !c.classList.contains("cc-note")) walk(c);
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
  function locateFinding(f) {
    var ps = paragraphs();
    var orig = state.snapshot ? state.snapshot[f.paragraphIndex] : null;
    var changed = ps.length !== state.count;
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

  function goTo(index, f) {
    var ps = paragraphs();
    var hit = f && f.matchText ? locateFinding(f) : null;
    if (hit && hit.none && hit.changed) { toast("The document changed since the review. Run the review again.", "err"); return; }
    var p = hit && !hit.none ? hit.p : ps[index];
    if (!p) { toast("That paragraph no longer exists. Run the review again.", "err"); return; }
    p.scrollIntoView({ behavior: "smooth", block: "center" });
    flash(p);
    if (hit && !hit.none) select(p, hit.loc.start, hit.loc.start + hit.loc.target.length);
  }

  /** Apply a finding as an in-document redline: <del>old</del><ins>new</ins>, or an appended <ins>. */
  function accept(f) {
    if (!f.action) return "missing";
    var hit = locateFinding(f);
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
    state.snapshot[f.paragraphIndex] = currentText(p);
    return "redlined";
  }

  /* ---------------- review ---------------- */

  function runReview() {
    var t = texts();
    if (!t.some(function (x) { return x.trim(); })) { toast("Load the demo lease or upload a .docx, .txt or .md contract first.", "err"); return; }
    state.result = Engine.review(t, { perspective: state.perspective });
    state.snapshot = t.slice();
    state.count = t.length;
    renderAll();
    var s = state.result.stats;
    toast("Review done: " + (s.high + s.medium + s.low) + " issues, " + s.missing + " missing", "ok");
  }

  function renderAll() {
    renderStats();
    renderResults();
    renderClauses();
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
        onclick: function () { state.filter[x[0]] = !state.filter[x[0]]; renderStats(); renderResults(); } }, h("b", { text: String(x[2]) }), h("span", { text: x[1] })));
    });
    box.appendChild(row);
  }

  function quote(f) {
    var q = h("blockquote", { class: "cc-quote" });
    var i = f.excerpt.indexOf(f.matchText);
    if (i === -1) { q.textContent = "\u201C" + f.excerpt + "\u201D"; return q; }
    q.appendChild(document.createTextNode("\u201C" + f.excerpt.slice(0, i)));
    q.appendChild(h("mark", { text: f.matchText }));
    q.appendChild(document.createTextNode(f.excerpt.slice(i + f.matchText.length) + "\u201D"));
    return q;
  }

  function findingCard(f) {
    var st = state.status[keyOf(f)];
    var actions = h("div", { class: "cc-actions" },
      h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { goTo(f.paragraphIndex, f); } }, "Go to"));
    if (!st) {
      if (f.action) actions.appendChild(h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: function () {
        var out = accept(f);
        if (out === "missing") { toast("Couldn't find that text. Run the review again.", "err"); return; }
        if (out === "moved") { toast("The document changed since the review. Run the review again.", "err"); return; }
        state.status[keyOf(f)] = out;
        save(); renderResults();
        toast(out === "resolved" ? "Already resolved" : "Redline applied", "ok");
      } }, "Accept"));
      actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { state.status[keyOf(f)] = "rejected"; renderResults(); } }, "Reject"));
    } else {
      actions.appendChild(h("span", { class: "cc-status" + (st === "rejected" ? " is-warn" : ""), text: st === "rejected" ? "Rejected" : st === "resolved" ? "\u2713 Already resolved" : "\u2713 Redlined" }));
    }
    var suggest = null;
    if (f.action && f.action.type === "replace") suggest = h("div", { class: "cc-suggest" }, h("strong", { text: "Suggested redline" }), h("del", { text: f.action.find }), " \u2192 ", h("ins", { text: f.action.replace }));
    else if (f.action) suggest = h("div", { class: "cc-suggest" }, h("strong", { text: "Suggested addition" }), h("ins", { text: f.action.text }));
    var riskFor = state.result.perspective === "neutral" ? h("span", { class: "cc-riskfor", text: "Risk for " + f.riskFor.map(cap).join(" / ") }) : null;
    return h("article", { class: "cc-card sev-" + f.severity + (st ? " is-done" : ""), "data-kind": "finding", "data-rule": f.ruleId },
      h("div", { class: "cc-card-head" },
        h("span", { class: "cc-badge " + f.severity, text: f.severity === "medium" ? "Med" : cap(f.severity) }),
        h("div", null, h("h3", { text: f.title }), h("p", { class: "cc-meta" },
          h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(f.paragraphIndex, f); } }, f.ref),
          " \u00B7 " + (f.section || f.categoryLabel), riskFor))),
      h("p", { class: "cc-why", text: f.why }),
      quote(f), suggest, actions);
  }

  function missingCard(m) {
    var st = state.status["m|" + m.id];
    var actions = h("div", { class: "cc-actions" });
    if (!st && m.templateText) actions.appendChild(h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: function () {
      var p = h("p");
      p.appendChild(h("ins", { class: "cc-ins", text: m.templateText }));
      if (paragraphs().length === state.count) state.count++; // appended at the end: earlier paragraphs keep their places
      editor.appendChild(p);
      state.status["m|" + m.id] = "inserted";
      save();
      p.scrollIntoView({ behavior: "smooth", block: "center" }); flash(p);
      renderResults();
      toast("Clause added at the end", "ok");
    } }, "Insert clause"));
    if (!st) actions.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () { state.status["m|" + m.id] = "rejected"; renderResults(); } }, "Reject"));
    else actions.appendChild(h("span", { class: "cc-status" + (st === "rejected" ? " is-warn" : ""), text: st === "rejected" ? "Rejected" : "\u2713 Inserted" }));
    return h("article", { class: "cc-card sev-" + m.severity + (st ? " is-done" : ""), "data-kind": "missing", "data-category": m.category },
      h("div", { class: "cc-card-head" }, h("span", { class: "cc-badge " + m.severity, text: m.severity === "medium" ? "Med" : cap(m.severity) }),
        h("div", null, h("h3", { text: m.title }))),
      h("p", { class: "cc-why", text: m.why }),
      m.templateText ? h("details", { class: "cc-template" }, h("summary", { text: "Preview: " + m.templateTitle }), h("p", { text: m.templateText })) : null,
      actions);
  }

  function renderResults() {
    var box = $("review-results");
    clear(box);
    if (!state.result) return;
    var list = state.result.findings.filter(function (f) { return state.filter[f.severity]; });
    if (!list.length) box.appendChild(h("p", { class: "cc-empty", text: state.result.findings.length ? "All issues are hidden by the filters." : "No issues flagged for this side." }));
    list.forEach(function (f) { box.appendChild(findingCard(f)); });
    if (state.filter.missing && state.result.missing.length) {
      box.appendChild(h("h3", { class: "cc-section", text: "Missing clauses" }));
      state.result.missing.forEach(function (m) { box.appendChild(missingCard(m)); });
    }
  }

  function renderClauses() {
    var box = $("clause-list");
    clear(box);
    if (!state.result) { box.appendChild(h("p", { class: "cc-empty", text: "Run a review to list detected clauses." })); return; }
    state.result.clauses.forEach(function (c) {
      var flagged = state.result.findings.filter(function (f) { return c.paragraphs.indexOf(f.paragraphIndex) !== -1; });
      var worst = flagged.some(function (f) { return f.severity === "high"; }) ? "high" : flagged.some(function (f) { return f.severity === "medium"; }) ? "medium" : flagged.length ? "low" : null;
      box.appendChild(h("button", { type: "button", class: "cc-clause", onclick: function () { goTo(c.first); } },
        h("span", { class: "cc-clause-ref", text: c.ref }),
        h("span", { class: "cc-clause-body" }, h("strong", { text: c.label }), h("span", { text: c.excerpt })),
        worst ? h("span", { class: "cc-badge " + worst, text: flagged.length + " issue" + (flagged.length === 1 ? "" : "s") }) : null));
    });
  }

  /* ---------------- ask + summary ---------------- */

  function ask(q) {
    q = String(q || "").trim();
    if (!q) return;
    var chat = $("ask-chat");
    chat.appendChild(h("div", { class: "cc-msg q", text: q }));
    var a = Engine.ask(q, texts(), { perspective: state.perspective });
    var msg = h("div", { class: "cc-msg a" + (a.found ? "" : " not-found"), text: a.answer });
    if (a.sources && a.sources.length) {
      var src = h("div", { class: "cc-sources" });
      a.sources.forEach(function (s) { src.appendChild(h("button", { type: "button", class: "cc-source", title: s.excerpt, onclick: function () { goTo(s.paragraphIndex); } }, s.ref || "\u00B6" + (s.paragraphIndex + 1))); });
      msg.appendChild(src);
    }
    chat.appendChild(msg);
    if (msg.scrollIntoView) msg.scrollIntoView({ block: "nearest" });
    $("ask-input").value = "";
  }

  function summary() {
    var out = $("summary-output");
    clear(out);
    var t = texts();
    if (!t.some(function (x) { return x.trim(); })) { out.appendChild(h("p", { class: "cc-empty", text: "Load or upload a contract first." })); return; }
    var s = Engine.summarize(t, { perspective: state.perspective });
    out.appendChild(h("p", { class: "doc-type", text: s.typeLabel + " \u00B7 reviewed for the " + s.perspectiveLabel.toLowerCase() + " side" }));
    var table = h("table", { class: "cc-kv" });
    s.fields.forEach(function (f) {
      table.appendChild(h("tr", null, h("th", { text: f.label }), h("td", null, f.value, " ", h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(f.paragraphIndex); } }, "\u2192"))));
    });
    out.appendChild(table);
    if (s.risks.length) {
      out.appendChild(h("h3", { class: "cc-section", text: "Key risks" }));
      var ul = h("ul", { class: "cc-risks" });
      s.risks.forEach(function (r) { ul.appendChild(h("li", null, h("span", { class: "cc-badge " + r.severity, text: r.severity === "medium" ? "Med" : cap(r.severity) }), " " + r.title + " ", h("button", { type: "button", class: "cc-ref", onclick: function () { goTo(r.paragraphIndex); } }, r.ref))); });
      out.appendChild(ul);
    }
    if (s.missing.length) out.appendChild(h("p", { class: "cc-missing", text: "Missing: " + s.missing.join(", ") }));
    out.appendChild(h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: function () {
      try {
        navigator.clipboard.writeText(s.text).then(function () { toast("Summary copied", "ok"); }, function () { toast("Copy not available in this browser", "err"); });
      } catch (e) { toast("Copy not available in this browser", "err"); }
    } }, "Copy summary"));
  }

  /* ---------------- export / upload / chrome ---------------- */

  function exportHtml() {
    var title = (paragraphs()[0] && currentText(paragraphs()[0]).trim()) || "Contract";
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
    var blob = new Blob(["<!DOCTYPE html>\n" + doc.documentElement.outerHTML], { type: "text/html" });
    var a = h("a", { href: URL.createObjectURL(blob), download: title.replace(/[^\w\- ]+/g, "").trim().replace(/\s+/g, "-").toLowerCase() + "-redline.html" });
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
  }

  /* .docx reader: a .docx is a zip; pull out word/document.xml with the browser's own
     DecompressionStream and keep each paragraph's visible text (tracked deletions skipped). */
  var W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
    var body = doc.getElementsByTagNameNS(W_NS, "body")[0];
    if (!body) throw new Error("That file isn't a Word document.");
    var SKIP = { del: 1, moveFrom: 1, delText: 1, instrText: 1, pPr: 1, rPr: 1, p: 1 };
    function text(p) {
      var out = "";
      (function walk(n) {
        for (var c = n.firstChild; c; c = c.nextSibling) {
          if (c.nodeType !== 1 || c.localName === "Fallback") continue;
          if (c.namespaceURI === W_NS) {
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
        if (c.namespaceURI === W_NS && c.localName === "p") lines.push(text(c));
        if (c.namespaceURI === W_NS && (c.localName === "del" || c.localName === "moveFrom" || c.localName === "sectPr")) continue;
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

  function addPerspectiveControl() {
    var intro = document.querySelector("#panel-review .panel-intro");
    var runBtn = $("btn-run-review-side");
    var sel = h("select", { id: "perspective", "aria-label": "Whose side you are reviewing for" });
    [["auto", "Auto (from document)"], ["tenant", "Tenant"], ["landlord", "Landlord"], ["buyer", "Buyer"], ["seller", "Seller"], ["neutral", "Both sides"]].forEach(function (o) {
      sel.appendChild(h("option", { value: o[0], text: o[1] }));
    });
    sel.addEventListener("change", function () { state.perspective = sel.value; if (state.result) runReview(); });
    intro.insertBefore(h("label", { class: "cc-field" }, h("span", { text: "Reviewing for" }), sel), runBtn);
  }

  function bind() {
    $("btn-load-demo").addEventListener("click", function () {
      var t = texts();
      var isDemo = t.length >= Demo.lease.paragraphs.length && Demo.lease.paragraphs.every(function (x, i) { return t[i] === x; });
      if (t.some(function (x) { return x.trim(); }) && !isDemo && !window.confirm("Replace the current document with the demo lease?")) return;
      setDocument(Demo.lease.paragraphs);
      runReview();
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
    $("btn-export").addEventListener("click", exportHtml);
    $("btn-print").addEventListener("click", function () { window.print(); });
    $("btn-run-review").addEventListener("click", function () { switchTab("review"); runReview(); });
    $("btn-run-review-side").addEventListener("click", runReview);
    $("btn-summary").addEventListener("click", summary);
    $("ask-form").addEventListener("submit", function (e) { e.preventDefault(); ask($("ask-input").value); });
    Array.prototype.forEach.call(document.querySelectorAll(".ask-chips .chip"), function (c) { c.addEventListener("click", function () { ask(c.getAttribute("data-q")); }); });
    Array.prototype.forEach.call(document.querySelectorAll(".sidebar-tabs .tab"), function (t) { t.addEventListener("click", function () { switchTab(t.getAttribute("data-tab")); }); });
    Array.prototype.forEach.call(document.querySelectorAll("#floating-toolbar [data-cmd]"), function (b) {
      b.addEventListener("mousedown", function (e) { e.preventDefault(); });
      b.addEventListener("click", function () { document.execCommand(b.getAttribute("data-cmd"), false, null); save(); });
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
    if (!Engine || !Demo) { document.body.insertBefore(h("p", { class: "cc-fatal", text: "Clause Craft engine failed to load (engine/clause-engine.js)." }), document.body.firstChild); return; }
    addPerspectiveControl();
    bind();
    if (!restore()) setDocument(Demo.lease.paragraphs);
    renderAll();
    switchTab("review");
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
