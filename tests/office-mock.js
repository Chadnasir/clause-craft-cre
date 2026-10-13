/*
 * Minimal Office.js / Word API simulator for testing the Clause Craft task pane without Word.
 * Served in place of https://appsforoffice.microsoft.com/lib/1/hosted/office.js by tests/run-ui-tests.mjs.
 *
 * Mirrors the parts of the real API the add-in uses, including its batching rules:
 *  - API calls are queued and only run on context.sync() (or when Word.run finishes);
 *  - properties can only be read after load() + sync(), otherwise PropertyNotLoaded is thrown;
 *  - tracked changes (TrackAll) record deletions/insertions instead of editing text in place;
 *  - WordApi 1.4 features (comments, tracking, getReviewedText) can be switched off.
 * The document lives in window.parent.MockWordHost when loaded inside tests/harness.html.
 */
(function () {
  "use strict";

  function createHost() {
    return {
      config: { api: "1.4", paragraphTextIncludesDeleted: true, searchIncludesDeleted: true },
      doc: { paragraphs: [{ runs: [], align: "Left", bold: false }], comments: [], selection: null, trackMode: "Off", settings: {} },
      log: [],
      render: function () {}
    };
  }

  var host = (function () {
    try { if (window.parent && window.parent !== window && window.parent.MockWordHost) return window.parent.MockWordHost; } catch (e) { /* cross-origin */ }
    return window.MockWordHost || (window.MockWordHost = createHost());
  })();
  var D = host.doc;

  function apiAtLeast(v) { return parseFloat(host.config.api) >= parseFloat(v); }
  function notLoaded(prop) {
    var e = new Error("The property '" + prop + "' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.");
    e.code = "PropertyNotLoaded";
    throw e;
  }
  function apiNotFound(name) { var e = new Error("This API is not supported in this version of Word: " + name); e.code = "ApiNotFound"; return e; }

  /* ---------------- document model helpers ---------------- */

  function fullText(p) { return p.runs.map(function (r) { return r.text; }).join(""); }
  function currentText(p) { return p.runs.filter(function (r) { return !r.del; }).map(function (r) { return r.text; }).join(""); }
  function originalText(p) { return p.runs.filter(function (r) { return !r.ins; }).map(function (r) { return r.text; }).join(""); }
  function visibleText(p) { return host.config.paragraphTextIncludesDeleted ? fullText(p) : currentText(p); }
  function plen(p) { return fullText(p).length; }

  function cloneFmt(f) { var o = {}; for (var k in f) o[k] = f[k]; return o; }

  /** Ensures a run boundary at offset; returns the index of the run that starts at offset. */
  function splitAt(p, offset) {
    var pos = 0;
    for (var i = 0; i < p.runs.length; i++) {
      var r = p.runs[i];
      if (offset === pos) return i;
      if (offset < pos + r.text.length) {
        var left = { text: r.text.slice(0, offset - pos), ins: r.ins, del: r.del, fmt: cloneFmt(r.fmt) };
        var right = { text: r.text.slice(offset - pos), ins: r.ins, del: r.del, fmt: cloneFmt(r.fmt) };
        p.runs.splice(i, 1, left, right);
        return i + 1;
      }
      pos += r.text.length;
    }
    return p.runs.length;
  }

  function runsIn(p, start, end) {
    var a = splitAt(p, start), b = splitAt(p, end);
    return { from: a, to: b, runs: p.runs.slice(a, b) };
  }

  function fmtBefore(p, offset) {
    var pos = 0, last = {};
    for (var i = 0; i < p.runs.length; i++) {
      if (pos >= offset) break;
      last = p.runs[i].fmt;
      pos += p.runs[i].text.length;
    }
    return cloneFmt(last);
  }

  function insertRun(p, offset, text, tracked) {
    var idx = splitAt(p, offset);
    p.runs.splice(idx, 0, { text: text, ins: !!tracked, del: false, fmt: fmtBefore(p, offset) });
    return { para: p, start: offset, end: offset + text.length };
  }

  function tracking() { return apiAtLeast("1.4") && D.trackMode !== "Off"; }

  function replaceRange(loc, text) {
    var p = loc.para;
    if (tracking()) {
      var span = runsIn(p, loc.start, loc.end);
      var removedLen = 0, i;
      // deleting text that was itself a tracked insertion removes it outright (as Word does)
      for (i = span.to - 1; i >= span.from; i--) {
        var r = p.runs[i];
        if (r.ins) { removedLen += r.text.length; p.runs.splice(i, 1); }
        else r.del = true;
      }
      var at = loc.end - removedLen;
      return insertRun(p, at, text, true);
    }
    var s = runsIn(p, loc.start, loc.end);
    var fmt = s.runs.length ? cloneFmt(s.runs[0].fmt) : fmtBefore(p, loc.start);
    p.runs.splice(s.from, s.to - s.from, { text: text, ins: false, del: false, fmt: fmt });
    return { para: p, start: loc.start, end: loc.start + text.length };
  }

  function applyFont(loc, key, value) {
    var s = runsIn(loc.para, loc.start, loc.end);
    s.runs.forEach(function (r) { r.fmt[key] = value; });
  }

  function newParagraph(text, tracked, like) {
    return { runs: text ? [{ text: text, ins: !!tracked, del: false, fmt: {} }] : [], align: like ? like.align : "Left", bold: false, tracked: !!tracked };
  }

  function unescapeSearch(s) { return String(s).replace(/\^\^/g, "^"); }

  /** Search inside [start,end) of paragraph p. Returns full-text locations. */
  function searchIn(p, start, end, needle, opts) {
    needle = unescapeSearch(needle);
    if (!needle) return [];
    if (needle.length > 255) { var e = new Error("Search string too long"); e.code = "InvalidArgument"; throw e; }
    var matchCase = opts && opts.matchCase;
    var map = [], hay = "", pos = 0;
    p.runs.forEach(function (r) {
      for (var k = 0; k < r.text.length; k++) {
        var abs = pos + k;
        if (abs >= start && abs < end && (host.config.searchIncludesDeleted || !r.del)) { map.push(abs); hay += r.text.charAt(k); }
      }
      pos += r.text.length;
    });
    var h = matchCase ? hay : hay.toLowerCase(), n = matchCase ? needle : needle.toLowerCase();
    var out = [], i = h.indexOf(n);
    while (i !== -1) {
      out.push({ para: p, start: map[i], end: map[i + n.length - 1] + 1 });
      i = h.indexOf(n, i + n.length);
    }
    return out;
  }

  function textOfLoc(loc) {
    var out = "", pos = 0;
    loc.para.runs.forEach(function (r) {
      for (var k = 0; k < r.text.length; k++) {
        var abs = pos + k;
        if (abs >= loc.start && abs < loc.end && (host.config.paragraphTextIncludesDeleted || !r.del)) out += r.text.charAt(k);
      }
      pos += r.text.length;
    });
    return out;
  }

  /* ---------------- proxies ---------------- */

  function RequestContext() {
    this._queue = [];
    this.document = new Document(this);
    this.trackedObjects = { add: function () {}, remove: function () {} };
  }
  RequestContext.prototype.load = function (obj, props) { obj.load(props); return obj; };
  RequestContext.prototype._push = function (fn) { this._queue.push(fn); };
  RequestContext.prototype.sync = function () {
    var q = this._queue;
    this._queue = [];
    try {
      for (var i = 0; i < q.length; i++) q[i]();
    } catch (e) {
      return Promise.reject(e);
    } finally {
      try { host.render(); } catch (e2) { /* ignore */ }
    }
    return Promise.resolve();
  };

  function ClientResult() { this._set = false; }
  Object.defineProperty(ClientResult.prototype, "value", { get: function () { if (!this._set) notLoaded("value"); return this._value; } });

  function Base(ctx) { this.context = ctx; this._loaded = {}; this._vals = {}; }
  Base.prototype._get = function (prop) { if (!this._loaded[prop]) notLoaded(prop); return this._vals[prop]; };
  function parseProps(props) {
    if (!props) return ["*"];
    if (typeof props === "string") return props.split(",").map(function (s) { return s.trim(); });
    if (Array.isArray(props)) return props;
    if (props.select) return parseProps(props.select);
    return Object.keys(props);
  }

  /* Font proxy: setters queue formatting ops on the owner's location. */
  function Font(owner) { this._owner = owner; }
  ["bold", "italic", "strikeThrough", "color", "underline", "highlightColor", "size", "name"].forEach(function (key) {
    Object.defineProperty(Font.prototype, key, {
      set: function (v) {
        var owner = this._owner;
        owner.context._push(function () {
          var loc = owner._resolve();
          if (key === "bold" && owner._isParagraph) owner._para().bold = v;
          applyFont(loc, key, v);
        });
      },
      get: function () { notLoaded("font." + key); }
    });
  });

  /* Range */
  function Range(ctx, resolver) { Base.call(this, ctx); this._resolver = resolver; this.font = new Font(this); }
  Range.prototype = Object.create(Base.prototype);
  Range.prototype._resolve = function () { var l = this._resolver(); if (!l) throw new Error("Range is not available"); return l; };
  Range.prototype._para = function () { return this._resolve().para; };
  Object.defineProperty(Range.prototype, "text", { get: function () { return this._get("text"); } });
  Object.defineProperty(Range.prototype, "isEmpty", { get: function () { return this._get("isEmpty"); } });
  Range.prototype.load = function (props) {
    var self = this, list = parseProps(props);
    this.context._push(function () {
      var loc = self._resolve();
      if (list.indexOf("*") !== -1 || list.indexOf("text") !== -1) { self._vals.text = textOfLoc(loc); self._loaded.text = true; }
      if (list.indexOf("*") !== -1 || list.indexOf("isEmpty") !== -1) { self._vals.isEmpty = loc.end <= loc.start; self._loaded.isEmpty = true; }
    });
    return this;
  };
  Range.prototype.search = function (text, opts) {
    var self = this;
    var coll = new RangeCollection(this.context);
    this.context._push(function () {
      var loc = self._resolve();
      coll._locs = searchIn(loc.para, loc.start, loc.end, text, opts);
    });
    return coll;
  };
  Range.prototype.insertText = function (text, where) {
    var self = this, holder = { loc: null };
    var r = new Range(this.context, function () { return holder.loc; });
    this.context._push(function () {
      var loc = self._resolve();
      if (where === "Replace") holder.loc = replaceRange(loc, text);
      else if (where === "After" || where === "End") holder.loc = insertRun(loc.para, loc.end, text, tracking());
      else if (where === "Before" || where === "Start") holder.loc = insertRun(loc.para, loc.start, text, tracking());
      else throw new Error("Invalid insertLocation " + where);
      host.log.push({ op: "insertText", where: where, text: text, tracked: tracking() });
    });
    return r;
  };
  Range.prototype.insertParagraph = function (text, where) {
    var self = this, holder = { p: null };
    var para = new Paragraph(this.context, function () { return holder.p; });
    this.context._push(function () {
      var loc = self._resolve();
      var idx = D.paragraphs.indexOf(loc.para);
      var np = newParagraph(text, tracking(), loc.para);
      if (where === "After") D.paragraphs.splice(idx + 1, 0, np);
      else if (where === "Before") D.paragraphs.splice(idx, 0, np);
      else throw new Error("Invalid insertLocation " + where);
      holder.p = np;
      host.log.push({ op: "insertParagraph", where: where, text: text, tracked: tracking() });
    });
    return para;
  };
  Range.prototype.insertComment = function (text) {
    if (!apiAtLeast("1.4")) throw apiNotFound("Range.insertComment");
    var self = this;
    this.context._push(function () {
      var loc = self._resolve();
      D.comments.push({ para: loc.para, start: loc.start, end: loc.end, text: text, anchor: textOfLoc(loc) });
      host.log.push({ op: "insertComment", text: text });
    });
    return { context: this.context };
  };
  Range.prototype.select = function () {
    var self = this;
    this.context._push(function () { var loc = self._resolve(); D.selection = { para: loc.para, start: loc.start, end: loc.end }; });
  };
  /* Paragraphs touched by this range (the mock's ranges never span paragraphs). */
  Object.defineProperty(Range.prototype, "paragraphs", {
    get: function () {
      var self = this;
      var coll = new ParagraphCollection(this.context);
      coll._filter = function () { return [self._resolve().para]; };
      return coll;
    }
  });
  Range.prototype.getReviewedText = function (version) {
    if (!apiAtLeast("1.4")) throw apiNotFound("Range.getReviewedText");
    var self = this, res = new ClientResult();
    this.context._push(function () { var p = self._para(); res._value = version === "Original" ? originalText(p) : currentText(p); res._set = true; });
    return res;
  };

  /* Paragraph */
  function Paragraph(ctx, getPara) {
    var self = this;
    Range.call(this, ctx, function () { var p = getPara(); return p ? { para: p, start: 0, end: plen(p) } : null; });
    this._getPara = getPara;
    this._isParagraph = true;
    Object.defineProperty(this, "alignment", {
      set: function (v) { ctx._push(function () { self._getPara().align = v; }); },
      get: function () { return self._get("alignment"); }
    });
  }
  Paragraph.prototype = Object.create(Range.prototype);
  Paragraph.prototype.load = function (props) {
    var self = this, list = parseProps(props);
    this.context._push(function () {
      var p = self._getPara();
      if (list.indexOf("*") !== -1 || list.indexOf("text") !== -1) { self._vals.text = visibleText(p); self._loaded.text = true; }
      if (list.indexOf("*") !== -1 || list.indexOf("alignment") !== -1) { self._vals.alignment = p.align; self._loaded.alignment = true; }
    });
    return this;
  };
  Paragraph.prototype.insertText = function (text, where) {
    var self = this, holder = { loc: null };
    var r = new Range(this.context, function () { return holder.loc; });
    this.context._push(function () {
      var p = self._getPara();
      if (where === "Replace") holder.loc = replaceRange({ para: p, start: 0, end: plen(p) }, text);
      else if (where === "End") holder.loc = insertRun(p, plen(p), text, tracking());
      else if (where === "Start") holder.loc = insertRun(p, 0, text, tracking());
      else throw new Error("Invalid insertLocation for Paragraph.insertText: " + where);
      host.log.push({ op: "paragraph.insertText", where: where, text: text, tracked: tracking() });
    });
    return r;
  };
  Paragraph.prototype.getRange = function (where) {
    var self = this;
    return new Range(this.context, function () {
      var p = self._getPara(), n = plen(p);
      if (where === "Start") return { para: p, start: 0, end: 0 };
      if (where === "End") return { para: p, start: n, end: n };
      return { para: p, start: 0, end: n };
    });
  };

  /* Collections */
  function RangeCollection(ctx) { Base.call(this, ctx); this._locs = null; this._items = null; }
  RangeCollection.prototype = Object.create(Base.prototype);
  Object.defineProperty(RangeCollection.prototype, "items", { get: function () { return this._get("items"); } });
  RangeCollection.prototype.load = function (props) {
    var self = this, list = parseProps(props);
    this.context._push(function () {
      self._vals.items = (self._locs || []).map(function (l) {
        var r = new Range(self.context, function () { return l; });
        if (list.some(function (x) { return x.indexOf("text") !== -1 || x === "*"; })) { r._vals.text = textOfLoc(l); r._loaded.text = true; }
        return r;
      });
      self._loaded.items = true;
    });
    return this;
  };

  function ParagraphCollection(ctx) { Base.call(this, ctx); }
  ParagraphCollection.prototype = Object.create(Base.prototype);
  Object.defineProperty(ParagraphCollection.prototype, "items", { get: function () { return this._get("items"); } });
  ParagraphCollection.prototype.load = function (props) {
    var self = this, list = parseProps(props);
    var wantText = list.some(function (x) { return x === "*" || x === "text" || x === "items/text"; });
    this.context._push(function () {
      self._vals.items = (self._filter ? self._filter() : D.paragraphs).map(function (p) {
        var para = new Paragraph(self.context, function () { return p; });
        if (wantText) { para._vals.text = visibleText(p); para._loaded.text = true; }
        return para;
      });
      self._loaded.items = true;
    });
    return this;
  };

  /* Body + Document */
  function Body(ctx) { Base.call(this, ctx); this.paragraphs = new ParagraphCollection(ctx); }
  Body.prototype = Object.create(Base.prototype);
  Object.defineProperty(Body.prototype, "text", { get: function () { return this._get("text"); } });
  Body.prototype.load = function (props) {
    var self = this;
    this.context._push(function () { self._vals.text = D.paragraphs.map(visibleText).join("\r"); self._loaded.text = true; });
    return this;
  };
  Body.prototype.insertParagraph = function (text, where) {
    var holder = { p: null };
    var para = new Paragraph(this.context, function () { return holder.p; });
    this.context._push(function () {
      var np = newParagraph(text, tracking());
      if (where === "End") D.paragraphs.push(np);
      else if (where === "Start") D.paragraphs.unshift(np);
      else throw new Error("Invalid insertLocation for Body.insertParagraph: " + where);
      holder.p = np;
      host.log.push({ op: "body.insertParagraph", where: where, text: text, tracked: tracking() });
    });
    return para;
  };
  Body.prototype.search = function (text, opts) {
    var coll = new RangeCollection(this.context);
    this.context._push(function () {
      coll._locs = [];
      D.paragraphs.forEach(function (p) { coll._locs = coll._locs.concat(searchIn(p, 0, plen(p), text, opts)); });
    });
    return coll;
  };

  function Document(ctx) {
    Base.call(this, ctx);
    var self = this;
    this.body = new Body(ctx);
    Object.defineProperty(this, "changeTrackingMode", {
      get: function () { if (!apiAtLeast("1.4")) throw apiNotFound("changeTrackingMode"); return self._get("changeTrackingMode"); },
      set: function (v) {
        if (!apiAtLeast("1.4")) throw apiNotFound("changeTrackingMode");
        ctx._push(function () { D.trackMode = v; host.log.push({ op: "trackMode", value: v }); });
      }
    });
  }
  Document.prototype = Object.create(Base.prototype);
  Document.prototype.load = function (props) {
    var self = this, list = parseProps(props);
    this.context._push(function () {
      if (list.indexOf("changeTrackingMode") !== -1) {
        if (!apiAtLeast("1.4")) throw apiNotFound("changeTrackingMode");
        self._vals.changeTrackingMode = D.trackMode; self._loaded.changeTrackingMode = true;
      }
    });
    return this;
  };
  Document.prototype.getSelection = function () {
    return new Range(this.context, function () {
      if (D.selection && D.paragraphs.indexOf(D.selection.para) !== -1) return D.selection;
      var last = D.paragraphs[D.paragraphs.length - 1];
      return { para: last, start: plen(last), end: plen(last) };
    });
  };

  /* ---------------- globals ---------------- */

  window.OfficeExtension = { Error: Error, Promise: Promise };
  window.Word = {
    InsertLocation: { replace: "Replace", start: "Start", end: "End", before: "Before", after: "After" },
    ChangeTrackingMode: { off: "Off", trackAll: "TrackAll", trackMineOnly: "TrackMineOnly" },
    run: function (batch) {
      var ctx = new RequestContext();
      return Promise.resolve().then(function () { return batch(ctx); }).then(function (out) {
        return ctx.sync().then(function () { return out; });
      });
    }
  };
  var inWord = host.config.host !== "none";
  window.Office = {
    HostType: { Word: "Word", Excel: "Excel" },
    context: {
      requirements: { isSetSupported: function (name, v) { return name === "WordApi" && apiAtLeast(v || "1.1"); } },
      document: {
        settings: {
          get: function (k) { return Object.prototype.hasOwnProperty.call(D.settings, k) ? D.settings[k] : null; },
          set: function (k, v) { D.settings[k] = v; },
          saveAsync: function (cb) { if (cb) cb({ status: "succeeded" }); }
        }
      },
      displayLanguage: "en-US"
    },
    onReady: function (cb) {
      var info = inWord ? { host: "Word", platform: "PC" } : { host: null, platform: null };
      setTimeout(function () { if (cb) cb(info); }, 0);
      return Promise.resolve(info);
    }
  };
})();
