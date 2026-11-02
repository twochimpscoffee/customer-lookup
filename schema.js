/* Two Chimps customer lookup - works out the spreadsheet's columns from its two header rows,
   so columns can be added, moved or inserted without changing the panel's code.
   Pure functions (no Office/Graph), so it can be tested in Node. */
(function (root) {
  "use strict";

  function letter(idx) {
    var s = "", n = idx + 1;
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function clean(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }

  var TICK_FONTS = /webdings|wingdings/i;
  var TICK_VALUE = /^(a{1,3}|ü{1,3}|A|s|q|x|aq|as|sa|v|va)$/;
  var DATE_FORMAT = /(d{1,4}.*m{1,4})|(m{1,4}.*d{1,4})|y{2,4}/i;

  function isFormula(v) { return typeof v === "string" && v.charAt(0) === "="; }

  function guessType(col, ctx) {
    var fmt = String(ctx.numberFormat[col.idx] || "");
    var font = String(ctx.fonts[col.idx] || "");
    var vals = ctx.samples.map(function (r) { return r[col.idx]; })
      .filter(function (v) { return v !== "" && v != null; });
    var label = norm(col.label);
    if (vals.length) {
      var ticks = vals.filter(function (v) { return TICK_VALUE.test(String(v).trim()); }).length;
      if (ticks / vals.length >= 0.8) return "tick";
      var nums = vals.filter(function (v) { return typeof v === "number"; }).length;
      if (nums / vals.length >= 0.8 && DATE_FORMAT.test(fmt) && !/general/i.test(fmt)) return "date";
      if (vals.some(function (v) { return String(v).length > 60; })) return "longtext";
      return "text";
    }
    // brand-new, still-empty column: go by its font, format and header wording
    if (TICK_FONTS.test(font)) return "tick";
    if (DATE_FORMAT.test(fmt) && !/general|@/i.test(fmt)) return "date";
    if (/\bdate\b/.test(label)) return "date";
    if (/\?$|\bsent\b|brochure|letter|leaflet|guide|magnet|sticker/.test(label)) return "tick";
    return "text";
  }

  /*
    input = {
      headers: [row1[], row2[]],         // header rows, as values
      formulaRows: [[...], [...]],       // one or two full rows of formulas (to spot phase columns)
      numberFormat: [...],               // one row of number formats
      fonts: {idx: fontName},            // font of one recent data row, per column (optional)
      samples: [[...], ...]              // some data rows (values) to guess column types
    }
    cfg = TCC_CONFIG
  */
  function buildSchema(input, cfg) {
    var h1 = input.headers[0] || [], h2 = input.headers[1] || [];
    var width = Math.max(h1.length, h2.length);
    var ignore = (cfg.ignoreHeaders || []).map(norm);
    var hide = (cfg.hideHeaders || []).map(norm);
    var ctx = { numberFormat: input.numberFormat || [], fonts: input.fonts || {}, samples: input.samples || [] };

    // 1. describe every column
    var cols = [], carry = "";
    for (var i = 0; i < width; i++) {
      var a = clean(h1[i]), b = clean(h2[i]);
      if (a && !/^(done|sent|can we send\??)$/i.test(a)) carry = a;
      var label;
      if (b) label = (carry && norm(carry) !== norm(b) && !/customer details/i.test(carry) && !a) ? carry + " – " + b
                   : (a && !/customer details/i.test(a) && norm(a) !== norm(b)) ? a + " – " + b : b;
      else label = a;
      var ignored = ignore.indexOf(norm(a)) >= 0 || ignore.indexOf(norm(b)) >= 0;
      cols.push({ idx: i, letter: letter(i), h1: a, h2: b, label: label, ignored: ignored,
                  formula: (input.formulaRows || []).some(function (r) { return isFormula(r[i]); }) });
    }

    // 2. core fields by header name
    var used = {}, core = [], missing = [];
    (cfg.coreFields || []).forEach(function (f) {
      var want = norm(f.header);
      var c = cols.filter(function (x) { return !x.ignored && !x.formula && (norm(x.h2) === want || norm(x.h1) === want); })[0];
      if (!c) { missing.push(f.header); return; }
      used[c.idx] = true;
      core.push({ key: f.key, idx: c.idx, letter: c.letter, label: f.label || f.header, type: f.type,
                  options: f.options, header: f.header });
    });

    // 3. phases: a run of formula columns, followed by the column where it's ticked off
    var phases = [];
    for (var j = 0; j < width; j++) {
      if (!cols[j].formula || used[j] || cols[j].ignored) continue;
      var start = j, k = j;
      while (k < width && cols[k].formula) { used[k] = true; k++; }
      var due = cols[start], canSend = null;
      for (var m = start + 1; m < k; m++) if (/can we send/i.test(cols[m].h1 + " " + cols[m].h2)) canSend = cols[m];
      var done = (k < width && !used[k] && !cols[k].ignored) ? cols[k] : null;
      if (done) used[done.idx] = true;
      var top = due.h1, sub = due.h2;
      var plabel = sub ? (/^phase/i.test(top) ? top + " – " + sub : sub) : top;
      phases.push({ label: plabel || due.letter,
                    due: { idx: due.idx, letter: due.letter },
                    done: done ? { idx: done.idx, letter: done.letter } : null,
                    canSend: canSend ? { idx: canSend.idx, letter: canSend.letter } : null });
      j = k;
    }

    // 4. everything else that has a header: shown automatically, type guessed
    var others = [];
    cols.forEach(function (c) {
      if (used[c.idx] || c.ignored || c.formula || !c.label) return;
      if (hide.indexOf(norm(c.label)) >= 0 || hide.indexOf(norm(c.h1)) >= 0 || hide.indexOf(norm(c.h2)) >= 0) return;
      var override = (cfg.fieldTypes || {})[norm(c.label)] || (cfg.fieldTypes || {})[norm(c.h2)] || (cfg.fieldTypes || {})[norm(c.h1)];
      others.push({ idx: c.idx, letter: c.letter, label: c.label, type: override || guessType(c, ctx) });
    });

    var lastIdx = 0;
    cols.forEach(function (c) { if (!c.ignored && (c.label || c.formula)) lastIdx = Math.max(lastIdx, c.idx); });
    var key = {};
    core.forEach(function (f) { key[f.key] = f; });

    return { cols: cols, core: core, byKey: key, phases: phases, others: others, missing: missing,
             lastLetter: letter(lastIdx), width: width,
             signature: JSON.stringify([h1.map(clean), h2.map(clean)]) };
  }

  var api = { buildSchema: buildSchema, letter: letter, norm: norm };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TCCSchema = api;
})(this);
