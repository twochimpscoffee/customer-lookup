/* Reads and writes the customer spreadsheet in SharePoint via Microsoft Graph.
   Nothing here assumes column letters - the caller passes them in from the schema. */
(function () {
  "use strict";
  var cfg = window.TCC_CONFIG;
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var itemBase = null;

  function shareId(url) {
    var b64 = btoa(unescape(encodeURIComponent(url)));
    return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  }

  var TIMEOUT_MS = 90000;
  var session = null;          // one Excel session per panel, so the workbook is opened once, not per request
  var wbPath = null;

  async function call(method, path, body, useSession) {
    var token = await TCCAuth.getToken();
    var headers = { Authorization: "Bearer " + token, "Content-Type": "application/json" };
    if (useSession && session) headers["workbook-session-id"] = session.id;
    var ctl = typeof AbortController !== "undefined" ? new AbortController() : null;
    var timer = ctl ? setTimeout(function () { ctl.abort(); }, TIMEOUT_MS) : null;
    var res;
    try {
      res = await fetch(path.indexOf("http") === 0 ? path : GRAPH + path, {
        method: method, headers: headers, body: body ? JSON.stringify(body) : undefined,
        signal: ctl ? ctl.signal : undefined
      });
    } catch (e) {
      if (e && e.name === "AbortError") {
        var te = new Error("The spreadsheet took too long to respond. It may still be recalculating - open it once in Excel, then try again.");
        te.status = "timeout"; throw te;
      }
      throw e;
    } finally { if (timer) clearTimeout(timer); }
    if (!res.ok) {
      var t = await res.text();
      var msg = t, code = "";
      try { var j = JSON.parse(t).error; msg = j.message; code = j.code || ""; } catch (e) {}
      var err = new Error(msg || ("Microsoft 365 error " + res.status));
      err.status = res.status; err.code = code;
      if (/session/i.test(code + " " + msg)) session = null;   // expired - a new one is made next time
      throw err;
    }
    if (session) session.at = Date.now();
    return res.status === 204 ? null : res.json();
  }

  async function ensureSession() {
    await base();
    if (session && Date.now() - session.at < 4 * 60000) return;
    var r = await call("POST", wbPath + "/createSession", { persistChanges: true });
    session = { id: r.id, at: Date.now() };
  }

  /* Workbook request through the session, retried once if the session had expired */
  async function wbCall(method, path, body) {
    await ensureSession();
    try { return await call(method, path, body, true); }
    catch (e) {
      if (session) throw e;
      await ensureSession();
      return call(method, path, body, true);
    }
  }

  async function findItem() {
    // Exact file ID (most reliable) - falls back to the sharing link
    if (cfg.driveId && cfg.itemId) {
      return call("GET", "/drives/" + cfg.driveId + "/items/" + cfg.itemId + "?$select=id,name,parentReference");
    }
    return call("GET", "/shares/" + shareId(cfg.workbookShareUrl) + "/driveItem?$select=id,name,parentReference");
  }

  async function base() {
    if (itemBase) return itemBase;
    var item = await findItem();
    wbPath = "/drives/" + item.parentReference.driveId + "/items/" + item.id + "/workbook";
    itemBase = wbPath + "/worksheets/" + encodeURIComponent(cfg.sheetName);
    window.TCCWorkbookName = item.name;
    return itemBase;
  }

  function rangePath(b, addr) { return b + "/range(address='" + addr + "')"; }

  // Chunked $batch (Graph allows 20 requests per batch)
  async function batch(reqs, retried) {
    var b = (await base()).replace(/^\//, "");
    await ensureSession();
    var out = [];
    for (var i = 0; i < reqs.length; i += 20) {
      var chunk = reqs.slice(i, i + 20).map(function (r, j) {
        var q = { id: String(i + j), method: r.method, url: "/" + b + r.url };
        q.headers = { "workbook-session-id": session.id };
        if (r.body) { q.body = r.body; q.headers["Content-Type"] = "application/json"; }
        // Excel prefers workbook writes one after another, so chain writes within a batch
        if (r.method !== "GET" && j > 0) q.dependsOn = [String(i + j - 1)];
        return q;
      });
      var res = await call("POST", GRAPH + "/$batch", { requests: chunk });
      res.responses.forEach(function (r) { out[Number(r.id)] = r; });
    }
    var bad = out.filter(function (r) { return r && r.status >= 400; })[0];
    if (bad) {
      var be = bad.body && bad.body.error ? bad.body.error : {};
      if (!retried && /session/i.test((be.code || "") + " " + (be.message || ""))) { session = null; return batch(reqs, true); }
      var e = new Error(be.message || "Update failed");
      e.status = bad.status; throw e;
    }
    return out;
  }
  function rangeUrl(addr, select) { return "/range(address='" + addr + "')" + (select ? "?$select=" + select : ""); }

  /* How big the sheet is: {lastCol: "CI", lastRow: 13048} */
  async function sheetSize() {
    var b = await base();
    var r = await wbCall("GET", b + "/usedRange(valuesOnly=true)?$select=address");
    var m = String(r.address).split("!").pop().match(/([A-Z]+)(\d+)$/);
    return { lastCol: m ? m[1] : "CI", lastRow: m ? Number(m[2]) : 20000 };
  }

  async function readHeaders(lastCol) {
    var b = await base();
    var r = await wbCall("GET", rangePath(b, "A1:" + lastCol + cfg.headerRows) + "?$select=values");
    return r.values;
  }

  /* Names + emails from their columns: {rows:[{row,name,email}], nextEmptyRow} */
  async function readIndex(nameCol, emailCol, lastRow) {
    var first = cfg.headerRows + 1;
    var res = await batch([
      { method: "GET", url: rangeUrl(nameCol + first + ":" + nameCol + lastRow, "values") },
      { method: "GET", url: rangeUrl(emailCol + first + ":" + emailCol + lastRow, "values") }
    ]);
    var names = res[0].body.values, emails = res[1].body.values;
    var list = [], lastUsed = first - 1;
    for (var i = 0; i < names.length; i++) {
      var name = names[i][0] == null ? "" : String(names[i][0]);
      var email = emails[i] && emails[i][0] != null ? String(emails[i][0]) : "";
      if (name || email) { lastUsed = first + i; list.push({ row: first + i, name: name, email: email }); }
    }
    return { rows: list, nextEmptyRow: lastUsed + 1 };
  }

  /* What's needed to work out the columns: formulas + formats of two rows, sample values */
  async function readSchemaInputs(lastCol, lastDataRow) {
    var first = cfg.headerRows + 1;
    var r1 = Math.max(first, lastDataRow), r2 = r1 + 1;
    var tailFrom = Math.max(first, lastDataRow - 299);
    var res = await batch([
      { method: "GET", url: rangeUrl("A" + r1 + ":" + lastCol + r2, "formulas,numberFormat") },
      { method: "GET", url: rangeUrl("A" + first + ":" + lastCol + Math.min(first + 299, lastDataRow), "values") },
      { method: "GET", url: rangeUrl("A" + tailFrom + ":" + lastCol + lastDataRow, "values") },
      // top rows too: some formula columns aren't filled all the way down the sheet
      { method: "GET", url: rangeUrl("A" + first + ":" + lastCol + (first + 2), "formulas") }
    ]);
    return {
      formulaRows: res[3].body.formulas.concat(res[0].body.formulas),
      belowRow: res[0].body.formulas[1] || [],
      numberFormat: res[0].body.numberFormat[0],
      samples: res[1].body.values.concat(res[2].body.values)
    };
  }

  /* Font name of one cell per column: {idx: "Webdings"} */
  async function readFonts(cols, row) {
    if (!cols.length) return {};
    var res = await batch(cols.map(function (c) {
      return { method: "GET", url: rangeUrl(c.letter + row) .replace(/\?.*$/, "") + "/format/font?$select=name" };
    }));
    var out = {};
    cols.forEach(function (c, i) { out[c.idx] = res[i].body.name; });
    return out;
  }

  /* Which columns are hidden in Excel: [true/false per column, from A] */
  async function readHidden(width) {
    var reqs = [];
    for (var i = 0; i < width; i++) {
      reqs.push({ method: "GET", url: rangeUrl(TCCSchema.letter(i) + "1", "columnHidden") });
    }
    var res = await batch(reqs);
    return res.map(function (r) { return !!(r && r.body && r.body.columnHidden); });
  }

  /* Full rows: {rowNumber: {values:[], text:[]}} */
  async function readRows(rowNumbers, lastCol) {
    var res = await batch(rowNumbers.map(function (n) {
      return { method: "GET", url: rangeUrl("A" + n + ":" + lastCol + n, "values,text") };
    }));
    var out = {};
    rowNumbers.forEach(function (n, i) { out[n] = { values: res[i].body.values[0], text: res[i].body.text[0] }; });
    return out;
  }

  /* changes: [{col, row, value, tick}] */
  async function writeCells(changes) {
    var reqs = [];
    changes.forEach(function (c) {
      var p = rangeUrl(c.col + c.row);
      reqs.push({ method: "PATCH", url: p, body: { values: [[c.value]] } });
      if (c.tick && c.value) reqs.push({ method: "PATCH", url: p + "/format/font", body: { name: cfg.tickFont } });
    });
    return batch(reqs);
  }

  window.TCCGraph = { sheetSize: sheetSize, readHeaders: readHeaders, readIndex: readIndex,
                      readSchemaInputs: readSchemaInputs, readFonts: readFonts, readHidden: readHidden,
                      readRows: readRows, writeCells: writeCells };
})();
