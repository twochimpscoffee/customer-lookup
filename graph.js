/* Reads and writes the customer spreadsheet in SharePoint via Microsoft Graph. */
(function () {
  "use strict";
  var cfg = window.TCC_CONFIG;
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var itemBase = null, headerCache = null;

  function shareId(url) {
    var b64 = btoa(unescape(encodeURIComponent(url)));
    return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  }

  async function call(method, path, body) {
    var token = await TCCAuth.getToken();
    var res = await fetch(path.indexOf("http") === 0 ? path : GRAPH + path, {
      method: method,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined
    });
    if (!res.ok) {
      var t = await res.text();
      var msg = t;
      try { msg = JSON.parse(t).error.message; } catch (e) {}
      var err = new Error(msg || ("Microsoft 365 error " + res.status));
      err.status = res.status;
      throw err;
    }
    return res.status === 204 ? null : res.json();
  }

  async function base() {
    if (itemBase) return itemBase;
    var item = await call("GET", "/shares/" + shareId(cfg.workbookShareUrl) + "/driveItem?$select=id,name,parentReference");
    itemBase = "/drives/" + item.parentReference.driveId + "/items/" + item.id +
               "/workbook/worksheets/" + encodeURIComponent(cfg.sheetName);
    window.TCCWorkbookName = item.name;
    return itemBase;
  }

  function rangePath(b, addr) { return b + "/range(address='" + addr + "')"; }

  // Chunked $batch (Graph allows 20 requests per batch)
  async function batch(reqs) {
    var out = [];
    for (var i = 0; i < reqs.length; i += 20) {
      var chunk = reqs.slice(i, i + 20).map(function (r, j) {
        var q = { id: String(i + j), method: r.method, url: r.url };
        if (r.body) { q.body = r.body; q.headers = { "Content-Type": "application/json" }; }
        // Excel prefers workbook writes one after another, so chain writes within a batch
        if (r.method !== "GET" && j > 0) q.dependsOn = [String(i + j - 1)];
        return q;
      });
      var res = await call("POST", GRAPH + "/$batch", { requests: chunk });
      res.responses.forEach(function (r) { out[Number(r.id)] = r; });
    }
    out.forEach(function (r) {
      if (r && r.status >= 400) {
        var e = new Error((r.body && r.body.error && r.body.error.message) || "Update failed");
        e.status = r.status; throw e;
      }
    });
    return out;
  }

  /* All names + emails: [{row, name, email}] */
  async function readIndex() {
    var b = await base();
    var r = await call("GET", rangePath(b, "A" + cfg.firstDataRow + ":B" + cfg.lastFormulaRow) + "?$select=values");
    var list = [], lastUsed = cfg.firstDataRow - 1;
    r.values.forEach(function (v, i) {
      var row = cfg.firstDataRow + i;
      var name = v[0] == null ? "" : String(v[0]);
      var email = v[1] == null ? "" : String(v[1]);
      if (name || email) { lastUsed = row; list.push({ row: row, name: name, email: email }); }
    });
    return { rows: list, nextEmptyRow: lastUsed + 1 };
  }

  async function readHeaders() {
    if (headerCache) return headerCache;
    var b = await base();
    var r = await call("GET", rangePath(b, "A1:" + cfg.lastColumn + "2") + "?$select=values");
    headerCache = r.values;
    return headerCache;
  }

  /* Full rows: returns {rowNumber: {values:[], text:[]}} */
  async function readRows(rowNumbers) {
    var b = (await base()).replace(/^\//, "");
    var res = await batch(rowNumbers.map(function (n) {
      return { method: "GET", url: "/" + rangePath(b, "A" + n + ":" + cfg.lastColumn + n) + "?$select=values,text" };
    }));
    var out = {};
    rowNumbers.forEach(function (n, i) {
      out[n] = { values: res[i].body.values[0], text: res[i].body.text[0] };
    });
    return out;
  }

  /* changes: [{col, row, value, tick}] */
  async function writeCells(changes) {
    var b = (await base()).replace(/^\//, "");
    var reqs = [];
    changes.forEach(function (c) {
      var p = "/" + rangePath(b, c.col + c.row);
      reqs.push({ method: "PATCH", url: p, body: { values: [[c.value]] } });
      if (c.tick && c.value) reqs.push({ method: "PATCH", url: p + "/format/font", body: { name: cfg.tickFont } });
    });
    return batch(reqs);
  }

  window.TCCGraph = { readIndex: readIndex, readHeaders: readHeaders, readRows: readRows, writeCells: writeCells };
})();
