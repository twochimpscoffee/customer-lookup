/* Two Chimps customer lookup - side panel logic */
(function () {
  "use strict";
  var cfg = window.TCC_CONFIG;
  var state = { parsed: null, index: null, headers: null, rows: {}, current: null, query: null };
  var $ = function (id) { return document.getElementById(id); };

  // ---------- helpers ----------
  function colIdx(col) { var n = 0; for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64); return n - 1; }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function serialToISO(v) {
    if (typeof v !== "number" || v < 1) return "";
    var d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  function isoToSerial(s) {
    if (!s) return "";
    var p = s.split("-");
    return (Date.UTC(+p[0], +p[1] - 1, +p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
  }
  function todaySerial() { var n = new Date(); return isoToSerial(n.getFullYear() + "-" + (n.getMonth() + 1) + "-" + n.getDate()); }
  function fmtDate(v) {
    var iso = serialToISO(v); if (!iso) return "";
    return new Date(iso + "T12:00:00Z").toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
  }
  function isTicked(v) { return v != null && String(v).trim() !== ""; }
  function isPlainTick(v) { var s = String(v || "").trim(); return s === "" || /^(a+|ü+|A)$/.test(s); }
  function headerFor(col) {
    var i = colIdx(col), h = state.headers || [[], []];
    var a = (h[0] && h[0][i]) || "", b = (h[1] && h[1][i]) || "";
    a = String(a).trim(); b = String(b).trim();
    return b || a || col;
  }
  function phaseLabel(p) {
    var i = colIdx(p.due), h = state.headers || [[], []];
    var top = String((h[0] && h[0][i]) || "").trim(), sub = String((h[1] && h[1][i]) || "").trim();
    return sub ? (top && /^phase/i.test(top) ? top + " – " + sub : sub || top) : top;
  }
  function recency(r) { var v = r.values; return v[colIdx("S")] || v[colIdx("R")] || 0; }

  function banner(kind, html) { var b = $("banner"); b.className = "banner " + kind; b.innerHTML = html; }
  function busy(msg) { banner("info", '<span class="spin"></span>' + esc(msg)); }
  function fail(e) {
    var m = e && e.message ? e.message : String(e);
    if (e && e.status === 403) m = "You don't have permission to open the customer spreadsheet. (" + m + ")";
    if (e && e.status === 404) m = "Couldn't find the customer spreadsheet. Check the link in config.js. (" + m + ")";
    banner("bad", "Something went wrong<small>" + esc(m) + "</small>");
  }

  // ---------- email ----------
  function readEmail() {
    return new Promise(function (resolve) {
      var item = Office.context.mailbox.item;
      if (!item) return resolve(null);
      item.body.getAsync(Office.CoercionType.Text, function (r) {
        var text = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "";
        resolve(TCCParse.parseOrderEmail(text, item.subject));
      });
    });
  }

  function showOrder(p) {
    if (!p || !p.found) { $("order").textContent = "No customer details found in this email – search manually below."; return; }
    var bits = [];
    if (p.orderNo) bits.push("Order #" + p.orderNo);
    if (p.orderType) bits.push(p.orderType.replace(/^New /, ""));
    $("order").innerHTML = esc(bits.join(" · ")) + "<br><b>" + esc(p.nameTidy || p.name) + "</b> · " + esc(p.email || "no email");
  }

  // ---------- lookup ----------
  async function lookup(query) {
    state.query = query;
    $("editor").classList.add("hidden");
    $("candidates").innerHTML = "";
    $("qName").value = query.name || ""; $("qEmail").value = query.email || "";
    busy("Looking up the customer spreadsheet…");
    try {
      var both = await Promise.all([TCCGraph.readIndex(), TCCGraph.readHeaders()]);
      state.index = both[0]; state.headers = both[1];
      var q = { name: query.name, email: cfg.placeholderEmails.indexOf(String(query.email || "").toLowerCase()) >= 0 ? "" : query.email };
      var rows = state.index.rows.map(function (r) {
        return cfg.placeholderEmails.indexOf(r.email.trim().toLowerCase()) >= 0 ? { row: r.row, name: r.name, email: "" } : r;
      });
      var res = TCCMatch.findMatches(rows, q, { max: 10 });
      if (res.status === "none") return showNone();
      busy("Loading " + res.candidates.length + " row" + (res.candidates.length > 1 ? "s" : "") + "…");
      var full = await TCCGraph.readRows(res.candidates.map(function (c) { return c.row; }));
      Object.assign(state.rows, full);
      res.candidates.forEach(function (c) { c.data = full[c.row]; });
      if (res.status === "exact") return openEditor(res.candidates[0].row, { kind: "ok", candidate: res.candidates[0], hidden: res.hiddenLookalikes });
      showCandidates(res);
    } catch (e) { fail(e); }
  }

  function showNone() {
    banner("bad", "Customer not found<small>No row in the spreadsheet matches this name or email.</small>");
    var c = $("candidates");
    c.innerHTML = '<button id="addNew">Add as a new customer</button>';
    $("addNew").onclick = function () { openNewForm(); };
  }

  function showCandidates(res) {
    var c = res.candidates;
    // most recent first
    c.sort(function (a, b) { return recency(b.data) - recency(a.data) || b.row - a.row; });
    var emails = {};
    c.forEach(function (x) { var e = x.email.trim().toLowerCase(); if (e) emails[e] = (emails[e] || 0) + 1; });
    if (res.status === "multiple") banner("warn", "Several possible matches<small>Pick the right customer. The most recent is at the top.</small>");
    else banner("warn", "Close match – please check<small>" + esc(c[0].reason) + ".</small>");
    var seen = {};
    $("candidates").innerHTML = c.map(function (x, i) {
      var v = x.data.values, e = x.email.trim().toLowerCase();
      var old = e && emails[e] > 1 && seen[e]; if (e) seen[e] = true;
      var when = v[colIdx("S")] ? "Last order " + fmtDate(v[colIdx("S")]) : (v[colIdx("R")] ? "1st order " + fmtDate(v[colIdx("R")]) : "");
      return '<div class="cand" tabindex="0" data-row="' + x.row + '">' +
        '<div class="n">' + esc(x.name) + '<span class="tag">Row ' + x.row + '</span>' +
        (i === 0 && res.status === "multiple" ? '<span class="tag">Most recent</span>' : "") +
        (old ? '<span class="tag old">Possibly old row</span>' : "") + "</div>" +
        '<div class="e">' + esc(x.email || "no email") + "</div>" +
        '<div class="m">' + esc([v[colIdx("U")], when].filter(Boolean).join(" · ")) + "</div>" +
        '<div class="why' + (x.level === "exact" ? "" : " warn") + '">' + esc(x.reason) + "</div></div>";
    }).join("") + '<button class="ghost" id="addNew">None of these – add new customer</button>';
    Array.prototype.forEach.call(document.querySelectorAll(".cand"), function (el) {
      var go = function () {
        var x = c.filter(function (y) { return y.row === +el.dataset.row; })[0];
        openEditor(x.row, { kind: x.level === "exact" ? "ok" : "warn", candidate: x, back: res });
      };
      el.onclick = go; el.onkeydown = function (ev) { if (ev.key === "Enter") go(); };
    });
    $("addNew").onclick = function () { openNewForm(); };
  }

  // ---------- editor ----------
  function fieldHtml(f, val, text) {
    var id = "f_" + f.col, v;
    if (f.type === "tick") {
      var extra = isPlainTick(val) ? "" : '<span class="extra">"' + esc(val) + '"</span>';
      return '<div class="field"><label class="tick"><input type="checkbox" id="' + id + '"' + (isTicked(val) ? " checked" : "") +
        '><span class="lbl">' + esc(f.label) + "</span>" + extra + "</label></div>";
    }
    var input;
    if (f.type === "date") { v = typeof val === "number" ? serialToISO(val) : ""; input = '<input type="date" id="' + id + '" value="' + v + '">' + (typeof val === "string" && val ? '<span class="extra">Currently: "' + esc(val) + '"</span>' : ""); }
    else if (f.type === "select") {
      var opts = f.options.slice(); if (val && opts.indexOf(val) < 0) opts.unshift(val);
      input = '<select id="' + id + '"><option value=""></option>' + opts.map(function (o) { return "<option" + (o === val ? " selected" : "") + ">" + esc(o) + "</option>"; }).join("") + "</select>";
    }
    else if (f.type === "longtext") input = '<textarea id="' + id + '">' + esc(val) + "</textarea>";
    else input = '<input id="' + id + '" value="' + esc(text != null && text !== "" ? text : val) + '">';
    return '<div class="field"><label for="' + id + '">' + esc(f.label) + "</label>" + input + "</div>";
  }

  function readField(f) {
    var el = $("f_" + f.col);
    if (f.type === "tick") return el.checked;
    if (f.type === "date") return el.value ? isoToSerial(el.value) : "";
    return el.value;
  }

  function sameValue(f, orig, now) {
    if (f.type === "tick") return isTicked(orig) === now;
    if (f.type === "date") return (typeof orig === "number" ? Math.round(orig) : "") === now || (now === "" && (orig === "" || orig == null));
    return String(orig == null ? "" : orig) === String(now);
  }

  function editableFields() {
    var list = cfg.fields.slice();
    cfg.phases.forEach(function (p) { list.push({ col: p.done, label: phaseLabel(p) + " – done", type: "tick", phase: p }); });
    return list;
  }

  function openEditor(row, ctx) {
    state.current = { row: row, ctx: ctx };
    var d = state.rows[row], v = d.values, t = d.text;
    var c = ctx.candidate;
    if (ctx.kind === "ok") banner("ok", "Customer found<small>Row " + row + " · " + esc(c.reason) + (ctx.hidden ? " · " + ctx.hidden + " similar-looking row(s) ignored" : "") + "</small>");
    else banner("warn", "Check this is the right customer<small>Row " + row + " · " + esc(c.reason) + "</small>");
    $("candidates").innerHTML = ctx.back ? '<button class="back" id="backBtn">← Back to all matches</button>' : "";
    if (ctx.back) $("backBtn").onclick = function () { $("editor").classList.add("hidden"); showCandidates(ctx.back); };

    // Due now
    var today = todaySerial(), due = [], older = [], win = cfg.dueWindowDays || 60;
    cfg.phases.forEach(function (p) {
      var dv = v[colIdx(p.due)], done = v[colIdx(p.done)], can = p.canSend ? String(v[colIdx(p.canSend)] || "").toLowerCase() : "";
      if (isTicked(done) || can === "no") return;
      if (typeof dv === "number" && dv > 0 && dv <= today) (dv >= today - win ? due : older).push({ p: p, when: "due " + fmtDate(dv) });
      else if (String(dv).toLowerCase() === "yes") due.push({ p: p, when: "due now" });
    });
    var li = function (x) {
      return '<li><label class="tick"><input type="checkbox" data-done="' + x.p.done + '"><span class="lbl">' + esc(phaseLabel(x.p)) +
        ' <span class="when">' + esc(x.when) + "</span></span></label></li>";
    };
    var dueHtml = (due.length ? '<ul class="due">' + due.map(li).join("") + "</ul>"
                              : '<div class="empty">Nothing due for this customer.</div>') +
      (older.length ? '<details style="margin-top:6px"><summary>' + older.length + " older, never ticked (more than " + win +
        ' days ago)</summary><ul class="due">' + older.map(li).join("") + "</ul></details>" : "");

    var main = cfg.fields.map(function (f) { return fieldHtml(f, v[colIdx(f.col)], t[colIdx(f.col)]); }).join("");
    var phaseTicks = cfg.phases.map(function (p) {
      var f = { col: p.done, label: phaseLabel(p) + " – done", type: "tick" };
      var dv = t[colIdx(p.due)];
      return fieldHtml(f, v[colIdx(p.done)]).replace("</span></label>", (dv ? ' <span class="when">(' + esc(dv) + ")</span>" : "") + "</span></label>");
    }).join("");

    $("editor").innerHTML =
      '<div class="card"><h3>Due now</h3>' + dueHtml + "</div>" +
      '<div class="card"><h3>Customer details</h3>' + main + "</div>" +
      '<details class="card"><summary><b>All phases</b></summary><div style="margin-top:8px">' + phaseTicks + "</div></details>" +
      '<div class="actions"><button id="saveBtn" disabled>Save changes</button><button class="ghost" id="reloadBtn">Reload</button></div>';
    $("editor").classList.remove("hidden");

    // "Due now" ticks mirror the matching "All phases" tick
    Array.prototype.forEach.call(document.querySelectorAll("[data-done]"), function (el) {
      el.onchange = function () { $("f_" + el.dataset.done).checked = el.checked; markChanged(); };
    });
    $("editor").oninput = markChanged; $("editor").onchange = markChanged;
    $("saveBtn").onclick = save;
    $("reloadBtn").onclick = function () { lookup(state.query); };
  }

  function pendingChanges() {
    var cur = state.current; if (!cur) return [];
    var v = state.rows[cur.row].values, t = state.rows[cur.row].text, out = [];
    editableFields().forEach(function (f) {
      var el = $("f_" + f.col); if (!el) return;
      var i = colIdx(f.col);
      // plain text boxes show the cell as displayed (keeps leading zeros on phone numbers)
      var orig = f.type === "text" && t[i] !== "" && t[i] != null ? t[i] : v[i], now = readField(f);
      var same = sameValue(f, orig, now);
      el.closest(".field").classList.toggle("changed", !same);
      if (!same) out.push({ f: f, orig: orig, now: now });
    });
    return out;
  }
  // Stop Excel turning "07921..." into a number and dropping the 0
  function keepText(f, v) { return (f.type === "text" && /^[0+]\d{6,}$/.test(String(v).replace(/\s/g, ""))) ? "'" + v : v; }
  function markChanged() {
    if (!$("saveBtn")) return; var n = pendingChanges().length; var b = $("saveBtn"); b.disabled = !n; b.textContent = n ? "Save " + n + " change" + (n > 1 ? "s" : "") : "Save changes"; }

  async function save() {
    var cur = state.current, changes = pendingChanges();
    if (!changes.length) return;
    $("saveBtn").disabled = true;
    busy("Checking nobody else has changed this row…");
    try {
      // Concurrency check: re-read the row and make sure the cells we're changing still hold what we loaded
      var fresh = (await TCCGraph.readRows([cur.row]))[cur.row];
      var loaded = state.rows[cur.row].values;
      var clash = changes.filter(function (c) { var i = colIdx(c.f.col); return String(fresh.values[i]) !== String(loaded[i]); });
      var idClash = String(fresh.values[0]) !== String(loaded[0]) || String(fresh.values[1]) !== String(loaded[1]);
      if (clash.length || idClash) {
        banner("bad", "Someone else just changed this row<small>Nothing was saved. Press Reload to see the latest, then make your change again." +
          (clash.length ? " Changed: " + esc(clash.map(function (c) { return c.f.label; }).join(", ")) : "") + "</small>");
        $("saveBtn").disabled = false; return;
      }
      busy("Saving…");
      await TCCGraph.writeCells(changes.map(function (c) {
        var val = c.f.type === "tick" ? (c.now ? cfg.tickValue : "") : keepText(c.f, c.now);
        return { col: c.f.col, row: cur.row, value: val, tick: c.f.type === "tick" };
      }));
      var after = (await TCCGraph.readRows([cur.row]))[cur.row];
      state.rows[cur.row] = after;
      var ctx = cur.ctx; ctx.kind = "ok";
      openEditor(cur.row, ctx);
      banner("ok", "Saved ✓<small>" + changes.length + " change" + (changes.length > 1 ? "s" : "") + " written to row " + cur.row + ".</small>");
    } catch (e) { fail(e); $("saveBtn").disabled = false; }
  }

  // ---------- new customer ----------
  function openNewForm() {
    var p = state.parsed || {}, q = state.query || {};
    var name = (q.name === p.name ? p.nameTidy : q.name) || "";
    var pre = { A: name, B: q.email || "", N: p.phone || "", U: TCCParse.guessFrequency(p), R: todaySerial() };
    var newFields = cfg.fields.filter(function (f) { return f.type !== "tick"; });
    state.current = null;
    $("editor").oninput = null; $("editor").onchange = null;
    banner("info", "New customer<small>Check the details, then add. It goes in the next empty row.</small>");
    $("candidates").innerHTML = "";
    $("editor").innerHTML = '<div class="card"><h3>Add new customer</h3>' +
      newFields.map(function (f) { return fieldHtml(f, pre[f.col] != null ? pre[f.col] : ""); }).join("") +
      (pre.U ? '<div class="empty">Frequency guessed from the order – please check.</div>' : "") +
      '</div><div class="actions"><button id="addBtn">Add customer</button><button class="ghost" id="cancelBtn">Cancel</button></div>';
    $("editor").classList.remove("hidden");
    $("cancelBtn").onclick = function () { lookup(state.query); };
    $("addBtn").onclick = function () { addCustomer(newFields); };
  }

  async function addCustomer(fields) {
    var vals = {}; fields.forEach(function (f) { vals[f.col] = readField(f); });
    if (!String(vals.A).trim()) { banner("bad", "A name is needed"); return; }
    $("addBtn").disabled = true;
    busy("Checking the spreadsheet again…");
    try {
      var idx = await TCCGraph.readIndex();
      var e = String(vals.B).trim().toLowerCase();
      var dupe = e && TCCMatch.isUsableEmail(e) && idx.rows.filter(function (r) { return r.email.trim().toLowerCase() === e; })[0];
      if (dupe) { banner("warn", "That email is already on row " + dupe.row + "<small>Nothing was added. Search again to open it.</small>"); $("addBtn").disabled = false; return; }
      var row = idx.nextEmptyRow;
      if (row > cfg.lastFormulaRow) { banner("bad", "The sheet's formulas only go down to row " + cfg.lastFormulaRow + "<small>Fill the formulas further down in Excel first.</small>"); return; }
      busy("Adding to row " + row + "…");
      await TCCGraph.writeCells(fields.filter(function (f) { return vals[f.col] !== ""; }).map(function (f) {
        return { col: f.col, row: row, value: keepText(f, vals[f.col]) };
      }));
      var full = await TCCGraph.readRows([row]);
      Object.assign(state.rows, full);
      openEditor(row, { kind: "ok", candidate: { reason: "New customer added", row: row } });
      banner("ok", "Customer added ✓<small>Row " + row + ".</small>");
    } catch (err) { fail(err); $("addBtn").disabled = false; }
  }

  // ---------- start ----------
  async function run() {
    state.parsed = await readEmail();
    showOrder(state.parsed);
    if (state.parsed && state.parsed.found) lookup({ name: state.parsed.name, email: state.parsed.email });
    else { banner("info", "Open the search above to look up a customer."); $("searchBox").open = true; }
  }

  Office.onReady(function () {
    $("searchForm").onsubmit = function (ev) { ev.preventDefault(); lookup({ name: $("qName").value.trim(), email: $("qEmail").value.trim() }); };
    $("foot").textContent = "Reads and writes: " + (cfg.workbookShareUrl.indexOf("PASTE") === 0 ? "(spreadsheet link not set)" : "customer spreadsheet") + ".";
    if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
      // pinned panel: refresh when a different email is selected
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { state.rows = {}; run(); });
    }
    run();
  });
})();
