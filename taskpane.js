/* Two Chimps customer lookup - side panel logic.
   Columns are found by header name (see schema.js), never by fixed letter. */
(function () {
  "use strict";
  var cfg = window.TCC_CONFIG;
  var state = { parsed: null, index: null, schema: null, size: null, rows: {}, current: null, query: null };
  var $ = function (id) { return document.getElementById(id); };
  var CACHE_KEY = "tccSchemaV2";

  // ---------- helpers ----------
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function serialToISO(v) {
    if (typeof v !== "number" || v < 1) return "";
    return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10);
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
  function isPlainTick(v) { var s = String(v == null ? "" : v).trim(); return s === "" || /^(a{1,3}|ü{1,3}|A)$/.test(s); }
  function core(key) { return state.schema && state.schema.byKey[key]; }
  function cell(data, key) { var f = core(key); return f ? data.values[f.idx] : ""; }
  // shown in the panel unless the column is hidden in Excel (name + email always shown)
  function shown(f) { return !(state.hidden && state.hidden[f.idx]) || f.key === "name" || f.key === "email"; }
  function phaseShown(p) { return p.done && !(state.hidden && state.hidden[p.done.idx]); }
  function recency(d) { return cell(d, "lastOrder") || cell(d, "posted") || 0; }

  function banner(kind, html) { var b = $("banner"); b.className = "banner " + kind; b.innerHTML = html; }
  function busy(msg) { banner("info", '<span class="spin"></span>' + esc(msg)); }
  function fail(e) {
    var m = e && e.message ? e.message : String(e);
    if (e && e.status === 403) m = "You don't have permission to open the customer spreadsheet. (" + m + ")";
    if (e && e.status === 404) m = "Couldn't find the customer spreadsheet. Check config.js. (" + m + ")";
    banner("bad", "Something went wrong<small>" + esc(m) + "</small>");
  }

  // ---------- email ----------
  function readEmail() {
    return new Promise(function (resolve) {
      var item = Office.context.mailbox.item;
      if (!item) return resolve(null);
      // HTML keeps the billing address on separate lines; plain text runs it all together
      item.body.getAsync(Office.CoercionType.Html, function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
          var p = TCCParse.parseOrderHtml(r.value, item.subject);
          if (p.found) return resolve(p);
        }
        item.body.getAsync(Office.CoercionType.Text, function (t) {
          resolve(TCCParse.parseOrderEmail(t.status === Office.AsyncResultStatus.Succeeded ? t.value : "", item.subject));
        });
      });
    });
  }

  function showOrder(p) {
    if (!p || !p.found) { $("order").textContent = "No customer details found in this email – search manually below."; return; }
    var bits = [];
    if (p.orderNo) bits.push("Order #" + p.orderNo);
    if (p.orderType) bits.push(p.orderType.replace(/^New /, ""));
    $("order").innerHTML = esc(bits.join(" · ")) + "<br><b>" + esc(p.nameTidy || p.name) + "</b> · " + esc(p.email || "no email") +
      (p.orderNo && cfg.clickDropUrl ? '<div class="cd"><button class="ghost small" id="cdBtn">Open in Click &amp; Drop</button>' +
        '<span id="cdMsg" class="cd-msg"></span></div>' : "");
    if (p.orderNo && cfg.clickDropUrl) $("cdBtn").onclick = function () { openClickDrop(p.orderNo); };
  }

  // ---------- Royal Mail Click & Drop ----------
  // Click & Drop's own page addresses use its internal IDs, so we copy the website order number
  // (searching "#452225" finds the order) and open Click & Drop, ready to paste into its search.
  function copyText(text) {
    function fallback() {
      var ta = document.createElement("textarea");
      ta.value = text; ta.setAttribute("readonly", ""); ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select();
      var ok = false; try { ok = document.execCommand("copy"); } catch (e) {}
      document.body.removeChild(ta);
      return Promise.resolve(ok);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text).then(function () { return true; }, fallback);
    }
    return fallback();
  }

  function openClickDrop(orderNo) {
    var ref = (cfg.clickDropSearchPrefix || "") + orderNo;
    copyText(ref).then(function (ok) {
      $("cdMsg").innerHTML = ok
        ? "Copied <b>" + esc(ref) + "</b>. In Click &amp; Drop, click Search, press Ctrl+V, then Enter."
        : "Search Click &amp; Drop for <b>" + esc(ref) + "</b>.";
    });
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) Office.context.ui.openBrowserWindow(cfg.clickDropUrl);
      else window.open(cfg.clickDropUrl, "_blank");
    } catch (e) { window.open(cfg.clickDropUrl, "_blank"); }
  }

  // ---------- spreadsheet layout ----------
  function loadCache(sig) {
    if (state.cache && state.cache.sig === sig) return state.cache.schema;
    try {
      var c = JSON.parse(sessionStorage.getItem(CACHE_KEY) || "null");
      if (c && c.sig === sig) { state.cache = c; return c.schema; }
    } catch (e) {}
    return null;
  }
  function saveCache(sig, schema) {
    state.cache = { sig: sig, schema: schema };
    try { sessionStorage.setItem(CACHE_KEY, JSON.stringify(state.cache)); } catch (e) {}
  }

  /* Reads headers + names/emails, and works out the columns (cached until the headers change) */
  async function loadSheet() {
    var size = await TCCGraph.sheetSize();
    var headers = await TCCGraph.readHeaders(size.lastCol);
    var pre = TCCSchema.buildSchema({ headers: headers }, cfg);
    var nameF = pre.byKey.name, emailF = pre.byKey.email;
    if (!nameF || !emailF) throw new Error("Couldn't find the '" + (!nameF ? "Name" : "Email address") + "' column in the header rows.");
    var got = await Promise.all([
      TCCGraph.readIndex(nameF.letter, emailF.letter, size.lastRow),
      TCCGraph.readHidden(pre.width)            // columns hidden in Excel are left out of the panel
    ]);
    var index = got[0];
    state.hidden = got[1];
    var schema = loadCache(pre.signature);
    if (!schema) {
      busy("Reading the spreadsheet's columns…");
      var lastData = index.nextEmptyRow - 1;
      var inputs = await TCCGraph.readSchemaInputs(size.lastCol, lastData);
      inputs.headers = headers;
      schema = TCCSchema.buildSchema(inputs, cfg);
      // brand-new empty columns: look at the font (Webdings = tick column)
      var empty = schema.others.filter(function (o) {
        return !inputs.samples.some(function (r) { return r[o.idx] !== "" && r[o.idx] != null; });
      });
      if (empty.length) {
        inputs.fonts = await TCCGraph.readFonts(empty, lastData);
        schema = TCCSchema.buildSchema(inputs, cfg);
      }
      var below = inputs.belowRow || [];
      schema.formulasBelow = below.some(function (v) { return typeof v === "string" && v.charAt(0) === "="; });
      saveCache(pre.signature, schema);
    }
    state.size = size; state.headers = headers; state.schema = schema; state.index = index;
    $("foot").textContent = "Reads and writes: " + (window.TCCWorkbookName || "customer spreadsheet") + ".";
  }

  // ---------- lookup ----------
  async function lookup(query) {
    state.query = query;
    $("editor").classList.add("hidden");
    $("candidates").innerHTML = "";
    $("qName").value = query.name || ""; $("qEmail").value = query.email || "";
    busy("Looking up the customer spreadsheet…");
    try {
      await loadSheet();
      var ph = cfg.placeholderEmails.map(function (e) { return e.toLowerCase(); });
      var q = { name: query.name, email: ph.indexOf(String(query.email || "").trim().toLowerCase()) >= 0 ? "" : query.email };
      var rows = state.index.rows.map(function (r) {
        return ph.indexOf(r.email.trim().toLowerCase()) >= 0 ? { row: r.row, name: r.name, email: "" } : r;
      });
      var res = TCCMatch.findMatches(rows, q, { max: 10 });
      if (res.status === "none") return showNone();
      busy("Loading " + res.candidates.length + " row" + (res.candidates.length > 1 ? "s" : "") + "…");
      var full = await TCCGraph.readRows(res.candidates.map(function (c) { return c.row; }), state.size.lastCol);
      Object.assign(state.rows, full);
      res.candidates.forEach(function (c) { c.data = full[c.row]; });
      if (res.status === "exact") return openEditor(res.candidates[0].row, { kind: "ok", candidate: res.candidates[0], hidden: res.hiddenLookalikes });
      showCandidates(res);
    } catch (e) { fail(e); }
  }

  function showNone() {
    banner("bad", "Customer not found<small>No row in the spreadsheet matches this name or email.</small>");
    $("candidates").innerHTML = '<button id="addNew">Add as a new customer</button>';
    $("addNew").onclick = function () { openNewForm(); };
  }

  function showCandidates(res) {
    var c = res.candidates;
    c.sort(function (a, b) { return recency(b.data) - recency(a.data) || b.row - a.row; });   // most recent first
    var emails = {};
    c.forEach(function (x) { var e = x.email.trim().toLowerCase(); if (e) emails[e] = (emails[e] || 0) + 1; });
    if (res.status === "multiple") banner("warn", "Several possible matches<small>Pick the right customer. The most recent is at the top.</small>");
    else banner("warn", "Close match – please check<small>" + esc(c[0].reason) + ".</small>");
    var seen = {};
    $("candidates").innerHTML = c.map(function (x, i) {
      var d = x.data, e = x.email.trim().toLowerCase();
      var old = e && emails[e] > 1 && seen[e]; if (e) seen[e] = true;
      var last = cell(d, "lastOrder"), first = cell(d, "posted");
      var when = last ? "Last order " + fmtDate(last) : (first ? "1st order " + fmtDate(first) : "");
      return '<div class="cand" tabindex="0" data-row="' + x.row + '">' +
        '<div class="n">' + esc(x.name) + '<span class="tag">Row ' + x.row + '</span>' +
        (i === 0 && res.status === "multiple" ? '<span class="tag">Most recent</span>' : "") +
        (old ? '<span class="tag old">Possibly old row</span>' : "") + "</div>" +
        '<div class="e">' + esc(x.email || "no email") + "</div>" +
        '<div class="m">' + esc([cell(d, "frequency"), when].filter(Boolean).join(" · ")) + "</div>" +
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

  // ---------- fields ----------
  /* f = {letter, label, type, options}. A tick cell holding real text (e.g. "PAYPAL CANT AMEND TIME")
     is shown as a text box so it can't be wiped by unticking. */
  function fieldHtml(f, val, text, extraLabel) {
    var id = "f_" + f.letter, v;
    var kind = f.type === "tick" && !isPlainTick(val) ? "text" : f.type;
    var lbl = esc(f.label) + (extraLabel ? ' <span class="when">' + esc(extraLabel) + "</span>" : "");
    if (kind === "tick") {
      return '<div class="field"><label class="tick"><input type="checkbox" data-kind="tick" id="' + id + '"' +
        (isTicked(val) ? " checked" : "") + '><span class="lbl">' + lbl + "</span></label></div>";
    }
    var input;
    if (kind === "date") {
      v = typeof val === "number" ? serialToISO(val) : "";
      input = '<input type="date" data-kind="date" id="' + id + '" value="' + v + '">' +
        (typeof val === "string" && val ? '<span class="extra">Currently: "' + esc(val) + '"</span>' : "");
    } else if (kind === "select") {
      var opts = (f.options || []).slice(); if (val && opts.indexOf(val) < 0) opts.unshift(val);
      input = '<select data-kind="select" id="' + id + '"><option value=""></option>' + opts.map(function (o) {
        return "<option" + (o === val ? " selected" : "") + ">" + esc(o) + "</option>"; }).join("") + "</select>";
    } else if (kind === "longtext") {
      input = '<textarea data-kind="longtext" id="' + id + '">' + esc(val) + "</textarea>";
    } else {
      input = '<input data-kind="text" id="' + id + '" value="' + esc(text != null && text !== "" ? text : val) + '">';
    }
    return '<div class="field"><label for="' + id + '">' + lbl + "</label>" + input + "</div>";
  }

  function readField(f) {
    var el = $("f_" + f.letter), kind = el.dataset.kind;
    if (kind === "tick") return el.checked;
    if (kind === "date") return el.value ? isoToSerial(el.value) : "";
    return el.value;
  }

  function sameValue(kind, orig, now) {
    if (kind === "tick") return isTicked(orig) === now;
    if (kind === "date") return (typeof orig === "number" ? Math.round(orig) : "") === now || (now === "" && (orig === "" || orig == null));
    return String(orig == null ? "" : orig) === String(now);
  }

  function allFields() {
    var s = state.schema, list = s.core.concat(s.others).filter(shown);
    s.phases.forEach(function (p) { if (phaseShown(p)) list.push({ idx: p.done.idx, letter: p.done.letter, label: p.label + " – done", type: "tick" }); });
    return list;
  }

  // ---------- editor ----------
  function openEditor(row, ctx) {
    state.current = { row: row, ctx: ctx };
    var s = state.schema, d = state.rows[row], v = d.values, t = d.text, c = ctx.candidate;
    if (ctx.kind === "ok") banner("ok", "Customer found<small>Row " + row + " · " + esc(c.reason) + (ctx.hidden ? " · " + ctx.hidden + " similar-looking row(s) ignored" : "") + "</small>");
    else banner("warn", "Check this is the right customer<small>Row " + row + " · " + esc(c.reason) + "</small>");
    $("candidates").innerHTML = ctx.back ? '<button class="back" id="backBtn">← Back to all matches</button>' : "";
    if (ctx.back) $("backBtn").onclick = function () { $("editor").classList.add("hidden"); showCandidates(ctx.back); };

    // Due now
    var today = todaySerial(), due = [], older = [], win = cfg.dueWindowDays || 60, lastDate = 0;
    s.phases.forEach(function (p) {
      var dv = v[p.due.idx];
      var dated = typeof dv === "number" && dv > 0;
      if (!phaseShown(p)) { if (dated) lastDate = dv; return; }
      var done = v[p.done.idx], can = p.canSend ? String(v[p.canSend.idx] || "").toLowerCase() : "";
      if (!isTicked(done) && can !== "no") {
        if (dated && dv <= today) (dv >= today - win ? due : older).push({ p: p, when: "due " + fmtDate(dv) });
        // "yes" phases have no date of their own - judge them by the dated phase before them
        else if (String(dv).toLowerCase() === "yes") {
          if (lastDate && lastDate < today - win) older.push({ p: p, when: "since " + fmtDate(lastDate) });
          else due.push({ p: p, when: "due now" });
        }
      }
      if (dated) lastDate = dv;
    });
    var li = function (x) {
      return '<li><label class="tick"><input type="checkbox" data-done="' + x.p.done.letter + '"><span class="lbl">' + esc(x.p.label) +
        ' <span class="when">' + esc(x.when) + "</span></span></label></li>";
    };
    var dueHtml = (due.length ? '<ul class="due">' + due.map(li).join("") + "</ul>"
                              : '<div class="empty">Nothing due for this customer.</div>') +
      (older.length ? '<details style="margin-top:6px"><summary>' + older.length + " older, never ticked (more than " + win +
        ' days ago)</summary><ul class="due">' + older.map(li).join("") + "</ul></details>" : "");

    var others = s.others.filter(shown);
    var main = s.core.filter(shown).map(function (f) { return fieldHtml(f, v[f.idx], t[f.idx]); }).join("");
    var more = others.map(function (f) { return fieldHtml(f, v[f.idx], t[f.idx]); }).join("");
    var phaseTicks = s.phases.filter(phaseShown).map(function (p) {
      var f = { letter: p.done.letter, label: p.label + " – done", type: "tick" };
      var dv = t[p.due.idx];
      return fieldHtml(f, v[p.done.idx], t[p.done.idx], dv ? "(" + dv + ")" : "");
    }).join("");

    $("editor").innerHTML =
      '<div class="card"><h3>Due now</h3>' + dueHtml + "</div>" +
      '<div class="card"><h3>Customer details</h3>' + main + "</div>" +
      (more ? '<details class="card" open><summary><b>More details</b> <span class="when">(' + others.length + ")</span></summary>" +
              '<div style="margin-top:8px">' + more + "</div></details>" : "") +
      '<details class="card"><summary><b>All phases</b></summary><div style="margin-top:8px">' + phaseTicks + "</div></details>" +
      '<div class="actions"><button id="saveBtn" disabled>Save changes</button><button class="ghost" id="reloadBtn">Reload</button></div>';
    $("editor").classList.remove("hidden");

    // "Due now" ticks mirror the matching "All phases" tick
    Array.prototype.forEach.call(document.querySelectorAll("[data-done]"), function (el) {
      el.onchange = function () { var t2 = $("f_" + el.dataset.done); if (t2 && t2.type === "checkbox") t2.checked = el.checked; markChanged(); };
    });
    $("editor").oninput = markChanged; $("editor").onchange = markChanged;
    $("saveBtn").onclick = save;
    $("reloadBtn").onclick = function () { lookup(state.query); };
  }

  function pendingChanges() {
    var cur = state.current; if (!cur) return [];
    var v = state.rows[cur.row].values, t = state.rows[cur.row].text, out = [];
    allFields().forEach(function (f) {
      var el = $("f_" + f.letter); if (!el) return;
      var kind = el.dataset.kind;
      // text boxes compare with the cell as displayed (keeps leading zeros on phone numbers)
      var orig = kind === "text" && t[f.idx] !== "" && t[f.idx] != null ? t[f.idx] : v[f.idx], now = readField(f);
      var same = sameValue(kind, orig, now);
      el.closest(".field").classList.toggle("changed", !same);
      if (!same) out.push({ f: f, kind: kind, now: now });
    });
    return out;
  }
  // Stop Excel turning "07921..." into a number and dropping the 0
  // Always save typed text as text: keeps leading zeros on phone numbers, and stops anything
  // starting with = + - @ (e.g. a customer name like "=HYPERLINK(...)") becoming a live formula.
  function keepText(kind, v) {
    if (kind === "tick" || kind === "date" || typeof v !== "string") return v;
    var s = v.replace(/\s/g, "");
    if (/^[=+\-@]/.test(v) || (kind === "text" && /^0\d{6,}$/.test(s))) return "'" + v;
    return v;
  }
  function markChanged() {
    if (!$("saveBtn")) return;
    var n = pendingChanges().length, b = $("saveBtn");
    b.disabled = !n; b.textContent = n ? "Save " + n + " change" + (n > 1 ? "s" : "") : "Save changes";
  }

  async function headersUnchanged() {
    var fresh = await TCCGraph.readHeaders(state.size.lastCol);
    return JSON.stringify(fresh) === JSON.stringify(state.headers);
  }

  async function save() {
    var cur = state.current, changes = pendingChanges();
    if (!changes.length) return;
    $("saveBtn").disabled = true;
    busy("Checking nobody else has changed this row…");
    try {
      if (!(await headersUnchanged())) {
        banner("bad", "The spreadsheet's columns have changed<small>Nothing was saved. Press Reload, then make your change again.</small>");
        $("saveBtn").disabled = false; return;
      }
      var fresh = (await TCCGraph.readRows([cur.row], state.size.lastCol))[cur.row];
      var loaded = state.rows[cur.row].values;
      var n = core("name").idx, e = core("email").idx;
      var clash = changes.filter(function (c) { return String(fresh.values[c.f.idx]) !== String(loaded[c.f.idx]); });
      var idClash = String(fresh.values[n]) !== String(loaded[n]) || String(fresh.values[e]) !== String(loaded[e]);
      if (clash.length || idClash) {
        banner("bad", "Someone else just changed this row<small>Nothing was saved. Press Reload to see the latest, then make your change again." +
          (clash.length ? " Changed: " + esc(clash.map(function (c) { return c.f.label; }).join(", ")) : "") + "</small>");
        $("saveBtn").disabled = false; return;
      }
      busy("Saving…");
      await TCCGraph.writeCells(changes.map(function (c) {
        var val = c.kind === "tick" ? (c.now ? cfg.tickValue : "") : keepText(c.kind, c.now);
        return { col: c.f.letter, row: cur.row, value: val, tick: c.kind === "tick" };
      }));
      state.rows[cur.row] = (await TCCGraph.readRows([cur.row], state.size.lastCol))[cur.row];
      var ctx = cur.ctx; ctx.kind = "ok";
      openEditor(cur.row, ctx);
      banner("ok", "Saved ✓<small>" + changes.length + " change" + (changes.length > 1 ? "s" : "") + " written to row " + cur.row + ".</small>");
    } catch (err) { fail(err); $("saveBtn").disabled = false; }
  }

  // ---------- new customer ----------
  function openNewForm() {
    var p = state.parsed || {}, q = state.query || {}, s = state.schema;
    var name = (q.name === p.name ? p.nameTidy : q.name) || "";
    var pre = { name: name, email: q.email || "", phone: p.phone || "", frequency: TCCParse.guessFrequency(p), posted: todaySerial() };
    var fields = s.core.filter(function (f) { return f.type !== "tick" && shown(f); });
    state.current = null;
    $("editor").oninput = null; $("editor").onchange = null;
    banner("info", "New customer<small>Check the details, then add. It goes in the next empty row.</small>");
    $("candidates").innerHTML = "";
    $("editor").innerHTML = '<div class="card"><h3>Add new customer</h3>' +
      fields.map(function (f) { return fieldHtml(f, pre[f.key] != null ? pre[f.key] : ""); }).join("") +
      (pre.frequency ? '<div class="empty">Frequency guessed from the order – please check.</div>' : "") +
      (s.formulasBelow === false ? '<div class="empty" style="color:var(--warn)">The formula columns don\'t seem to be filled down past the last customer – fill them down in Excel so the phases work for new rows.</div>' : "") +
      '</div><div class="actions"><button id="addBtn">Add customer</button><button class="ghost" id="cancelBtn">Cancel</button></div>';
    $("editor").classList.remove("hidden");
    $("cancelBtn").onclick = function () { lookup(state.query); };
    $("addBtn").onclick = function () { addCustomer(fields); };
  }

  async function addCustomer(fields) {
    var vals = {}; fields.forEach(function (f) { vals[f.key] = { f: f, kind: $("f_" + f.letter).dataset.kind, v: readField(f) }; });
    if (!String(vals.name.v).trim()) { banner("bad", "A name is needed"); return; }
    $("addBtn").disabled = true;
    busy("Checking the spreadsheet again…");
    try {
      if (!(await headersUnchanged())) {
        banner("bad", "The spreadsheet's columns have changed<small>Nothing was added. Press Cancel and try again.</small>");
        $("addBtn").disabled = false; return;
      }
      var idx = await TCCGraph.readIndex(core("name").letter, core("email").letter, state.size.lastRow + 50);
      var e = String(vals.email.v).trim().toLowerCase();
      var dupe = e && TCCMatch.isUsableEmail(e) && idx.rows.filter(function (r) { return r.email.trim().toLowerCase() === e; })[0];
      if (dupe) { banner("warn", "That email is already on row " + dupe.row + "<small>Nothing was added. Search again to open it.</small>"); $("addBtn").disabled = false; return; }
      var row = idx.nextEmptyRow;
      busy("Adding to row " + row + "…");
      await TCCGraph.writeCells(Object.keys(vals).filter(function (k) { return vals[k].v !== ""; }).map(function (k) {
        return { col: vals[k].f.letter, row: row, value: keepText(vals[k].kind, vals[k].v) };
      }));
      Object.assign(state.rows, await TCCGraph.readRows([row], state.size.lastCol));
      openEditor(row, { kind: "ok", candidate: { reason: "New customer added", row: row } });
      banner("ok", "Customer added ✓<small>Row " + row + ".</small>");
    } catch (err) { fail(err); $("addBtn").disabled = false; }
  }

  // ---------- start ----------
  // Only order emails are looked up; anything else (e.g. a weather email) just shows a quiet note
  function isOrderEmail(item) {
    var subj = String(item.subject || "");
    var from = String((item.from && item.from.emailAddress) || "").toLowerCase();
    var senders = (cfg.orderEmailSenders || []).map(function (x) { return x.toLowerCase(); });
    if (senders.indexOf(from) >= 0) return true;
    return (cfg.orderEmailSubjects || []).some(function (re) { return new RegExp(re, "i").test(subj); });
  }

  function resetPanel() {
    state.current = null;
    $("editor").classList.add("hidden"); $("editor").innerHTML = "";
    $("candidates").innerHTML = "";
  }

  async function run() {
    resetPanel();
    var item = Office.context.mailbox.item;
    if (!item) {
      $("order").textContent = "";
      banner("info", "Select an email to look up its customer.");
      return;
    }
    if (!isOrderEmail(item)) {
      state.parsed = null;
      $("order").textContent = "Not an order email.";
      banner("info", "Nothing to look up<small>Open Search manually if you need a customer.</small>");
      $("searchBox").open = false;
      return;
    }
    state.parsed = await readEmail();
    showOrder(state.parsed);
    if (state.parsed && state.parsed.found) lookup({ name: state.parsed.name, email: state.parsed.email });
    else { banner("info", "Open the search above to look up a customer."); $("searchBox").open = true; }
  }

  Office.onReady(function () {
    $("searchForm").onsubmit = function (ev) { ev.preventDefault(); lookup({ name: $("qName").value.trim(), email: $("qEmail").value.trim() }); };
    if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
      // pinned panel: refresh when a different email is selected
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { state.rows = {}; run(); });
    }
    run();
  });
})();
