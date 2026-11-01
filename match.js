/* Two Chimps customer lookup - matching logic.
   Pure functions, no Office/Graph dependencies, so it can be tested in Node. */
(function (root) {
  "use strict";

  // Emails that are placeholders, not real addresses (e.g. "noths" for Not On The High Street)
  function isUsableEmail(e) {
    return !!e && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
  }

  function normEmail(e) {
    return String(e || "").trim().toLowerCase();
  }

  function normName(n) {
    return String(n || "")
      .toLowerCase()
      .replace(/[^a-z0-9\s'-]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function tokens(n) {
    return normName(n).split(" ").filter(function (t) { return t.length > 1; });
  }

  function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    var prev = new Array(b.length + 1), cur = new Array(b.length + 1);
    for (var j = 0; j <= b.length; j++) prev[j] = j;
    for (var i = 1; i <= a.length; i++) {
      cur[0] = i;
      for (var k = 1; k <= b.length; k++) {
        cur[k] = Math.min(prev[k] + 1, cur[k - 1] + 1, prev[k - 1] + (a[i - 1] === b[k - 1] ? 0 : 1));
      }
      var t = prev; prev = cur; cur = t;
    }
    return prev[b.length];
  }

  function similarity(a, b) {
    if (!a || !b) return 0;
    var d = levenshtein(a, b);
    return 1 - d / Math.max(a.length, b.length);
  }

  // How alike two names are, 0..1. Handles "Bethan Dykes" vs "Bethan Dykes / Peter Dykes"
  // and small typos ("Jenney" vs "Jenny").
  function nameScore(a, b) {
    var na = normName(a), nb = normName(b);
    if (!na || !nb) return 0;
    if (na === nb) return 1;
    var whole = similarity(na, nb);
    var ta = tokens(a), tb = tokens(b);
    if (!ta.length || !tb.length) return whole;
    // every token of the shorter name is (nearly) present in the longer one
    var short = ta.length <= tb.length ? ta : tb, long = short === ta ? tb : ta;
    var hits = 0;
    short.forEach(function (t) {
      if (long.some(function (u) { return u === t || (t.length > 3 && similarity(t, u) >= 0.8); })) hits++;
    });
    var tokenScore = hits / short.length;
    // need surname-ish agreement: a single shared first name isn't enough
    if (short.length === 1 && long.length > 1) tokenScore *= 0.6;
    return Math.max(whole, tokenScore * 0.95);
  }

  function emailScore(a, b) {
    if (!isUsableEmail(a) || !isUsableEmail(b)) return 0;
    if (a === b) return 1;
    var la = a.split("@")[0], lb = b.split("@")[0];
    var s = similarity(la.replace(/[._-]/g, ""), lb.replace(/[._-]/g, ""));
    return s >= 0.85 ? s * 0.9 : 0;
  }

  /*
    rows: array of {row, name, email}
    query: {name, email}
    returns {status, candidates:[{row, name, email, reason, level}]}
      status: "exact" | "multiple" | "warning" | "none"
      level:  "exact" | "email" | "name" | "close"
  */
  function findMatches(rows, query, opts) {
    opts = opts || {};
    var max = opts.max || 10;
    var qEmail = normEmail(query.email);
    var qName = normName(query.name);
    var useEmail = isUsableEmail(qEmail);
    var out = [];

    rows.forEach(function (r) {
      var rEmail = normEmail(r.email);
      var rEmailOk = isUsableEmail(rEmail);
      var emailSame = useEmail && rEmailOk && rEmail === qEmail;
      var ns = nameScore(qName, r.name);
      var nameSame = qName && normName(r.name) === qName;
      var c = null;

      if (emailSame && nameSame) c = { level: "exact", score: 3, reason: "Name and email match" };
      else if (emailSame && ns >= 0.75) c = { level: "email", score: 2.6, reason: "Email matches, name is slightly different" };
      else if (emailSame) c = { level: "email", score: 2.4, reason: "Email matches, but the name is different" };
      else if (nameSame) {
        c = { level: "name", score: 2.0,
              reason: rEmailOk ? "Name matches, but the email is different" : "Name matches, no email on this row" };
      } else {
        var es = useEmail ? emailScore(qEmail, rEmail) : 0;
        if (es > 0 && ns >= 0.6) c = { level: "close", score: 1.5 + es / 10, reason: "Similar email and name" };
        else if (es > 0) c = { level: "close", score: 1.2 + es / 10, reason: "Very similar email address" };
        else if (ns >= 0.85) c = { level: "close", score: 1 + ns / 10, reason: "Similar name" };
      }
      if (c) {
        c.row = r.row; c.name = r.name; c.email = r.email;
        out.push(c);
      }
    });

    out.sort(function (a, b) { return b.score - a.score || b.row - a.row; });

    var status;
    var exact = out.filter(function (c) { return c.level === "exact"; });
    var strong = out.filter(function (c) { return c.level !== "close"; });
    if (!out.length) status = "none";
    else if (exact.length === 1 && out.length === 1) status = "exact";
    else if (exact.length >= 1 && strong.length === 1) status = "exact"; // one strong match plus faint lookalikes
    else if (strong.length > 1) status = "multiple";
    else status = "warning";

    // if exact, drop the faint lookalikes from view but keep a count
    var hidden = 0;
    if (status === "exact") { hidden = out.length - 1; out = [exact[0]]; }
    return { status: status, candidates: out.slice(0, max), total: out.length + hidden, hiddenLookalikes: hidden };
  }

  var api = { findMatches: findMatches, nameScore: nameScore, normName: normName,
              normEmail: normEmail, isUsableEmail: isUsableEmail };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TCCMatch = api;
})(this);
