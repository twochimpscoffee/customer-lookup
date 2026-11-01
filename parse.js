/* Two Chimps customer lookup - reads the WooCommerce order email text. */
(function (root) {
  "use strict";

  var EMAIL_RE = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/;
  var PHONE_RE = /^\+?[\d\s()-]{9,}$/;

  function cleanLines(text) {
    return String(text || "").replace(/\r/g, "").split("\n")
      .map(function (l) { return l.replace(/ /g, " ").trim(); })
      .filter(function (l) { return l.length; });
  }

  function titleCase(s) {
    return s.toLowerCase().replace(/(^|[\s'-])([a-z])/g, function (m, p, c) { return p + c.toUpperCase(); });
  }

  /* returns {name, email, phone, orderNo, orderDate, orderType, renewEvery, isSubscription, isSample} */
  function parseOrderEmail(text, subject) {
    var lines = cleanLines(text);
    var res = { name: "", email: "", phone: "", orderNo: "", orderDate: "", orderType: "",
                renewEvery: "", isSubscription: false, isSample: false, found: false };

    var bi = lines.findIndex(function (l) { return /^billing address$/i.test(l); });
    if (bi >= 0) {
      var end = lines.findIndex(function (l, i) { return i > bi && /^shipping address$|^congratulations/i.test(l); });
      var block = lines.slice(bi + 1, end > 0 ? end : bi + 12);
      if (block.length) res.name = block[0];
      block.forEach(function (l) {
        var m = l.match(EMAIL_RE);
        if (m && !res.email) res.email = m[0];
        else if (!m && PHONE_RE.test(l) && !res.phone) res.phone = l.replace(/\s+/g, "");
      });
    }
    // fallback: "order from Thomas JENNEY."
    if (!res.name) {
      var f = String(text).match(/order from ([^.\n]+)\./i);
      if (f) res.name = f[1].trim();
    }
    // WooCommerce often shouts surnames ("Thomas JENNEY") - tidy for display/new rows
    res.nameTidy = /\b[A-Z]{3,}\b/.test(res.name) ? titleCase(res.name) : res.name;

    var o = String(text).match(/\[Order #(\d+)\]\s*\(([^)]+)\)/);
    if (o) { res.orderNo = o[1]; res.orderDate = o[2]; }
    else {
      var s = String(subject || "").match(/\((\d{4,})\)/);
      if (s) res.orderNo = s[1];
    }
    var t = String(subject || "").match(/\]\s*(New [^(]+?)\s*\(/i);
    res.orderType = t ? t[1] : "";
    var r = String(text).match(/Renew Every:\s*\n?\s*([^\n]+)/i);
    if (r) res.renewEvery = r[1].trim();
    res.isSubscription = /subscription/i.test(text) || /subscription/i.test(subject || "");
    res.isSample = /\bsample/i.test(text);
    res.found = !!(res.name || res.email);
    return res;
  }

  /* Best guess at the Frequency dropdown value for a new customer row */
  function guessFrequency(p) {
    if (p.isSubscription && p.renewEvery) {
      var m = p.renewEvery.toLowerCase();
      if (/^(1 )?week$|^1 weeks?$|every week/.test(m)) return "Weekly";
      if (/2 weeks|fortnight/.test(m)) return "Fortnightly";
      if (/4 weeks|month/.test(m)) return "Monthly";
    }
    if (p.isSample && !p.isSubscription) return "Sample";
    if (!p.isSubscription) return "Once";
    return "";
  }

  var api = { parseOrderEmail: parseOrderEmail, guessFrequency: guessFrequency };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TCCParse = api;
})(this);
