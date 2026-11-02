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

  // Turn the email's HTML into one-thing-per-line text. Outlook's plain-text version of these
  // emails runs the billing address into everything else, so we use the HTML instead.
  var BLOCK = /<\/?(p|div|tr|td|th|h[1-6]|li|ul|ol|table|tbody|thead|address|section|header|footer)\b[^>]*>/gi;
  function htmlToText(html) {
    var s = String(html || "")
      .replace(/<(script|style|head)\b[\s\S]*?<\/\1>/gi, " ")
      .replace(/<!--[\s\S]*?-->/g, " ")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(BLOCK, "\n")
      .replace(/<[^>]+>/g, "");
    var ent = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", pound: "£" };
    s = s.replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, function (m, e) {
      if (e[0] === "#") return String.fromCharCode(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return ent[e.toLowerCase()] != null ? ent[e.toLowerCase()] : m;
    });
    return s.replace(/[ \t ]+/g, " ");
  }

  var OWN_EMAIL = /@twochimpscoffee\.com$|noreply|no-reply|wordpress@/i;

  function titleCase(s) {
    return s.toLowerCase().replace(/(^|[\s'-])([a-z])/g, function (m, p, c) { return p + c.toUpperCase(); });
  }

  /* returns {name, email, phone, orderNo, orderDate, orderType, renewEvery, isSubscription, isSample} */
  function parseOrderEmail(text, subject) {
    var lines = cleanLines(text);
    var res = { name: "", email: "", phone: "", orderNo: "", orderDate: "", orderType: "",
                renewEvery: "", isSubscription: false, isSample: false, found: false };

    var bi = lines.findIndex(function (l) { return /^billing address:?$/i.test(l); });
    if (bi >= 0) {
      var end = lines.findIndex(function (l, i) { return i > bi && /^shipping address:?$|^congratulations/i.test(l); });
      var block = lines.slice(bi + 1, end > 0 ? end : bi + 12);
      if (block.length) res.name = block[0];
      block.forEach(function (l) {
        var m = l.match(EMAIL_RE);
        if (m && !res.email) res.email = m[0];
        else if (!m && PHONE_RE.test(l) && !res.phone) res.phone = l.replace(/\s+/g, "");
      });
    }
    // fallbacks: "order from Thomas JENNEY." / "order from Megan Ross:"
    if (!res.name) {
      var f = String(text).match(/order from ([^:.\n\[]{2,60}?)\s*[:.\n\[]/i);
      if (f) res.name = f[1].trim();
    }
    if (!res.email) {
      var all = String(text).match(new RegExp(EMAIL_RE.source, "g")) || [];
      var cust = all.filter(function (e) { return !OWN_EMAIL.test(e); });
      if (cust.length) res.email = cust[0];
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

  function parseOrderHtml(html, subject) { return parseOrderEmail(htmlToText(html), subject); }

  var api = { parseOrderEmail: parseOrderEmail, parseOrderHtml: parseOrderHtml, htmlToText: htmlToText,
              guessFrequency: guessFrequency };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TCCParse = api;
})(this);
