/* AlizaPDF for Word - POC
 * analyzer.js: בדיקת צורת מסמך לפי הוראת מנהל בתי המשפט בדבר צורת מסמך ומבנהו
 * (מכוח תקנה 31 לתקנות סדר הדין האזרחי, תשע"ט-2018).
 * פונקציה טהורה: מקבלת XML (OOXML) ומחזירה דו"ח. עובדת בדפדפן (DOMParser) וב-Node (xmldom).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.AlizaAnalyzer = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
  var PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

  // הכללים (הוראת צורת מסמך, כפי שצוטטה דרך LawMate - לאמת מול הנוסח הרשמי)
  var RULES = {
    allowedFonts: ["david", "raanana", "רעננה", "calibri", "david libre"],
    bodySizeHalfPts: 24,      // 12 נק'
    headingMaxHalfPts: 28,    // 14 נק' לכותרות
    lineTwips: 360,           // שורה וחצי (240 = בודד)
    marginTwips: 1417,        // 2.5 ס"מ
    marginTol: 40,
    a4: { w: 11906, h: 16838, tol: 40 },
    annexThreshold: 5         // מעל 5 נספחים: תוכן עניינים ודף מקדים
  };

  var HEB = /[֐-׿]/;

  var ParserCtor = (typeof DOMParser !== "undefined") ? DOMParser : null;
  function setDOMParser(ctor) { ParserCtor = ctor; }
  function parseXml(str) {
    if (!ParserCtor) throw new Error("DOMParser not available");
    return new ParserCtor().parseFromString(str, "application/xml");
  }

  function localName(n) { return n.localName || (n.nodeName || "").replace(/^.*:/, ""); }

  function childrenNamed(el, name) {
    var out = [];
    if (!el) return out;
    for (var i = 0; i < el.childNodes.length; i++) {
      var c = el.childNodes[i];
      if (c.nodeType === 1 && localName(c) === name) out.push(c);
    }
    return out;
  }
  function firstChild(el, name) { var c = childrenNamed(el, name); return c.length ? c[0] : null; }
  function attr(el, name) {
    if (!el) return null;
    var v = el.getAttributeNS ? el.getAttributeNS(W_NS, name) : null;
    if (!v) v = el.getAttribute("w:" + name);
    return v || null;
  }
  function descendants(el, name) {
    var out = [];
    if (!el) return out;
    var all = el.getElementsByTagNameNS ? el.getElementsByTagNameNS(W_NS, name) : el.getElementsByTagName("w:" + name);
    for (var i = 0; i < all.length; i++) out.push(all[i]);
    return out;
  }

  // חילוץ חלקים מחבילת pkg:package (הפלט של body.getOoxml())
  function extractParts(packageXml) {
    var doc = parseXml(packageXml);
    var parts = {};
    var all = doc.getElementsByTagNameNS ? doc.getElementsByTagNameNS(PKG_NS, "part") : doc.getElementsByTagName("pkg:part");
    for (var i = 0; i < all.length; i++) {
      var p = all[i];
      var name = p.getAttributeNS ? (p.getAttributeNS(PKG_NS, "name") || p.getAttribute("pkg:name")) : p.getAttribute("pkg:name");
      var xmlData = null;
      for (var j = 0; j < p.childNodes.length; j++) {
        var c = p.childNodes[j];
        if (c.nodeType === 1 && localName(c) === "xmlData") { xmlData = c; break; }
      }
      if (name && xmlData) {
        var rootEl = null;
        for (var k = 0; k < xmlData.childNodes.length; k++) if (xmlData.childNodes[k].nodeType === 1) { rootEl = xmlData.childNodes[k]; break; }
        parts[name] = rootEl;
      }
    }
    return parts;
  }

  // ---------- סגנונות ----------
  function buildStyleIndex(stylesRoot, themeRoot) {
    var idx = { styles: {}, defaultParaStyle: null, docDefaults: { rPr: null, pPr: null }, theme: { cs: null, latin: null } };
    if (stylesRoot) {
      var dd = firstChild(stylesRoot, "docDefaults");
      if (dd) {
        var rd = firstChild(dd, "rPrDefault"); if (rd) idx.docDefaults.rPr = firstChild(rd, "rPr");
        var pd = firstChild(dd, "pPrDefault"); if (pd) idx.docDefaults.pPr = firstChild(pd, "pPr");
      }
      var styles = childrenNamed(stylesRoot, "style");
      for (var i = 0; i < styles.length; i++) {
        var s = styles[i];
        var id = attr(s, "styleId");
        if (!id) continue;
        var nameEl = firstChild(s, "name");
        var based = firstChild(s, "basedOn");
        idx.styles[id] = {
          id: id,
          type: attr(s, "type"),
          name: nameEl ? attr(nameEl, "val") : id,
          basedOn: based ? attr(based, "val") : null,
          rPr: firstChild(s, "rPr"),
          pPr: firstChild(s, "pPr"),
          isDefault: attr(s, "default") === "1" || attr(s, "default") === "true"
        };
        if (idx.styles[id].isDefault && idx.styles[id].type === "paragraph") idx.defaultParaStyle = id;
      }
    }
    idx.theme = { minor: themeFonts(themeRoot, "minorFont"), major: themeFonts(themeRoot, "majorFont") };
    return idx;
  }

  // גופני ערכת הנושא: עברית נקראת מ-a:cs, ואם ריק מ-a:font script="Hebr"
  function themeFonts(themeRoot, which) {
    var out = { cs: null, latin: null };
    if (!themeRoot) return out;
    var el = themeRoot.getElementsByTagNameNS ? themeRoot.getElementsByTagNameNS(A_NS, which) : themeRoot.getElementsByTagName("a:" + which);
    if (!el || !el.length) return out;
    var get = function (n) { return el[0].getElementsByTagNameNS ? el[0].getElementsByTagNameNS(A_NS, n) : el[0].getElementsByTagName("a:" + n); };
    var cs = get("cs"), lat = get("latin"), fonts = get("font");
    if (lat && lat.length) out.latin = lat[0].getAttribute("typeface") || null;
    if (cs && cs.length) out.cs = cs[0].getAttribute("typeface") || null;
    if (!out.cs && fonts) for (var i = 0; i < fonts.length; i++) if (fonts[i].getAttribute("script") === "Hebr") { out.cs = fonts[i].getAttribute("typeface") || null; break; }
    return out;
  }
  function themeFontFor(idx, themeAttr, hebrew) {
    if (!themeAttr) return null;
    var set = /^major/i.test(themeAttr) ? idx.theme.major : idx.theme.minor;
    return hebrew ? set.cs : set.latin;
  }

  function styleChain(idx, styleId) {
    var chain = [], seen = {};
    var cur = styleId;
    while (cur && idx.styles[cur] && !seen[cur]) {
      seen[cur] = true;
      chain.push(idx.styles[cur]);
      cur = idx.styles[cur].basedOn;
    }
    return chain;
  }

  // קריאת גופן מתוך rPr: מחזיר {cs, ascii, csTheme}
  function fontsFromRPr(rPr) {
    if (!rPr) return null;
    var rf = firstChild(rPr, "rFonts");
    if (!rf) return null;
    return {
      cs: attr(rf, "cs"), ascii: attr(rf, "ascii") || attr(rf, "hAnsi"),
      csTheme: attr(rf, "cstheme"), asciiTheme: attr(rf, "asciiTheme")
    };
  }
  function sizeFromRPr(rPr, hebrew) {
    if (!rPr) return null;
    var el = firstChild(rPr, hebrew ? "szCs" : "sz");
    if (!el && hebrew) el = firstChild(rPr, "sz");
    return el ? parseInt(attr(el, "val"), 10) : null;
  }

  // פתרון גופן וגודל לריצה, לפי שרשרת: ריצה -> סגנון תו -> סגנון פסקה -> ברירות מחדל -> ערכת נושא
  function resolveRun(idx, runRPr, paraStyleId, hebrew) {
    var sources = [];
    if (runRPr) sources.push(runRPr);
    var rStyle = runRPr ? firstChild(runRPr, "rStyle") : null;
    if (rStyle) { var ch = styleChain(idx, attr(rStyle, "val")); for (var i = 0; i < ch.length; i++) if (ch[i].rPr) sources.push(ch[i].rPr); }
    var pch = styleChain(idx, paraStyleId || idx.defaultParaStyle);
    for (var j = 0; j < pch.length; j++) if (pch[j].rPr) sources.push(pch[j].rPr);
    if (paraStyleId && paraStyleId !== idx.defaultParaStyle) {
      var dch = styleChain(idx, idx.defaultParaStyle);
      for (var k = 0; k < dch.length; k++) if (dch[k].rPr) sources.push(dch[k].rPr);
    }
    if (idx.docDefaults.rPr) sources.push(idx.docDefaults.rPr);

    var font = null, fontSource = "unknown", size = null;
    for (var s = 0; s < sources.length; s++) {
      if (font === null) {
        var f = fontsFromRPr(sources[s]);
        if (f) {
          var v = hebrew ? f.cs : f.ascii;
          var t = hebrew ? f.csTheme : f.asciiTheme;
          if (v) { font = v; fontSource = "explicit"; }
          else if (t) { font = themeFontFor(idx, t, hebrew); fontSource = "theme"; if (!font) font = "(ערכת נושא: " + t + ")"; }
        }
      }
      if (size === null) { var sz = sizeFromRPr(sources[s], hebrew); if (sz) size = sz; }
      if (font !== null && size !== null) break;
    }
    if (size === null) size = 20; // ברירת המחדל של Word: 10 נק'
    if (font === null) { font = hebrew ? (idx.theme.minor.cs || "Times New Roman") : (idx.theme.minor.latin || "Times New Roman"); fontSource = "default"; }
    return { font: font, size: size, fontSource: fontSource };
  }

  function resolveLineSpacing(idx, pPr, paraStyleId) {
    var sources = [];
    if (pPr) sources.push(pPr);
    var ch = styleChain(idx, paraStyleId || idx.defaultParaStyle);
    for (var i = 0; i < ch.length; i++) if (ch[i].pPr) sources.push(ch[i].pPr);
    if (idx.docDefaults.pPr) sources.push(idx.docDefaults.pPr);
    for (var s = 0; s < sources.length; s++) {
      var sp = firstChild(sources[s], "spacing");
      if (sp && attr(sp, "line")) return { line: parseInt(attr(sp, "line"), 10), rule: attr(sp, "lineRule") || "auto" };
    }
    return { line: 240, rule: "auto" };
  }

  function isHeadingStyle(idx, pPr, styleId) {
    var ch = styleChain(idx, styleId);
    for (var i = 0; i < ch.length; i++) {
      var id = (ch[i].id || "").toLowerCase(), nm = (ch[i].name || "").toLowerCase();
      if (/^(heading|title|subtitle)/.test(id) || /^(heading|title|subtitle)/.test(nm) || /כותרת/.test(nm)) return true;
      if (ch[i].pPr && firstChild(ch[i].pPr, "outlineLvl")) return true;
    }
    if (pPr && firstChild(pPr, "outlineLvl")) return true;
    return false;
  }

  function fontAllowed(name) {
    if (!name) return false;
    var n = String(name).toLowerCase().trim();
    for (var i = 0; i < RULES.allowedFonts.length; i++) if (n === RULES.allowedFonts[i]) return true;
    return false;
  }

  function paraText(p) {
    var ts = descendants(p, "t"), s = "";
    for (var i = 0; i < ts.length; i++) s += ts[i].textContent || "";
    return s;
  }

  // ---------- עמוד ----------
  function checkSectPr(sectPr) {
    var pgSz = firstChild(sectPr, "pgSz"), pgMar = firstChild(sectPr, "pgMar");
    var res = { a4: null, margins: null, pgSz: null, pgMar: null, problems: [] };
    if (pgSz) {
      var w = parseInt(attr(pgSz, "w"), 10), h = parseInt(attr(pgSz, "h"), 10);
      res.pgSz = { w: w, h: h, orient: attr(pgSz, "orient") || "portrait" };
      res.a4 = Math.abs(w - RULES.a4.w) <= RULES.a4.tol && Math.abs(h - RULES.a4.h) <= RULES.a4.tol;
      if (!res.a4) res.problems.push("גודל העמוד אינו A4 לאורך (" + (w / 567).toFixed(1) + "x" + (h / 567).toFixed(1) + " ס\"מ)");
    } else { res.problems.push("לא נמצאה הגדרת גודל עמוד"); }
    if (pgMar) {
      var m = { top: parseInt(attr(pgMar, "top"), 10), bottom: parseInt(attr(pgMar, "bottom"), 10), left: parseInt(attr(pgMar, "left"), 10), right: parseInt(attr(pgMar, "right"), 10) };
      res.pgMar = m;
      var bad = [];
      var names = { top: "עליון", bottom: "תחתון", left: "שמאלי", right: "ימני" };
      for (var k in m) if (isNaN(m[k]) || Math.abs(m[k] - RULES.marginTwips) > RULES.marginTol) bad.push(names[k] + " " + (isNaN(m[k]) ? "?" : (m[k] / 567).toFixed(2)) + " ס\"מ");
      res.margins = bad.length === 0;
      if (bad.length) res.problems.push("שוליים שאינם 2.5 ס\"מ: " + bad.join(", "));
    } else { res.problems.push("לא נמצאה הגדרת שוליים"); }
    return res;
  }

  // ---------- זיהוי נספחים ----------
  var ANNEX_RE = /נספח(?:ים)?\s*(?:מס(?:פר|')?\s*)?([א-ת]{1,2}(?:['׳"״])?|\d{1,3})(?![א-ת\d])/g;
  function detectAnnexes(texts) {
    var seen = {}, list = [];
    var SPLIT = /[,;:()\n]|(?:^|[^\d])\.(?!\d)/;
    for (var i = 0; i < texts.length; i++) {
      var t = texts[i], m, prevEnd = 0;
      ANNEX_RE.lastIndex = 0;
      while ((m = ANNEX_RE.exec(t)) !== null) {
        var label = m[1].replace(/['׳"״]/g, "");
        var matchEnd = m.index + m[0].length;
        // סינון: "נספח" + מילה שאינה סימון (למשל "נספח של")
        if (/^(של|זה|זו|את|או|על|כל|אל|גם|אם|מן|הנ|לא|יש|בו|בה)$/.test(label)) { prevEnd = matchEnd; continue; }
        if (!seen[label]) {
          seen[label] = true;
          var start = Math.max(0, m.index - 40), end = Math.min(t.length, matchEnd + 60);
          // הסעיפים שלפני ההפניה (מאז ההפניה הקודמת), לצורך ניחוש שם הנספח
          var segs = t.substring(prevEnd, m.index).split(SPLIT);
          var name = null, clause = "";
          for (var q = segs.length - 1; q >= 0 && q >= segs.length - 3; q--) {
            var g = guessAnnexName(segs[q], label, prevEnd > 0 && q === 0);
            if (g !== "נספח " + label) { name = g; clause = segs[q]; break; }
          }
          list.push({ label: label, firstMention: t.substring(start, end).trim(), clauseBefore: clause.trim(), paragraphIndex: i, guessedName: name || ("נספח " + label) });
        }
        prevEnd = matchEnd;
      }
    }
    return list;
  }

  // ניחוש שם הנספח: הסעיף שלפני "נספח X", בלי מילות הצירוף ("המצורף כ", "מצ"ב", "מסומן")
  function guessAnnexName(clause, label, continuation) {
    var s = (clause || "").trim();
    s = s.replace(/\s*(?:ה?מצורפ(?:ת|ים|ות)?|מצ["״]ב|רצ["״]ב|ה?מסומ(?:ן|נת|נים|נות)|ה?מצ(?:ורף|"ב)|אשר\s+צורף|שצורף|צורף|צורפה)\s*(?:בזאת|בזה|לעיל|להלן)?\s*(?:כ|ב|ל)?$/, "");
    s = s.replace(/\s+(?:כ|ב|ל|כנספח|בנספח)$/, "");
    s = s.replace(/^(?:ו|וכן|כן|גם|ואילו|וגם|כמו כן|בנוסף|אף|וכן גם)\s+/, "");
    if (continuation) s = s.replace(/^ו(?=[א-ת]{3,})/, "");
    s = s.replace(/^(?:העתק|עותק|צילום)\s+(?:של\s+)?/, "");
    s = s.replace(/\s+(?:ו|וכן)$/, "").trim();
    if (s.length < 3 || s.length > 80) return "נספח " + label;
    return s;
  }

  // ---------- מספור עמודים ----------
  function hasPageField(xmlStrings) {
    for (var i = 0; i < xmlStrings.length; i++) {
      var s = xmlStrings[i] || "";
      if (/w:instr="[^"]*\bPAGE\b/i.test(s)) return true;
      if (/<w:instrText[^>]*>[^<]*\bPAGE\b[^<]*<\/w:instrText>/i.test(s)) return true;
    }
    return false;
  }

  // ---------- הבדיקה הראשית ----------
  function analyze(input) {
    // input: { documentXml | packageXml, stylesXml?, themeXml?, footerXmls?: string[] }
    var docRoot = null, stylesRoot = null, themeRoot = null;
    if (input.packageXml) {
      var parts = extractParts(input.packageXml);
      docRoot = parts["/word/document.xml"] || null;
      stylesRoot = parts["/word/styles.xml"] || null;
      themeRoot = parts["/word/theme/theme1.xml"] || null;
    }
    if (!docRoot && input.documentXml) docRoot = parseXml(input.documentXml).documentElement;
    if (!stylesRoot && input.stylesXml) stylesRoot = parseXml(input.stylesXml).documentElement;
    if (!themeRoot && input.themeXml) themeRoot = parseXml(input.themeXml).documentElement;
    if (!docRoot) throw new Error("לא נמצא document.xml");

    var idx = buildStyleIndex(stylesRoot, themeRoot);
    var body = firstChild(docRoot, "body");
    var paragraphs = descendants(body, "p");
    var report = {
      stylesAvailable: !!stylesRoot,
      pageSetup: [], fonts: { checked: 0, violations: [] }, sizes: { violations: [] },
      spacing: { checked: 0, violations: [] }, pageNumbers: null, annexes: [], paragraphCount: 0, texts: []
    };

    // sectPr: בגוף (אחרון) ובפסקאות (מעברי מקטע)
    var sects = descendants(body, "sectPr");
    for (var s = 0; s < sects.length; s++) report.pageSetup.push(checkSectPr(sects[s]));

    var texts = [];
    for (var i = 0; i < paragraphs.length; i++) {
      var p = paragraphs[i];
      var text = paraText(p);
      texts.push(text);
      if (!text.trim()) continue;
      report.paragraphCount++;
      var pPr = firstChild(p, "pPr");
      var pStyle = pPr ? firstChild(pPr, "pStyle") : null;
      var styleId = pStyle ? attr(pStyle, "val") : null;
      var heading = isHeadingStyle(idx, pPr, styleId);
      var snippet = text.length > 50 ? text.substring(0, 50) + "..." : text;

      // ריווח
      var ls = resolveLineSpacing(idx, pPr, styleId);
      report.spacing.checked++;
      var okSpacing = (ls.rule === "auto" && Math.abs(ls.line - RULES.lineTwips) <= 8);
      if (!okSpacing && !heading) {
        report.spacing.violations.push({ paragraph: i + 1, snippet: snippet, line: ls.line, rule: ls.rule, display: ls.rule === "auto" ? (ls.line / 240).toFixed(2) + " שורות" : (ls.line / 20) + " נק' (" + ls.rule + ")" });
      }

      // ריצות
      var runs = childrenNamed(p, "r");
      // ריצות בתוך היפר-קישוק/smartTag וכדומה
      var extra = descendants(p, "r");
      if (extra.length > runs.length) runs = extra;
      var seenFonts = {};
      for (var r = 0; r < runs.length; r++) {
        var rt = paraText(runs[r]);
        if (!rt.trim()) continue;
        var hebrew = HEB.test(rt);
        var rPr = firstChild(runs[r], "rPr");
        var res = resolveRun(idx, rPr, styleId, hebrew);
        report.fonts.checked++;
        var key = res.font + "|" + res.size;
        if (seenFonts[key]) continue;
        seenFonts[key] = true;
        if (!fontAllowed(res.font)) {
          report.fonts.violations.push({ paragraph: i + 1, snippet: snippet, font: res.font, hebrew: hebrew, source: res.fontSource });
        }
        var maxSz = heading ? RULES.headingMaxHalfPts : RULES.bodySizeHalfPts;
        var minSz = heading ? RULES.bodySizeHalfPts : RULES.bodySizeHalfPts;
        if (res.size > maxSz || res.size < minSz) {
          report.sizes.violations.push({ paragraph: i + 1, snippet: snippet, size: res.size / 2, heading: heading });
        }
      }
    }
    report.texts = texts;
    report.annexes = detectAnnexes(texts);
    report.needsTocAndCovers = report.annexes.length > RULES.annexThreshold;
    if (input.footerXmls || input.headerXmls) {
      report.pageNumbers = hasPageField((input.footerXmls || []).concat(input.headerXmls || []));
    }

    // סיכום
    var issues = [];
    for (var ps = 0; ps < report.pageSetup.length; ps++) issues = issues.concat(report.pageSetup[ps].problems);
    if (report.fonts.violations.length) issues.push("גופן שאינו David/רעננה/Calibri ב-" + report.fonts.violations.length + " מקומות");
    if (report.sizes.violations.length) issues.push("גודל גופן שאינו 12 (או 14 לכותרת) ב-" + report.sizes.violations.length + " מקומות");
    if (report.spacing.violations.length) issues.push("ריווח שורות שאינו שורה וחצי ב-" + report.spacing.violations.length + " פסקאות");
    if (report.pageNumbers === false) issues.push("לא נמצא מספור עמודים בכותרת התחתונה או העליונה");
    if (report.needsTocAndCovers) issues.push("יותר מ-5 נספחים (" + report.annexes.length + "): נדרשים תוכן עניינים ודף מקדים לכל נספח");
    report.issues = issues;
    report.ok = issues.length === 0;
    return report;
  }

  return { analyze: analyze, extractParts: extractParts, detectAnnexes: detectAnnexes, RULES: RULES, hasPageField: hasPageField, setDOMParser: setDOMParser };
});
