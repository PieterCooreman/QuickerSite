/* QuickerSite Developer Guidelines - behaviour */
(function () {
  "use strict";

  /* ---------- mobile sidebar ---------- */
  var aside = document.querySelector("aside");
  var toggle = document.querySelector(".menu-toggle");
  if (toggle) toggle.addEventListener("click", function () { aside.classList.toggle("open"); });
  document.querySelectorAll("nav.toc a").forEach(function (a) {
    a.addEventListener("click", function () { aside.classList.remove("open"); });
  });

  /* ---------- scroll-spy ---------- */
  var links = Array.prototype.slice.call(document.querySelectorAll("nav.toc a[href^='#']"));
  var map = {};
  links.forEach(function (a) {
    var id = a.getAttribute("href").slice(1);
    var el = document.getElementById(id);
    if (el) map[id] = a;
  });
  var ids = Object.keys(map);
  function spy() {
    var pos = window.scrollY + 120, current = null;
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (el && el.offsetTop <= pos) current = ids[i];
    }
    links.forEach(function (a) { a.classList.remove("active"); });
    if (current && map[current]) map[current].classList.add("active");
  }
  window.addEventListener("scroll", spy, { passive: true });
  spy();

  /* ---------- lightweight VBScript / HTML highlighter ---------- */
  var KW = ("dim|set|if|then|else|elseif|end|function|sub|for|each|to|step|next|do|while|wend|loop|" +
    "select|case|true|false|null|nothing|new|and|or|not|is|exit|byref|byval|public|private|class|" +
    "property|get|let|on|error|resume|goto|with|redim|preserve|const|option|explicit|in").split("|");
  // matches a COMPLETE word (used via .test on an already-isolated token); no /g flag
  var kwRe = new RegExp("^(" + KW.join("|") + ")$", "i");

  function esc(s) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

  function highlight(code, lang) {
    var out = [];
    code.split("\n").forEach(function (line) {
      out.push(hlLine(line));
    });
    return out.join("\n");
  }
  // Single-pass tokenizer over the RAW line. Each char is classified exactly
  // once; token text is HTML-escaped on output so injected <span> markup can
  // never be re-processed. Handles VBScript "" escaped quotes, ' comments,
  // <,>,& inside strings, numbers, keywords and [SHORTCODE] tokens.
  function span(cls, text) { return "<span class='" + cls + "'>" + esc(text) + "</span>"; }

  function hlLine(line) {
    var html = "";
    var i = 0, n = line.length;
    while (i < n) {
      var c = line[i];

      // string literal: "...", with "" as an embedded quote
      if (c === '"') {
        var str = '"';
        i++;
        while (i < n) {
          if (line[i] === '"') {
            if (line[i + 1] === '"') { str += '""'; i += 2; continue; } // escaped quote
            str += '"'; i++; break;                                      // closing quote
          }
          str += line[i]; i++;
        }
        html += span("tok-str", str);
        continue;
      }

      // comment: ' to end of line (we're outside any string here)
      if (c === "'") { html += span("tok-com", line.slice(i)); break; }

      // [SHORTCODE] token
      if (c === "[") {
        var m = /^\[[A-Z0-9_:]+\]/.exec(line.slice(i));
        if (m) { html += span("tok-fn", m[0]); i += m[0].length; continue; }
      }

      // number
      if (/[0-9]/.test(c) && !/[A-Za-z_]/.test(line[i - 1] || "")) {
        var num = /^[0-9]+\.?[0-9]*/.exec(line.slice(i))[0];
        html += span("tok-num", num); i += num.length; continue;
      }

      // identifier / keyword
      if (/[A-Za-z_]/.test(c)) {
        var word = /^[A-Za-z_][A-Za-z0-9_]*/.exec(line.slice(i))[0];
        if (kwRe.test(word)) html += span("tok-kw", word);
        else html += esc(word);
        i += word.length; continue;
      }

      // anything else: a single escaped char
      html += esc(c); i++;
    }
    return html;
  }

  document.querySelectorAll("pre > code").forEach(function (code) {
    var lang = code.getAttribute("data-lang") || "vbs";
    var raw = code.textContent;
    if (lang !== "none") code.innerHTML = highlight(raw, lang);
    // copy button
    var pre = code.parentNode;
    var btn = document.createElement("button");
    btn.className = "copy-btn"; btn.type = "button"; btn.textContent = "copy";
    btn.addEventListener("click", function () {
      copyText(raw); btn.textContent = "copied!";
      setTimeout(function () { btn.textContent = "copy"; }, 1200);
    });
    pre.appendChild(btn);
  });

  function copyText(t) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(t).catch(function () { legacy(t); });
    } else legacy(t);
    function legacy(s) {
      var ta = document.createElement("textarea"); ta.value = s;
      document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); } catch (e) {}
      document.body.removeChild(ta);
    }
  }
  window.__qsCopy = copyText;

  function toast(msg) {
    var t = document.getElementById("toast");
    if (!t) { t = document.createElement("div"); t.id = "toast"; t.className = "toast"; document.body.appendChild(t); }
    t.textContent = msg; t.classList.add("show");
    setTimeout(function () { t.classList.remove("show"); }, 1300);
  }
  window.__qsToast = toast;

  /* ============================================================
     SIMPLIFIED PROMPT BUILDER
     Produces a compact, accurate spec an AI agent can follow.
     ============================================================ */
  var $ = function (id) { return document.getElementById(id); };
  if (!$("pbGoal")) return; // builder not on page

  var caps = {};
  document.querySelectorAll("#pbCaps .pill").forEach(function (p) {
    p.addEventListener("click", function () {
      var c = p.getAttribute("data-c");
      caps[c] = !caps[c];
      p.classList.toggle("on", caps[c]);
    });
  });
  $("pbParams").addEventListener("change", function () {
    $("pbParamBox").style.display = this.checked ? "" : "none";
  });

  var DOCS_URL = "https://pietercooreman.github.io/QuickerSite/";

  // Reference block depends on whether the target agent can browse the web.
  function referenceBlock(canBrowse) {
    if (canBrowse) {
      return "REFERENCE (your agent can browse the web):\n" +
        "- Full QuickerSite developer guidelines, class reference, helper functions and the COMPLETE\n" +
        "  database/data model are published at: " + DOCS_URL + "\n" +
        "- If you need exact table columns, class signatures or more examples than this prompt provides,\n" +
        "  FETCH that URL (chapter 8 = database overview, chapter 5 = classes). Do not guess schema.";
    }
    return "REFERENCE (offline / no web access assumed):\n" +
      "- Do NOT fetch any URL and do NOT rely on memory of QuickerSite internals.\n" +
      "- Use ONLY the facts in this prompt. If a table/column or class member you need is NOT listed below,\n" +
      "  STOP and ask for it (or state the assumption explicitly) instead of inventing it.\n" +
      "- Full docs (to copy schema/classes from) live at: " + DOCS_URL + " (chapter 8 = database).";
  }

  var RULES =
    "QUICKERSITE CUSTOM SCRIPT - RULES (do not violate):\n" +
    "- The code is stored as a constant of type QS_VBScript and is invoked in page text as [NAME] or [NAME(args)].\n" +
    "- QuickerSite wraps your code as:  function CustomFunction(<params>) ... end function  and returns its value.\n" +
    "- It also rewrites every `Response.Write x` in the body to `CustomFunction = CustomFunction & x`.\n" +
    "  => Build output with Response.Write OR by assigning to CustomFunction. Return a STRING FRAGMENT only.\n" +
    "- Do NOT emit <html>/<head>/<body>, do NOT call Response.End, do NOT set headers. Output is spliced into a page.\n" +
    "- `Option Explicit` is ON: declare every variable with Dim.\n" +
    "- Production HIDES runtime errors (script silently outputs nothing). Be defensive: guard nulls, check rs.eof,\n" +
    "  coerce params with convertGetal()/convertStr(), and degrade to empty output on failure.\n" +
    "- HTML-encode/sanitize every dynamic value and ALL Request(...) input with sanitize()/Server.HTMLEncode.\n" +
    "- In scope: db (db.Execute(sql), db.GetDynamicRS), cId (customer id), customer, selectedPage, logon, l(\"key\").\n" +
    "- SQL: never concatenate raw input; escape strings with cleanUp(), scope tenant rows with iCustomerID = cId.\n" +
    "- Close recordsets (rs.Close) and Set every object = Nothing.";

  var CAPTXT = {
    db: "Reads/writes the database via db.Execute(sql) or db.GetDynamicRS. Use cleanUp() on strings, convertGetal() on numbers, and scope by iCustomerID = cId.",
    auth: "May branch on backsite login: Session(Application(\"QS_CMS_iCustomerID\") & \"isAUTHENTICATED\") = true.",
    customer: "May read the in-scope `customer` object (cls_customer): customer.siteName, customer.sDatumFormat, customer.sUrl, etc.",
    page: "May read the in-scope `selectedPage` object (cls_page): selectedPage.sTitle, selectedPage.iId, selectedPage.listitems(true), etc.",
    request: "May read Request.Form(...) / Request.QueryString(...) - sanitize ALL of it before use or output.",
    session: "May read/write Session(...) state (prefix keys with cId where appropriate).",
    application: "May cache in Application(...) keyed by cId (e.g. \"MYKEY\" & cId).",
    fso: "May use Server.CreateObject(\"Scripting.FileSystemObject\"); guard paths and Set Nothing after.",
    http: "May use Server.CreateObject(\"MSXML2.ServerXMLHTTP\") for outbound calls; wrap in On Error Resume Next.",
    constants: "May embed other QuickerSite shortcodes; they resolve recursively."
  };

  // crude heuristic: does the goal sound like it needs the database?
  var DB_HINT = /\b(shop|product|catalog|cart|order|list|news|article|post|member|contact|form|submission|booking|event|poll|gallery|feed|search|filter|recent|latest|count|database|table|query|store|inventory)\b/i;

  function build() {
    var goal = ($("pbGoal").value || "").trim();
    var name = ($("pbName").value || "").trim().toUpperCase().replace(/[\[\]\s]/g, "");
    var hasP = $("pbParams").checked;
    var sig = ($("pbSig").value || "").trim();
    var notes = ($("pbNotes").value || "").trim();
    var schema = ($("pbSchema").value || "").trim();
    var out = $("pbOut").value;
    var extra = ($("pbExtra").value || "").trim();
    var canBrowse = (document.querySelector("input[name=pbAgent]:checked") || {}).value === "online";

    var capLines = [];
    Object.keys(CAPTXT).forEach(function (k) { if (caps[k]) capLines.push("- " + CAPTXT[k]); });
    var noCaps = capLines.length === 0;

    var L = [];
    L.push("You are writing a custom ASP/VBScript snippet for the QuickerSite CMS.");
    L.push("");
    L.push(referenceBlock(canBrowse));
    L.push("");
    L.push(RULES);
    L.push("");
    L.push("TASK");
    L.push("Goal: " + (goal || "(describe what the script should do)"));
    L.push("Shortcode name: " + (name || "(CHOOSE_NAME)") + "  -> used as [" + (name || "NAME") + (hasP ? "(args)" : "") + "]");
    if (hasP) {
      L.push("Parameters: " + (sig || "(define the VBScript argument list)"));
      if (notes) L.push("Parameter notes: " + notes.replace(/\n/g, " / "));
    } else {
      L.push("Parameters: none");
    }

    var outTxt = out === "text" ? "plain text" :
      out === "json" ? "a JSON string (AJAX/endpoint use)" :
        out === "none" ? "no visible output (side-effect only; return empty string)" :
          "an HTML fragment to embed in a page";
    L.push("Output: " + outTxt + ".");

    L.push("");
    L.push("ALLOWED CAPABILITIES");
    if (noCaps) {
      // FIX (issue #1): never assert "no DB/IO" as a hard rule, because it can
      // contradict the Goal (e.g. "add a shop"). Instead let the model use the
      // minimum it needs, and warn when the goal looks data-driven.
      L.push("- Use the MINIMUM capabilities the Goal requires. Prefer pure string building, but you MAY");
      L.push("  read the database (db.Execute), customer and selectedPage if the Goal cannot be met otherwise.");
      L.push("- When you use the database, follow the SQL rules above (escape input, scope by iCustomerID = cId).");
      if (goal && DB_HINT.test(goal)) {
        L.push("- NOTE: this Goal looks data-driven. Do NOT hardcode fake/sample data to avoid the database;");
        L.push("  query the real QuickerSite tables instead. If you lack the exact schema, ask for it / state assumptions.");
      }
    } else {
      L.push(capLines.join("\n"));
    }

    // Schema: include whenever provided (not only when the db pill is on),
    // since with no caps selected the model may still need it.
    if (schema) {
      L.push("");
      L.push("DATABASE SCHEMA (authoritative - use these names exactly; do not invent others)");
      L.push(schema.split("\n").map(function (x) { return "- " + x; }).join("\n"));
    } else if (!canBrowse && (caps.db || (noCaps && goal && DB_HINT.test(goal)))) {
      // FIX (issue #2): offline agent + likely-DB task + no schema pasted.
      L.push("");
      L.push("DATABASE SCHEMA");
      L.push("- (none provided). You are offline, so do NOT guess table/column names.");
      L.push("  Ask the user to paste the relevant tables from chapter 8 of " + DOCS_URL + ",");
      L.push("  or state clearly which columns you are assuming.");
    }

    if (extra) {
      L.push("");
      L.push("EXTRA REQUIREMENTS");
      L.push(extra.split("\n").map(function (x) { return "- " + x; }).join("\n"));
    }

    // FIX (issue #3): give an explicit, copy-paste-ready output template instead
    // of an abstract 1-2-3 list, so the model does not wrap it in markdown fences.
    var paramsExample = hasP ? (sig || "iCount") : "none";
    L.push("");
    L.push("OUTPUT FORMAT (reply with EXACTLY this layout - plain text, NO markdown code fences):");
    L.push("");
    L.push("===== PARAMETERS =====");
    L.push(paramsExample);
    L.push("");
    L.push("===== FUNCTION BODY =====");
    L.push("<the VBScript for the main code box.");
    L.push(" Do NOT include the function CustomFunction(...) / end function wrapper.");
    L.push(" Do NOT include <% %> tags. Just the body statements.>");
    L.push("");
    L.push("===== GLOBAL CODE =====");
    L.push("<helper Functions/Classes, or the single word: none>");
    L.push("");
    L.push("===== NOTES =====");
    L.push("<assumptions you made, plus a one-line test instruction: paste the 3 blocks into");
    L.push(" bs_constantEdit.asp, click Test!" + (hasP ? ", supply sample args in double quotes" : "") +
      ", confirm \"TEST OK!\", Save, then insert [" + (name || "NAME") + (hasP ? "(...)" : "") + "] in a page.>");
    L.push("");
    L.push("Rules for the output: keep the five ===== headers exactly as shown; put real content");
    L.push("between them; do not add any text before the first header or after the last block.");

    $("pbResult").textContent = L.join("\n");
  }

  $("pbBuild").addEventListener("click", build);
  $("pbCopy").addEventListener("click", function () {
    var t = $("pbResult").textContent;
    if (t && t.indexOf("Press") !== 0) { copyText(t); toast("Prompt copied!"); }
  });
  $("pbExample").addEventListener("click", function () {
    $("pbGoal").value = "Show the N most recent online list items of the current list page as a simple <ul>.";
    $("pbName").value = "RECENTITEMS";
    $("pbParams").checked = true; $("pbParamBox").style.display = "";
    $("pbSig").value = "iCount";
    $("pbNotes").value = "iCount = max items (integer, default 5 if blank/zero)";
    caps = { db: false, page: true };
    document.querySelectorAll("#pbCaps .pill").forEach(function (p) {
      p.classList.toggle("on", !!caps[p.getAttribute("data-c")]);
    });
    $("pbOut").value = "html";
    $("pbExtra").value = "Use selectedPage.listitems(true). Output nothing if there are no items.";
    var offline = document.querySelector("input[name=pbAgent][value=offline]");
    if (offline) offline.checked = true;
    build();
  });
})();
