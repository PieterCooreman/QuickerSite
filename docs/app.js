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
  var kwRe = new RegExp("\\b(" + KW.join("|") + ")\\b", "gi");

  function esc(s) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

  function highlight(code, lang) {
    // tokenise by lines to keep comments/strings sane
    var out = [];
    code.split("\n").forEach(function (line) {
      out.push(hlLine(line, lang));
    });
    return out.join("\n");
  }
  function hlLine(line, lang) {
    var safe = esc(line);
    // comment (VBScript ' )  -- only if not preceding code is a string; good enough
    var comIdx = findComment(line);
    var codePart = safe, comPart = "";
    if (comIdx > -1) {
      codePart = esc(line.slice(0, comIdx));
      comPart = "<span class='tok-com'>" + esc(line.slice(comIdx)) + "</span>";
    }
    // strings
    codePart = codePart.replace(/&quot;[^&]*?&quot;|"[^"]*"/g, function (m) {
      return "<span class='tok-str'>" + m + "</span>";
    });
    // numbers
    codePart = codePart.replace(/\b(\d+\.?\d*)\b/g, "<span class='tok-num'>$1</span>");
    // keywords (avoid inside already-tagged spans is hard; acceptable here)
    codePart = codePart.replace(kwRe, function (m) { return "<span class='tok-kw'>" + m + "</span>"; });
    // bracket shortcodes [NAME]
    codePart = codePart.replace(/\[([A-Z0-9_:]+)\]/g, "<span class='tok-fn'>[$1]</span>");
    return codePart + comPart;
  }
  function findComment(line) {
    var inStr = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (c === '"') inStr = !inStr;
      if (c === "'" && !inStr) return i;
    }
    return -1;
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

  function build() {
    var goal = ($("pbGoal").value || "").trim();
    var name = ($("pbName").value || "").trim().toUpperCase().replace(/[\[\]\s]/g, "");
    var hasP = $("pbParams").checked;
    var sig = ($("pbSig").value || "").trim();
    var notes = ($("pbNotes").value || "").trim();
    var schema = ($("pbSchema").value || "").trim();
    var out = $("pbOut").value;
    var extra = ($("pbExtra").value || "").trim();

    var L = [];
    L.push("You are writing a custom ASP/VBScript snippet for the QuickerSite CMS.");
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

    var capLines = [];
    Object.keys(CAPTXT).forEach(function (k) { if (caps[k]) capLines.push("- " + CAPTXT[k]); });
    L.push("");
    L.push("ALLOWED CAPABILITIES");
    L.push(capLines.length ? capLines.join("\n") : "- Plain string building only (no DB / IO).");

    if (caps.db && schema) {
      L.push("");
      L.push("DATABASE SCHEMA");
      L.push(schema.split("\n").map(function (x) { return "- " + x; }).join("\n"));
    }
    if (extra) {
      L.push("");
      L.push("EXTRA REQUIREMENTS");
      L.push(extra.split("\n").map(function (x) { return "- " + x; }).join("\n"));
    }

    L.push("");
    L.push("DELIVERABLE");
    L.push("Return three labelled blocks ready to paste into bs_constantEdit.asp:");
    L.push("1) Parameters  - the exact Parameters-field string (" + (hasP ? "as above" : "write `none`") + ").");
    L.push("2) Function body - the VBScript for the main code box. Do NOT include the function/end function");
    L.push("   wrapper or <% %> tags; QuickerSite adds them.");
    L.push("3) Global Code - helper Functions/Classes, or `none`.");
    L.push("Then give a one-line test note (paste into bs_constantEdit.asp, click Test!" +
      (hasP ? ", supply sample args in double quotes" : "") + ", confirm \"TEST OK!\", Save, insert [" +
      (name || "NAME") + (hasP ? "(...)" : "") + "]).");
    L.push("State any assumptions; do not invent DB columns or helpers that were not given.");

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
    build();
  });
})();
