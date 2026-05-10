javascript:(function(){
"use strict";

// ─── Guard ───────────────────────────────────────────────────
if(window.__EXECIQ_P1__){console.warn("[ExecIQ] Already running.");return;}
window.__EXECIQ_P1__ = true;

var VERSION = "5.0";
// xlsx-js-style: drop-in replacement for SheetJS with full cell style support
// Same API, same XLSX global — fills, fonts, borders, alignment all apply in Excel
var SHEETJS = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";

// ─────────────────────────────────────────────────────────────
// STATIC TYPE HINTS
// The only thing that stays static: how to interpret a field's
// raw value. Everything else — which fields exist, what they're
// called, whether they're enabled — is discovered dynamically.
//
// Pattern rules applied in order. First match wins.
// Patterns tested against the UPPERCASE backend field name.
// ─────────────────────────────────────────────────────────────
var TYPE_RULES = [
  // Currency — stored value fields (not rate/currency metadata)
  { test: function(f){ return /STOREDVALUE$/.test(f) && !/RATE$/.test(f); }, type: "currency" },
  // Probability / percent fields
  { test: function(f){ return /PROBABILITY|PERCENT|MARKUP|FEEPERCENT|FEECIPERCENT/.test(f); }, type: "percent" },
  // All date fields
  { test: function(f){ return /^DT|DATE$|DATE\d*$|STARTDATE|ENDDATE|COMPLETIONDATE|ORIGDATE|BIDDATE/.test(f); }, type: "date" },
  // Currency — fee and cost fields
  { test: function(f){ return /FEE$|COST$|BUDGET$|ACTUAL$|REVENUE|MARGIN|^IFIRMFEE|^IFACTOREDFEE|^ICOST|^IFEE$|^IMARKETBUDGET|^IMARKETACTUAL|LEADMONEY\d+$/.test(f); }, type: "currency" },
  // Integer / count fields
  { test: function(f){ return /DAYSINSTAGE|ROWCOUNT|TOTALRECORDS|SUBCOUNT|^ISIZE$|WORKHOURS|MANAGEMENTUNITS|LEADNUMBER\d+$|LEADPERCENT\d+$/.test(f); }, type: "number" },
  // Long text / narrative
  { test: function(f){ return /^TX|DESCRIPTION$|STRATEGY$|HURDLES$|ANALYSIS$|NARRATIVE$|NOTES$|LEADLONGTEXT\d+$/.test(f); }, type: "longtext" },
  // HTML fields that need stripping
  { test: function(f){ return /OWNERCONTACT$|CONTACT$/.test(f) && !/ID$/.test(f); }, type: "html" },
  // Boolean / flag fields
  { test: function(f){ return /^CHR|^SF\d|SUBMITTED$|IND$|CHECK$|IMPORTED/.test(f); }, type: "flag" },
  // Default
  { test: function(){ return true; }, type: "text" }
];

// ─────────────────────────────────────────────────────────────
// PREFERRED NAME LIST FIELDS
// The server returns both ID lists (PRIMARYCATEGORYLIST) and resolved name lists
// (PRIMARYCATEGORYNAMELIST). We suppress the raw ID lists and use the name lists.
var PREFER_NAMELIST = new Set([
  "PRIMARYCATEGORYLIST",
  "SECONDARYCATEGORYLIST",
]);

// ─────────────────────────────────────────────────────────────
// STRUCTURAL NOISE — the ONLY fields suppressed by name.
// These are fields Unanet always injects into every response
// regardless of enabled state. They are never user-configured
// and have no analytical value. Keep this list minimal.
// Suppression by field name ENDS here. Everything else is
// controlled by enabledOppFields from firmData.cfc.
// ─────────────────────────────────────────────────────────────
var STRUCTURAL_NOISE = new Set([
  // Pagination artifacts — always in response, never real data
  "ROWNUMBER", "TOTALRECORDS",
  // Session/tenant IDs — internal plumbing
  "FIRMID", "CFIRMID",
  // Soft-delete and migration flags — not user data
  "DELETERECORD", "OLD_ID", "IMPORTEDRECORD",
  // Currency session metadata — exchange rate display artifacts
  "SELECTEDCURRENCY", "SELECTEDCURRENCYABBR", "SELECTEDCURRENCYSYMBOL", "SELECTEDRATE",
  "BASECURRENCY", "BASECURRENCYABBR", "BASECURRENCYSYMBOL",
  // Internal FK ID fields — we already surface the resolved NAME version of these
  // e.g. STAGEID is noise because STAGENAME is what we want
  // e.g. ROLEID is noise because ROLENAME is what we want
  "STAGEID", "ROLEID", "ISTATEID", "ICOUNTRYID",
  "OWNERCONTACTID", "OWNER_CRID", "OWNERCFIRMID", "CLIENT_CRID",
  "MASTERLEADID", "MASTERIND",
  "IPROJECTROLEID", "ISUBMITTALTYPEID",
  "OPPLEADCREATEMETHODID", "OPPLEADCREATEMETHODSUBID", "OPPLEADDODGEDATAID",
  "ILEADID",  // internal DB primary key — OppNumber (VCHLEADNUMBER) is the user-facing ID
  // Redundant computed location string — we have individual city/state/country fields
  "LOCATION",
  // Sales cycle — internal config flag, not user data
  "SALESCYCLE",
  // Raw category ID lists — server also returns resolved CATEGORYNAMELIST variants
  // We always prefer the name list; suppress the ID list to avoid duplicate columns
  "PRIMARYCATEGORYLIST",
  "SECONDARYCATEGORYLIST",
  // Calculated dupes — suffixed _CALC variants duplicate the primary field
  // (detected by suffix pattern below)
]);

// Suffix patterns for currency metadata triplets — these are structural,
// not content, and appear on every currency field automatically
function isStructuralNoise(backendKey){
  var upper = backendKey.toUpperCase();
  if(STRUCTURAL_NOISE.has(upper)) return true;
  // STOREDRATE and STOREDCURRENCY suffix variants — exchange rate metadata
  if(/STOREDRATE$|STOREDCURRENCY$/.test(upper)) return true;
  // _CALC suffix — calculated dupes of primary fields
  if(/_CALC$/.test(upper)) return true;
  // STAFFROLE_{id}ID — internal contact ID companion to STAFFROLE_{id} name field
  // We keep the name field, suppress the ID
  if(/^STAFFROLE_\d+ID$/.test(upper)) return true;
  // PRIMARYCATEGORYLIST/SECONDARYCATEGORYLIST are in STRUCTURAL_NOISE above
  // PREFER_NAMELIST check kept as belt-and-suspenders
  if(PREFER_NAMELIST && PREFER_NAMELIST.has(upper)) return true;
  return false;
}

function getType(fieldName){
  var f = fieldName.toUpperCase();
  for(var i = 0; i < TYPE_RULES.length; i++){
    if(TYPE_RULES[i].test(f)) return TYPE_RULES[i].type;
  }
  return "text";
}

// ─────────────────────────────────────────────────────────────
// FIELD ORDERING — controls column order in the output sheet.
// Fields not listed here appear alphabetically after these.
// This is display preference only — not a whitelist.
// ─────────────────────────────────────────────────────────────
var COLUMN_ORDER = [
  "VCHLEADNUMBER","VCHPROJECTNAME","COMPANY","OWNERCOMPANY","OWNER",
  "STAGENAME","__STATUS","ACTIVEIND","DAYSINSTAGE",
  "IFIRMFEE","IFACTOREDFEE","IPROBABILITY","TOTALESTIMATEDFEESTOREDVALUE","ICOST",
  "ROLENAME","PROSPECTTYPES","CLIENTTYPES","CONTRACTTYPES","DELIVERYMETHOD","SERVICETYPES",
  "SUBMITTALTYPENAME","SOLICITATIONNUMBER","CHRFPREC","CHPROPOSALSUB","NAICSCODES",
  "DTRFPDATE","DTQUALSDATE","DTPROPOSALDATE","DTPRESENTATIONDATE",
  "DTSTARTDATE","ESTIMATEDSTARTDATE","ESTIMATEDCOMPLETIONDATE","DTCLOSEDATE",
  "SHORTLISTDATE","OPPLEADBIDDATE","OPPLEADORIGDATE",
  "PRECONSTARTDATE","PRECONENDDATE","DESIGNSTARTDATE","DESIGNCOMPLETIONDATE",
  "CONSTRUCTIONSTARTDATE","CONSTRUCTIONCOMPLETIONDATE",
  "CREATEDATE","MODDATE",
  "OFFICELIST","DIVISIONLIST","STUDIOLIST","PRACTICEAREALIST","TERRITORYLIST","OFFICEDIVISIONLIST",
  "PRIMARYCATEGORYLIST","SECONDARYCATEGORYLIST",
  "OWNERCONTACT",
  "VCHADDRESS1","VCHCITY","STATEABRV","VCHPOSTALCODE","VCHCOUNTRY",
  "SCORE","REDZONESCORE","FUNDPROBABILITY",
  "DESCRIPTION","TXCOMMENTS","TXNOTE","VCHNEXTACTION",
  "PROJECTSTRATEGY","CHALLENGESHURDLES","COMPETITIONANALYSIS",
  "IMARKETBUDGET","IMARKETACTUAL",
];

// ─────────────────────────────────────────────────────────────
// FIELDS NEEDING ID RESOLUTION
// Maps backend field name → which lookup map to use.
// Built dynamically at runtime — this just names the lookup key.
// ─────────────────────────────────────────────────────────────
var ID_RESOLUTION = {
  "OFFICELIST":              "firmOrg",
  "DIVISIONLIST":            "firmOrg",
  "STUDIOLIST":              "firmOrg",
  "PRACTICEAREALIST":        "firmOrg",
  "TERRITORYLIST":           "firmOrg",
  "OFFICEDIVISIONLIST":      "firmOrg",
  // Use NAMELIST variants (already resolved by server) — suppress raw ID lists below
  "PRIMARYCATEGORYLIST":     "priCat",      // fallback if NAMELIST unavailable
  "SECONDARYCATEGORYLIST":   "secCat",      // fallback if NAMELIST unavailable
  "CONTRACTTYPES":           "contract",
  "CLIENTTYPES":             "clientType",
  "PROSPECTTYPES":           "prospect",
  "DELIVERYMETHOD":          "delivery",
  "STAGEID":                 "stage",
};


// ─────────────────────────────────────────────────────────────
// UI MODULE
// ─────────────────────────────────────────────────────────────
var UI = (function(){
  var elStatus, elFill, elLog, elBtn;

  var CSS = [
    "#iq1{position:fixed;top:16px;right:16px;z-index:2147483647;width:430px;",
    "background:#0a0e17;color:#e2e8f0;font-family:'Segoe UI',system-ui,sans-serif;",
    "font-size:13px;border:1px solid #1e3a5f;border-radius:10px;",
    "box-shadow:0 20px 60px rgba(0,0,0,.85);overflow:hidden;",
    "animation:iqIn .25s cubic-bezier(.16,1,.3,1);}",
    "@keyframes iqIn{from{opacity:0;transform:translateY(-12px)}to{opacity:1;transform:none}}",
    "#iq1 *{box-sizing:border-box;}",
    "#iq1-hd{background:#0d1526;border-bottom:1px solid #1e3a5f;padding:12px 14px;",
    "display:flex;align-items:center;gap:10px;}",
    "#iq1-logo{width:34px;height:34px;border-radius:8px;flex-shrink:0;",
    "background:linear-gradient(135deg,#1a6cf6,#7c3aed);display:flex;",
    "align-items:center;justify-content:center;font-weight:900;font-size:13px;",
    "color:#fff;letter-spacing:-.5px;}",
    "#iq1-title{flex:1;}",
    "#iq1-title h3{margin:0;font-size:13px;font-weight:700;color:#f1f5f9;}",
    "#iq1-title small{color:#64748b;font-size:10px;}",
    "#iq1-close{background:none;border:none;color:#475569;cursor:pointer;",
    "font-size:20px;line-height:1;padding:2px 6px;border-radius:4px;}",
    "#iq1-close:hover{color:#f87171;}",
    "#iq1-bd{padding:12px 14px;}",
    "#iq1-status{background:#0d1526;border:1px solid #1e3a5f;border-radius:6px;",
    "padding:8px 10px;margin-bottom:8px;font-size:11px;color:#64748b;min-height:30px;}",
    ".sq-ok{color:#34d399;}.sq-er{color:#f87171;}.sq-wn{color:#fbbf24;}",
    "#iq1-prog{height:4px;background:#1e293b;border-radius:2px;margin-bottom:8px;overflow:hidden;}",
    "#iq1-fill{height:100%;width:0%;background:linear-gradient(90deg,#1a6cf6,#7c3aed);",
    "border-radius:2px;transition:width .3s ease;}",
    "#iq1-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:8px;}",
    ".iq1-stat{background:#0d1526;border:1px solid #1e3a5f;border-radius:6px;",
    "padding:6px 8px;text-align:center;}",
    ".iq1-sv{font-size:18px;font-weight:700;color:#60a5fa;}",
    ".iq1-sl{font-size:9px;color:#475569;margin-top:1px;text-transform:uppercase;letter-spacing:.5px;}",
    "#iq1-log{background:#020617;border:1px solid #0f172a;border-radius:6px;",
    "padding:6px 8px;height:115px;overflow-y:auto;margin-bottom:10px;",
    "font-size:10px;font-family:'Consolas',monospace;color:#334155;}",
    "#iq1-log::-webkit-scrollbar{width:4px;}",
    "#iq1-log::-webkit-scrollbar-track{background:#020617;}",
    "#iq1-log::-webkit-scrollbar-thumb{background:#1e293b;border-radius:2px;}",
    "#iq1-log .ls{color:#34d399;}#iq1-log .le{color:#f87171;}#iq1-log .lw{color:#fbbf24;}",
    "#iq1-btn{width:100%;padding:9px;border:none;border-radius:6px;cursor:pointer;",
    "background:linear-gradient(135deg,#16a34a,#15803d);color:#fff;",
    "font-size:12px;font-weight:700;letter-spacing:.3px;transition:all .15s;}",
    "#iq1-btn:hover:not(:disabled){filter:brightness(1.1);}",
    "#iq1-btn:disabled{background:#1e293b;color:#475569;cursor:not-allowed;}",
    "#iq1-refresh{width:100%;padding:7px;border:1px solid #1e3a5f;border-radius:6px;",
    "cursor:pointer;background:transparent;color:#64748b;font-size:11px;",
    "font-weight:600;letter-spacing:.3px;transition:all .15s;margin-top:5px;display:none;}",
    "#iq1-refresh:hover{border-color:#2e75b6;color:#e2e8f0;}",
    "#iq1-filters{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:8px;}",
    "#iq1-filters label{font-size:9px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;display:block;margin-bottom:3px;}",
    "#iq1-filters select{width:100%;background:#0d1526;border:1px solid #1e3a5f;border-radius:4px;",
    "color:#e2e8f0;font-size:11px;padding:4px 6px;cursor:pointer;}",
    "#iq1-filters select:focus{outline:none;border-color:#2e75b6;}",
    "#iq1-window{display:flex;gap:4px;margin-bottom:8px;flex-wrap:wrap;}",
    ".iq1-win{flex:1;min-width:60px;padding:4px 2px;border:1px solid #1e3a5f;border-radius:4px;",
    "background:#0d1526;color:#64748b;font-size:10px;cursor:pointer;text-align:center;transition:all .15s;}",
    ".iq1-win:hover{border-color:#2e75b6;color:#e2e8f0;}",
    ".iq1-win.active{background:#1a6cf6;border-color:#1a6cf6;color:#fff;font-weight:700;}"
  ].join("");

  function mount(){
    var s = document.createElement("style");
    s.id = "iq1-css"; s.textContent = CSS;
    document.head.appendChild(s);
    var el = document.createElement("div");
    el.id = "iq1";
    el.innerHTML =
      '<div id="iq1-hd">' +
        '<div id="iq1-logo">IQ</div>' +
        '<div id="iq1-title">' +
          '<h3>ExecIQ Data Extractor <span style="font-weight:400;color:#475569">v' + VERSION + '</span></h3>' +
          '<small>Unanet CRM · Full Pipeline Extract · Dynamic Discovery</small>' +
        '</div>' +
        '<button id="iq1-close">×</button>' +
      '</div>' +
      '<div id="iq1-bd">' +
        '<div id="iq1-status"><span class="sq-ok">● Initializing...</span></div>' +
        '<div id="iq1-prog"><div id="iq1-fill"></div></div>' +
        '<div id="iq1-stats">' +
          '<div class="iq1-stat"><div class="iq1-sv" id="sv-opps">--</div><div class="iq1-sl">Opportunities</div></div>' +
          '<div class="iq1-stat"><div class="iq1-sv" id="sv-cols">--</div><div class="iq1-sl">Fields</div></div>' +
          '<div class="iq1-stat"><div class="iq1-sv" id="sv-cust">--</div><div class="iq1-sl">Custom Fields</div></div>' +
        '</div>' +
        '<div style="font-size:9px;color:#475569;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;">Date Filtering Options</div>' +
        '<div id="iq1-filters">' +
          '<div><label>Date Field</label>' +
            '<select id="iq1-datefield">' +
              '<option value="CREATEDATE">Date Created</option>' +
              '<option value="MODDATE">Last Modified</option>' +
              '<option value="DTSTARTDATE">Estimated Award Date</option>' +
              '<option value="DTCLOSEDATE">Actual Close Date</option>' +
              '<option value="ESTIMATEDSTARTDATE">Estimated PoP Start</option>' +
            '</select></div>' +
          '<div><label>Time Window</label>' +
            '<div id="iq1-window">' +
              '<button class="iq1-win" data-years="0">YTD</button>' +
              '<button class="iq1-win" data-years="2">2 Years</button>' +
              '<button class="iq1-win active" data-years="3">3 Years</button>' +
              '<button class="iq1-win" data-years="999">All Time</button>' +
            '</div></div>' +
        '</div>' +
        '<div id="iq1-log"></div>' +
        '<button id="iq1-btn" disabled>Preparing...</button>' +
        '<button id="iq1-refresh">↺  Refresh with New Settings</button>' +
      '</div>';
    document.body.appendChild(el);
    elStatus = document.getElementById("iq1-status");
    elFill   = document.getElementById("iq1-fill");
    elLog    = document.getElementById("iq1-log");
    elBtn    = document.getElementById("iq1-btn");
    document.getElementById("iq1-close").onclick = destroy;

    // Wire window preset buttons
    document.querySelectorAll(".iq1-win").forEach(function(btn){
      btn.onclick = function(){
        document.querySelectorAll(".iq1-win").forEach(function(b){ b.classList.remove("active"); });
        btn.classList.add("active");
      };
    });
  }

  function showRefresh(fn){
    var el = document.getElementById("iq1-refresh");
    if(el){
      el.style.display = "block";
      el.onclick = function(){
        // Reset guard and UI state for a fresh run
        window.__EXECIQ_P1__ = false;
        el.style.display = "none";
        var btn = document.getElementById("iq1-btn");
        if(btn){ btn.disabled = true; btn.textContent = "Preparing..."; btn.onclick = null; }
        var log = document.getElementById("iq1-log");
        if(log) log.innerHTML = "";
        ["sv-opps","sv-cols","sv-cust"].forEach(function(id){
          var e = document.getElementById(id); if(e) e.textContent = "--";
        });
        fn();
      };
    }
  }

  function destroy(){
    window.__EXECIQ_P1__ = false;
    document.getElementById("iq1")?.remove();
    document.getElementById("iq1-css")?.remove();
  }

  function status(msg, type){
    type = type || "ok";
    elStatus.innerHTML = '<span class="sq-' + type + '">' +
      (type==="ok"?"●":type==="er"?"✖":"⚠") + " " + msg + "</span>";
  }

  function prog(pct){ elFill.style.width = Math.min(100,Math.max(0,pct)) + "%"; }

  function log(msg, cls){
    var t = new Date().toLocaleTimeString("en-US",{hour12:false});
    var d = document.createElement("div");
    if(cls) d.className = cls;
    d.textContent = "[" + t + "] " + msg;
    elLog.appendChild(d);
    elLog.scrollTop = elLog.scrollHeight;
  }

  function stat(id, val){ var e=document.getElementById(id); if(e) e.textContent=val; }

  function enableExport(fn){
    elBtn.disabled = false;
    elBtn.textContent = "⬇  Export Opportunity Data (.xlsx)";
    elBtn.onclick = fn;
  }

  function getFilterSettings(){
    var dateField = document.getElementById("iq1-datefield");
    var activeWin = document.querySelector(".iq1-win.active");
    return {
      dateField: dateField ? dateField.value : "CREATEDATE",
      dateFieldLabel: dateField ? dateField.options[dateField.selectedIndex].text : "Date Created",
      years: activeWin ? parseInt(activeWin.getAttribute("data-years")) : 3
    };
  }

  return { mount, destroy, status, prog, log, stat, enableExport, getFilterSettings, showRefresh };
})();

// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────
async function fetchJSON(url, opts){
  try{
    var r = await fetch(url, opts||{credentials:"include",headers:{"X-Requested-With":"XMLHttpRequest"}});
    if(!r.ok) return null;
    var t = await r.text();
    var s = t.indexOf("{"); if(s<0) s=t.indexOf("["); if(s<0) return null;
    return JSON.parse(t.slice(s));
  }catch(e){ return null; }
}

async function postForm(url, params){
  return fetchJSON(url,{
    method:"POST", credentials:"include",
    headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest"},
    body: new URLSearchParams(params).toString()
  });
}

function parseCFC(data){
  if(!data) return [];
  if(data.response) data=data.response;
  if(data.Response) data=data.Response;
  // Handle both uppercase COLUMNS/DATA and lowercase columns/data (role.cfc variant)
  var cols = data.COLUMNS || data.columns;
  var rows = data.DATA    || data.data;
  if(Array.isArray(cols) && Array.isArray(rows)){
    return rows.map(function(row){
      if(!Array.isArray(row)) return row;
      var o={}; cols.forEach(function(c,i){o[c]=row[i];}); return o;
    });
  }
  if(Array.isArray(rows)) return rows;
  if(Array.isArray(data)) return data;
  return [];
}

function buildLookup(records, idField, nameField){
  var map={};
  records.forEach(function(r){
    // Try provided field names first, then common fallbacks
    var id = String(
      r[idField] || r[idField.toLowerCase()] ||
      r.ID || r.id || r.STAGEID || r.ROLEID || r.CATEGORYID ||
      r.STAFFROLEID || ""
    ).trim();
    var name = String(
      r[nameField] || r[nameField.toLowerCase()] ||
      r.NAME || r.name || r.LABEL || r.label ||
      r.DISPLAYNAME || r.displayName ||
      r.STAGENAME || r.ROLENAME || r.CATEGORYNAME ||
      r.STAFFROLENAME || r.TYPENAME || r.VALUE || r.value || ""
    ).trim();
    if(id && id!=="0" && name) map[id]=name;
  });
  return map;
}

function resolveIDs(val, map){
  if(val===null||val===undefined||val==="") return "";
  var s=String(val).trim(); if(!s) return "";
  if((map||{})[s]) return map[s];
  return s.split(/[,|]/).map(function(x){
    x=x.trim(); return (map||{})[x]||x;
  }).filter(Boolean).join(", ");
}

function fmtDate(val){
  // Returns a real Date object so SheetJS writes a proper Excel date serial number.
  // Handles all date formats observed in Unanet CRM responses:
  //   MM/DD/YYYY         — custom fields, some standard fields
  //   YYYY-MM-DD[T...]   — ISO format, datetime stamps
  //   Month, DD YYYY ... — ColdFusion default date format
  if(!val||val==="") return null;
  var s=String(val).trim(), d, parts, m;

  // MM/DD/YYYY or M/D/YYYY  (custom fields return this format)
  if(/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)){
    parts = s.slice(0,10).split("/");
    d = new Date(parseInt(parts[2]), parseInt(parts[0])-1, parseInt(parts[1]));
  }
  // YYYY-MM-DD (ISO — standard fields, datetime stamps)
  else if(/^\d{4}-\d{2}-\d{2}/.test(s)){
    parts = s.slice(0,10).split("-");
    d = new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]));
  }
  // Month, DD YYYY HH:MM:SS (ColdFusion default: "May, 04 2026 00:00:00")
  else if((m=s.match(/^(\w+),\s*(\d+)\s+(\d{4})/))){
    d = new Date(m[1]+" "+m[2]+","+m[3]);
  }
  // DD-Mon-YYYY ("04-May-2026") — occasionally seen
  else if((m=s.match(/^(\d{1,2})-(\w{3})-(\d{4})/))){
    d = new Date(m[2]+" "+m[1]+","+m[3]);
  }
  else {
    return null;
  }
  if(!d||isNaN(d.getTime())) return null;
  return d;
}

function stripHTML(str){
  if(!str) return "";
  return String(str).replace(/<br\s*\/?>/gi," | ").replace(/<[^>]+>/g,"").replace(/\s*\|\s*/g," | ").trim();
}

function classifyStatus(stageName, activeInd){
  var s=String(stageName||"").toLowerCase();
  if(String(activeInd)==="2"){
    if(/\bwon\b|award|executed/.test(s)) return "Won";
    if(/\blost\b|loss|no.go|dead|declined/.test(s)) return "Lost";
    return "Closed";
  }
  return "Active";
}

// ─────────────────────────────────────────────────────────────
// ENDPOINT DISCOVERY — all from page resource entries
// ─────────────────────────────────────────────────────────────
async function findOppBase(){
  var segs=window.location.pathname.split("/").filter(function(s){return s&&!s.includes(".");});
  var candidates=["/"];
  var p="";
  for(var i=0;i<segs.length;i++){ p+="/"+segs[i]; candidates.push(p+"/"); }
  ["/contact/","/contact/opportunity/"].forEach(function(c){ if(!candidates.includes(c)) candidates.push(c); });
  var results=await Promise.all(candidates.map(async function(base){
    try{
      var r=await fetch(base+"oppActions.cfm",{
        method:"POST",credentials:"include",
        headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest"},
        body:"action=getOpportunityGridData&json=1&start=0&limit=1&ActiveInd=0&visibleColumns=VCHPROJECTNAME"
      });
      if(r.status===404) return null;
      var t=await r.text();
      return (t.includes("ROWCOUNT")||t.includes("DATA"))?base:null;
    }catch(e){return null;}
  }));
  return results.find(function(r){return r!==null;})||null;
}

function findURLByPattern(pattern){
  var entries=performance.getEntriesByType("resource");
  for(var i=0;i<entries.length;i++){
    var url=entries[i].name;
    if(url.includes(window.location.host)&&pattern.test(url))
      return url.split("?")[0];
  }
  return null;
}

function findLookupBase(){
  var patterns=[/stage\.cfc/i,/oppData\.cfc/i,/role\.cfc/i,/contractType\.cfc/i,
                /deliveryMethod\.cfc/i,/clientType\.cfc/i,/firmOrg\.cfc/i,/staffTeam\.cfc/i];
  for(var p=0;p<patterns.length;p++){
    var url=findURLByPattern(patterns[p]);
    if(url) return url.replace(/[^\/]+\.cfc$/i,"");
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
// STEP 1: PROBE — discover which fields this client actually has
// Makes a limit=1 call and returns the set of field keys returned
// ─────────────────────────────────────────────────────────────
async function probeAvailableFields(oppBase){
  UI.log("Probing available fields from oppActions.cfm...");

  // Request every field we know about from the canonical schema
  // plus wildcard patterns to catch anything new
  // The server will only return fields that exist for this client
  var probeColumns = [
    // Identity
    "VCHLEADNUMBER","VCHPROJECTNAME","ILEADID",
    // Client / Owner
    "COMPANY","OWNERCOMPANY","OWNER","OWNERCONTACT","OWNERCONTACTID",
    // Stage
    "STAGENAME","STAGEID","ACTIVEIND","DAYSINSTAGE","SALESCYCLE",
    // Financials
    "IFIRMFEE","IFACTOREDFEE","IPROBABILITY","TOTALESTIMATEDFEESTOREDVALUE",
    "ICOST","IMARKETBUDGET","IMARKETACTUAL","IFEE",
    "FACTOREDFEESTOREDVALUE","GROSSMARGINDOLLARSSTD","GROSSMARGINPERCENTSTD",
    "GROSSREVENUESTD","FACTOREDCOSTSTD","MARKUP","LABORDIFFERENTIAL",
    "ESTIMATEDCOSTSTOREDVALUE","IFIRMFEEORIGSTOREDVALUE",
    "TOTALESTIMATEDFEESTOREDVALUE","FIRMESTIMATEDFEESTOREDVALUE",
    // Firm Orgs
    "OFFICELIST","DIVISIONLIST","STUDIOLIST","PRACTICEAREALIST",
    "TERRITORYLIST","OFFICEDIVISIONLIST",
    // Classification
    "PRIMARYCATEGORYLIST","PRIMARYCATEGORYNAMELIST",
    "SECONDARYCATEGORYLIST","SECONDARYCATEGORYNAMELIST","CONTRACTTYPES",
    "DELIVERYMETHOD","CLIENTTYPES","SERVICETYPES","PROSPECTTYPES",
    "ROLENAME","ROLEID","SUBMITTALTYPENAME","SOLICITATIONNUMBER",
    "CHRFPREC","CHPROPOSALSUB","NAICSCODES","SF330FORM","SF255FORM",
    "OPPTYPE","ISUBMITTALTYPEID","IPROJECTROLEID",
    // People
    "BUSINESSDEVELOPERID","PRINCIPALINCHARGEID","REFERREDBY","REFERREDBYID",
    "CREATORNAME","DECISIONMAKERDESCRIPTION",
    // Dates — actual
    "DTCLOSEDATE","SHORTLISTDATE","CONTRACTDATE","CREATEDATE","MODDATE",
    // Dates — estimated / forward
    "DTRFPDATE","DTQUALSDATE","DTPROPOSALDATE","DTPRESENTATIONDATE",
    "DTSTARTDATE","ESTIMATEDSTARTDATE","ESTIMATEDCOMPLETIONDATE",
    "PRECONSTARTDATE","PRECONENDDATE","DESIGNSTARTDATE","DESIGNCOMPLETIONDATE",
    "CONSTRUCTIONSTARTDATE","CONSTRUCTIONCOMPLETIONDATE",
    "OPPLEADBIDDATE","OPPLEADORIGDATE","DTONCALLSTART","DTONCALLEND",
    // Location
    "VCHADDRESS1","VCHADDRESS2","VCHCITY","STATEABRV","ISTATEID",
    "VCHPOSTALCODE","VCHPOSTALCODEEXT","VCHCOUNTRY","ICOUNTRYID",
    "REGIONID","COUNTY","LOCATION",
    // Scores / intel
    "IPROBABILITY","IPROJECTPROBABILITY","FUNDPROBABILITY","SCORE","REDZONESCORE",
    // Narrative
    "DESCRIPTION","TXCOMMENTS","TXNOTE","TSNOTES","VCHNEXTACTION",
    "PROJECTNARRATIVE","PROJECTSTRATEGY","CHALLENGESHURDLES","COMPETITIONANALYSIS",
    // Size / hours
    "ISIZE","VCHSIZEUNIT","ESTIMATEDMANAGEMENTUNITS","MANAGEMENTUNITRETURN",
    "WORKHOURSCONSTRUCTION","WORKHOURSDESIGN","WORKHOURSENGINEER",
    "OPPSELFPERFORMHOURS",
    // Marketing
    "IMARKETBUDGET","IMARKETACTUAL",
    // Lead/custom flex fields — dates
    "LEADDATE1","LEADDATE2","LEADDATE3","LEADDATE4","LEADDATE5",
    // Lead/custom flex fields — money
    "LEADMONEY1","LEADMONEY2","LEADMONEY3","LEADMONEY4","LEADMONEY5",
    // Lead/custom flex fields — numbers
    "LEADNUMBER1","LEADNUMBER2","LEADNUMBER3","LEADNUMBER4","LEADNUMBER5",
    // Lead/custom flex fields — percent
    "LEADPERCENT1","LEADPERCENT2","LEADPERCENT3","LEADPERCENT4","LEADPERCENT5",
    // Lead/custom flex fields — text
    "LEADSHORTTEXT1","LEADSHORTTEXT2","LEADSHORTTEXT3","LEADSHORTTEXT4","LEADSHORTTEXT5",
    "LEADLONGTEXT1","LEADLONGTEXT2","LEADLONGTEXT3","LEADLONGTEXT4","LEADLONGTEXT5",
    // Lead/custom flex fields — dropdowns
    "LEADVALUELISTID1","LEADVALUELISTNAME1","LEADVALUELISTID2","LEADVALUELISTNAME2",
    "LEADVALUELISTID3","LEADVALUELISTNAME3","LEADVALUELISTID4","LEADVALUELISTNAME4",
    "LEADVALUELISTID5","LEADVALUELISTNAME5",
    // System
    "SUBCOUNT","MASTERLEADID","MASTERIND","DELETERECORD","OLD_ID","IMPORTEDRECORD",
    "FIRMID","CFIRMID","ROWNUMBER","TOTALRECORDS",
    "SELECTEDCURRENCY","SELECTEDCURRENCYABBR","SELECTEDCURRENCYSYMBOL","SELECTEDRATE",
    "BASECURRENCY","BASECURRENCYABBR","BASECURRENCYSYMBOL",
    "OPPLEADCREATEMETHODID","OPPLEADCREATEMETHODSUBID","OPPLEADDODGEDATAID",
    // Staff team role columns — pattern STAFFROLE_{roleId}
    // We include a broad set of known role ID ranges; the actual IDs are
    // resolved from staffTeam.cfc after the probe. Any that return data
    // are captured; the full set is added in the schema build step.
    "STAFFROLE_253241","STAFFROLE_253242","STAFFROLE_253243",
    "STAFFROLE_258412","STAFFROLE_258413","STAFFROLE_258414","STAFFROLE_259155",
  ];

  var bodyParts = [
    "action=getOpportunityGridData","json=1","start=0","limit=1",
    "ActiveInd=0","SalesCycle=NaN",
    "officeId=0","divisionId=0","studioId=0","practiceAreaId=0",
    "territoryId=0","stageId=0","priCatId=0","secCatId=0",
    "masterSub=0","staffRoleId=0","dateCreated=0","dateModified=0",
    "dateCreatedModified=0","filteredSearch=0","search="
  ];
  probeColumns.forEach(function(c){
    // UUIDs must NOT be encoded — the server requires raw UUID format with hyphens
    // Standard field names are safe to encode
    var isUUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(c);
    bodyParts.push("visibleColumns=" + (isUUID ? c.toLowerCase() : encodeURIComponent(c)));
  });

  var data = await fetchJSON(oppBase + "oppActions.cfm", {
    method:"POST", credentials:"include",
    headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest"},
    body: bodyParts.join("&")
  });

  if(!data || !Array.isArray(data.DATA) || !data.DATA.length){
    UI.log("Probe returned no records — using full column list", "lw");
    return new Set(probeColumns);
  }

  // The server returns ONLY the fields that exist for this client
  // Extract the actual keys from the first record
  var availableFields = new Set(Object.keys(data.DATA[0]));
  UI.log("✓ Probe: server returned " + availableFields.size + " fields for this client", "ls");
  return availableFields;
}

// ─────────────────────────────────────────────────────────────
// STEP 2: BUILD CLIENT CONFIG from firmData.cfc
// Returns the complete per-client configuration dictionary
// ─────────────────────────────────────────────────────────────
async function buildClientConfig(firmDataBase){
  var config = {
    oppLabels:     {},   // defaultFieldName → clientLabel  (from oppLabels)
    firmOrgLabels: {},   // firmOrgId → clientLabel          (from labels)
    customFields:  [],   // [{uuid, label, type}]            (from customFieldConfigs)
    enabledFields: new Set(), // field names enabled for this client (from enabledOppFields)
    lookups: {
      firmOrg: {}, stage: {}, prospect: {}, role: {},
      contract: {}, delivery: {}, clientType: {}, submittal: {},
      priCat: {}, secCat: {}
    }
  };

  UI.log("Calling firmData.cfc (no method param — full config payload)...");

  // Try bare call first — returns full oppLabels, customFieldConfigs, enabledOppFields
  var fd = await fetchJSON(firmDataBase, {
    credentials:"include", headers:{"X-Requested-With":"XMLHttpRequest"}
  });

  // Fallback to GetFirmOrgData variant (partial — org data only)
  if(!fd){
    UI.log("Bare call failed, trying GetFirmOrgData...", "lw");
    fd = await fetchJSON(firmDataBase + "?method=GetFirmOrgData", {
      credentials:"include", headers:{"X-Requested-With":"XMLHttpRequest"}
    });
  }

  if(!fd){
    UI.log("firmData.cfc unavailable — labels will be raw field names", "lw");
    return config;
  }

  // ── oppLabels: opp-specific field label remapping ──────────
  var oppLabelRows = parseCFC(fd.oppLabels || {});
  oppLabelRows.forEach(function(r){
    var key = String(r.FIELDNAME || r.fieldName || "").trim();
    var val = String(r.CUSTOMFIELDNAME || r.customFieldName || "").trim();
    if(key && val) config.oppLabels[key] = val;
  });
  UI.log("✓ Opp labels: " + Object.keys(config.oppLabels).length, "ls");

  // ── labels: firm org hierarchy label remapping ─────────────
  var labelRows = parseCFC(fd.labels || {});
  labelRows.forEach(function(r){
    var id  = String(r.FIRMORGID || "").trim();
    var lbl = String(r.CUSTOMFIRMORGNAME || r.FIRMORGNAME || "").trim();
    if(id && lbl) config.firmOrgLabels[id] = lbl;
  });
  UI.log("✓ Firm Org labels: " + Object.keys(config.firmOrgLabels).length, "ls");

  // ── enabledOppFields: which fields are active for this client
  var enabledArr = Array.isArray(fd.enabledOppFields) ? fd.enabledOppFields : [];
  enabledArr.forEach(function(f){
    // Each entry has {name, enabled, ...}
    var name    = String(f.name || f.NAME || f.fieldName || "").trim().toLowerCase();
    var enabled = f.enabled !== false && f.enabled !== 0 && f.enabled !== "false";
    if(name && enabled) config.enabledFields.add(name);
  });
  UI.log("✓ Enabled fields: " + config.enabledFields.size, "ls");

  // ── customFieldConfigs: UUID-keyed custom fields ───────────
  var cfArr = Array.isArray(fd.customFieldConfigs) ? fd.customFieldConfigs : [];
  cfArr.forEach(function(cf){
    var uuid  = String(cf.DefinitionId || cf.ExternalId || cf.externalId || "").trim();
    var label = String(cf.Label || cf.label || "").trim();
    if(!uuid || !label) return;

    // Presence in customFieldConfigs is the enabled signal — Unanet does not
    // include deactivated custom fields here. GridSettings.EnabledInTheGrid is
    // a grid display preference only and is intentionally ignored: a user may
    // hide a custom field from their grid view while the field remains active
    // and populated on the record. We include all custom fields found here.

    // Map FieldType to our internal type system
    // FieldType values: Date, NumberCurrency, NumberInteger, NumberDecimal,
    //                   NumberPercent, ShortText, LongText, ValueList,
    //                   SelectSingle, SelectMultiple, TextHyperlink, TextRich
    var rawType = String(cf.FieldType || cf.fieldType || cf.CustomFieldTypeName || "text").toLowerCase();
    var type = rawType.includes("currency") || rawType.includes("money") ? "currency"
             : rawType.includes("percent")                               ? "percent"
             : rawType.includes("date")                                  ? "date"
             : rawType.includes("number") || rawType.includes("decimal") ||
               rawType.includes("integer")                               ? "number"
             : rawType.includes("longtext") || rawType.includes("long") ||
               rawType.includes("rich")                                  ? "longtext"
             : rawType.includes("selectsingle") || rawType.includes("select") ? "select"
             : "text";

    config.customFields.push({uuid:uuid, label:label, type:type});
  });
  UI.log("✓ Custom fields: " + config.customFields.length, "ls");

  // ── Firm Org ID → Name lookup (offices, divisions, etc.) ───
  var orgSections = ["offices","divisions","studios","practiceAreas","territories","officeDivisions","regions"];
  orgSections.forEach(function(key){
    var section = fd[key]; if(!section) return;
    var rows = Array.isArray(section) ? section
             : Array.isArray(section.DATA) ? section.DATA : [];
    rows.forEach(function(row){
      var id, name;
      if(Array.isArray(row)){ id=String(row[0]||"").trim(); name=String(row[1]||"").trim(); }
      else{ id=String(row.FIRMORGID||row.ID||row.id||"").trim(); name=String(row.FIRMORGNAME||row.NAME||row.name||"").trim(); }
      if(id && id!=="0" && name) config.lookups.firmOrg[id]=name;
    });
  });
  UI.log("✓ Firm Org lookup: " + Object.keys(config.lookups.firmOrg).length + " entries", "ls");

  // ── Pre-seed lookups from firmData where available ────────────
  // These serve as fallbacks when .cfc endpoints return 404 on some instances.
  // opportunityContactRoles = role lookup (Prime/Sub/JV)
  var ocr = parseCFC(fd.opportunityContactRoles || {});
  if(ocr.length){
    ocr.forEach(function(r){
      var id=String(r.ROLEID||r.roleid||"").trim();
      var name=String(r.ROLENAME||r.rolename||"").trim();
      if(id&&name) config.lookups.roleFromFirmData = config.lookups.roleFromFirmData||{};
      if(id&&name) config.lookups.roleFromFirmData[id]=name;
    });
  }

  // staffRoles in firmData (some instances expose this here too)
  var fsr = parseCFC(fd.staffRoles || {});
  if(fsr.length){
    config.lookups.staffRolesFromFirmData = {};
    fsr.forEach(function(r){
      var id=String(r.STAFFROLEID||r.staffroleid||"").trim();
      var name=String(r.STAFFROLENAME||r.staffrolename||"").trim();
      if(id&&name) config.lookups.staffRolesFromFirmData[id]=name;
    });
    UI.log("✓ Staff roles from firmData: " + Object.keys(config.lookups.staffRolesFromFirmData).length, "ls");
  }

  // ── Build SelectSingle custom field value maps ────────────────
  // customFieldConfigs with FieldType=SelectSingle contain SelectValues:
  // [{Key: uuid, Value: displayName}] — build lookup map per field UUID
  config.lookups.customSelectValues = {};
  var cfArr2 = Array.isArray(fd.customFieldConfigs) ? fd.customFieldConfigs : [];
  cfArr2.forEach(function(cf){
    var uuid = String(cf.DefinitionId||"").trim().toLowerCase();
    var sv = Array.isArray(cf.SelectValues) ? cf.SelectValues : [];
    if(uuid && sv.length){
      var map = {};
      sv.forEach(function(item){
        var k = String(item.Key||item.key||"").trim().toLowerCase();
        var v = String(item.Value||item.value||"").trim();
        if(k&&v) map[k]=v;
      });
      if(Object.keys(map).length) config.lookups.customSelectValues[uuid] = map;
    }
  });
  if(Object.keys(config.lookups.customSelectValues).length){
    UI.log("✓ Custom select value maps: " + Object.keys(config.lookups.customSelectValues).length, "ls");
  }

  return config;
}

// ─────────────────────────────────────────────────────────────
// STEP 3: LOAD LOOKUP TABLES from .cfc endpoints
// ─────────────────────────────────────────────────────────────
async function loadLookupTables(oppBase, config){
  UI.log("Loading reference lookup tables...");
  // staffRoles is used to resolve STAFFROLE_{id} field labels
  config.lookups.staffRoles = {};

  var lookupBase = findLookupBase();
  if(lookupBase){
    UI.log("✓ Lookup base from resources: " + lookupBase, "ls");
  } else {
    lookupBase = oppBase;
    UI.log("Lookup base not in resources — using oppBase fallback", "lw");
  }

  async function getLookup(file, method){
    var bases = (lookupBase !== oppBase) ? [lookupBase, oppBase] : [oppBase];
    for(var b=0; b<bases.length; b++){
      var url = bases[b] + file;
      var d = await postForm(url, {method:method});
      if(d && (Array.isArray(d)||d.DATA||d.COLUMNS)) return d;
      d = await fetchJSON(url + "?method=" + method);
      if(d && (Array.isArray(d)||d.DATA||d.COLUMNS)) return d;
    }
    return null;
  }

  // Fire only the lookups that actually contribute unique data.
  // Eliminated:
  //   stage        — STAGENAME already in opp record, no ID resolution needed
  //   role         — firmData.opportunityContactRoles has same data (pre-seeded above)
  //   clientType   — 404s on all tested instances
  //   submittalType — server returns SUBMITTALTYPENAME (resolved), not an ID
  //   primaryCategory / secondaryCategory — server returns CATEGORYNAMELIST (resolved)
  var results = await Promise.all([
    getLookup("oppData.cfc",       "getProspectTypes"),   // 0: prospect types
    getLookup("contractType.cfc",  "getContractTypes"),   // 1: contract types
    getLookup("deliveryMethod.cfc","getDeliveryMethods"), // 2: delivery methods
    getLookup("staffTeam.cfc",     "getStaffTeamRoles"),  // 3: staff role ID → name
  ]);

  config.lookups.prospect = buildLookup(parseCFC(results[0]), "ID",              "DISPLAYNAME");
  config.lookups.contract = buildLookup(parseCFC(results[1]), "CONTRACTTYPEID",  "CONTRACTNAME");
  config.lookups.delivery = buildLookup(parseCFC(results[2]), "DELIVERYMETHODID","DELIVERYMETHODNAME");

  // Role lookup — firmData.opportunityContactRoles pre-seeded in buildClientConfig
  if(config.lookups.roleFromFirmData && Object.keys(config.lookups.roleFromFirmData).length){
    config.lookups.role = config.lookups.roleFromFirmData;
    UI.log("✓ Role lookup: " + Object.keys(config.lookups.role).length + " roles (firmData)", "ls");
  }

  // Category lookups — server resolves names via CATEGORYNAMELIST fields directly
  // Keep empty maps so ID_RESOLUTION fallback works on instances where NAMELIST unavailable
  config.lookups.priCat = {};
  config.lookups.secCat = {};
  config.lookups.clientType = {};
  config.lookups.submittal  = {};

  // Staff role ID → role name
  var staffRoleRaw = parseCFC(results[3]);
  config.lookups.staffRoles = buildLookup(staffRoleRaw, "STAFFROLEID", "STAFFROLENAME");
  if(!Object.keys(config.lookups.staffRoles).length && config.lookups.staffRolesFromFirmData){
    config.lookups.staffRoles = config.lookups.staffRolesFromFirmData;
    UI.log("  Staff roles: using firmData fallback", "lw");
  }
  UI.log("✓ Staff roles: " + Object.keys(config.lookups.staffRoles).length, "ls");

  var total = Object.values(config.lookups).reduce(function(t,m){
    return t + (m && typeof m==="object" && !Array.isArray(m) ? Object.keys(m).length : 0);
  }, 0);
  UI.log("✓ Total lookup entries: " + total, "ls");
  return config;
}

// ─────────────────────────────────────────────────────────────
// STEP 4: BUILD DYNAMIC FIELD SCHEMA
//
// INCLUSION LOGIC — in priority order:
//
// 1. CUSTOM FIELDS (UUID keys from customFieldConfigs):
//    Always included if they appeared in the probe response.
//    Custom fields were explicitly created by the client — they
//    are always intentional.
//
// 2. STANDARD FIELDS — included if ALL of the following are true:
//    a) The field appeared in the probe response (server returned it)
//    b) The field is NOT structural noise (pagination/session artifacts)
//    c) The field IS in enabledOppFields from firmData.cfc
//       — OR — enabledOppFields was empty/unavailable (fallback: include all)
//
// There is no other suppression. No name-pattern filtering.
// No guessing about what might be meaningful.
// If the client enabled it, we include it.
// ─────────────────────────────────────────────────────────────
function buildFieldSchema(availableFields, config){
  var schema = [];
  var enabledKnown = config.enabledFields.size > 0;

  // Build a fast lookup: enabledFields entries → set of lowercase strings
  // enabledOppFields uses Unanet's internal field name which often matches
  // the backend key directly (e.g. "firmfee", "probability", "officeid")
  // We normalise both sides for matching
  var enabledLower = config.enabledFields; // already lowercase Set

  // Core identity fields that must always be present for the sheet to be useful
  // regardless of enabled state — these are structural to the extract, not analytical
  // Core identity fields that must always be present for the sheet to be useful.
  // These pass through regardless of enabled state because without them
  // the output isn't a readable pipeline report.
  // Note: ILEADID is intentionally excluded — it's an internal DB key.
  // VCHLEADNUMBER (Opp Number) is the user-facing identifier.
  var ALWAYS_INCLUDE = new Set([
    "VCHLEADNUMBER","VCHPROJECTNAME","STAGENAME","ACTIVEIND",
    "COMPANY","OWNERCOMPANY","OWNER","DAYSINSTAGE",
    "IFIRMFEE","IFACTOREDFEE","IPROBABILITY"
  ]);

  // ── Standard fields ──────────────────────────────────────────
  availableFields.forEach(function(backendKey){
    var upper = backendKey.toUpperCase();

    // Skip UUID custom fields — handled in the next block
    if(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(backendKey)) return;

    // Skip structural noise — pagination artifacts and currency session metadata
    // This is the ONLY name-based exclusion. See STRUCTURAL_NOISE definition.
    if(isStructuralNoise(upper)) return;

    // Enabled gate — the primary inclusion filter
    var isAlways  = ALWAYS_INCLUDE.has(upper);
    var isEnabled = !enabledKnown || isAlways || isFieldEnabled(upper, backendKey, enabledLower);

    if(!isEnabled){
      return; // disabled fields are silently excluded
    }

    var label       = resolveFieldLabel(upper, config.oppLabels, config.firmOrgLabels);
    var type        = getType(upper);
    var resolveWith = ID_RESOLUTION[upper] || null;

    schema.push({
      backendKey: backendKey,
      upper:      upper,
      label:      label,
      type:       type,
      isCustom:   false,
      resolve:    resolveWith
    });
  });

  // ── Custom fields (UUID keys) ─────────────────────────────────
  // Source of truth: customFieldConfigs from firmData.cfc.
  // ALL custom fields are included in schema regardless of whether
  // the server returned data for them. If the server doesn't return
  // a UUID (Unanet limitation on certain LEAD slot mappings), the
  // column will exist but be empty — which is correct behaviour.
  config.customFields.forEach(function(cf){
    var uuidLower = cf.uuid.toLowerCase().trim();
    schema.push({
      backendKey: uuidLower,
      upper:      "CUSTOM_" + cf.label.toUpperCase().replace(/[^A-Z0-9]/g,"_"),
      label:      cf.label,
      type:       cf.type,
      isCustom:   true,
      resolve:    null
    });
    UI.log("  ✓ Custom field: " + cf.label + " [" + cf.type + "]", "ls");
  });

  // ── Sort by COLUMN_ORDER preference, then alphabetically ─────
  schema.sort(function(a, b){
    var ai = COLUMN_ORDER.indexOf(a.upper);
    var bi = COLUMN_ORDER.indexOf(b.upper);
    if(ai === -1 && bi === -1) return a.label.localeCompare(b.label);
    if(ai === -1) return 1;
    if(bi === -1) return -1;
    return ai - bi;
  });

  var customCount = schema.filter(function(f){ return f.isCustom; }).length;
  UI.log("✓ Field schema: " + schema.length + " fields (" + customCount + " custom)", "ls");
  if(!enabledKnown){
    UI.log("  ℹ enabledOppFields unavailable — all non-noise fields included", "lw");
  }

  return schema;
}

// Resolve STAFFROLE placeholder labels using the staffRoles lookup
// Called after both buildFieldSchema and loadLookupTables are complete
function resolveStaffRoleLabels(schema, staffRolesLookup){
  var resolved = 0;
  schema.forEach(function(field){
    if(field.label && field.label.indexOf("__STAFFROLE__") === 0){
      var roleId = field.label.replace("__STAFFROLE__", "");
      var roleName = staffRolesLookup[roleId];
      if(roleName){
        field.label = roleName;
        resolved++;
      } else {
        // Role ID not in lookup — use a readable fallback
        field.label = "Staff Role " + roleId;
      }
    }
  });
  if(resolved > 0) UI.log("✓ Resolved " + resolved + " staff role labels", "ls");
  return schema;
}

// Build STAFFROLE visibleColumns from the staffRoles lookup
// Returns array of STAFFROLE_{id} column keys for all known roles
function buildStaffRoleColumns(staffRolesLookup){
  return Object.keys(staffRolesLookup).map(function(roleId){
    return "STAFFROLE_" + roleId;
  });
}

// Match a backend field key against the enabledOppFields set.
// enabledOppFields uses Unanet's internal short names (e.g. "firmfee", "probability")
// which often don't match the backend key exactly (e.g. "IFIRMFEE", "IPROBABILITY").
//
// STRATEGY: direct match, then prefix/suffix stripping only.
// No substring matching — it produces false positives (e.g. "cost" matching
// "factoredcoststd"). If we can't match by direct key or known prefix/suffix
// transformations, we don't include the field. The enabledOppFields list is
// comprehensive — if a field is enabled it will match via one of these paths.
function isFieldEnabled(upper, backendKey, enabledLower){
  var lower = backendKey.toLowerCase();

  // 1. Direct match
  if(enabledLower.has(lower)) return true;

  // 2. Strip common prefixes Unanet adds to backend keys
  //    and check the result against enabledOppFields
  var strippedPrefixes = [
    lower.replace(/^i(?=[a-z])/,""),           // IFIRMFEE → firmfee
    lower.replace(/^vch/,""),                 // VCHCITY → city
    lower.replace(/^dt(?=[a-z])/,""),         // DTCLOSEDATE → closedate
    lower.replace(/^chr/,""),                 // CHRFPREC → fprec
    lower.replace(/^tx/,""),                  // TXCOMMENTS → comments
    lower.replace(/^vch/,"").replace(/list$/,""), // VCHLIST variants
  ];
  for(var i=0; i<strippedPrefixes.length; i++){
    if(strippedPrefixes[i] && strippedPrefixes[i] !== lower && enabledLower.has(strippedPrefixes[i])){
      return true;
    }
  }

  // 3. Strip common suffixes and check
  var strippedSuffixes = [
    lower.replace(/list$/,""),      // OFFICELIST → office
    lower.replace(/list$/,"id"),    // OFFICELIST → officeid
    lower.replace(/name$/,""),      // STAGENAME → stage
    lower.replace(/name$/,"id"),    // STAGENAME → stageid
    lower.replace(/date$/,""),      // CLOSEDATE → close
    lower.replace(/date$/,""),      // ESTIMATEDCOMPLETIONDATE → estimatedcompletion
  ];
  // Also try removing 'estimated' prefix for eststartdate/estcompletiondate variants
  var estStripped = lower.replace(/^estimated/,"est").replace(/^estimated/,"");
  if(estStripped && estStripped !== lower) strippedSuffixes.push(estStripped);
  for(var j=0; j<strippedSuffixes.length; j++){
    if(strippedSuffixes[j] && strippedSuffixes[j] !== lower && enabledLower.has(strippedSuffixes[j])){
      return true;
    }
  }

  // 4. Combined prefix + suffix strip
  var core = lower
    .replace(/^i(?=[a-z])/,"")
    .replace(/^vch/,"")
    .replace(/^dt(?=[a-z])/,"")
    .replace(/list$/,"")
    .replace(/name$/,"");
  if(core && core !== lower && enabledLower.has(core)) return true;

  // 5. Known field mappings that don't fit the pattern rules above
  var KNOWN_MAPPINGS = {
    "estimatedstartdate":      "eststartdate",
    "estimatedcompletiondate": "estcompletiondate",
    "primarycategorylist":     "projectcategoryid",
    "primarycategorynamelist": "projectcategoryid",    // name list variant
    "secondarycategorylist":   "secondarycategoryid",
    "secondarycategorynamelist":"secondarycategoryid", // name list variant
    "practicearealist":        "practiceareaid",
    "studiolist":              "studioid",
    "officelist":              "officeid",
    "territorylist":           "territoryid",
    "divisionlist":            "divisionid",
    "contracttypes":           "contracttype",
    "clienttypes":             "clienttypeid",
    "prospecttypes":           "prospecttypes",
    "servicetypes":            "servicetypeid",
    "naicscodes":              "naicscode",
    "txcomments":              "txcomments",
    "txnote":                  "notes",
    "vchnextaction":           "nextaction",
    "vchaddress1":             "address1",
    "vchaddress2":             "address2",
    "vchcity":                 "city",
    "stateabrv":               "state",
    "vchpostalcode":           "postalcode",
    "vchcountry":              "country",
    "rolename":                "roleid",
    "submittaltypename":       "submittaltype",
    "chrfprec":                "chrfprec",
    "chproposalsub":           "proposalsubmitted",
    "solicitationnumber":      "solicitationnumber",
    "daysinstage":             "daysinstage",
    "createdate":              "leadnumber",   // always include
    "moddate":                 "leadnumber",   // always include
  };
  var mapped = KNOWN_MAPPINGS[lower];
  if(mapped && enabledLower.has(mapped)) return true;

  return false;
}

// Maps backend Firm Org field → firmOrgId for firmOrgLabels lookup
// firmOrgLabels is keyed by firmOrgId (1=Offices, 2=Divisions, 3=Studios,
// 4=Practice Areas, 5=Components, 6=OfficeDivision, 7=Territories)
var FIRM_ORG_ID_MAP = {
  "OFFICELIST":        "1",
  "DIVISIONLIST":      "2",
  "STUDIOLIST":        "3",
  "PRACTICEAREALIST":  "4",
  "OFFICEDIVISIONLIST":"6",
  "TERRITORYLIST":     "7",
};

// Resolve a backend field name to the client's configured label
function resolveFieldLabel(upperKey, oppLabels, firmOrgLabels){
  // oppLabels is keyed by Unanet's default UI label name (e.g. "Win Probability")
  // firmOrgLabels is keyed by firmOrgId (e.g. "1" → "Location", "4" → "Market Sector")
  // We need to map from backend key → default label → client label

  // ── Firm Org fields: check firmOrgLabels first ──────────────
  // These are configured at the org hierarchy level, not in oppLabels
  if(FIRM_ORG_ID_MAP[upperKey]){
    var orgId = FIRM_ORG_ID_MAP[upperKey];
    if(firmOrgLabels[orgId]) return firmOrgLabels[orgId];
  }

  var DEFAULT_UI_LABELS = {
    "VCHLEADNUMBER":           "Opportunity Number",
    "VCHPROJECTNAME":          "Opportunity Name",
    "COMPANY":                 "Client Company",
    "OWNERCOMPANY":            "Owner Company",
    "OWNER":                   "Opportunity Owner",
    "STAGENAME":               "Stage",
    "ACTIVEIND":               "Status",
    "DAYSINSTAGE":             "Days in Stage",
    "IFIRMFEE":                "Firm Estimated Fee",
    "IFACTOREDFEE":            "Factored Fee",
    "IPROBABILITY":            "Win Probability",
    "IPROJECTPROBABILITY":     "Project Probability",
    "TOTALESTIMATEDFEESTOREDVALUE": "Total Estimated Fee",
    "ICOST":                   "Estimated Cost",
    "IMARKETBUDGET":           "Marketing Cost - Budget",
    "IMARKETACTUAL":           "Marketing Cost - Actual",
    "OFFICELIST":              "Offices",
    "DIVISIONLIST":            "Divisions",
    "STUDIOLIST":              "Studios",
    "PRACTICEAREALIST":        "Practice Areas",
    "TERRITORYLIST":           "Territories",
    "OFFICEDIVISIONLIST":      "Office Division",
    "PRIMARYCATEGORYLIST":     "Primary Categories",
    "PRIMARYCATEGORYNAMELIST": "Primary Categories",  // resolved name version — preferred
    "SECONDARYCATEGORYLIST":   "Secondary Categories",
    "SECONDARYCATEGORYNAMELIST":"Secondary Categories", // resolved name version — preferred
    "CONTRACTTYPES":           "Contract Type",
    "DELIVERYMETHOD":          "Delivery Method",
    "CLIENTTYPES":             "Client Types",
    "SERVICETYPES":            "Service Types",
    "PROSPECTTYPES":           "Prospect Types",
    "ROLENAME":                "Opportunity Role",
    "SUBMITTALTYPENAME":       "Submittal Type",
    "SOLICITATIONNUMBER":      "Solicitation Number",
    "CHRFPREC":                "Bid",
    "CHPROPOSALSUB":           "Proposal Submitted",
    "NAICSCODES":              "NAICS Codes",
    "OWNERCONTACT":            "Primary Contact",
    "DTCLOSEDATE":             "Close Date",
    "DTRFPDATE":               "Expected RFP Date",
    "DTQUALSDATE":             "Quals Due Date",
    "DTPROPOSALDATE":          "Proposal Due Date",
    "DTPRESENTATIONDATE":      "Presentation Date",
    "DTSTARTDATE":             "Estimated Selection Date", // oppLabels key — resolves to client label e.g. "Expected Award Date", "Estimated Award Date"
    "ESTIMATEDSTARTDATE":      "Estimated Start Date",
    "ESTIMATEDCOMPLETIONDATE": "Estimated Completion Date",
    "PRECONSTARTDATE":         "PreCon Start Date",
    "PRECONENDDATE":           "PreCon End Date",
    "DESIGNSTARTDATE":         "Design Start Date",
    "DESIGNCOMPLETIONDATE":    "Design Completion Date",
    "CONSTRUCTIONSTARTDATE":   "Construction Start Date",
    "CONSTRUCTIONCOMPLETIONDATE":"Construction Completion Date",
    "SHORTLISTDATE":           "Shortlist Date",
    "OPPLEADBIDDATE":          "Lead Bid Date",
    "OPPLEADORIGDATE":         "Lead Origination Date",
    "CREATEDATE":              "Date Created",
    "MODDATE":                 "Last Modified",
    "VCHADDRESS1":             "Address",
    "VCHADDRESS2":             "Address 2",
    "VCHCITY":                 "City",
    "STATEABRV":               "State",
    "VCHPOSTALCODE":           "Postal Code",
    "VCHCOUNTRY":              "Country",
    "COUNTY":                  "County",
    "REGIONID":                "Region",
    "SCORE":                   "Score",
    "REDZONESCORE":            "Red Zone Score",
    "FUNDPROBABILITY":         "Funding Probability",
    "DESCRIPTION":             "Description",
    "TXCOMMENTS":              "Comments",
    "TXNOTE":                  "Notes",
    "TSNOTES":                 "TS Notes",
    "VCHNEXTACTION":           "Next Action",
    "PROJECTSTRATEGY":         "Opportunity Strategy",
    "CHALLENGESHURDLES":       "Challenges / Hurdles",
    "COMPETITIONANALYSIS":     "Competition Analysis",
    "PROJECTNARRATIVE":        "Project Narrative",
    "ISIZE":                   "Estimated Size",
    "VCHSIZEUNIT":             "Size Unit",
    "WORKHOURSCONSTRUCTION":   "Construction Work Hours",
    "WORKHOURSDESIGN":         "Design Work Hours",
    "WORKHOURSENGINEER":       "Engineering Work Hours",
    "OPPSELFPERFORMHOURS":     "Self Perform Hours",
    "ESTIMATEDMANAGEMENTUNITS":"Estimated Management Units",
    "MANAGEMENTUNITRETURN":    "Management Unit Return",
    "SUBCOUNT":                "Sub-Opportunity Count",
    "REFERREDBY":              "Referred By",
    "BUSINESSDEVELOPERID":     "Business Developer",
    "CREATORNAME":             "Created By",
    "DECISIONMAKERDESCRIPTION":"Decision Maker",
    "SF330FORM":               "SF330 Form",
    "SF255FORM":               "SF255 Form",
    "MARKUP":                  "Markup",
    "LABORDIFFERENTIAL":       "Labor Differential",
    "GROSSMARGINDOLLARSSTD":   "Gross Margin ($)",
    "GROSSMARGINPERCENTSTD":   "Gross Margin (%)",
    "GROSSREVENUESTD":         "Gross Revenue",
    "FACTOREDCOSTSTD":         "Factored Cost",
    "IFEE":                    "Total Estimated Fee",
  };

  // Check if oppLabels has a custom label for this field's default UI name
  var defaultLabel = DEFAULT_UI_LABELS[upperKey];
  if(defaultLabel && oppLabels[defaultLabel]){
    return oppLabels[defaultLabel];
  }

  // Fall back to the default label if we have one
  if(defaultLabel) return defaultLabel;

  // For STAFFROLE_{id} fields — resolve role name from staffRoles lookup
  // oppLabels won't have these; we need to use the staffRoles lookup
  // The lookup is passed in config but resolveFieldLabel doesn't have it
  // So we return a placeholder that gets resolved after schema build
  var staffMatch = upperKey.match(/^STAFFROLE_(\d+)$/);
  if(staffMatch){
    // Return placeholder — resolved to role name in buildFieldSchema
    return "__STAFFROLE__" + staffMatch[1];
  }
  // Skip STAFFROLE_{id}ID — already in structural noise
  if(/^STAFFROLE_\d+ID$/.test(upperKey)) return null;

  // For LEAD* flex fields — generate readable label
  var leadMatch = upperKey.match(/^LEAD(DATE|MONEY|NUMBER|PERCENT|SHORTTEXT|LONGTEXT|VALUELISTNAME)(\d+)$/);
  if(leadMatch){
    var typeNames = {DATE:"Date",MONEY:"Value",NUMBER:"Number",PERCENT:"Percent",
                    SHORTTEXT:"Text",LONGTEXT:"Long Text",VALUELISTNAME:"List"};
    return "Custom " + (typeNames[leadMatch[1]]||leadMatch[1]) + " " + leadMatch[2];
  }
  // Skip LEADVALUELISTID (internal IDs — we use VALUELISTNAME instead)
  if(/^LEADVALUELISTID\d+$/.test(upperKey)) return null;

  // Last resort: clean up the raw backend name
  return upperKey.replace(/_/g," ").replace(/([A-Z])([A-Z][a-z])/g,"$1 $2").toLowerCase()
    .replace(/\b\w/g,function(c){return c.toUpperCase();});
}

// ─────────────────────────────────────────────────────────────
// STEP 5: FETCH ALL OPPORTUNITIES
// No filters, no pagination, columns from schema only
// ─────────────────────────────────────────────────────────────
async function fetchAllOpportunities(oppBase, schema, customFieldUUIDs, filterSettings){
  var dateField  = (filterSettings && filterSettings.dateField)  || "CREATEDATE";
  var dateLabel  = (filterSettings && filterSettings.dateFieldLabel) || "Date Created";
  var years      = (filterSettings && filterSettings.years !== undefined) ? filterSettings.years : 3;
  var allTime    = years >= 999;

  // Compute the cutoff date
  var cutoffDate = null;
  if(!allTime){
    cutoffDate = new Date();
    if(years === 0){
      // Current year — Jan 1 of this year
      cutoffDate = new Date(cutoffDate.getFullYear(), 0, 1);
    } else {
      // Rolling N years back from today
      cutoffDate.setFullYear(cutoffDate.getFullYear() - years);
      cutoffDate.setHours(0,0,0,0);
    }
  }

  // Server-side date filtering only works for CREATEDATE and MODDATE
  // All other fields require client-side filtering after fetch
  var serverSideFilter = !allTime && (dateField === "CREATEDATE" || dateField === "MODDATE");
  var clientSideFilter = !allTime && !serverSideFilter;

  var windowLabel = allTime ? "All Time"
    : years === 0 ? "Current Year"
    : "Last " + years + " Years";
  UI.log("Fetching opportunities — " + dateLabel + " · " + windowLabel + "...");

  // Build the cutoff timestamp for server-side filtering
  // oppActions.cfm accepts dateCreated/dateModified as Unix timestamps (ms)
  var cutoffTs = cutoffDate ? cutoffDate.getTime() : 0;

  // Build the fixed base parameters — these never change between pages
  var baseParams = [
    "action=getOpportunityGridData","json=1","sort=STAGEID","dir=ASC",
    "selectedCurrency=USD","view=0",
    "ActiveInd=0",            // ALL statuses — never filter by status
    "SalesCycle=NaN",
    "officeId=0","divisionId=0","studioId=0","practiceAreaId=0",
    "territoryId=0","stageId=0","priCatId=0","secCatId=0",
    "masterSub=0","staffRoleId=0",
    // Server-side date filters — 0 means no filter
    "dateCreated=" + (serverSideFilter && dateField==="CREATEDATE" ? cutoffTs : 0),
    "dateModified=" + (serverSideFilter && dateField==="MODDATE" ? cutoffTs : 0),
    "dateCreatedModified=0","filteredSearch=0","search="
  ];

  // Build visibleColumns list from schema
  // UUIDs must NOT be encoded — server requires raw hyphens
  var colParams = [];
  schema.forEach(function(f){
    if(f.isCustom){
      colParams.push("visibleColumns=" + f.backendKey.toLowerCase());
    } else {
      colParams.push("visibleColumns=" + encodeURIComponent(f.backendKey));
    }
  });

  // Add any custom UUIDs not already in schema
  var schemaUUIDs = new Set(schema.filter(function(f){ return f.isCustom; })
    .map(function(f){ return f.backendKey.toLowerCase(); }));
  customFieldUUIDs.forEach(function(uuid){
    var uuidLower = uuid.toLowerCase();
    if(!schemaUUIDs.has(uuidLower)){
      colParams.push("visibleColumns=" + uuidLower);
    }
  });

  // ── Pagination strategy ───────────────────────────────────────
  // Unanet's grid endpoint has an implicit server-side page size limit
  // (observed at ~100 records on large instances regardless of limit= value).
  // We fetch in pages of 100 and concatenate until we have all records.
  // The ROWCOUNT field in the first response tells us the total.
  var PAGE_SIZE = 100;
  var allRecords = [];
  var totalExpected = null;
  var start = 0;
  var page = 0;
  var maxPages = 100; // safety ceiling: 100 pages × 100 records = 10,000 opps

  while(page < maxPages){
    var pageParams = baseParams.concat([
      "start=" + start,
      "limit=" + PAGE_SIZE
    ]).concat(colParams);

    var data = await fetchJSON(oppBase + "oppActions.cfm", {
      method:"POST", credentials:"include",
      headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest"},
      body: pageParams.join("&")
    });

    if(!data || !Array.isArray(data.DATA)){
      if(page === 0){
        UI.log("No data returned from oppActions.cfm", "le");
        return null;
      }
      // Subsequent page failure — stop here with what we have
      UI.log("⚠ Page " + (page+1) + " failed — stopping with " + allRecords.length + " records", "lw");
      break;
    }

    var pageRecords = data.DATA;
    allRecords = allRecords.concat(pageRecords);

    // First page: get the total record count
    if(page === 0){
      totalExpected = parseInt(data.ROWCOUNT) || pageRecords.length;
      UI.log("✓ Total opportunities on server: " + totalExpected, "ls");
      if(totalExpected <= PAGE_SIZE){
        // All records came back in one page — done
        UI.log("✓ Single page fetch complete", "ls");
        break;
      }
    }

    UI.log("  Page " + (page+1) + ": " + pageRecords.length + " records (total so far: " + allRecords.length + "/" + totalExpected + ")", "ls");

    // Check if we have everything
    if(allRecords.length >= totalExpected){
      break;
    }

    // Check if last page returned fewer records than requested (end of data)
    if(pageRecords.length < PAGE_SIZE){
      break;
    }

    start += PAGE_SIZE;
    page++;
  }

  UI.log("✓ " + allRecords.length + " of " + (totalExpected||allRecords.length) + " opportunities received", "ls");

  if(totalExpected && allRecords.length < totalExpected){
    UI.log("⚠ Expected " + totalExpected + " but only received " + allRecords.length, "lw");
  }

  // ── Client-side date filter ───────────────────────────────────
  if(clientSideFilter && cutoffDate){
    var before = allRecords.length;
    var noDateCount = 0;
    allRecords = allRecords.filter(function(rec){
      var raw = rec[dateField] || rec[dateField.toLowerCase()] || "";
      if(!raw || raw === ""){
        noDateCount++;
        return false; // exclude records with no date in the selected field
      }
      // Parse the date value
      var d = null;
      var s = String(raw).trim();
      if(/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)){
        var p=s.slice(0,10).split("/"); d=new Date(parseInt(p[2]),parseInt(p[0])-1,parseInt(p[1]));
      } else if(/^\d{4}-\d{2}-\d{2}/.test(s)){
        var p=s.slice(0,10).split("-"); d=new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));
      } else {
        var m=s.match(/^(\w+),\s*(\d+)\s+(\d{4})/);
        if(m) d=new Date(m[1]+" "+m[2]+","+m[3]);
      }
      if(!d || isNaN(d.getTime())) { noDateCount++; return false; }
      return d >= cutoffDate;
    });
    var excluded = before - allRecords.length;
    if(noDateCount > 0){
      UI.log("⚠ " + noDateCount + " records have no " + dateLabel + " and were excluded from this extract.", "lw");
    }
    if(excluded > noDateCount){
      UI.log("  Filtered " + (excluded - noDateCount) + " records outside the " + windowLabel + " window.", "ls");
    }
    UI.log("✓ " + allRecords.length + " records after date filter", "ls");
  }

  // Build a combined data object matching the original single-page structure
  var combinedData = {
    DATA: allRecords,
    ROWCOUNT: allRecords.length
  };

  // Log any custom field UUIDs that the server did not return
  var returnedKeys = new Set();
  if(allRecords.length > 0){
    Object.keys(allRecords[0]).forEach(function(k){ returnedKeys.add(k.toLowerCase()); });
  }
  customFieldUUIDs.forEach(function(uuid){
    if(!returnedKeys.has(uuid.toLowerCase())){
      UI.log("  ℹ Server did not return UUID: ..."+uuid.slice(-8)+" (Unanet slot limitation)", "lw");
    }
  });

  return combinedData;
}

// ─────────────────────────────────────────────────────────────
// STEP 6: NORMALIZE RECORDS
// Apply type formatting, ID resolution, and label mapping
// entirely from the dynamic schema — no hardcoded field names
// ─────────────────────────────────────────────────────────────
function normalizeRecord(opp, schema, config){
  var row = {};

  // Computed meta fields
  row["__Status"]   = classifyStatus(opp.STAGENAME||"", opp.ACTIVEIND);
  row["__Fee"]      = parseFloat(opp.IFIRMFEE)||0;
  row["__Weighted"] = parseFloat(opp.IFACTOREDFEE)||0;
  row["__Pwin"]     = (opp.IPROBABILITY!=null&&opp.IPROBABILITY!=="")
                      ? parseFloat(opp.IPROBABILITY)/100 : null;
  row["__Stage"]    = String(opp.STAGENAME||"");

  // Build a lowercase key → value map of the raw opp record once
  // so UUID lookups are case-insensitive without repeated scanning
  var oppLower = {};
  Object.keys(opp).forEach(function(k){ oppLower[k.toLowerCase()] = opp[k]; });

  schema.forEach(function(field){
    // For custom UUID fields: always look up by lowercase key
    // For standard fields: try original, then uppercase, then lowercase
    var val;
    if(field.isCustom){
      val = oppLower[field.backendKey.toLowerCase()];
    } else {
      val = opp[field.backendKey];
      if(val===undefined||val===null||val===""){
        val = opp[field.backendKey.toUpperCase()] ||
              opp[field.backendKey.toLowerCase()] || "";
      }
    }

    if(val===null||val===undefined||val==="") return;

    var display;

    // Custom field — UUID key may be returned in different case by server
    if(field.isCustom){
      if(field.type==="date"){
        display = fmtDate(val);
      } else if(field.type==="currency"){
        display = parseFloat(val)||0;
      } else if(field.type==="percent"){
        display = (parseFloat(val)||0)/100;
      } else if(field.type==="number"){
        display = parseFloat(val)||0;
      } else if(field.type==="select"){
        // SelectSingle — val is a UUID key, resolve to display value
        var selectMap = (config.lookups.customSelectValues||{})[field.backendKey.toLowerCase()];
        if(selectMap){
          display = selectMap[String(val).trim().toLowerCase()] || String(val).trim();
        } else {
          display = String(val).trim();
        }
      } else {
        display = String(val).trim();
      }
      // Truncate long text — Excel hard limit is 32,767 chars per cell
      if(display && typeof display === "string" && display.length > 2000){
        display = display.slice(0, 2000) + "…";
      }
      // Store even if empty string so the column appears — only skip null/undefined
      if(display!==null&&display!==undefined) row[field.label] = display;
      return;
    }

    // Standard field — apply resolution and type formatting
    if(field.resolve){
      // CATEGORYNAMELIST fields are already resolved by the server — treat as text
      if(field.upper.endsWith("CATEGORYNAMELIST")){
        display = String(val).trim();
      } else {
        // ID resolution field
        display = resolveIDs(val, config.lookups[field.resolve]||{});
        if(!display) display = String(val).trim();
      }
    } else if(field.type==="date"){
      display = fmtDate(val);
    } else if(field.type==="currency"){
      display = parseFloat(val)||0;
    } else if(field.type==="percent"){
      // Probability comes in as 0-100, convert to decimal for Excel %
      display = (parseFloat(val)||0)/100;
    } else if(field.type==="number"){
      display = parseFloat(val)||0;
    } else if(field.type==="html"){
      display = stripHTML(String(val)).split("|")[0].trim();
    } else if(field.type==="flag"){
      // Normalize boolean flags
      var sv = String(val);
      display = sv==="1"||sv.toLowerCase()==="yes"||sv.toLowerCase()==="true" ? "Yes"
              : sv==="0"||sv.toLowerCase()==="no"||sv.toLowerCase()==="false" ? "No"
              : sv;
    } else {
      display = String(val).trim();
    }

    // Truncate long text fields — Excel hard limit is 32,767 chars per cell
    // Long text fields are useful for reference but not for analysis;
    // truncate at 2,000 chars which is ample for any meaningful content
    if(display && typeof display === "string" && display.length > 2000){
      display = display.slice(0, 2000) + "…";
    }

    if(display!==""&&display!=null&&display!==undefined){
      // Use client label as the row key — this is what appears in the sheet header
      row[field.label] = display;
    }
  });

  return row;
}


// ─────────────────────────────────────────────────────────────
// EXEC SUMMARY SHEET
// Derived entirely from cleanRows — no additional data calls.
// All labels used here are the client's configured labels from
// the schema, not hardcoded strings, so they reflect whatever
// the client has renamed their fields to.
// ─────────────────────────────────────────────────────────────
function buildExecSummarySheet(rows, schema, config){

  // ── Field label resolver ─────────────────────────────────────
  function clientLabel(backendKey){
    var upper = backendKey.toUpperCase();
    for(var i=0;i<schema.length;i++){
      if(schema[i].upper===upper||schema[i].backendKey.toUpperCase()===upper)
        return schema[i].label;
    }
    return null;
  }

  var lbl = {
    ourFee:      clientLabel("IFIRMFEE")      || "Our Fee",
    weighted:    clientLabel("IFACTOREDFEE")  || "Weighted Value",
    pwin:        clientLabel("IPROBABILITY")  || "Pwin",
    stage:       clientLabel("STAGENAME")     || "Stage",
    client:      clientLabel("COMPANY")       || "Client Company",
    daysInStage: clientLabel("DAYSINSTAGE")   || "Days in Stage",
    estAward:    clientLabel("DTSTARTDATE")   || "Estimated Award Date",
  };

  // ── Calculations ─────────────────────────────────────────────
  function sum(arr,key){ return arr.reduce(function(t,r){ var v=r[key]; return t+(v!=null&&v!==""?parseFloat(v)||0:0); },0); }
  function avg(arr,key){ var vs=arr.filter(function(r){ return r[key]!=null&&r[key]!==""; }); if(!vs.length) return null; return vs.reduce(function(t,r){ return t+(parseFloat(r[key])||0); },0)/vs.length; }
  function pct(n,d){ return d>0?n/d:null; }
  function fmtCur(v){ if(v==null||v==="") return ""; return "$"+Math.round(v).toLocaleString("en-US"); }
  function fmtPct(v){ if(v==null||v==="") return ""; return (v*100).toFixed(1)+"%"; }
  function fmtNum(v){ if(v==null||v==="") return ""; return Math.round(v).toLocaleString("en-US"); }

  var today = new Date(); today.setHours(0,0,0,0);
  var d90   = new Date(today.getTime()+90*24*60*60*1000);

  var active   = rows.filter(function(r){ return r["__Status"]==="Active"; });
  var won      = rows.filter(function(r){ return r["__Status"]==="Won"; });
  var lost     = rows.filter(function(r){ return r["__Status"]==="Lost"; });
  var closed   = rows.filter(function(r){ return r["__Status"]==="Closed"; });
  var resolved = won.concat(lost);

  var totalFee      = sum(rows,   lbl.ourFee);
  var activeFee     = sum(active, lbl.ourFee);
  var wonFee        = sum(won,    lbl.ourFee);
  var lostFee       = sum(lost,   lbl.ourFee);
  var weightedTotal = sum(active, lbl.weighted);

  var stagnant60 = active.filter(function(r){ return (parseFloat(r[lbl.daysInStage])||0)>=60; });
  var stagnant90 = active.filter(function(r){ return (parseFloat(r[lbl.daysInStage])||0)>=90; });

  var forecast90 = active.filter(function(r){
    var v=r[lbl.estAward]; if(!v) return false;
    var d=v instanceof Date?v:new Date(v);
    return !isNaN(d.getTime())&&d>=today&&d<=d90;
  });
  var overdue = active.filter(function(r){
    var v=r[lbl.estAward]; if(!v) return false;
    var d=v instanceof Date?v:new Date(v);
    return !isNaN(d.getTime())&&d<today;
  });
  var noDate = active.filter(function(r){ var v=r[lbl.estAward]; return !v||v===""; });
  var noPwin = active.filter(function(r){ var v=r[lbl.pwin]; return v===null||v===undefined||v===""; });

  var stageGroups={};
  rows.forEach(function(r){
    var s=String(r[lbl.stage]||"Unknown"), st=r["__Status"]||"";
    if(!stageGroups[s]) stageGroups[s]={stage:s,status:st,count:0,fee:0,weighted:0};
    stageGroups[s].count++;
    stageGroups[s].fee+=parseFloat(r[lbl.ourFee])||0;
    stageGroups[s].weighted+=parseFloat(r[lbl.weighted])||0;
  });
  var stageList=Object.values(stageGroups).sort(function(a,b){
    var ord={Active:0,Won:1,Lost:2,Closed:3};
    var ao=ord[a.status]!=null?ord[a.status]:9, bo=ord[b.status]!=null?ord[b.status]:9;
    return ao!==bo?ao-bo:b.fee-a.fee;
  });

  var clientGroups={};
  rows.forEach(function(r){
    var c=String(r[lbl.client]||"Unknown").trim();
    if(!clientGroups[c]) clientGroups[c]={count:0,fee:0,weighted:0,won:0,lost:0,active:0};
    clientGroups[c].count++;
    clientGroups[c].fee+=parseFloat(r[lbl.ourFee])||0;
    clientGroups[c].weighted+=parseFloat(r[lbl.weighted])||0;
    if(r["__Status"]==="Won")    clientGroups[c].won++;
    if(r["__Status"]==="Lost")   clientGroups[c].lost++;
    if(r["__Status"]==="Active") clientGroups[c].active++;
  });
  var topClients=Object.entries(clientGroups)
    .sort(function(a,b){ return b[1].fee-a[1].fee; })
    .slice(0,10);

  // ── Sheet builder ────────────────────────────────────────────
  // Single-column layout: each section stacks vertically.
  // All text: Arial 10pt, left-aligned.
  // Col A = label/name, Col B = value, Col C = value2, etc.
  // Currency/percent are pre-formatted as strings so alignment is consistent.

  var aoa   = [];  // array of arrays
  var meta  = [];  // parallel array: {sectionHdr, tableHdr, altRow, bold, fills:[colIdx:color]}

  function push(row, m){
    aoa.push(row);
    meta.push(m||{});
  }

  // Colors: 6-char RRGGBB for xlsx-js-style (no alpha prefix)
  var C_NAVY  = "1F3864";
  var C_BLUE  = "2E75B6";
  var C_LTBLU = "D6E4F0";
  var C_WHITE = "FFFFFF";
  var C_AMBER = "FFF2CC";
  var C_GREEN = "E2EFDA";
  var C_RED   = "FCE4D6";
  var C_DGREY = "F2F2F2";

  // ── TITLE ────────────────────────────────────────────────────
  push(["ExecIQ  |  Executive Summary", "", "", "", "", "Generated:", new Date().toLocaleString("en-US")],
    {titleRow: true});
  push([], {});

  // ── Section helper ───────────────────────────────────────────
  function section(title){
    push([title], {sectionHdr: true});
  }

  // ── Table header helper ──────────────────────────────────────
  function tblHdr(cols){
    push(cols, {tableHdr: true});
  }

  // ── Data row helper ──────────────────────────────────────────
  function row(cols, fillColor, bold){
    push(cols, {fill: fillColor||null, bold:bold||false});
  }

  function spacer(){
    push([], {});
  }

  // ── SECTION 1: PIPELINE OVERVIEW — horizontal KPI tiles ────────
  // All 8 metrics displayed as KPI tiles across two rows of 4.
  // Each tile: label (small, top), spacer, value (large, bold), spacer.
  // Cols A–D = 4 tiles per row, uniform width.

  var avgPwin     = avg(active, lbl.pwin);
  var avgDealAll  = pct(totalFee, rows.length);
  var avgDealAct  = pct(activeFee, active.length);

  // ── KPI Row 1: 4 tiles ───────────────────────────────────────
  push([
    "TOTAL OPPORTUNITIES",
    "ACTIVE OPPORTUNITIES",
    "TOTAL GROSS PIPELINE (" + lbl.ourFee.toUpperCase() + ")",
    "ACTIVE PIPELINE (" + lbl.ourFee.toUpperCase() + ")"
  ], {kpiLabel: true});
  push([fmtNum(rows.length), fmtNum(active.length), fmtCur(totalFee), fmtCur(activeFee)],
    {kpiValue: true});

  // ── KPI Row 2: 4 tiles ───────────────────────────────────────
  push([
    "WEIGHTED PIPELINE (" + lbl.weighted.toUpperCase() + ")",
    "AVERAGE " + lbl.pwin.toUpperCase(),
    "AVERAGE DEAL SIZE (ALL OPPS)",
    "AVERAGE ACTIVE DEAL SIZE"
  ], {kpiLabel: true});
  push([fmtCur(weightedTotal), fmtPct(avgPwin), fmtCur(avgDealAll), fmtCur(avgDealAct)],
    {kpiValue: true});

  spacer();

  // ── OVERALL PIPELINE HEALTH — horizontal summary bar ─────────
  // Pre-compute signal statuses so we can count critical/elevated
  // before rendering the risks table below
  var activeCount = active.length || 1;
  var top3Fee = topClients.slice(0,3).reduce(function(t,e){ return t+e[1].fee; }, 0);
  var top3Pct = pct(top3Fee, activeFee);

  function getSignalStatus(key, impactPct){
    if(impactPct === null || impactPct === undefined) return "🟢 Healthy";
    if(impactPct >= 0.50) return "🔴 Critical";
    if(impactPct >= 0.25) return "🟠 Elevated";
    if(impactPct >= 0.10) return "🟡 Watch";
    return "🟢 Healthy";
  }

  var signalStatuses = [
    getSignalStatus("stagnant60",    pct(stagnant60.length, activeCount)),
    getSignalStatus("stagnant90",    pct(stagnant90.length, activeCount)),
    getSignalStatus("overdue",       pct(overdue.length,    activeCount)),
    getSignalStatus("noDate",        pct(noDate.length,     activeCount)),
    getSignalStatus("noPwin",        pct(noPwin.length,     activeCount)),
    getSignalStatus("concentration", top3Pct),
  ];

  var criticalCount  = signalStatuses.filter(function(s){ return s === "🔴 Critical"; }).length;
  var elevatedCount  = signalStatuses.filter(function(s){ return s === "🟠 Elevated"; }).length;

  var overallHealth = criticalCount > 0        ? "🔴 At Risk"
                    : elevatedCount >= 2        ? "🟠 Needs Attention"
                    : "🟢 Healthy";

  var healthFill = criticalCount > 0   ? "FCE4D6"
                 : elevatedCount >= 2  ? "FFEB9C"
                 : C_GREEN;

  // Three-cell horizontal layout: Critical | Elevated | Overall Health
  push(["OVERALL PIPELINE HEALTH"], {sectionHdr: true});
  push([
    "🔴 Critical Signals:  " + criticalCount,
    "🟠 Elevated Signals:  " + elevatedCount,
    "Overall Health:  " + overallHealth
  ], {fill: healthFill, riskRow: true});
  spacer();

  // ── SECTION 2: PIPELINE RISKS & ALERTS ──────────────────────
  section("PIPELINE RISKS & ALERTS");
  tblHdr(["Signal", "Count", "% of Pipeline", "Status", "Insight"]);

  // Per-signal insight copy keyed by status
  var RISK_INSIGHTS = {
    "stagnant60": {
      "🔴 Critical": "Over half of active pipeline has stalled beyond 60 days, indicating significant pipeline stagnation.",
      "🟠 Elevated": "A large portion of pipeline is not progressing, which may impact near-term conversion.",
      "🟡 Watch":    "Some deals are aging beyond expected timelines — monitor progression closely.",
      "🟢 Healthy":  "Pipeline progression is within expected timeframes."
    },
    "stagnant90": {
      "🔴 Critical": "A significant share of pipeline is effectively stalled (>90 days), reducing likelihood of conversion.",
      "🟠 Elevated": "A meaningful portion of deals may be at risk due to extended inactivity.",
      "🟡 Watch":    "Some older deals may require re-engagement or reassessment.",
      "🟢 Healthy":  "Minimal long-term stagnation observed."
    },
    "overdue": {
      "🔴 Critical": "A large portion of pipeline is past expected award dates, signaling unreliable forecasting.",
      "🟠 Elevated": "Decision timelines are slipping across multiple opportunities.",
      "🟡 Watch":    "Some deals are extending beyond expected timelines.",
      "🟢 Healthy":  "Decision timelines are largely on track."
    },
    "noDate": {
      "🔴 Critical": "A significant portion of pipeline lacks expected close dates, limiting forecast reliability.",
      "🟠 Elevated": "Forecast visibility is reduced due to missing timing data.",
      "🟡 Watch":    "Some opportunities lack estimated award dates.",
      "🟢 Healthy":  "Most opportunities have defined timelines."
    },
    "noPwin": {
      "🔴 Critical": "Many active opportunities lack win probability, limiting pipeline accuracy.",
      "🟠 Elevated": "Probability scoring is incomplete across a portion of pipeline.",
      "🟡 Watch":    "Some opportunities are missing probability inputs.",
      "🟢 Healthy":  "Probability coverage is strong across pipeline."
    },
    "concentration": {
      "🔴 Critical": "Pipeline is highly concentrated across a few clients, increasing revenue risk exposure.",
      "🟠 Elevated": "A large share of pipeline is tied to a small number of clients.",
      "🟡 Watch":    "Moderate concentration across key clients.",
      "🟢 Healthy":  "Pipeline is well diversified across clients."
    }
  };

  function getSignalInsight(key, status){
    var cfg = RISK_INSIGHTS[key];
    if(!cfg || !cfg[status]) return "—";
    return cfg[status];
  }

  function riskFill(status, alt){
    if(status === "🔴 Critical") return "FCE4D6";   // red
    if(status === "🟠 Elevated") return "FFEB9C";   // darker amber/orange
    if(status === "🟡 Watch")    return "FFF2CC";   // lighter yellow
    return alt ? C_LTBLU : C_WHITE;
  }

  // Client concentration uses top 3 clients combined
  var signals = [
    ["Stagnant > 60 Days in Stage",                    stagnant60.length, pct(stagnant60.length, activeCount), "stagnant60"],
    ["Stagnant > 90 Days in Stage",                    stagnant90.length, pct(stagnant90.length, activeCount), "stagnant90"],
    ["Overdue Decisions (past " + lbl.estAward + ")",  overdue.length,    pct(overdue.length,    activeCount), "overdue"],
    ["No " + lbl.estAward + " Set (active)",           noDate.length,     pct(noDate.length,     activeCount), "noDate"],
    ["Active Opps with No " + lbl.pwin,                noPwin.length,     pct(noPwin.length,     activeCount), "noPwin"],
    ["Top 3 Client Concentration (% of Active Pipeline)", null,            top3Pct,                            "concentration"],
  ];

  signals.forEach(function(sig, i){
    var label   = sig[0];
    var count   = sig[1];
    var impact  = sig[2];
    var key     = sig[3];
    var status  = getSignalStatus(key, impact);
    var insight = getSignalInsight(key, status);
    var fill    = riskFill(status, i%2===1);
    push([label,
           count !== null ? fmtNum(count) : "—",
           fmtPct(impact),
           status,
           insight
          ], {fill: fill, riskRow: true});
  });
  spacer();

  // ── SECTION 3: 90-DAY FORECAST ──────────────────────────────
  section("90-DAY FORECAST  (" + lbl.estAward + " within next 90 days)");
  tblHdr(["Metric", "Value"]);
  row(["Opportunities in Window",                fmtNum(forecast90.length)],             C_WHITE);
  row([lbl.ourFee + " in Window",               fmtCur(sum(forecast90,lbl.ourFee))],    C_LTBLU);
  row(["Weighted Value in Window",               fmtCur(sum(forecast90,lbl.weighted))], C_WHITE);
  row(["Avg " + lbl.pwin + " in Window",        fmtPct(avg(forecast90,lbl.pwin))],      C_LTBLU);
  spacer();

  // ── SECTION 4: WIN / LOSS METRICS ───────────────────────────
  section("WIN / LOSS METRICS");
  tblHdr(["Metric", "Value"]);
  row(["Won Opportunities",        fmtNum(won.length)],                                  C_WHITE);
  row(["Lost Opportunities",       fmtNum(lost.length)],                                 C_LTBLU);
  row(["Won Revenue",              fmtCur(wonFee)],                                      C_GREEN);
  row(["Lost Revenue",             fmtCur(lostFee)],                                     C_RED);
  row(["Win Rate (by count)",      fmtPct(pct(won.length,resolved.length))],             C_WHITE);
  row(["Win Rate (by value)",      fmtPct(pct(wonFee,wonFee+lostFee))],                 C_LTBLU);
  spacer();

  // ── SECTION 5: STATUS BREAKDOWN ─────────────────────────────
  section("STATUS BREAKDOWN");
  tblHdr(["Status", "Count", "% of Total", lbl.ourFee, lbl.weighted]);
  var statusData = [
    ["Active",  active.length,  pct(active.length,rows.length),  activeFee,               weightedTotal],
    ["Won",     won.length,     pct(won.length,rows.length),     wonFee,                  null],
    ["Lost",    lost.length,    pct(lost.length,rows.length),    lostFee,                 null],
    ["Closed",  closed.length,  pct(closed.length,rows.length),  sum(closed,lbl.ourFee),  null],
  ];
  statusData.forEach(function(s,i){
    var fill = s[0]==="Won"?C_GREEN:s[0]==="Lost"?C_RED:i%2===0?C_WHITE:C_LTBLU;
    row([s[0], fmtNum(s[1]), fmtPct(s[2]), fmtCur(s[3]), s[4]!=null?fmtCur(s[4]):"—"], fill);
  });
  // Total row
  row(["TOTAL", fmtNum(rows.length), fmtPct(1), fmtCur(totalFee), "—"], C_DGREY, true);
  spacer();

  // ── SECTION 6: STAGE DISTRIBUTION ───────────────────────────
  section("STAGE DISTRIBUTION");
  tblHdr(["Stage", "Status", "Count", "% of Total", lbl.ourFee, lbl.weighted]);
  stageList.forEach(function(sg,i){
    var fill = sg.status==="Won"?C_GREEN:sg.status==="Lost"?C_RED:i%2===0?C_WHITE:C_LTBLU;
    row([sg.stage, sg.status, fmtNum(sg.count), fmtPct(pct(sg.count,rows.length)),
         fmtCur(sg.fee), fmtCur(sg.weighted)], fill);
  });
  spacer();

  // ── SECTION 7: TOP 10 CLIENTS BY PIPELINE ───────────────────
  section("TOP 10 CLIENTS BY PIPELINE");
  tblHdr(["Client", "Total Opps", "% of Active Pipeline", lbl.ourFee, lbl.weighted, "Active", "Won", "Lost"]);
  topClients.forEach(function(e,i){
    var fill = i%2===0?C_WHITE:C_LTBLU;
    row([e[0], fmtNum(e[1].count), fmtPct(pct(e[1].fee,activeFee)),
         fmtCur(e[1].fee), fmtCur(e[1].weighted),
         fmtNum(e[1].active), fmtNum(e[1].won), fmtNum(e[1].lost)], fill);
  });

  // ── Build worksheet ──────────────────────────────────────────
  var ws = XLSX.utils.aoa_to_sheet(aoa);

  // Apply cell-level styling
  // xlsx-js-style supports full cell styling via the .s property.
  // All styles below are applied directly to each cell.
  aoa.forEach(function(rowData, ri){
    var m = meta[ri];
    var numCols = rowData.length;

    // Determine row fill color
    var rowFill      = m.fill        || null;
    var isSectionHdr = m.sectionHdr  || false;
    var isTblHdr     = m.tableHdr    || false;
    var isTitleRow   = m.titleRow    || false;
    var isBold       = m.bold        || false;
    var isKpiLabel   = m.kpiLabel    || false;
    var isKpiValue   = m.kpiValue    || false;
    var isKpiSpacer  = m.kpiSpacer   || false;

    // KPI tile rows span exactly 4 columns (A–D), one tile per column
    var isKpiRow = isKpiLabel || isKpiValue || isKpiSpacer;

    var isRiskRow = m.riskRow || false;

    // Style all columns up to maxCols so merged cells get consistent fills
    var styleCols = (isSectionHdr||isTblHdr||isTitleRow) ? 8
                  : isKpiRow  ? 4
                  : isRiskRow ? 8   // extend risk row fills across all 8 cols
                  : Math.max(numCols,1);

    for(var ci=0; ci<styleCols; ci++){
      var addr = XLSX.utils.encode_cell({r:ri, c:ci});
      if(!ws[addr]) ws[addr] = {v:"", t:"s"};

      var fill, fclr, fbold, fsize;

      if(isKpiRow){
        // KPI tiles: navy background, white text throughout
        fill  = C_NAVY;
        fclr  = "FFFFFF";
        fbold = true;                  // bold on both label and value rows
        fsize = isKpiValue ? 22 : 12;  // 22pt value, 12pt label
      } else if(isSectionHdr || isTitleRow){
        fill  = C_NAVY;  fclr = "FFFFFF";  fbold = true;  fsize = 10;
      } else if(isTblHdr){
        fill  = C_BLUE;  fclr = "FFFFFF";  fbold = true;  fsize = 10;
      } else {
        fill  = rowFill; fclr = "000000";  fbold = isBold; fsize = 10;
      }

      // White outside border around each tile (label + value cells together).
      // Each tile occupies 2 rows (label row + value row) × 1 column.
      // We draw the outside edges only — no border between label and value cells.
      var border = {};
      if(isKpiRow){
        var thick = {style:"thick", color:{rgb:"FFFFFF"}};
        // Left edge of tile
        border.left = thick;
        // Right edge of tile (also serves as left edge of next tile)
        border.right = thick;
        // Top edge — only on label rows (top of tile)
        if(isKpiLabel) border.top = thick;
        // Bottom edge — only on value rows (bottom of tile)
        if(isKpiValue) border.bottom = thick;
        // No border between label and value rows within the same tile
      }

      ws[addr].s = {
        font: {
          name:  "Arial",
          sz:    fsize || 10,
          bold:  fbold,
          color: {rgb: fclr}
        },
        fill: fill ? {
          patternType: "solid",
          fgColor: {rgb: fill}
        } : {patternType:"none"},
        alignment: {
          horizontal: isKpiRow ? "center" : "left",
          vertical:   "center",
          wrapText:   isKpiLabel ? true : false,
          shrinkToFit: false
        },
        border: border
      };
    }
  });

  // Merge section header rows across all columns
  var maxCols = 0;
  aoa.forEach(function(r){ if(r.length>maxCols) maxCols=r.length; });
  if(!ws["!merges"]) ws["!merges"]=[];
  aoa.forEach(function(rowData, ri){
    var m = meta[ri];
    if(m.sectionHdr||m.titleRow){
      ws["!merges"].push({s:{r:ri,c:0},e:{r:ri,c:maxCols-1}});
    }
    // Merge insight cell (col E, index 4) across E–H so full text is visible
    // Excel won't overflow text into formatted cells — merge is the reliable fix
    if(m.riskRow){
      ws["!merges"].push({s:{r:ri,c:4},e:{r:ri,c:7}});
    }
  });

  // Row heights
  ws["!rows"] = aoa.map(function(rowData, ri){
    var m = meta[ri];
    if(m.titleRow)   return {hpt:22};
    if(m.sectionHdr) return {hpt:18};
    if(m.tableHdr)   return {hpt:16};
    if(m.kpiLabel)   return {hpt:52};   // label row — 12pt bold, wrapped
    if(m.kpiValue)   return {hpt:52};   // value row — 22pt bold
    if(m.kpiSpacer)  return {hpt:0};    // removed — zero height fallback
    if(!rowData||!rowData.length||!rowData[0]) return {hpt:8}; // spacer
    return {hpt:15};
  });

  // Column widths — KPI tile columns (A-D) sized generously for wrapped labels
  // and large value text. Data table columns below use the same widths.
  ws["!cols"] = [
    {wch:38}, // A
    {wch:38}, // B
    {wch:38}, // C
    {wch:38}, // D
    {wch:38}, // E — insight starts here, overflows into F-H
    {wch:18}, // F
    {wch:18}, // G
    {wch:22}, // H
  ];

  // Sheet ref
  ws["!ref"] = "A1:" + XLSX.utils.encode_cell({r:aoa.length-1, c:maxCols-1});

  return ws;
}


// ─────────────────────────────────────────────────────────────
// STEP 7: BUILD EXCEL SHEET
// ─────────────────────────────────────────────────────────────
function buildOpportunitySheet(rows, schema){
  // Build headers from schema labels, in schema order
  // Include __Status which is always computed
  var orderedLabels = [];

  // Always-first columns
  ["Opportunity Number","Opportunity Name","Client Company","Owner Company",
   "Opportunity Owner","Stage","Status","Days in Stage"].forEach(function(l){
    if(!orderedLabels.includes(l)) orderedLabels.push(l);
  });

  // Schema-ordered columns
  schema.forEach(function(f){
    var lbl = f.isCustom ? f.label : f.label;
    if(lbl && !orderedLabels.includes(lbl)) orderedLabels.push(lbl);
  });

  // Add Status if not already
  if(!orderedLabels.includes("Status")) orderedLabels.splice(6,0,"Status");

  // Collect all labels that actually appear in data
  var seenLabels = {};
  rows.forEach(function(row){
    Object.keys(row).forEach(function(k){
      if(!k.startsWith("__")) seenLabels[k]=1;
    });
  });
  // Always include custom field labels even if no records have data for them
  // (server may not return certain UUIDs but column should still exist)
  schema.forEach(function(f){
    if(f.isCustom && f.label) seenLabels[f.label] = 1;
  });

  // Final ordered column list — only labels that have data
  var finalCols = orderedLabels.filter(function(l){ return seenLabels[l]; });
  // Append any data labels not in the ordered list (shouldn't happen but safety net)
  Object.keys(seenLabels).forEach(function(l){
    if(!finalCols.includes(l)) finalCols.push(l);
  });

  // Build column type map for formatting
  var colTypes = {};
  schema.forEach(function(f){ colTypes[f.label] = f.type; });
  colTypes["Status"] = "text";

  var headers = finalCols;
  var EXCEL_CHAR_LIMIT = 32000; // Excel limit is 32,767 — use 32,000 as safe ceiling
  var data = rows.map(function(row){
    return finalCols.map(function(col){
      if(col==="Status") return row["__Status"]||"";
      var v = row[col];
      if(v==null) return "";
      // Hard safety net for Excel cell character limit
      if(typeof v === "string" && v.length > EXCEL_CHAR_LIMIT){
        return v.slice(0, EXCEL_CHAR_LIMIT) + "…";
      }
      return v;
    });
  });

  var ws = XLSX.utils.aoa_to_sheet([headers].concat(data));

  // Apply number formats — currency, percent, and date
  // Date cells contain real Date objects (written as Excel serial numbers by SheetJS).
  // Without an explicit format code they display as numbers — we must apply mm/dd/yyyy.
  data.forEach(function(_, ri){
    finalCols.forEach(function(col, ci){
      var addr = XLSX.utils.encode_cell({r:ri+1, c:ci});
      if(!ws[addr]) return;
      var type = colTypes[col]||"text";
      if(type==="currency")     ws[addr].z = '"$"#,##0';
      else if(type==="percent") ws[addr].z = "0%";
      else if(type==="date" && ws[addr].v instanceof Date) ws[addr].z = "mm/dd/yyyy";
    });
  });

  ws["!autofilter"] = {ref:"A1:"+XLSX.utils.encode_col(headers.length-1)+"1"};
  ws["!freeze"]     = {xSplit:0, ySplit:1};
  ws["!cols"]       = headers.map(function(h){
    return {wch: Math.max(10, Math.min(50, String(h).length+4))};
  });

  return ws;
}

// ─────────────────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────────────────
async function main(isRefresh){
  if(!isRefresh) UI.mount();
  UI.prog(5);
  UI.status("Starting — locating CRM endpoints...");

  // 1. Find oppActions.cfm
  var oppBase = await findOppBase();
  if(!oppBase){
    UI.status("Cannot locate oppActions.cfm — navigate to the Opportunities page and try again.", "er");
    return;
  }
  UI.log("✓ oppBase: " + oppBase, "ls");
  UI.prog(10);

  // 2. Find firmData.cfc
  var firmDataURL = findURLByPattern(/firmData\.cfc/i);
  var firmDataBase = firmDataURL ? firmDataURL : oppBase + "firmData.cfc";
  UI.log((firmDataURL ? "✓ firmData.cfc from resources: " : "firmData.cfc constructed: ") + firmDataBase, firmDataURL?"ls":"lw");
  UI.prog(15);

  // 3. Probe: discover which fields this client's server returns
  UI.status("Probing available fields...");
  var availableFields = await probeAvailableFields(oppBase);
  UI.prog(28);

  // 4. Build client config from firmData.cfc
  UI.status("Loading client configuration...");
  var config = await buildClientConfig(firmDataBase);
  UI.prog(42);

  // 5. Load lookup tables
  UI.status("Loading lookup tables...");
  config = await loadLookupTables(oppBase, config);
  UI.prog(55);

  // 6. Build dynamic field schema
  UI.status("Building field schema...");
  var schema = buildFieldSchema(availableFields, config);

  // 6a. Add STAFFROLE fields to schema dynamically from staffRoles lookup
  // These fields follow a dynamic naming pattern (STAFFROLE_{roleId}) that
  // can't be known until the staffRoles lookup is loaded.
  // We add them to the schema here and request them in the full fetch.
  if(Object.keys(config.lookups.staffRoles).length > 0){
    var staffCols = buildStaffRoleColumns(config.lookups.staffRoles);
    staffCols.forEach(function(colKey){
      var upper = colKey.toUpperCase();
      // Only add if not already in schema from probe
      var alreadyIn = schema.some(function(f){ return f.upper === upper; });
      if(!alreadyIn){
        var roleId = colKey.replace("STAFFROLE_","").replace("staffrole_","");
        var roleName = config.lookups.staffRoles[roleId] || ("Staff Role " + roleId);
        schema.push({
          backendKey: colKey,
          upper:      upper,
          label:      roleName,
          type:       "text",
          isCustom:   false,
          resolve:    null
        });
      }
    });
    // Resolve any __STAFFROLE__ placeholders from the probe phase
    schema = resolveStaffRoleLabels(schema, config.lookups.staffRoles);
    UI.log("✓ Staff role columns added: " + staffCols.length, "ls");
  }

  // Enabled-field filtering is now handled inside buildFieldSchema.
  // No secondary filter needed here.
  var customUUIDs = config.customFields.map(function(cf){ return cf.uuid; });
  UI.stat("sv-cust", customUUIDs.length);
  UI.log("Requesting " + schema.length + " fields (" + customUUIDs.length + " custom)", "ls");
  UI.prog(60);

  // 7. Fetch all opportunities
  var filterSettings = UI.getFilterSettings();
  UI.log("Date filter: " + filterSettings.dateFieldLabel + " · " +
    (filterSettings.years >= 999 ? "All Time" :
     filterSettings.years === 0  ? "YTD" :
     "Last " + filterSettings.years + " Years"), "ls");
  UI.status("Fetching all opportunities...");
  var oppData = await fetchAllOpportunities(oppBase, schema, customUUIDs, filterSettings);
  if(!oppData){
    UI.status("Failed to retrieve opportunity data.", "er");
    return;
  }
  UI.stat("sv-opps", oppData.DATA.length);
  UI.prog(72);

  // 8. Normalize records
  UI.status("Normalizing " + oppData.DATA.length + " records...");
  var cleanRows = [];
  var normErrors = 0;
  oppData.DATA.forEach(function(opp, idx){
    try{
      cleanRows.push(normalizeRecord(opp, schema, config));
    }catch(e){
      normErrors++;
      UI.log("⚠ Row " + idx + " error: " + e.message, "lw");
      cleanRows.push({
        "Opportunity Number": String(opp.VCHLEADNUMBER||""),
        "Opportunity Name":   String(opp.VCHPROJECTNAME||""),
        "Stage":              String(opp.STAGENAME||""),
        "__Status":           classifyStatus(opp.STAGENAME||"", opp.ACTIVEIND),
        "__Fee":              parseFloat(opp.IFIRMFEE)||0,
        "__Weighted":         parseFloat(opp.IFACTOREDFEE)||0,
        "__Pwin":             opp.IPROBABILITY!=null ? parseFloat(opp.IPROBABILITY)/100 : null,
        "__Stage":            String(opp.STAGENAME||""),
        "Parse Error":        e.message
      });
    }
  });
  if(normErrors) UI.log("⚠ " + normErrors + " rows had errors — included with partial data", "lw");
  UI.log("✓ Normalization complete", "ls");
  UI.prog(82);

  UI.stat("sv-cols", schema.length);
  UI.log("✓ " + cleanRows.length + " rows × " + schema.length + " schema fields", "ls");
  UI.prog(87);

  // 9. Load SheetJS
  UI.status("Loading export library...");
  try{
    await Promise.race([
      new Promise(function(res,rej){
        if(window.XLSX){res();return;}
        var s=document.createElement("script"); s.src=SHEETJS;
        s.onerror=function(){rej(new Error("SheetJS CDN failed"));};
        s.onload=function(){
          var n=0, iv=setInterval(function(){
            if(window.XLSX){clearInterval(iv);res();}
            else if(++n>40){clearInterval(iv);rej(new Error("XLSX not available"));}
          },100);
        };
        document.head.appendChild(s);
      }),
      new Promise(function(_,rej){setTimeout(function(){rej(new Error("SheetJS timeout"));},20000);})
    ]);
  }catch(e){
    UI.status("Export library failed: " + e.message, "er");
    return;
  }
  UI.log("✓ SheetJS ready", "ls");
  UI.prog(93);

  // 10. Build workbook
  UI.status("Building workbook...");
  var wb = XLSX.utils.book_new();

  UI.log("Building Exec Summary...");
  XLSX.utils.book_append_sheet(wb, buildExecSummarySheet(cleanRows, schema, config), "Exec Summary");

  UI.log("Building Opportunity Data...");
  XLSX.utils.book_append_sheet(wb, buildOpportunitySheet(cleanRows, schema), "Opportunity Data");

  UI.log("✓ Workbook ready — 2 sheets", "ls");
  UI.prog(100);
  UI.status(cleanRows.length + " opportunities ready for export");
  UI.log("✓ Done — click Export to download.", "ls");

  // Show Refresh button so user can rerun with different date settings
  UI.showRefresh(function(){ main(true); });

  // 11. Export
  UI.enableExport(function(){
    try{
      var buf = XLSX.write(wb,{bookType:"xlsx",type:"array",compression:false});
      var blob = new Blob([buf],{type:"application/octet-stream"});
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "ExecIQ_Report_" + new Date().toISOString().slice(0,10) + ".xlsx";
      document.body.appendChild(a); a.click();
      setTimeout(function(){URL.revokeObjectURL(a.href); a.remove();},1500);
      UI.log("✓ Download started", "ls");
      UI.status("Downloaded — check your Downloads folder");
    }catch(e){
      UI.log("Export error: " + e.message, "le");
      UI.status("Export failed", "er");
    }
  });
}

main();

})();
