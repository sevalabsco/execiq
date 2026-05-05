javascript:(function(){
"use strict";

// ─── Guard ───────────────────────────────────────────────────
if(window.__EXECIQ_P1__){console.warn("[ExecIQ] Already running.");return;}
window.__EXECIQ_P1__ = true;

var VERSION = "3.7";
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
  "OFFICELIST":          "firmOrg",
  "DIVISIONLIST":        "firmOrg",
  "STUDIOLIST":          "firmOrg",
  "PRACTICEAREALIST":    "firmOrg",
  "TERRITORYLIST":       "firmOrg",
  "OFFICEDIVISIONLIST":  "firmOrg",
  "PRIMARYCATEGORYLIST": "priCat",
  "SECONDARYCATEGORYLIST":"secCat",
  "CONTRACTTYPES":       "contract",
  "CLIENTTYPES":         "clientType",
  "PROSPECTTYPES":       "prospect",
  "DELIVERYMETHOD":      "delivery",
  "STAGEID":             "stage",
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
    "#iq1-btn:disabled{background:#1e293b;color:#475569;cursor:not-allowed;}"
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
        '<div id="iq1-log"></div>' +
        '<button id="iq1-btn" disabled>Preparing...</button>' +
      '</div>';
    document.body.appendChild(el);
    elStatus = document.getElementById("iq1-status");
    elFill   = document.getElementById("iq1-fill");
    elLog    = document.getElementById("iq1-log");
    elBtn    = document.getElementById("iq1-btn");
    document.getElementById("iq1-close").onclick = destroy;
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

  return { mount, destroy, status, prog, log, stat, enableExport };
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
    "PRIMARYCATEGORYLIST","SECONDARYCATEGORYLIST","CONTRACTTYPES",
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

  // Fire all lookups in parallel
  var results = await Promise.all([
    getLookup("stage.cfc",          "get"),
    getLookup("stage.cfc",          "getList"),          // some instances use getList
    getLookup("oppData.cfc",        "getProspectTypes"),
    getLookup("role.cfc",           "getOpportunityAvailableRoles"),
    getLookup("role.cfc",           "getList"),
    getLookup("contractType.cfc",   "getContractTypes"),
    getLookup("deliveryMethod.cfc", "getDeliveryMethods"),
    getLookup("clientType.cfc",     "getClientTypes"),
    getLookup("submittalType.cfc",  "getSubmittalTypes"),
    getLookup("primaryCategory.cfc","getList"),
    getLookup("secondaryCategory.cfc","getList"),
    getLookup("staffTeam.cfc",      "getStaffTeamRoles"),  // staff role ID → name
  ]);

  // Merge both stage method variants — use whichever returned JSON
  function pickJSON(a, b){
    var ar=parseCFC(a), br=parseCFC(b);
    return ar.length ? ar : br;
  }
  config.lookups.stage      = buildLookup(pickJSON(results[0],results[1]), "STAGEID",           "STAGENAME");
  config.lookups.prospect   = buildLookup(parseCFC(results[2]),            "ID",                "DISPLAYNAME");

  // role.cfc returns {data:[{roleid,rolename}]} — parseCFC now handles lowercase
  var roleRaw = pickJSON(results[3], results[4]);
  config.lookups.role = buildLookup(roleRaw, "ROLEID", "ROLENAME");
  // Merge firmData role fallback if .cfc returned nothing
  if(!Object.keys(config.lookups.role).length && config.lookups.roleFromFirmData){
    config.lookups.role = config.lookups.roleFromFirmData;
    UI.log("  Role lookup: using firmData fallback", "lw");
  }

  config.lookups.contract   = buildLookup(parseCFC(results[5]), "CONTRACTTYPEID","CONTRACTNAME");
  config.lookups.delivery   = buildLookup(parseCFC(results[6]), "DELIVERYMETHODID","DELIVERYMETHODNAME");
  config.lookups.clientType = buildLookup(parseCFC(results[7]), "ID",             "DISPLAYNAME");
  config.lookups.submittal  = buildLookup(parseCFC(results[8]), "SUBMITTALTYPEID","SUBMITTALTYPENAME");

  // Primary/Secondary categories — getList returns HTML on some instances, try getJSON
  config.lookups.priCat = buildLookup(parseCFC(results[9]),  "CATEGORYID", "CATEGORYNAME");
  config.lookups.secCat = buildLookup(parseCFC(results[10]), "CATEGORYID", "CATEGORYNAME");
  // If categories came back empty (HTML response), log and continue — IDs will show as-is
  if(!Object.keys(config.lookups.priCat).length){
    UI.log("  Primary category lookup empty — IDs will show raw (no JSON endpoint found)", "lw");
  }

  // Staff role ID → role name
  var staffRoleRaw = parseCFC(results[11]);
  config.lookups.staffRoles = buildLookup(staffRoleRaw, "STAFFROLEID", "STAFFROLENAME");
  // Merge firmData staff roles fallback
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
    "secondarycategorylist":   "secondarycategoryid",
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
    "SECONDARYCATEGORYLIST":   "Secondary Categories",
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
async function fetchAllOpportunities(oppBase, schema, customFieldUUIDs){
  UI.log("Fetching all opportunities (no filters, no pagination)...");

  // Build the fixed base parameters — these never change between pages
  var baseParams = [
    "action=getOpportunityGridData","json=1","sort=STAGEID","dir=ASC",
    "selectedCurrency=USD","view=0",
    "ActiveInd=0",            // ALL statuses — never filter by status
    "SalesCycle=NaN",
    "officeId=0","divisionId=0","studioId=0","practiceAreaId=0",
    "territoryId=0","stageId=0","priCatId=0","secCatId=0",
    "masterSub=0","staffRoleId=0","dateCreated=0","dateModified=0",
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
      // Store even if empty string so the column appears — only skip null/undefined
      if(display!==null&&display!==undefined) row[field.label] = display;
      return;
    }

    // Standard field — apply resolution and type formatting
    if(field.resolve){
      // ID resolution field
      display = resolveIDs(val, config.lookups[field.resolve]||{});
      if(!display) display = String(val).trim();
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

  // ── SECTION 1: PIPELINE OVERVIEW ────────────────────────────
  section("PIPELINE OVERVIEW");
  tblHdr(["Metric", "Value"]);
  row(["Total Opportunities",                   fmtNum(rows.length)],                    C_WHITE);
  row(["Active Opportunities",                  fmtNum(active.length)],                  C_LTBLU);
  row(["Total Gross Pipeline (" + lbl.ourFee + ")", fmtCur(totalFee)],                   C_WHITE);
  row(["Active Pipeline (" + lbl.ourFee + ")",  fmtCur(activeFee)],                      C_LTBLU);
  row(["Weighted Pipeline (" + lbl.weighted + ")", fmtCur(weightedTotal)],               C_WHITE);
  row(["Average " + lbl.pwin + " (active opps)", fmtPct(avg(active,lbl.pwin))],          C_LTBLU);
  row(["Average Deal Size (all opps)",           fmtCur(pct(totalFee,rows.length))],      C_WHITE);
  row(["Average Active Deal Size",               fmtCur(pct(activeFee,active.length))],  C_LTBLU);
  spacer();

  // ── SECTION 2: WIN / LOSS METRICS ───────────────────────────
  section("WIN / LOSS METRICS");
  tblHdr(["Metric", "Value"]);
  row(["Won Opportunities",        fmtNum(won.length)],                                  C_WHITE);
  row(["Lost Opportunities",       fmtNum(lost.length)],                                 C_LTBLU);
  row(["Won Revenue",              fmtCur(wonFee)],                                      C_GREEN);
  row(["Lost Revenue",             fmtCur(lostFee)],                                     C_RED);
  row(["Win Rate (by count)",      fmtPct(pct(won.length,resolved.length))],             C_WHITE);
  row(["Win Rate (by value)",      fmtPct(pct(wonFee,wonFee+lostFee))],                 C_LTBLU);
  spacer();

  // ── SECTION 3: 90-DAY FORECAST ──────────────────────────────
  section("90-DAY FORECAST  (" + lbl.estAward + " within next 90 days)");
  tblHdr(["Metric", "Value"]);
  row(["Opportunities in Window",                fmtNum(forecast90.length)],             C_WHITE);
  row([lbl.ourFee + " in Window",               fmtCur(sum(forecast90,lbl.ourFee))],    C_LTBLU);
  row(["Weighted Value in Window",               fmtCur(sum(forecast90,lbl.weighted))], C_WHITE);
  row(["Avg " + lbl.pwin + " in Window",        fmtPct(avg(forecast90,lbl.pwin))],      C_LTBLU);
  spacer();

  // ── SECTION 4: PIPELINE HEALTH SIGNALS ──────────────────────
  section("PIPELINE HEALTH SIGNALS");
  tblHdr(["Signal", "Count", "Status"]);

  function healthStatus(val, warnThreshold, critThreshold){
    if(val===null||val===0) return "✓ OK";
    if(critThreshold!=null&&val>=critThreshold) return "⚠ Review";
    if(val>=warnThreshold) return "△ Watch";
    return "✓ OK";
  }

  var s60fill  = stagnant60.length > 0 ? C_AMBER : C_WHITE;
  var s90fill  = stagnant90.length > 0 ? C_AMBER : C_LTBLU;
  var ovfill   = overdue.length    > 0 ? C_AMBER : C_WHITE;
  var ndfill   = noDate.length     > 0 ? C_AMBER : C_LTBLU;
  var npfill   = noPwin.length     > 5 ? C_AMBER : C_WHITE;
  var concfill = (topClientPct!=null&&topClientPct>0.33) ? C_AMBER : C_LTBLU;

  var topClientPct = pct(topClients.length?topClients[0][1].fee:0, activeFee);

  row(["Stagnant > 60 Days in Stage",  fmtNum(stagnant60.length), healthStatus(stagnant60.length,1,10)], s60fill);
  row(["Stagnant > 90 Days in Stage",  fmtNum(stagnant90.length), healthStatus(stagnant90.length,1,5)],  s90fill);
  row(["Overdue Decisions (past " + lbl.estAward + ")", fmtNum(overdue.length), healthStatus(overdue.length,1,5)], ovfill);
  row(["No " + lbl.estAward + " Set (active)", fmtNum(noDate.length), healthStatus(noDate.length,1,10)], ndfill);
  row(["Active Opps with No " + lbl.pwin, fmtNum(noPwin.length), healthStatus(noPwin.length,1,10)],      npfill);
  row(["Top Client Concentration",     fmtPct(topClientPct),
    topClientPct!=null&&topClientPct>0.33?"⚠ Review":topClientPct!=null&&topClientPct>0.20?"△ Watch":"✓ OK"], concfill);
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
  tblHdr(["Client", "Total Opps", "Active", "Won", "Lost", lbl.ourFee, lbl.weighted, "% of Active Pipeline"]);
  topClients.forEach(function(e,i){
    var fill = i%2===0?C_WHITE:C_LTBLU;
    row([e[0], fmtNum(e[1].count), fmtNum(e[1].active), fmtNum(e[1].won), fmtNum(e[1].lost),
         fmtCur(e[1].fee), fmtCur(e[1].weighted), fmtPct(pct(e[1].fee,activeFee))], fill);
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
    var rowFill = m.fill || null;
    var isSectionHdr = m.sectionHdr || false;
    var isTblHdr     = m.tableHdr   || false;
    var isTitleRow   = m.titleRow   || false;
    var isBold       = m.bold       || false;

    // Style all columns up to maxCols so merged cells get consistent fills
    var styleCols = (isSectionHdr||isTblHdr||isTitleRow) ? 8 : Math.max(numCols,1);
    for(var ci=0; ci<styleCols; ci++){
      var addr = XLSX.utils.encode_cell({r:ri, c:ci});
      if(!ws[addr]) ws[addr] = {v:"", t:"s"};

      // Number format — all values are pre-formatted strings, so General is fine
      // (we formatted them as strings in fmtCur/fmtPct so Excel treats as text)

      // Apply styles via .s property
      var fill  = isSectionHdr ? C_NAVY : isTblHdr ? C_BLUE : isTitleRow ? C_NAVY : rowFill;
      var fclr  = (isSectionHdr||isTblHdr||isTitleRow) ? "FFFFFF" : "000000";
      var fbold = isSectionHdr||isTblHdr||isTitleRow||isBold;

      ws[addr].s = {
        font: {
          name:  "Arial",
          sz:    10,
          bold:  fbold,
          color: {rgb: fclr}
        },
        fill: fill ? {
          patternType: "solid",
          fgColor: {rgb: fill}
        } : {patternType:"none"},
        alignment: {
          horizontal: "left",
          vertical:   "center",
          wrapText:   false
        }
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
  });

  // Row heights
  ws["!rows"] = aoa.map(function(rowData, ri){
    var m = meta[ri];
    if(m.titleRow)   return {hpt:22};
    if(m.sectionHdr) return {hpt:18};
    if(m.tableHdr)   return {hpt:16};
    if(!rowData||!rowData.length||!rowData[0]) return {hpt:8}; // spacer
    return {hpt:15};
  });

  // Column widths
  ws["!cols"] = [
    {wch:48}, // A — label / name
    {wch:18}, // B
    {wch:16}, // C
    {wch:16}, // D
    {wch:16}, // E
    {wch:18}, // F
    {wch:18}, // G
    {wch:22}, // H — % of pipeline
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
  var data = rows.map(function(row){
    return finalCols.map(function(col){
      if(col==="Status") return row["__Status"]||"";
      var v = row[col];
      return v!=null ? v : "";
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
async function main(){
  UI.mount();
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
  UI.status("Fetching all opportunities...");
  var oppData = await fetchAllOpportunities(oppBase, schema, customUUIDs);
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
