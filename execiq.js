javascript:(function (){"use strict";if (window.__EXECIQ_V7__){console.warn("[ExecIQ]Already running.");return;}window.__EXECIQ_V7__ = true;

// CONFIGURATION
var SHEETJS = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

// Field configuration (standard fields)
var CF ={"OpportunityNumber":{l:"Opp Number",t:"text",e:1 },"OpportunityName":{l:"Opportunity Name",t:"text",e:1 },"ClientCompany":{l:"Client",t:"text",e:1 },"OwnerCompany":{l:"Owner Company",t:"text",e:1 },"Stage":{l:"Stage",t:"text",e:1 },"OpportunityStatus":{l:"Status",t:"text",e:1 },"FirmEstimatedFee":{l:"Our Fee",t:"currency",e:1 },"FactoredFee":{l:"Weighted Value",t:"currency",e:1 },"WinProbability":{l:"Pwin",t:"percent",e:1 },"EstimatedCost":{l:"Vehicle or Prime Ceiling",t:"currency",e:1 },"TotalEstimatedFee":{l:"Our Ceiling Value",t:"currency",e:1 },"PrimaryContact":{l:"Primary Contact",t:"text",e:1 },"ClientTypes":{l:"Priority",t:"text",e:1 },"ProspectTypes":{l:"Growth Type",t:"text",e:1 },"OpportunityRole":{l:"Our Anticipated Role",t:"text",e:1 },"SubmittalType":{l:"Submittal Type",t:"text",e:1 },"ContractType":{l:"Contract Type",t:"text",e:1 },"DeliveryMethod":{l:"Delivery Method",t:"text",e:1 },"ServiceTypes":{l:"Service Types",t:"text",e:1 },"Offices":{l:"Office",t:"text",e:1 },"Studios":{l:"Service Area",t:"text",e:1 },"PracticeAreas":{l:"Market Sector",t:"text",e:1 },"Divisions":{l:"Division",t:"text",e:1 },"Territories":{l:"Territory",t:"text",e:1 },"OfficeDivision":{l:"Office Division",t:"text",e:0 },"PrimaryCategory":{l:"Project Type",t:"text",e:0 },"SecondaryCategory":{l:"Set Aside/Competition Type",t:"text",e:0 },"QualsDueDate":{l:"RFI Response Due Date",t:"date",e:1 },"ExpectedRFPDate":{l:"Expected RFP Release Date",t:"date",e:1 },"RFPReceived":{l:"RFP Received",t:"text",e:1 },"ProposalDueDate":{l:"RFP Response Due Date",t:"date",e:1 },"PresentationDate":{l:"Expected RFI Release Date",t:"date",e:1 },"EstimatedSelectionDate":{l:"Estimated Award Date",t:"date",e:1 },"CloseDate":{l:"Actual Award/Close Date",t:"date",e:1 },"EstimatedStartDate":{l:"Estimated PoP Start Date",t:"date",e:1 },"EstimatedCompletionDate":{l:"Estimated PoP End Date",t:"date",e:1 },"SolicitationNumber":{l:"Solicitation Number",t:"text",e:1 },"Comments":{l:"Comments",t:"text",e:1 },"Notes":{l:"Notes",t:"text",e:1 },"OpportunityDescription":{l:"Description",t:"text",e:1 },"LeadBidDate":{l:"Lead Bid Date",t:"date",e:1 },"LeadOriginationDate":{l:"Lead Origination Date",t:"date",e:1 },"DaysinStage":{l:"Days in Stage",t:"number",e:1 },"DateCreated":{l:"Created Date",t:"date",e:1 },"DateModified":{l:"Last Modified",t:"date",e:1 },"Address1":{l:"Address 1",t:"text",e:1 },"City":{l:"City",t:"text",e:1 },"StateProv":{l:"State",t:"text",e:1 },"Country":{l:"Country",t:"text",e:1 },"PostalCode":{l:"Postal Code",t:"text",e:1 },};

// Backend field name to Compass field name mapping
var BF ={"VCHLEADNUMBER":"OpportunityNumber","VCHPROJECTNAME":"OpportunityName","COMPANY":"ClientCompany","OWNERCOMPANY":"OwnerCompany","OWNER":"OwnerCompany","STAGENAME":"Stage","ACTIVEIND":"OpportunityStatus","DAYSINSTAGE":"DaysinStage","IFIRMFEE":"FirmEstimatedFee","IFACTOREDFEE":"FactoredFee","IPROBABILITY":"WinProbability","IFEE":"TotalEstimatedFee","OFFICELIST":"Offices","STUDIOLIST":"Studios","PRACTICEAREALIST":"PracticeAreas","DIVISIONLIST":"Divisions","TERRITORYLIST":"Territories","OFFICEDIVISIONLIST":"OfficeDivision","PRIMARYCATEGORYLIST":"PrimaryCategory","SECONDARYCATEGORYLIST":"SecondaryCategory","CONTRACTTYPES":"ContractType","DELIVERYMETHOD":"DeliveryMethod","CLIENTTYPES":"ClientTypes","PROSPECTTYPES":"ProspectTypes","SERVICETYPES":"ServiceTypes","SUBMITTALTYPENAME":"SubmittalType","ROLENAME":"OpportunityRole","OWNERCONTACTID":"PrimaryContact","DTPROPOSALDATE":"ProposalDueDate","DTSTARTDATE":"EstimatedSelectionDate","DTCLOSEDATE":"CloseDate","DTQUALSDATE":"QualsDueDate","DTRFPDATE":"ExpectedRFPDate","DTPRESENTATIONDATE":"PresentationDate","ESTIMATEDSTARTDATE":"EstimatedStartDate","ESTIMATEDCOMPLETIONDATE":"EstimatedCompletionDate","CREATEDATE":"DateCreated","MODDATE":"DateModified","CHRFPREC":"RFPReceived","SOLICITATIONNUMBER":"SolicitationNumber","TXCOMMENTS":"Comments","TXNOTE":"Notes","VCHADDRESS1":"Address1","VCHCITY":"City","STATEABRV":"StateProv","VCHCOUNTRY":"Country","VCHPOSTALCODE":"PostalCode","DESCRIPTION":"OpportunityDescription",};

var STANDARD_COLUMNS =["VCHLEADNUMBER","VCHPROJECTNAME","COMPANY","OWNERCOMPANY","STAGENAME","ACTIVEIND","DAYSINSTAGE","IFIRMFEE","IFACTOREDFEE","IPROBABILITY","IFEE","OFFICELIST","STUDIOLIST","PRACTICEAREALIST","DIVISIONLIST","TERRITORYLIST","OFFICEDIVISIONLIST","PRIMARYCATEGORYLIST","SECONDARYCATEGORYLIST","CONTRACTTYPES","DELIVERYMETHOD","CLIENTTYPES","PROSPECTTYPES","SERVICETYPES","SUBMITTALTYPENAME","ROLENAME","OWNERCONTACTID","DTPROPOSALDATE","DTSTARTDATE","DTCLOSEDATE","DTQUALSDATE","DTRFPDATE","ESTIMATEDSTARTDATE","ESTIMATEDCOMPLETIONDATE","DTPRESENTATIONDATE","CREATEDATE","MODDATE","CHRFPREC","SOLICITATIONNUMBER","TXCOMMENTS","TXNOTE","VCHADDRESS1","VCHCITY","STATEABRV","VCHCOUNTRY","VCHPOSTALCODE","DESCRIPTION",];

var SUPPRESS = new Set(["CFIRMID","OWNERCOMPANY","OWNER","STAGEID","ROLEID","FIRMID","ILEADID","STAGENAME","COMPANY","VCHLEADNUMBER","VCHPROJECTNAME","OWNERCONTACT","ROWNUMBER","TOTALRECORDS","ACTIVEIND","SELECTEDCURRENCYSYMBOL","SELECTEDRATE","SELECTEDCURRENCY","BASECURRENCYSYMBOL","BASECURRENCYABBR","SELECTEDCURRENCYABBR","BASECURRENCY","ICOUNTRYID","ISTATEID","LOCATION","MASTERLEADID","SALESCYCLE",]);

var KEY_ORDER =["OpportunityNumber","OpportunityName","ClientCompany","OwnerCompany","Stage","OpportunityStatus","FirmEstimatedFee","FactoredFee","WinProbability",];

// Firm Org Field Mapping
var FIRM_ORG_FIELDS = {
  "Offices": { backendField: "OFFICELIST", defaultLabel: "Office" },
  "Studios": { backendField: "STUDIOLIST", defaultLabel: "Service Area" },
  "PracticeAreas": { backendField: "PRACTICEAREALIST", defaultLabel: "Market Sector" },
  "Divisions": { backendField: "DIVISIONLIST", defaultLabel: "Division" },
  "Territories": { backendField: "TERRITORYLIST", defaultLabel: "Territory" },
  "OfficeDivision": { backendField: "OFFICEDIVISIONLIST", defaultLabel: "Office Division" }
};

// Global config object
var userConfig = null;

// Configuration Manager
var ConfigManager = (function(){
  var CONFIG_KEY = "execiq_v7_config";
  
  function load(){
    try{
      var stored = localStorage.getItem(CONFIG_KEY);
      if (!stored) return null;
      return JSON.parse(stored);
    }catch(e){
      return null;
    }
  }
  
  function save(config){
    try{
      config.lastUpdated = new Date().toISOString();
      localStorage.setItem(CONFIG_KEY, JSON.stringify(config));
      return true;
    }catch(e){
      console.error("[ExecIQ] Config save failed:", e);
      return false;
    }
  }
  
  function clear(){
    try{
      localStorage.removeItem(CONFIG_KEY);
      return true;
    }catch(e){
      return false;
    }
  }
  
  return { load, save, clear };
})();

// Firm Org Discovery Engine
function discoverFirmOrgs(rows){
  var usage = {};
  
  // Initialize tracking for each Firm Org field
  Object.keys(FIRM_ORG_FIELDS).forEach(function(field){
    usage[field] = {
      field: field,
      label: CF[field] ? CF[field].l : FIRM_ORG_FIELDS[field].defaultLabel,
      count: 0,
      sampleValues: [],
      enabled: false,
      priority: 999,
      includeInRisk: false
    };
  });
  
  // Scan all rows
  rows.forEach(function(row){
    Object.keys(FIRM_ORG_FIELDS).forEach(function(field){
      var val = row[field];
      if (val && String(val).trim() !== ""){
        usage[field].count++;
        if (usage[field].sampleValues.length < 5){
          var samples = String(val).split(",").map(function(s){ return s.trim(); });
          samples.forEach(function(s){
            if (s && usage[field].sampleValues.indexOf(s) === -1 && usage[field].sampleValues.length < 5){
              usage[field].sampleValues.push(s);
            }
          });
        }
      }
    });
  });
  
  // Filter to active fields and auto-prioritize
  var active = Object.keys(usage)
    .filter(function(k){ return usage[k].count > 0; })
    .sort(function(a, b){ return usage[b].count - usage[a].count; });
  
  active.forEach(function(field, idx){
    usage[field].enabled = true;
    usage[field].priority = idx + 1;
    usage[field].includeInRisk = idx < 3; // Top 3 in risk analysis by default
  });
  
  return usage;
}

// Configuration UI Module
var ConfigUI = (function(){
  var CSS = "#iqcfg{position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:2147483648;width:680px;max-height:90vh;background:#0d1117;color:#e6edf3;font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;border:1px solid #30363d;border-radius:12px;box-shadow:0 25px 80px rgba(0,0,0,.9);overflow:hidden;}#iqcfg-overlay{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.7);z-index:2147483647;}#iqcfg-hd{background:#161b22;border-bottom:1px solid #30363d;padding:16px 20px;}#iqcfg-hd h2{margin:0;font-size:16px;font-weight:700;color:#f0f6fc;}#iqcfg-hd p{margin:6px 0 0;font-size:12px;color:#8b949e;}#iqcfg-bd{padding:20px;max-height:calc(90vh - 140px);overflow-y:auto;}#iqcfg-bd::-webkit-scrollbar{width:8px;}#iqcfg-bd::-webkit-scrollbar-track{background:#161b22;}#iqcfg-bd::-webkit-scrollbar-thumb{background:#30363d;border-radius:4px;}#iqcfg-bd::-webkit-scrollbar-thumb:hover{background:#484f58;}.cfg-section{margin-bottom:24px;}.cfg-section h3{margin:0 0 12px;font-size:13px;font-weight:700;color:#58a6ff;text-transform:uppercase;letter-spacing:0.5px;}.cfg-field{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:12px 14px;margin-bottom:10px;}.cfg-field-hd{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;}.cfg-field-name{font-weight:600;color:#f0f6fc;font-size:13px;}.cfg-field-stats{font-size:11px;color:#8b949e;}.cfg-field-samples{font-size:11px;color:#8b949e;margin-bottom:8px;line-height:1.5;}.cfg-field-controls{display:flex;gap:12px;align-items:center;}.cfg-label{font-size:11px;color:#8b949e;margin-right:4px;}.cfg-select{background:#0d1117;border:1px solid:#30363d;color:#e6edf3;padding:4px 8px;border-radius:4px;font-size:12px;}.cfg-checkbox{margin:0;}.cfg-radio-group{display:flex;gap:16px;flex-wrap:wrap;}.cfg-radio{display:flex;align-items:center;gap:6px;font-size:12px;color:#e6edf3;}.cfg-radio input{margin:0;}.cfg-input{background:#0d1117;border:1px solid:#30363d;color:#e6edf3;padding:6px 10px;border-radius:4px;font-size:12px;width:80px;text-align:center;}#iqcfg-ft{background:#161b22;border-top:1px solid #30363d;padding:14px 20px;display:flex;gap:10px;justify-content:flex-end;}.cfg-btn{padding:8px 16px;border:none;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;}.cfg-btn-primary{background:linear-gradient(135deg,#238636,#2ea043);color:#fff;}.cfg-btn-primary:hover{filter:brightness(1.12);}.cfg-btn-secondary{background:#21262d;color:#e6edf3;}.cfg-btn-secondary:hover{background:#30363d;}.cfg-inactive{opacity:0.5;}.cfg-field-inactive .cfg-field-controls{display:none;}";
  
  var modalEl = null;
  var overlayEl = null;
  var discoveredOrgs = null;
  var resolveCallback = null;
  
  function mount(firmOrgs){
    discoveredOrgs = firmOrgs;
    
    // Create overlay
    overlayEl = document.createElement("div");
    overlayEl.id = "iqcfg-overlay";
    document.body.appendChild(overlayEl);
    
    // Create modal
    modalEl = document.createElement("div");
    modalEl.id = "iqcfg";
    
    var activeOrgs = Object.keys(firmOrgs).filter(function(k){ return firmOrgs[k].count > 0; });
    var inactiveOrgs = Object.keys(firmOrgs).filter(function(k){ return firmOrgs[k].count === 0; });
    
    var orgFieldsHTML = activeOrgs.map(function(field){
      var org = firmOrgs[field];
      return '<div class="cfg-field" data-field="' + field + '">' +
        '<div class="cfg-field-hd">' +
          '<div class="cfg-field-name">' + org.label + '</div>' +
          '<div class="cfg-field-stats">' + org.count + ' opportunities</div>' +
        '</div>' +
        '<div class="cfg-field-samples">Examples: ' + org.sampleValues.join(", ") + '</div>' +
        '<div class="cfg-field-controls">' +
          '<label class="cfg-label">Priority:</label>' +
          '<select class="cfg-select cfg-priority" data-field="' + field + '">' +
            Array.from({length: activeOrgs.length}, function(_, i){
              return '<option value="' + (i+1) + '"' + (org.priority === i+1 ? ' selected' : '') + '>' + (i+1) + '</option>';
            }).join('') +
          '</select>' +
          '<label class="cfg-label" style="margin-left:12px;">' +
            '<input type="checkbox" class="cfg-checkbox cfg-risk" data-field="' + field + '"' + (org.includeInRisk ? ' checked' : '') + '> ' +
            'Include in Risk Analysis' +
          '</label>' +
        '</div>' +
      '</div>';
    }).join('');
    
    var inactiveHTML = inactiveOrgs.length > 0 ? 
      '<div style="margin-top:12px;padding:10px;background:#0d1117;border:1px solid #21262d;border-radius:6px;font-size:11px;color:#8b949e;">' +
        '✓ Not used in your opportunities: ' + inactiveOrgs.map(function(f){ return firmOrgs[f].label; }).join(", ") +
      '</div>' : '';
    
    modalEl.innerHTML = 
      '<div id="iqcfg-hd">' +
        '<h2>ExecIQ v7.0 Configuration</h2>' +
        '<p>Configure your analysis dimensions and reporting preferences</p>' +
      '</div>' +
      '<div id="iqcfg-bd">' +
        '<div class="cfg-section">' +
          '<h3>Firm Organization Fields</h3>' +
          '<p style="margin:0 0 14px;font-size:12px;color:#8b949e;">We found ' + activeOrgs.length + ' Firm Org fields in use. Prioritize them and select which to include in risk analysis.</p>' +
          orgFieldsHTML +
          inactiveHTML +
        '</div>' +
        '<div class="cfg-section">' +
          '<h3>Geographic Analysis</h3>' +
          '<p style="margin:0 0 10px;font-size:12px;color:#8b949e;">Select primary field for geographic breakdown:</p>' +
          '<div class="cfg-radio-group">' +
            '<label class="cfg-radio"><input type="radio" name="geo" value="City"> City</label>' +
            '<label class="cfg-radio"><input type="radio" name="geo" value="StateProv" checked> State</label>' +
            '<label class="cfg-radio"><input type="radio" name="geo" value="Country"> Country</label>' +
          '</div>' +
        '</div>' +
        '<div class="cfg-section">' +
          '<h3>Risk Thresholds</h3>' +
          '<p style="margin:0 0 10px;font-size:12px;color:#8b949e;">Set concentration risk levels (% of total pipeline):</p>' +
          '<div style="display:flex;gap:20px;align-items:center;">' +
            '<div><label class="cfg-label">Critical (🔴):</label> > <input type="number" class="cfg-input" id="cfg-crit" value="50" min="1" max="100"> %</div>' +
            '<div><label class="cfg-label">Warning (🟡):</label> > <input type="number" class="cfg-input" id="cfg-warn" value="33" min="1" max="100"> %</div>' +
          '</div>' +
        '</div>' +
        '<div class="cfg-section">' +
          '<h3>Forecast Settings</h3>' +
          '<p style="margin:0 0 10px;font-size:12px;color:#8b949e;">Configure revenue forecast parameters:</p>' +
          '<div style="margin-bottom:12px;">' +
            '<label class="cfg-label" style="display:block;margin-bottom:6px;">Forecast Window:</label>' +
            '<div class="cfg-radio-group">' +
              '<label class="cfg-radio"><input type="radio" name="forecast" value="30"> 30 days</label>' +
              '<label class="cfg-radio"><input type="radio" name="forecast" value="60"> 60 days</label>' +
              '<label class="cfg-radio"><input type="radio" name="forecast" value="90" checked> 90 days</label>' +
              '<label class="cfg-radio"><input type="radio" name="forecast" value="180"> 180 days</label>' +
            '</div>' +
          '</div>' +
          '<div>' +
            '<label class="cfg-label" style="display:block;margin-bottom:6px;">Pipeline Scope:</label>' +
            '<div class="cfg-radio-group">' +
              '<label class="cfg-radio"><input type="radio" name="scope" value="active" checked> Active Only</label>' +
              '<label class="cfg-radio"><input type="radio" name="scope" value="active+won"> Active + Won</label>' +
              '<label class="cfg-radio"><input type="radio" name="scope" value="all"> All Opportunities</label>' +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>' +
      '<div id="iqcfg-ft">' +
        '<button class="cfg-btn cfg-btn-secondary" id="cfg-defaults">Use Smart Defaults</button>' +
        '<button class="cfg-btn cfg-btn-primary" id="cfg-save">Save & Continue</button>' +
      '</div>';
    
    // Add CSS
    var styleEl = document.createElement("style");
    styleEl.id = "iqcfg-css";
    styleEl.textContent = CSS;
    document.head.appendChild(styleEl);
    
    document.body.appendChild(modalEl);
    
    // Bind events
    document.getElementById("cfg-save").onclick = handleSave;
    document.getElementById("cfg-defaults").onclick = handleDefaults;
  }
  
  function handleDefaults(){
    // Just use the auto-detected settings
    handleSave();
  }
  
  function handleSave(){
    var config = {
      version: "7.0",
      firmOrgs: [],
      geographic: {
        field: document.querySelector('input[name="geo"]:checked').value,
        label: CF[document.querySelector('input[name="geo"]:checked').value].l
      },
      risk: {
        criticalThreshold: parseFloat(document.getElementById("cfg-crit").value) / 100,
        warningThreshold: parseFloat(document.getElementById("cfg-warn").value) / 100
      },
      forecast: {
        windowDays: parseInt(document.querySelector('input[name="forecast"]:checked').value),
        scope: document.querySelector('input[name="scope"]:checked').value
      }
    };
    
    // Collect Firm Org settings
    Object.keys(discoveredOrgs).forEach(function(field){
      if (discoveredOrgs[field].count > 0){
        var priorityEl = document.querySelector('.cfg-priority[data-field="' + field + '"]');
        var riskEl = document.querySelector('.cfg-risk[data-field="' + field + '"]');
        
        config.firmOrgs.push({
          field: field,
          label: discoveredOrgs[field].label,
          priority: priorityEl ? parseInt(priorityEl.value) : 999,
          includeInRisk: riskEl ? riskEl.checked : false,
          oppCount: discoveredOrgs[field].count,
          enabled: true
        });
      }
    });
    
    // Sort by priority
    config.firmOrgs.sort(function(a, b){ return a.priority - b.priority; });
    
    // Save to localStorage
    ConfigManager.save(config);
    
    // Close modal
    destroy();
    
    // Resolve promise
    if (resolveCallback) resolveCallback(config);
  }
  
  function destroy(){
    if (modalEl) modalEl.remove();
    if (overlayEl) overlayEl.remove();
    document.getElementById("iqcfg-css")?.remove();
    modalEl = null;
    overlayEl = null;
  }
  
  function show(firmOrgs){
    return new Promise(function(resolve){
      resolveCallback = resolve;
      mount(firmOrgs);
    });
  }
  
  return { show };
})();

// UI Module (progress tracker)
var UI = (function(){
  var elSt,elFill,elLog,elBtn,elSettings;
  var CSS = "#iq7{position:fixed;top:16px;right:16px;z-index:2147483647;width:440px;background:#0d1117;color:#e6edf3;font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;border:1px solid #30363d;border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,.8);overflow:hidden;animation:iq7in .28s cubic-bezier(.16,1,.3,1)}@keyframes iq7in{from{opacity:0;transform:translateY(-14px)}to{opacity:1;transform:none}}#iq7 *{box-sizing:border-box}#iq7-hd{background:#161b22;border-bottom:1px solid #30363d;padding:13px 16px;display:flex;align-items:center;gap:10px}#iq7-logo{width:36px;height:36px;border-radius:9px;flex-shrink:0;background:linear-gradient(135deg,#388bfd,#a371f7);display:flex;align-items:center;justify-content:center;font-weight:800;font-size:14px;color:#fff}#iq7-ttl{flex:1}#iq7-ttl h3{margin:0;font-size:14px;font-weight:700;color:#f0f6fc}#iq7-ttl small{color:#8b949e;font-size:11px}#iq7-acts{display:flex;gap:6px}#iq7-settings{background:#21262d;border:1px solid #30363d;color:#8b949e;padding:4px 10px;border-radius:6px;cursor:pointer;font-size:11px;transition:all .15s}#iq7-settings:hover{background:#30363d;color:#e6edf3}#iq7-x{background:none;border:none;color:#8b949e;cursor:pointer;font-size:22px;line-height:1;padding:2px 5px;border-radius:4px}#iq7-x:hover{color:#f85149}#iq7-bd{padding:14px 16px}#iq7-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px}.iq7-kpi{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:8px 10px;text-align:center}.iq7-kv{font-size:20px;font-weight:700;color:#388bfd}.iq7-kl{font-size:10px;color:#8b949e;margin-top:2px}#iq7-st{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:9px 12px;margin-bottom:10px;font-size:12px;color:#8b949e;line-height:1.5;min-height:34px}.ok{color:#3fb950}.er{color:#f85149}.wn{color:#d29922}#iq7-prog{height:5px;background:#21262d;border-radius:3px;margin-bottom:10px;overflow:hidden}#iq7-fill{height:100%;width:0%;background:linear-gradient(90deg,#388bfd,#a371f7);border-radius:3px;transition:width .35s ease}#iq7-log{background:#010409;border:1px solid #21262d;border-radius:8px;padding:7px 10px;height:120px;overflow-y:auto;margin-bottom:13px;font-size:11px;font-family:monospace;color:#484f58}#iq7-log .ls{color:#3fb950}#iq7-log .le{color:#f85149}#iq7-log .lw{color:#d29922}#iq7-btn{width:100%;padding:10px;border:none;border-radius:8px;cursor:pointer;background:linear-gradient(135deg,#238636,#2ea043);color:#fff;font-size:13px;font-weight:700;transition:all .15s}#iq7-btn:hover:not(:disabled){filter:brightness(1.12)}#iq7-btn:disabled{background:#21262d;color:#6e7681;cursor:not-allowed;opacity:.6}";
  
  function mount(){
    var s = document.createElement("style");
    s.id = "iq7-css";
    s.textContent = CSS;
    document.head.appendChild(s);
    var el = document.createElement("div");
    el.id = "iq7";
    el.innerHTML = '<div id="iq7-hd"><div id="iq7-logo">IQ</div><div id="iq7-ttl"><h3>ExecIQ Extractor <span style="font-weight:400;font-size:10px;color:#8b949e">v7.0</span></h3><small>Unanet CRM — Advanced Analytics Suite</small></div><div id="iq7-acts"><button id="iq7-settings">⚙️ Settings</button><button id="iq7-x">×</button></div></div><div id="iq7-bd"><div id="iq7-kpis"><div class="iq7-kpi"><div class="iq7-kv" id="kv-o">--</div><div class="iq7-kl">Opportunities</div></div><div class="iq7-kpi"><div class="iq7-kv" id="kv-f">--</div><div class="iq7-kl">Fields</div></div><div class="iq7-kpi"><div class="iq7-kv" id="kv-r">0</div><div class="iq7-kl">Lookups</div></div></div><div id="iq7-st"><span class="ok">Starting...</span></div><div id="iq7-prog"><div id="iq7-fill"></div></div><div id="iq7-log"></div><button id="iq7-btn" disabled>Export Excel Report</button></div>';
    document.body.appendChild(el);
    elSt = document.getElementById("iq7-st");
    elFill = document.getElementById("iq7-fill");
    elLog = document.getElementById("iq7-log");
    elBtn = document.getElementById("iq7-btn");
    elSettings = document.getElementById("iq7-settings");
    document.getElementById("iq7-x").onclick = destroy;
  }
  
  function destroy(){
    window.__EXECIQ_V7__ = false;
    document.getElementById("iq7")?.remove();
    document.getElementById("iq7-css")?.remove();
  }
  
  function status(html,type){
    type = type || "ok";
    elSt.innerHTML = '<span class="' + type + '">' + (type === "ok" ? "●":type === "er" ? "✖":"⚠")+ " " + html + "</span>";
  }
  
  function prog(pct){
    elFill.style.width = Math.min(100,Math.max(0,pct))+ "%";
  }
  
  function log(msg,cls){
    var t = new Date().toLocaleTimeString("en-US",{hour12:false});
    var d = document.createElement("div");
    if (cls)d.className = cls;
    d.textContent = "[" + t + "] " + msg;
    elLog.appendChild(d);
    elLog.scrollTop = elLog.scrollHeight;
  }
  
  function kpi(id,val){
    var e = document.getElementById(id);
    if (e)e.textContent = val;
  }
  
  function enableExport(fn){
    elBtn.disabled = false;
    elBtn.textContent = "Export Excel + JSON";
    elBtn.onclick = fn;
  }
  
  function bindSettings(fn){
    if (elSettings) elSettings.onclick = fn;
  }
  
  return{mount,destroy,status,prog,log,kpi,enableExport,bindSettings };
})();

// Helper Functions
async function fetchJSON(url,opts){
  try{
    var r = await fetch(url,opts ||{credentials:"include",headers:{"X-Requested-With":"XMLHttpRequest"}});
    if (!r.ok)return null;
    var t = await r.text();
    var s = t.indexOf("{");
    if (s < 0)s = t.indexOf("[");
    if (s < 0)return null;
    return JSON.parse(t.slice(s));
  }catch(e){
    return null;
  }
}

async function postForm(url,params){
  return fetchJSON(url,{
    method:"POST",
    credentials:"include",
    headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest" },
    body:new URLSearchParams(params).toString()
  });
}

function parseCFC(data){
  if (!data)return [];
  if (data.response)data = data.response;
  if (data.Response)data = data.Response;
  if (Array.isArray(data.COLUMNS)&& Array.isArray(data.DATA)){
    return data.DATA.map(function(row){
      if (!Array.isArray(row))return row;
      var o = {};
      data.COLUMNS.forEach(function(c,i){o[c]= row[i];});
      return o;
    });
  }
  if (Array.isArray(data.DATA))return data.DATA;
  if (Array.isArray(data))return data;
  return [];
}

function buildMap(records){
  var map = {};
  records.forEach(function(r){
    var id = r.FIRMORGID || r.ID || r.id || r.STAGEID || r.ROLEID || r.CATEGORYID || r.CONTRACTTYPEID || r.CLIENTTYPEID || r.PROSPECTTYPEID || r.SUBMITTALTYPEID || "";
    var name = r.FIRMORGNAME || r.NAME || r.name || r.STAGENAME || r.ROLENAME || r.CATEGORYNAME || r.TYPENAME || r.DESCRIPTION || r.LABEL || r.CONTRACTNAME || r.CLIENTTYPENAME || r.PROSPECTTYPENAME || r.SUBMITTALTYPENAME || r.ROLENAME || r.CATEGORYNAME || "";
    id = String(id).trim();
    name = String(name).trim();
    if (id && id !== "0" && name)map[id]= name;
  });
  return map;
}

function resolve(val,map){
  if (val === null || val === undefined || val === "")return "";
  var s = String(val).trim();
  if (!s)return "";
  if (map[s])return map[s];
  return s.split(",").map(function(x){
    x = x.trim();
    return map[x]|| x;
  }).filter(Boolean).join(",");
}

function fmtDate(val){
  if (!val || val === "")return "";
  var s = String(val).trim();
  var d;
  if (/^\d{4}-\d{2}-\d{2}/.test(s))d = new Date(s);
  else{
    var m = s.match(/^(\w+),\s*(\d+)\s+(\d{4})/);
    if (!m)return "";
    d = new Date(m[1]+ " " + m[2]+ "," + m[3]);
  }
  if (!d || isNaN(d.getTime()))return "";
  return String(d.getMonth()+ 1).padStart(2,"0")+ "/" + String(d.getDate()).padStart(2,"0")+ "/" + d.getFullYear();
}

function stripHTML(str){
  if (!str)return "";
  return String(str).replace(/<br\s*\/?>/gi," | ").replace(/<[^>]+>/g,"").replace(/\s*\|\s*/g," | ").trim();
}

function classifyStatus(stageName,activeInd){
  var s = String(stageName || "").toLowerCase();
  if (String(activeInd)=== "2"){
    if (/\bwon\b|award|executed/.test(s))return "Won";
    if (/\blost\b|loss|no.go|dead|declined/.test(s))return "Lost";
    return "Closed";
  }
  return "Active";
}

function basePath(url){
  if (!url)return null;
  var m = url.match(/^https?:\/\/[^\/]+((?:\/[^\/]+)*\/)([^\/\?#]+)/);
  return m ? m[1]:null;
}

async function findOppBase(){
  var segs = window.location.pathname.split("/").filter(function(s){return s && !s.includes(".");});
  var candidates = ["/"];
  var p = "";
  for (var i = 0;i < segs.length;i++){
    p += "/" + segs[i];
    candidates.push(p + "/");
  }
  ["/contact/","/contact/opportunity/"].forEach(function(c){
    if (!candidates.includes(c))candidates.push(c);
  });
  var results = await Promise.all(candidates.map(async function(base){
    try{
      var r = await fetch(base + "oppActions.cfm",{
        method:"POST",
        credentials:"include",
        headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest" },
        body:"action=getOpportunityGridData&json=1&start=0&limit=1&ActiveInd=0&SalesCycle=NaN&officeId=0&divisionId=0&studioId=0&practiceAreaId=0&territoryId=0&stageId=0&priCatId=0&secCatId=0&masterSub=0&staffRoleId=0&dateCreated=0&dateModified=0&visibleColumns=VCHPROJECTNAME"
      });
      if (r.status === 404)return null;
      var t = await r.text();
      return (t.includes("ROWCOUNT")|| t.includes("DATA"))? base:null;
    }catch(e){
      return null;
    }
  }));
  return results.find(function(r){return r !== null;})|| null;
}

async function loadCustomFields(customFieldsURL){
  UI.log("Loading ALL custom field definitions...");
  var uuids = [];
  var sources = [];
  if (customFieldsURL){
    sources.push({url:customFieldsURL,method:"GET" });
    var cfBase = customFieldsURL.split("/customfieldsdata.cfc")[0]+ "/";
    sources.push({
      url:cfBase + "customfieldsdata.cfc?method=getCustomFieldsByTypeAsJSON",
      method:"GET"
    });
    sources.push({
      url:cfBase + "customfieldsdata.cfc",
      method:"POST",
      params:{method:"getCustomFieldsByType",type:"lead"}
    });
  }
  var allRecords = [];
  for (var si = 0;si < sources.length;si++){
    var src2 = sources[si];
    var data = src2.method === "GET" ? await fetchJSON(src2.url):await postForm(src2.url,src2.params);
    if (!data)continue;
    if (data.response)data = data.response;
    if (data.Response)data = data.Response;
    var records = [];
    if (Array.isArray(data))records = data;
    else if (Array.isArray(data.COLUMNS)&& Array.isArray(data.DATA))records = parseCFC(data);
    else if (Array.isArray(data.DATA))records = data.DATA;
    else if (data.customFields)records = data.customFields;
    else if (data.CustomFields)records = data.CustomFields;
    else if (data.items)records = data.items;
    UI.log("Source " + (si+1)+ ": " + records.length + " custom fields");
    allRecords = allRecords.concat(records);
  }
  var seen = {};
  allRecords.forEach(function(cf){
    var key = cf.EXTERNALID || cf.externalId || cf.ExternalId || cf.FieldKey || cf.fieldKey || cf.FIELDKEY || cf.VCHFIELDKEY || cf.Key || cf.key || cf.UUID || cf.uuid || cf.FIELDID || cf.fieldId || "";
    var label = cf.LABEL || cf.label || cf.FieldLabel || cf.fieldLabel || cf.FIELDLABEL || cf.VCHFIELDLABEL || cf.NAME || cf.name || cf.CUSTOMLABEL || key;
    var rawType = String(cf.CUSTOMFIELDTYPENAME || cf.customFieldTypeName || cf.FieldType || cf.fieldType || cf.FIELDTYPE || cf.VCHFIELDTYPE || "text").toLowerCase();
    var type = rawType.includes("currency")|| rawType.includes("money")? "currency":
      rawType.includes("percent")? "percent":
      rawType.includes("date")? "date":
      rawType.includes("number")|| rawType.includes("decimal")? "number":"text";
    key = String(key).trim();
    label = String(label).trim();
    if (!key || !label || seen[key])return;
    seen[key]= 1;
    var compassKey = "CF_" + key.replace(/[^a-zA-Z0-9]/g,"_");
    CF[compassKey]={l:label,t:type,e:1 };
    BF[key]= compassKey;
    BF[key.toLowerCase()]= compassKey;
    BF[key.toUpperCase()]= compassKey;
    uuids.push(key);
    UI.log("✓ " + label + " [" + type + "]","ls");
  });
  UI.log(uuids.length + " custom fields registered",uuids.length > 0 ? "ls":"lw");
  return uuids;
}

function applyFmt(ws,addr,fmt){
  if (ws[addr])ws[addr].z = fmt;
}

function buildOppSheet(rows,cols){
  if (!rows.length)return XLSX.utils.aoa_to_sheet([["No data"]]);
  var hdrs = cols.map(function(c){return c.label;});
  var data = rows.map(function(row){
    return cols.map(function(c){
      var v = row[c.key];
      return v != null ? v:"";
    });
  });
  var ws = XLSX.utils.aoa_to_sheet([hdrs].concat(data));
  data.forEach(function(_,ri){
    cols.forEach(function(c,ci){
      var addr = XLSX.utils.encode_cell({r:ri + 1,c:ci });
      if (c.type === "currency")applyFmt(ws,addr,'"$"#,##0');
      else if (c.type === "percent")applyFmt(ws,addr,"0%");
    });
  });
  ws["!autofilter"]={ref:"A1:" + XLSX.utils.encode_col(hdrs.length - 1)+ "1" };
  ws["!cols"]= hdrs.map(function(h){return{wch:Math.max(10,Math.min(40,String(h).length + 4))};});
  return ws;
}

function buildExecutiveDashboard(rows){
  try{
    function sumF(arr,k){return arr.reduce(function(t,r){var v = r[k];return t + (v != null && v !== "" ? parseFloat(v)|| 0:0);},0);}
    
    var active = rows.filter(function(r){return r._status === "Active";});
    var won = rows.filter(function(r){return r._status === "Won";});
    var lost = rows.filter(function(r){return r._status === "Lost";});
    var closed = rows.filter(function(r){return r._status === "Closed";});
    
    var totalPipeline = sumF(rows,"_fee");
    var weightedPipeline = sumF(rows,"_wgt");
    var closedN = won.length + lost.length;
    var winRate = closedN > 0 ? won.length / closedN : 0;
    var avgDealSize = rows.length > 0 ? totalPipeline / rows.length : 0;
    
    var prows = rows.filter(function(r){return r._prob != null;});
    var avgP = prows.length ? prows.reduce(function(t,r){return t + r._prob;},0)/ prows.length:0;
    
    // Forecast calculation (using config)
    var forecastDays = userConfig ? userConfig.forecast.windowDays : 90;
    var today = new Date();
    var forecastEnd = new Date(today.getTime() + forecastDays * 24 * 60 * 60 * 1000);
    
    var forecastOpps = rows.filter(function(r){
      if (!r.EstimatedSelectionDate) return false;
      var closeDate = new Date(r.EstimatedSelectionDate);
      return closeDate >= today && closeDate <= forecastEnd && r._status === "Active";
    });
    var forecastRevenue = sumF(forecastOpps, "_wgt");
    
    // Risk Analysis
    var riskData = [];
    
    // Client concentration
    var clientGroups = {};
    rows.forEach(function(r){
      var c = String(r.ClientCompany || "Unknown");
      if (!clientGroups[c]) clientGroups[c] = {name: c, fee: 0};
      clientGroups[c].fee += r._fee || 0;
    });
    var topClient = Object.values(clientGroups).sort(function(a,b){return b.fee - a.fee;})[0];
    if (topClient){
      var pct = totalPipeline > 0 ? topClient.fee / totalPipeline : 0;
      var risk = pct > (userConfig ? userConfig.risk.criticalThreshold : 0.5) ? "🔴 CRITICAL" :
                 pct > (userConfig ? userConfig.risk.warningThreshold : 0.33) ? "🟡 WARNING" : "🟢 HEALTHY";
      riskData.push(["Client", topClient.name, pct, risk]);
    }
    
    // Firm Org concentration
    if (userConfig && userConfig.firmOrgs){
      userConfig.firmOrgs.forEach(function(orgCfg){
        if (!orgCfg.includeInRisk) return;
        var groups = {};
        rows.forEach(function(r){
          var val = String(r[orgCfg.field] || "Unknown");
          val.split(",").forEach(function(v){
            v = v.trim();
            if (!groups[v]) groups[v] = {name: v, fee: 0};
            groups[v].fee += r._fee || 0;
          });
        });
        var top = Object.values(groups).sort(function(a,b){return b.fee - a.fee;})[0];
        if (top){
          var pct = totalPipeline > 0 ? top.fee / totalPipeline : 0;
          var risk = pct > (userConfig.risk.criticalThreshold) ? "🔴 CRITICAL" :
                     pct > (userConfig.risk.warningThreshold) ? "🟡 WARNING" : "🟢 HEALTHY";
          riskData.push([orgCfg.label, top.name, pct, risk]);
        }
      });
    }
    
    // Stage concentration
    var stageGroups = {};
    rows.forEach(function(r){
      var s = String(r._stage || "Unknown");
      if (!stageGroups[s]) stageGroups[s] = {name: s, fee: 0};
      stageGroups[s].fee += r._fee || 0;
    });
    var topStage = Object.values(stageGroups).sort(function(a,b){return b.fee - a.fee;})[0];
    if (topStage){
      var pct = totalPipeline > 0 ? topStage.fee / totalPipeline : 0;
      var risk = pct > (userConfig ? userConfig.risk.criticalThreshold : 0.5) ? "🔴 CRITICAL" :
                 pct > (userConfig ? userConfig.risk.warningThreshold : 0.33) ? "🟡 WARNING" : "🟢 HEALTHY";
      riskData.push(["Stage", topStage.name, pct, risk]);
    }
    
    // Geographic concentration
    if (userConfig && userConfig.geographic){
      var geoGroups = {};
      rows.forEach(function(r){
        var g = String(r[userConfig.geographic.field] || "Unknown");
        if (!geoGroups[g]) geoGroups[g] = {name: g, fee: 0};
        geoGroups[g].fee += r._fee || 0;
      });
      var topGeo = Object.values(geoGroups).sort(function(a,b){return b.fee - a.fee;})[0];
      if (topGeo){
        var pct = totalPipeline > 0 ? topGeo.fee / totalPipeline : 0;
        var risk = pct > (userConfig.risk.criticalThreshold) ? "🔴 CRITICAL" :
                   pct > (userConfig.risk.warningThreshold) ? "🟡 WARNING" : "🟢 HEALTHY";
        riskData.push(["Geography (" + userConfig.geographic.label + ")", topGeo.name, pct, risk]);
      }
    }
    
    var ts = new Date().toLocaleString("en-US");
    var aoa = [
      ["ExecIQ Executive Dashboard", "", "Generated:", ts],
      [],
      ["KEY PERFORMANCE INDICATORS"],
      ["Total Opportunities", rows.length],
      ["Active Opportunities", active.length],
      ["Gross Pipeline (Our Fee)", totalPipeline],
      ["Weighted Pipeline", weightedPipeline],
      ["Average Pwin", avgP],
      ["Average Deal Size", avgDealSize],
      [],
      ["WIN/LOSS METRICS"],
      ["Win Rate (Won / Closed)", winRate],
      ["Won Opportunities", won.length],
      ["Lost Opportunities", lost.length],
      ["Won Revenue", sumF(won,"_fee")],
      [],
      ["FORECAST (" + forecastDays + "-DAY)"],
      ["Forecasted Opportunities", forecastOpps.length],
      ["Forecasted Revenue (Weighted)", forecastRevenue],
      [],
      ["CONCENTRATION RISK ANALYSIS"],
      ["Category", "Top Entity", "% of Pipeline", "Risk Level"]
    ];
    
    riskData.forEach(function(r){ aoa.push(r); });
    
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    
    // Format currency cells
    [[5,1],[6,1],[8,1],[14,1],[18,1]].forEach(function(rc){
      applyFmt(ws,XLSX.utils.encode_cell({r:rc[0],c:rc[1]}),'"$"#,##0');
    });
    
    // Format percent cells
    [[7,1],[11,1]].forEach(function(rc){
      applyFmt(ws,XLSX.utils.encode_cell({r:rc[0],c:rc[1]}),"0%");
    });
    
    // Format risk analysis percent column
    for (var i = 0; i < riskData.length; i++){
      applyFmt(ws, XLSX.utils.encode_cell({r: 22 + i, c: 2}), "0%");
    }
    
    ws["!cols"]= [{wch:32},{wch:20},{wch:14},{wch:16}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Dashboard error:" + e.message]]);
  }
}

function buildSummarySheet(rows){
  try{
    function sumF(arr,k){return arr.reduce(function(t,r){var v = r[k];return t + (v != null && v !== "" ? parseFloat(v)|| 0:0);},0);}
    var active = rows.filter(function(r){return r._status === "Active";});
    var won = rows.filter(function(r){return r._status === "Won";});
    var lost = rows.filter(function(r){return r._status === "Lost";});
    var closed = rows.filter(function(r){return r._status === "Closed";});
    var prows = rows.filter(function(r){return r._prob != null;});
    var avgP = prows.length ? prows.reduce(function(t,r){return t + r._prob;},0)/ prows.length:0;
    var closedN = won.length + lost.length;
    var ts = new Date().toLocaleString("en-US");
    var aoa =[
      ["ExecIQ Opportunity Report","","Generated:",ts],
      [],
      ["PIPELINE SUMMARY"],
      ["Total Opportunities",rows.length],
      ["Gross Pipeline (Our Fee)",sumF(rows,"_fee")],
      ["Weighted Pipeline",sumF(rows,"_wgt")],
      ["Average Pwin",avgP],
      [],
      ["STATUS BREAKDOWN","Count","% of Total","Gross Fee"],
      ["Active / Open",active.length,active.length / (rows.length || 1),sumF(active,"_fee")],
      ["Won",won.length,won.length / (rows.length || 1),sumF(won,"_fee")],
      ["Lost",lost.length,lost.length / (rows.length || 1),sumF(lost,"_fee")],
      ["Closed (Other)",closed.length,closed.length / (rows.length || 1),sumF(closed,"_fee")],
      [],
      ["WIN RATE"],
      ["Win Rate (Won / Closed)",closedN > 0 ? won.length / closedN:0],
      ["Won Revenue",sumF(won,"_fee")],
    ];
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    [[4,1],[5,1],[9,3],[10,3],[11,3],[12,3],[16,1]].forEach(function(rc){applyFmt(ws,XLSX.utils.encode_cell({r:rc[0],c:rc[1]}),'"$"#,##0');});
    [[6,1],[9,2],[10,2],[11,2],[12,2],[15,1]].forEach(function(rc){applyFmt(ws,XLSX.utils.encode_cell({r:rc[0],c:rc[1]}),"0%");});
    ws["!cols"]= [{wch:32},{wch:18},{wch:12},{wch:18}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Summary error:" + e.message]]);
  }
}

function buildClientSheet(rows){
  try{
    var groups={};
    var hasOwner=false;
    var totalPipeline = rows.reduce(function(t,r){return t + (r._fee || 0);}, 0);
    
    rows.forEach(function(row){
      var client=String(row.ClientCompany||"Uncategorized").trim();
      var owner=String(row.OwnerCompany||"").trim();
      if(owner&&owner!==client){hasOwner=true;}
      var k=client||"Uncategorized";
      if(!groups[k]){groups[k]={client:k,count:0,active:0,won:0,lost:0,fee:0,wgt:0,lastMod:null};}
      var g=groups[k];
      g.count++;
      if(row._status==="Active"){g.active++;}
      if(row._status==="Won"){g.won++;}
      if(row._status==="Lost"){g.lost++;}
      g.fee+=row._fee||0;
      g.wgt+=row._wgt||0;
      var mod=row.DateModified;
      if(mod&&(!g.lastMod||mod>g.lastMod)){g.lastMod=mod;}
    });
    
    var sorted=Object.values(groups).sort(function(a,b){return b.fee-a.fee;});
    var hdr=["Client","Total Opps","Active","Won","Lost","Win Rate","Pipeline","% of Total","Weighted","Avg Opp Size","Last Activity"];
    var data=sorted.map(function(g){
      var closed=g.won+g.lost;
      var winRate=closed>0?g.won/closed:null;
      var avgSize=g.count>0?g.fee/g.count:0;
      var pctTotal = totalPipeline > 0 ? g.fee / totalPipeline : 0;
      return[g.client,g.count,g.active,g.won,g.lost,winRate,g.fee,pctTotal,g.wgt,avgSize,g.lastMod||""];
    });
    
    var ws=XLSX.utils.aoa_to_sheet([hdr].concat(data));
    for(var r=1;r<=data.length;r++){
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:5}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:6}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:7}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:8}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:9}),'"$"#,##0');
    }
    ws["!autofilter"]={ref:"A1:K1"};
    ws["!cols"]=[{wch:32},{wch:10},{wch:8},{wch:6},{wch:6},{wch:10},{wch:16},{wch:12},{wch:16},{wch:14},{wch:12}];
    return{sheet:ws,hasOwner:hasOwner};
  }catch(e){
    return{sheet:XLSX.utils.aoa_to_sheet([["Client error:"+e.message]]),hasOwner:false};
  }
}

function buildOwnerSheet(rows){
  try{
    var groups={};
    var totalPipeline = rows.reduce(function(t,r){return t + (r._fee || 0);}, 0);
    
    rows.forEach(function(row){
      var owner=String(row.OwnerCompany||"Uncategorized").trim();
      var k=owner||"Uncategorized";
      if(!groups[k]){groups[k]={owner:k,count:0,active:0,won:0,lost:0,fee:0,wgt:0,lastMod:null};}
      var g=groups[k];
      g.count++;
      if(row._status==="Active"){g.active++;}
      if(row._status==="Won"){g.won++;}
      if(row._status==="Lost"){g.lost++;}
      g.fee+=row._fee||0;
      g.wgt+=row._wgt||0;
      var mod=row.DateModified;
      if(mod&&(!g.lastMod||mod>g.lastMod)){g.lastMod=mod;}
    });
    
    var sorted=Object.values(groups).sort(function(a,b){return b.fee-a.fee;});
    var hdr=["Owner Company","Total Opps","Active","Won","Lost","Win Rate","Pipeline","% of Total","Weighted","Avg Opp Size","Last Activity"];
    var data=sorted.map(function(g){
      var closed=g.won+g.lost;
      var winRate=closed>0?g.won/closed:null;
      var avgSize=g.count>0?g.fee/g.count:0;
      var pctTotal = totalPipeline > 0 ? g.fee / totalPipeline : 0;
      return[g.owner,g.count,g.active,g.won,g.lost,winRate,g.fee,pctTotal,g.wgt,avgSize,g.lastMod||""];
    });
    
    var ws=XLSX.utils.aoa_to_sheet([hdr].concat(data));
    for(var r=1;r<=data.length;r++){
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:5}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:6}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:7}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:8}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:9}),'"$"#,##0');
    }
    ws["!autofilter"]={ref:"A1:K1"};
    ws["!cols"]=[{wch:32},{wch:10},{wch:8},{wch:6},{wch:6},{wch:10},{wch:16},{wch:12},{wch:16},{wch:14},{wch:12}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Owner error:"+e.message]]);
  }
}

function buildStageSheet(rows){
  try{
    var groups = {};
    rows.forEach(function(row){
      var k = row._stage || "Unknown";
      if (!groups[k])groups[k]={stage:k,status:row._status || "",count:0,fee:0,wgt:0,ps:0,pn:0 };
      var g = groups[k];
      g.count++;
      g.fee += row._fee || 0;
      g.wgt += row._wgt || 0;
      if (row._prob != null){g.ps += row._prob;g.pn++;}
    });
    var ord ={Active:0,Won:1,Lost:2,Closed:3 };
    var sorted = Object.values(groups).sort(function(a,b){return ((ord[a.status]|| 9)- (ord[b.status]|| 9))|| b.count - a.count;});
    var hdr = ["Stage","Status","Count","% of Total","Gross Fee","Weighted Fee","Avg Pwin"];
    var data = sorted.map(function(g){return [g.stage,g.status,g.count,g.count / (rows.length || 1),g.fee,g.wgt,g.pn > 0 ? g.ps / g.pn:null];});
    var tots = ["TOTAL","",rows.length,1,sorted.reduce(function(s,g){return s+g.fee;},0),sorted.reduce(function(s,g){return s+g.wgt;},0),""];
    var ws = XLSX.utils.aoa_to_sheet([hdr].concat(data).concat([tots]));
    for (var r = 1;r <= data.length + 1;r++){
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:3}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:4}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:5}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:6}),"0%");
    }
    ws["!autofilter"]={ref:"A1:G1" };
    ws["!cols"]= [{wch:30},{wch:10},{wch:8},{wch:10},{wch:16},{wch:16},{wch:10}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Stage error:" + e.message]]);
  }
}

function buildPipelineDetailSheet(rows, cols){
  // Build comprehensive detail view with ALL fields
  return buildOppSheet(rows, cols);
}

function buildFirmOrgSheet(rows, fieldKey, fieldLabel){
  try{
    var groups = {};
    var total = rows.reduce(function(t,r){return t + (r._fee || 0);},0);
    rows.forEach(function(row){
      var val = String(row[fieldKey]|| "Uncategorized");
      val.split(",").map(function(s){return s.trim();}).forEach(function(v){
        var k = v || "Uncategorized";
        if (!groups[k])groups[k]={count:0,active:0,won:0,lost:0,fee:0,wgt:0};
        groups[k].count++;
        if (row._status === "Active") groups[k].active++;
        if (row._status === "Won") groups[k].won++;
        if (row._status === "Lost") groups[k].lost++;
        groups[k].fee += row._fee || 0;
        groups[k].wgt += row._wgt || 0;
      });
    });
    
    var sorted = Object.entries(groups).sort(function(a,b){return b[1].fee - a[1].fee;});
    var hdr = [fieldLabel,"Total Opps","Active","Won","Lost","Win Rate","Pipeline","% of Total","Weighted","Avg Opp Size"];
    var data = sorted.map(function(e){
      var g = e[1];
      var closed = g.won + g.lost;
      var winRate = closed > 0 ? g.won / closed : null;
      var avgSize = g.count > 0 ? g.fee / g.count : 0;
      var pctTotal = total > 0 ? g.fee / total : 0;
      return [e[0],g.count,g.active,g.won,g.lost,winRate,g.fee,pctTotal,g.wgt,avgSize];
    });
    
    var ws = XLSX.utils.aoa_to_sheet([hdr].concat(data));
    data.forEach(function(_,ri){
      var r = ri + 1;
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:5}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:6}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:7}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:8}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:9}),'"$"#,##0');
    });
    ws["!autofilter"]={ref:"A1:J1" };
    ws["!cols"]= [{wch:32},{wch:10},{wch:8},{wch:6},{wch:6},{wch:10},{wch:16},{wch:12},{wch:16},{wch:14}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Firm Org error:" + e.message]]);
  }
}

function buildGeographicSheet(rows, geoField, geoLabel){
  try{
    var groups = {};
    var total = rows.reduce(function(t,r){return t + (r._fee || 0);},0);
    rows.forEach(function(row){
      var val = String(row[geoField]|| "Uncategorized").trim();
      var k = val || "Uncategorized";
      if (!groups[k])groups[k]={count:0,active:0,won:0,lost:0,fee:0,wgt:0};
      groups[k].count++;
      if (row._status === "Active") groups[k].active++;
      if (row._status === "Won") groups[k].won++;
      if (row._status === "Lost") groups[k].lost++;
      groups[k].fee += row._fee || 0;
      groups[k].wgt += row._wgt || 0;
    });
    
    var sorted = Object.entries(groups).sort(function(a,b){return b[1].fee - a[1].fee;});
    var hdr = [geoLabel,"Total Opps","Active","Won","Lost","Win Rate","Pipeline","% of Total","Weighted","Avg Opp Size"];
    var data = sorted.map(function(e){
      var g = e[1];
      var closed = g.won + g.lost;
      var winRate = closed > 0 ? g.won / closed : null;
      var avgSize = g.count > 0 ? g.fee / g.count : 0;
      var pctTotal = total > 0 ? g.fee / total : 0;
      return [e[0],g.count,g.active,g.won,g.lost,winRate,g.fee,pctTotal,g.wgt,avgSize];
    });
    
    var ws = XLSX.utils.aoa_to_sheet([hdr].concat(data));
    data.forEach(function(_,ri){
      var r = ri + 1;
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:5}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:6}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:7}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:8}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:9}),'"$"#,##0');
    });
    ws["!autofilter"]={ref:"A1:J1" };
    ws["!cols"]= [{wch:32},{wch:10},{wch:8},{wch:6},{wch:6},{wch:10},{wch:16},{wch:12},{wch:16},{wch:14}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Geographic error:" + e.message]]);
  }
}

function buildForecastSheet(rows){
  try{
    if (!userConfig) return XLSX.utils.aoa_to_sheet([["Configuration required"]]);
    
    var forecastDays = userConfig.forecast.windowDays;
    var today = new Date();
    var forecastEnd = new Date(today.getTime() + forecastDays * 24 * 60 * 60 * 1000);
    
    // Filter opportunities within forecast window
    var forecastOpps = rows.filter(function(r){
      if (!r.EstimatedSelectionDate) return false;
      var closeDate = new Date(r.EstimatedSelectionDate);
      return closeDate >= today && closeDate <= forecastEnd && r._status === "Active";
    });
    
    // Sort by expected close date
    forecastOpps.sort(function(a,b){
      var da = new Date(a.EstimatedSelectionDate || 0);
      var db = new Date(b.EstimatedSelectionDate || 0);
      return da - db;
    });
    
    var hdr = ["Opp Number","Opportunity Name","Client","Expected Close","Our Fee","Pwin","Weighted Value","Stage"];
    var data = forecastOpps.map(function(r){
      return [
        r.OpportunityNumber || "",
        r.OpportunityName || "",
        r.ClientCompany || "",
        r.EstimatedSelectionDate || "",
        r._fee || 0,
        r._prob || 0,
        r._wgt || 0,
        r._stage || ""
      ];
    });
    
    var totalFee = forecastOpps.reduce(function(t,r){return t + (r._fee || 0);}, 0);
    var totalWgt = forecastOpps.reduce(function(t,r){return t + (r._wgt || 0);}, 0);
    
    var summary = [
      ["FORECAST SUMMARY (" + forecastDays + " Days)",""],
      ["Total Opportunities", forecastOpps.length],
      ["Gross Pipeline", totalFee],
      ["Weighted Pipeline", totalWgt],
      ["Forecast Period", today.toLocaleDateString() + " to " + forecastEnd.toLocaleDateString()],
      []
    ];
    
    var ws = XLSX.utils.aoa_to_sheet(summary.concat([hdr]).concat(data));
    
    // Format summary
    applyFmt(ws, XLSX.utils.encode_cell({r:2,c:1}), '"$"#,##0');
    applyFmt(ws, XLSX.utils.encode_cell({r:3,c:1}), '"$"#,##0');
    
    // Format data
    data.forEach(function(_,ri){
      var r = ri + 7; // Offset for summary
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:4}),'"$"#,##0');
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:5}),"0%");
      applyFmt(ws,XLSX.utils.encode_cell({r:r,c:6}),'"$"#,##0');
    });
    
    ws["!autofilter"]={ref:"A7:H7"};
    ws["!cols"]=[{wch:14},{wch:32},{wch:28},{wch:14},{wch:14},{wch:8},{wch:16},{wch:24}];
    return ws;
  }catch(e){
    return XLSX.utils.aoa_to_sheet([["Forecast error:" + e.message]]);
  }
}

async function main(){
  UI.mount();
  UI.prog(3);
  UI.status("Starting extraction...");
  
  // Check for existing config
  userConfig = ConfigManager.load();
  
  UI.status("Finding oppActions.cfm...");
  var oppBase = await findOppBase();
  if (!oppBase){
    UI.status("Cannot find oppActions.cfm. Navigate to Opportunities page.","er");
    return;
  }
  UI.log("✓ oppBase: " + oppBase,"ls");
  UI.prog(10);
  
  var entries = performance.getEntriesByType("resource");
  var firmDataURL = null;
  var customFieldsURL = null;
  var lookupBase = null;
  entries.forEach(function(e){
    var url = e.name;
    if (!url.includes(window.location.host))return;
    if (!firmDataURL && /firmData\.cfc/i.test(url)){
      firmDataURL = url.split("?")[0]+ "?method=GetFirmOrgData";
    }
    if (!customFieldsURL && /customfieldsdata\.cfc/i.test(url)){
      customFieldsURL = url;
    }
    if (!lookupBase && /\b(stage|contractType|role|primaryCategory|secondaryCategory)\.cfc/i.test(url)){
      lookupBase = basePath(url);
    }
  });
  
  UI.log("firmData: " + (firmDataURL ? "found":"NOT FOUND"),firmDataURL ? "ls":"lw");
  UI.log("customFields: " + (customFieldsURL ? "found":"NOT FOUND"),customFieldsURL ? "ls":"lw");
  UI.log("lookupBase: " + (lookupBase ? "found":"NOT FOUND"),lookupBase ? "ls":"lw");
  UI.prog(18);
  
  UI.status("Loading custom field definitions...");
  var customUUIDs = await loadCustomFields(customFieldsURL);
  UI.prog(30);
  
  UI.status("Loading lookup tables...");
  var firmOrgMap = {},stageMap = {},contractMap = {},clientTypeMap = {},prospectMap = {},priCatMap = {},secCatMap = {},roleMap = {},submittalMap = {};
  if (firmDataURL){
    var fd = await fetchJSON(firmDataURL);
    if (fd && typeof fd === "object"){
      var orgKeys = ["divisions","practiceAreas","territories","offices","studios","officeDivisions","regions","departments"];
      orgKeys.forEach(function(key){
        var section = fd[key];
        if (!section)return;
        var rows = section.DATA || section.data;
        if (!Array.isArray(rows))return;
        rows.forEach(function(row){
          if (!Array.isArray(row))return;
          var id = String(row[0]|| "").trim();
          var name = String(row[1]|| "").trim();
          if (id && id !== "0" && name)firmOrgMap[id]= name;
        });
      });
      UI.log("✓ Org lookups: " + Object.keys(firmOrgMap).length,"ls");
    }
  }
  if (lookupBase){
    async function getLookup(base,file,method){
      var url = base + file;
      var d = await postForm(url,{method:method });
      if (d && (Array.isArray(d)|| d.DATA || d.data || d.COLUMNS))return d;
      d = await fetchJSON(url + "?method=" + method);
      if (d && (Array.isArray(d)|| d.DATA || d.data || d.COLUMNS))return d;
      return null;
    }
    var [stR,coR,clR,prR,pcR,scR,roR,suR]= await Promise.all([
      getLookup(lookupBase,"stage.cfc","getList"),
      getLookup(lookupBase,"contractType.cfc","getContractTypes"),
      getLookup(lookupBase,"clientType.cfc","getClientTypes"),
      getLookup(lookupBase,"prospectType.cfc","getProspectTypes"),
      getLookup(lookupBase,"primaryCategory.cfc","getList"),
      getLookup(lookupBase,"secondaryCategory.cfc","getList"),
      getLookup(lookupBase,"role.cfc","getList"),
      getLookup(lookupBase,"submittalType.cfc","getSubmittalTypes"),
    ]);
    stageMap = buildMap(parseCFC(stR));
    contractMap = buildMap(parseCFC(coR));
    clientTypeMap = buildMap(parseCFC(clR));
    prospectMap = buildMap(parseCFC(prR));
    priCatMap = buildMap(parseCFC(pcR));
    secCatMap = buildMap(parseCFC(scR));
    roleMap = buildMap(parseCFC(roR));
    submittalMap= buildMap(parseCFC(suR));
    var totalLookups = Object.keys(firmOrgMap).length + Object.keys(stageMap).length + Object.keys(contractMap).length + Object.keys(clientTypeMap).length;
    UI.kpi("kv-r",totalLookups);
    UI.log("✓ Total lookups: " + totalLookups,"ls");
  }
  UI.prog(45);
  
  var allColumns = STANDARD_COLUMNS.concat(customUUIDs);
  UI.log("Requesting " + allColumns.length + " fields (" + customUUIDs.length + " custom)","ls");
  
  UI.status("Fetching opportunities...");
  var bodyParts =["action=getOpportunityGridData","json=1","sort=STAGEID","dir=ASC","selectedCurrency=USD","start=0","limit=9999","view=0","ActiveInd=0","SalesCycle=NaN","officeId=0","divisionId=0","studioId=0","practiceAreaId=0","territoryId=0","stageId=0","priCatId=0","secCatId=0","masterSub=0","staffRoleId=0","dateCreated=0","dateModified=0","dateCreatedModified=0","filteredSearch=0","search=" ];
  allColumns.forEach(function(c){
    var isUUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(c);
    bodyParts.push("visibleColumns=" + (isUUID ? c.toLowerCase():encodeURIComponent(c)));
  });
  var oppData = null;
  try{
    var r = await fetch(oppBase + "oppActions.cfm",{
      method:"POST",
      credentials:"include",
      headers:{"Content-Type":"application/x-www-form-urlencoded","X-Requested-With":"XMLHttpRequest" },
      body:bodyParts.join("&")
    });
    var text = await r.text();
    var s = text.indexOf("{");
    if (s < 0)s = text.indexOf("[");
    if (s >= 0)oppData = JSON.parse(text.slice(s));
  }catch(e){
    UI.log("Fetch error:" + e.message,"le");
  }
  if (!oppData || (!Array.isArray(oppData.DATA)&& !Array.isArray(oppData))){
    UI.status("No data from oppActions.cfm","er");
    return;
  }
  var opps = Array.isArray(oppData.DATA)? oppData.DATA:oppData;
  UI.log("✓ Loaded " + opps.length + " opportunities","ls");
  UI.kpi("kv-o",opps.length);
  UI.prog(60);
  
  UI.status("Processing records...");
  var seenCompass = {};
  var cleanRows = opps.map(function(opp){
    var row = {};
    row._status = classifyStatus(opp.STAGENAME || "",opp.ACTIVEIND);
    row._prob = (opp.IPROBABILITY != null && opp.IPROBABILITY !== "")? parseFloat(opp.IPROBABILITY)/ 100:null;
    row._fee = (opp.IFIRMFEE != null && opp.IFIRMFEE !== "")? parseFloat(opp.IFIRMFEE)|| 0:0;
    row._wgt = (opp.IFACTOREDFEE != null && opp.IFACTOREDFEE !== "")? parseFloat(opp.IFACTOREDFEE)|| 0:0;
    row._stage = String(opp.STAGENAME || "");
    row["OpportunityNumber"]= String(opp.VCHLEADNUMBER || "");
    row["OpportunityName"]= String(opp.VCHPROJECTNAME || "");
    row["ClientCompany"]= String(opp.COMPANY || "");
    row["OwnerCompany"]= String(opp.OWNERCOMPANY || opp.OWNER || "");
    row["Stage"]= String(opp.STAGENAME || "");
    row["OpportunityStatus"]= row._status;
    row["FirmEstimatedFee"]= row._fee;
    row["FactoredFee"]= row._wgt;
    row["WinProbability"]= row._prob;
    ["OpportunityNumber","OpportunityName","ClientCompany","OwnerCompany","Stage","OpportunityStatus","FirmEstimatedFee","FactoredFee","WinProbability"].forEach(function(k){seenCompass[k]= 1;});
    Object.keys(opp).forEach(function(bf){
      if (SUPPRESS.has(bf))return;
      var val = opp[bf];
      if (val === null || val === undefined || val === "")return;
      var compassKey = BF[bf]|| BF[bf.toUpperCase()]|| BF[bf.toLowerCase()];
      if (!compassKey)return;
      if (["OpportunityNumber","OpportunityName","ClientCompany","OwnerCompany","Stage","OpportunityStatus","FirmEstimatedFee","FactoredFee","WinProbability"].includes(compassKey))return;
      var cfg = CF[compassKey];
      if (!cfg)return;
      var type = cfg.t;
      var display;
      if (compassKey === "PrimaryContact"){
        display = stripHTML(opp.OWNERCONTACT || "").split("|")[0].trim();
      }else if (compassKey === "Offices"){
        display = resolve(val,firmOrgMap);
      }else if (compassKey === "Studios"){
        display = resolve(val,firmOrgMap);
      }else if (compassKey === "PracticeAreas"){
        display = resolve(val,firmOrgMap);
      }else if (compassKey === "Divisions"){
        display = resolve(val,firmOrgMap);
      }else if (compassKey === "Territories"){
        display = resolve(val,firmOrgMap);
      }else if (compassKey === "OfficeDivision"){
        display = resolve(val,firmOrgMap);
      }else if (compassKey === "PrimaryCategory"){
        display = resolve(val,priCatMap);
      }else if (compassKey === "SecondaryCategory"){
        display = resolve(val,secCatMap);
      }else if (compassKey === "ContractType"){
        display = resolve(val,contractMap);
      }else if (compassKey === "ClientTypes"){
        display = resolve(val,clientTypeMap)|| String(val);
      }else if (compassKey === "ProspectTypes"){
        display = resolve(val,prospectMap)|| String(val);
      }else if (compassKey === "OpportunityRole"){
        display = resolve(val,roleMap)|| String(opp.ROLENAME || val);
      }else if (compassKey === "SubmittalType"){
        display = String(val);
      }else if (compassKey === "RFPReceived"){
        display = String(val)=== "1" ? "Yes":String(val)=== "0" ? "No":String(val);
      }else if (type === "date"){
        display = fmtDate(val);
      }else if (type === "currency"){
        display = parseFloat(val)|| 0;
      }else if (type === "percent"){
        display = (parseFloat(val)|| 0)/ 100;
      }else{
        display = String(val).trim();
      }
      if (display != null && display !== ""){
        row[compassKey]= display;
        seenCompass[compassKey]= 1;
      }
    });
    return row;
  });
  
  var activeCols = [];
  KEY_ORDER.forEach(function(ck){
    if (seenCompass[ck]&& CF[ck])activeCols.push({key:ck,label:CF[ck].l,type:CF[ck].t });
  });
  Object.keys(CF).forEach(function(ck){
    if (!KEY_ORDER.includes(ck)&& seenCompass[ck])activeCols.push({key:ck,label:CF[ck].l,type:CF[ck].t });
  });
  
  UI.kpi("kv-f",activeCols.length);
  UI.log("✓ " + cleanRows.length + " rows, " + activeCols.length + " columns","ls");
  UI.prog(70);
  
  // Discover Firm Orgs
  UI.status("Discovering Firm Organization fields...");
  var firmOrgUsage = discoverFirmOrgs(cleanRows);
  UI.log("✓ Firm Org discovery complete","ls");
  
  // Show config UI if needed
  if (!userConfig){
    UI.status("Waiting for configuration...","wn");
    userConfig = await ConfigUI.show(firmOrgUsage);
    UI.log("✓ Configuration saved","ls");
  }
  
  // Bind settings button to reconfigure
  UI.bindSettings(async function(){
    UI.log("Opening configuration...","ls");
    userConfig = await ConfigUI.show(firmOrgUsage);
    UI.log("✓ Configuration updated","ls");
  });
  
  UI.prog(75);
  UI.status("Building Excel workbook...");
  
  try{
    await Promise.race([
      new Promise(function(res,rej){
        if (window.XLSX){res();return;}
        var s = document.createElement("script");
        s.src = SHEETJS;
        s.onerror = function(){rej(new Error("SheetJS CDN failed"));};
        s.onload = function(){
          var n = 0;
          var iv = setInterval(function(){
            if (window.XLSX){clearInterval(iv);res();}
            else if (++n > 30){clearInterval(iv);rej(new Error("XLSX global not found"));}
          },100);
        };
        document.head.appendChild(s);
      }),
      new Promise(function(_,rej){setTimeout(function(){rej(new Error("SheetJS timeout"));},20000);})
    ]);
  }catch(se){
    UI.status("SheetJS failed:" + se.message,"er");
    return;
  }
  UI.log("✓ SheetJS ready","ls");
  UI.prog(85);
  
  var wb = XLSX.utils.book_new();
  var sheetCount = 0;
  
  try{
    UI.log("Sheet " + (++sheetCount) + ": Executive Dashboard...");
    XLSX.utils.book_append_sheet(wb,buildExecutiveDashboard(cleanRows),"Executive Dashboard");
  }catch(e){UI.log("Executive Dashboard error:" + e.message,"le");}
  
  try{
    UI.log("Sheet " + (++sheetCount) + ": Stage Analysis...");
    XLSX.utils.book_append_sheet(wb,buildStageSheet(cleanRows),"Stage Analysis");
  }catch(e){UI.log("Stage Analysis error:" + e.message,"le");}
  
  try{
    UI.log("Sheet " + (++sheetCount) + ": Client Analysis...");
    var clientResult=buildClientSheet(cleanRows);
    XLSX.utils.book_append_sheet(wb,clientResult.sheet,"Client Analysis");
    if(clientResult.hasOwner){
      UI.log("Sheet " + (++sheetCount) + ": Owner Analysis...");
      XLSX.utils.book_append_sheet(wb,buildOwnerSheet(cleanRows),"Owner Analysis");
    }
  }catch(e){UI.log("Client/Owner error:" + e.message,"le");}
  
  try{
    UI.log("Sheet " + (++sheetCount) + ": Pipeline Detail...");
    XLSX.utils.book_append_sheet(wb,buildPipelineDetailSheet(cleanRows,activeCols),"Pipeline Detail");
  }catch(e){UI.log("Pipeline Detail error:" + e.message,"le");}
  
  // Dynamic Firm Org sheets
  if (userConfig && userConfig.firmOrgs){
    userConfig.firmOrgs.forEach(function(orgCfg){
      if (!orgCfg.enabled) return;
      try{
        UI.log("Sheet " + (++sheetCount) + ": " + orgCfg.label + " Analysis...");
        var sheetName = (orgCfg.label + " Analysis").slice(0,31);
        XLSX.utils.book_append_sheet(wb,buildFirmOrgSheet(cleanRows,orgCfg.field,orgCfg.label),sheetName);
      }catch(e){UI.log(orgCfg.label + " error:" + e.message,"le");}
    });
  }
  
  try{
    UI.log("Sheet " + (++sheetCount) + ": Geographic Breakdown...");
    var geoField = userConfig ? userConfig.geographic.field : "StateProv";
    var geoLabel = userConfig ? userConfig.geographic.label : "State";
    XLSX.utils.book_append_sheet(wb,buildGeographicSheet(cleanRows,geoField,geoLabel),"Geographic Breakdown");
  }catch(e){UI.log("Geographic error:" + e.message,"le");}
  
  try{
    UI.log("Sheet " + (++sheetCount) + ": Forecast...");
    XLSX.utils.book_append_sheet(wb,buildForecastSheet(cleanRows),"Forecast");
  }catch(e){UI.log("Forecast error:" + e.message,"le");}
  
  UI.prog(100);
  UI.status(cleanRows.length + " opportunities ready");
  UI.log("✓ Complete! " + sheetCount + " sheets created. Click Export","ls");
  
  UI.enableExport(function(){
    try{
      var buf = XLSX.write(wb,{bookType:"xlsx",type:"array",compression:false });
      var blob = new Blob([buf],{type:"application/octet-stream" });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "ExecIQ_Report_" + new Date().toISOString().slice(0,10)+ ".xlsx";
      document.body.appendChild(a);
      a.click();
      setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},1500);
      setTimeout(function(){
        var exportData = {
          metadata: {
            extractedAt: new Date().toISOString(),
            opportunityCount: cleanRows.length,
            fieldCount: activeCols.length,
            customFieldCount: customUUIDs.length,
            configuration: userConfig
          },
          fieldDefinitions: CF,
          opportunities: cleanRows
        };
        var jsonBlob = new Blob([JSON.stringify(exportData, null, 2)], {type: 'application/json'});
        var a2 = document.createElement('a');
        a2.href = URL.createObjectURL(jsonBlob);
        a2.download = 'ExecIQ_Data_' + new Date().toISOString().slice(0,10) + '.json';
        document.body.appendChild(a2);
        a2.click();
        setTimeout(function(){URL.revokeObjectURL(a2.href);a2.remove();},1500);
        UI.log("✓ Downloads started!","ls");
        UI.status("Downloaded Excel + JSON");
      },2000);
    }catch(e){
      UI.log("Export error:" + e.message,"le");
      UI.status("Export failed","er");
    }
  });
}

main();
})()
