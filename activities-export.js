// Unanet Activities Export Tool
// Fetches all activities from Unanet CRM and exports to Excel

(async function() {
    // SHEETJS with styling support - same library as ExecIQ
    var SHEETJS = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
    
    // Create filter dialog first
    const filterOverlay = document.createElement('div');
    filterOverlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.7);
        z-index: 999999;
        display: flex;
        justify-content: center;
        align-items: center;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    `;
    
    const filterDialog = document.createElement('div');
    filterDialog.style.cssText = `
        background: white;
        padding: 30px;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        min-width: 500px;
        max-width: 600px;
        position: relative;
    `;
    
    const filterTitle = document.createElement('h2');
    filterTitle.textContent = 'Unanet Activities Export - Filters';
    filterTitle.style.cssText = 'margin: 0 0 20px 0; color: #333; font-size: 20px;';
    
    const filterForm = document.createElement('div');
    filterForm.style.cssText = 'margin: 20px 0;';
    
    // Helper to create form field
    function createField(label, inputElement) {
        const field = document.createElement('div');
        field.style.cssText = 'margin-bottom: 15px;';
        
        const labelEl = document.createElement('label');
        labelEl.textContent = label;
        labelEl.style.cssText = 'display: block; margin-bottom: 5px; color: #555; font-size: 14px; font-weight: 500;';
        
        field.appendChild(labelEl);
        field.appendChild(inputElement);
        return field;
    }
    
    // Date range filter
    const dateRangeSelect = document.createElement('select');
    dateRangeSelect.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';
    dateRangeSelect.innerHTML = `
        <option value="all">All dates</option>
        <option value="last30">Last 30 days</option>
        <option value="last60">Last 60 days</option>
        <option value="last90" selected>Last 90 days</option>
        <option value="last6months">Last 6 months</option>
        <option value="last12months">Last 12 months</option>
        <option value="currentYear">Current year</option>
        <option value="next90">Next 90 days (future)</option>
        <option value="last90toNext90">Last 90 days to Next 90 days</option>
        <option value="custom">Custom date range...</option>
    `;
    
    // Custom date inputs (hidden by default)
    const customDateDiv = document.createElement('div');
    customDateDiv.style.cssText = 'display: none; margin-top: 10px; padding: 10px; background: #f5f5f5; border-radius: 4px;';
    
    const startDateInput = document.createElement('input');
    startDateInput.type = 'date';
    startDateInput.style.cssText = 'width: 48%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; margin-right: 4%;';
    
    const endDateInput = document.createElement('input');
    endDateInput.type = 'date';
    endDateInput.style.cssText = 'width: 48%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';
    
    const customLabel = document.createElement('div');
    customLabel.textContent = 'Start Date → End Date';
    customLabel.style.cssText = 'font-size: 12px; color: #666; margin-bottom: 5px;';
    
    customDateDiv.appendChild(customLabel);
    customDateDiv.appendChild(startDateInput);
    customDateDiv.appendChild(endDateInput);
    
    dateRangeSelect.addEventListener('change', () => {
        if (dateRangeSelect.value === 'custom') {
            customDateDiv.style.display = 'block';
        } else {
            customDateDiv.style.display = 'none';
        }
    });
    
    // Owner filter
    const ownerInput = document.createElement('input');
    ownerInput.type = 'text';
    ownerInput.placeholder = 'e.g., John Smith (leave blank for all)';
    ownerInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';
    
    // Company filter
    const companyInput = document.createElement('input');
    companyInput.type = 'text';
    companyInput.placeholder = 'e.g., Acme Corp (leave blank for all)';
    companyInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';
    
    // Opportunity filter
    const opportunityInput = document.createElement('input');
    opportunityInput.type = 'text';
    opportunityInput.placeholder = 'e.g., (25-0044) Security Assessment (leave blank for all)';
    opportunityInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';
    
    // Build form
    filterForm.appendChild(createField('Date Range:', dateRangeSelect));
    filterForm.appendChild(customDateDiv);
    filterForm.appendChild(createField('Owner (partial match):', ownerInput));
    filterForm.appendChild(createField('Company (partial match):', companyInput));
    filterForm.appendChild(createField('Opportunity (partial match):', opportunityInput));
    
    // Buttons
    const buttonDiv = document.createElement('div');
    buttonDiv.style.cssText = 'margin-top: 25px; display: flex; gap: 10px; justify-content: flex-end;';
    
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText = 'padding: 10px 20px; background: #f0f0f0; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;';
    cancelBtn.onmouseover = () => cancelBtn.style.background = '#e0e0e0';
    cancelBtn.onmouseout = () => cancelBtn.style.background = '#f0f0f0';
    
    const exportBtn = document.createElement('button');
    exportBtn.textContent = 'Export Activities';
    exportBtn.style.cssText = 'padding: 10px 20px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;';
    exportBtn.onmouseover = () => exportBtn.style.background = '#45a049';
    exportBtn.onmouseout = () => exportBtn.style.background = '#4CAF50';
    
    buttonDiv.appendChild(cancelBtn);
    buttonDiv.appendChild(exportBtn);
    
    filterDialog.appendChild(filterTitle);
    filterDialog.appendChild(filterForm);
    filterDialog.appendChild(buttonDiv);
    filterOverlay.appendChild(filterDialog);
    document.body.appendChild(filterOverlay);
    
    // Handle cancel
    cancelBtn.onclick = () => {
        document.body.removeChild(filterOverlay);
    };
    
    // Handle export
    exportBtn.onclick = async () => {
        // Get filter values
        const filters = {
            dateRange: dateRangeSelect.value,
            customStartDate: startDateInput.value,
            customEndDate: endDateInput.value,
            owner: ownerInput.value.trim().toLowerCase(),
            company: companyInput.value.trim().toLowerCase(),
            opportunity: opportunityInput.value.trim().toLowerCase()
        };
        
        // Calculate date range
        let startDate = null;
        let endDate = null;
        
        const now = new Date();
        
        switch(filters.dateRange) {
            case 'last30':
                startDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
                endDate = now;
                break;
            case 'last60':
                startDate = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000);
                endDate = now;
                break;
            case 'last90':
                startDate = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
                endDate = now;
                break;
            case 'last6months':
                startDate = new Date(now.getTime() - 180 * 24 * 60 * 60 * 1000);
                endDate = now;
                break;
            case 'last12months':
                startDate = new Date(now.getTime() - 365 * 24 * 60 * 60 * 1000);
                endDate = now;
                break;
            case 'currentYear':
                startDate = new Date(now.getFullYear(), 0, 1);
                endDate = new Date(now.getFullYear(), 11, 31);
                break;
            case 'next90':
                startDate = now;
                endDate = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);
                break;
            case 'last90toNext90':
                startDate = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
                endDate = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);
                break;
            case 'custom':
                if (filters.customStartDate) startDate = new Date(filters.customStartDate);
                if (filters.customEndDate) endDate = new Date(filters.customEndDate);
                break;
            case 'all':
            default:
                startDate = null;
                endDate = null;
                break;
        }
        
        // Remove filter dialog
        document.body.removeChild(filterOverlay);
        
        // Start export process
        await runExport(filters, startDate, endDate);
    };
    
    // ═══════════════════════════════════════════════════════════
    // ACTIVITIES ANALYSIS BUILDER
    // ═══════════════════════════════════════════════════════════
    
    function buildActivitiesAnalysis(activities) {
        const now = new Date();
        const day30 = new Date(now - 30 * 24 * 60 * 60 * 1000);
        const day60 = new Date(now - 60 * 24 * 60 * 60 * 1000);
        const day90 = new Date(now - 90 * 24 * 60 * 60 * 1000);
        const day180 = new Date(now - 180 * 24 * 60 * 60 * 1000);
        
        // Helper: Split semicolon-delimited fields into arrays
        function splitField(val) {
            if (!val) return [];
            return val.split(';').map(s => s.trim()).filter(s => s);
        }
        
        // ─── 1. BUILD CLIENT ANALYSIS ───────────────────────────
        const clientMap = {};
        
        activities.forEach(act => {
            const companies = splitField(act['Company(ies)']);
            const owners = splitField(act['Owner(s)']);
            const contacts = splitField(act['Contact(s)']);
            const opps = splitField(act['Opportunity(ies)']);
            const callType = act['Call Type'] || '';
            const startDate = act['Start Date'];
            
            companies.forEach(company => {
                if (!company) return;
                
                if (!clientMap[company]) {
                    clientMap[company] = {
                        total: 0,
                        last30: 0,
                        last60: 0,
                        last90: 0,
                        owners: new Set(),
                        contacts: new Set(),
                        opps: new Set(),
                        meetings: 0,
                        meetings90: 0,
                        callTypes: {},
                        lastActivity: null,
                        lastOwner: '',
                        lastCallType: '',
                        activities: []
                    };
                }
                
                const client = clientMap[company];
                client.total++;
                client.activities.push(act);
                
                if (startDate >= day30) client.last30++;
                if (startDate >= day60) client.last60++;
                if (startDate >= day90) client.last90++;
                
                owners.forEach(o => client.owners.add(o));
                contacts.forEach(c => client.contacts.add(c));
                opps.forEach(op => client.opps.add(op));
                
                if (callType.toLowerCase().includes('meeting')) {
                    client.meetings++;
                    if (startDate >= day90) client.meetings90++;
                }
                
                client.callTypes[callType] = (client.callTypes[callType] || 0) + 1;
                
                if (!client.lastActivity || startDate > client.lastActivity) {
                    client.lastActivity = startDate;
                    client.lastOwner = owners[0] || '';
                    client.lastCallType = callType;
                }
            });
        });
        
        // ─── 2. BUILD OWNER ANALYSIS ─────────────────────────────
        const ownerMap = {};
        
        activities.forEach(act => {
            const owners = splitField(act['Owner(s)']);
            const companies = splitField(act['Company(ies)']);
            const contacts = splitField(act['Contact(s)']);
            const callType = act['Call Type'] || '';
            const startDate = act['Start Date'];
            
            owners.forEach(owner => {
                if (!owner) return;
                
                if (!ownerMap[owner]) {
                    ownerMap[owner] = {
                        total: 0,
                        companies: new Set(),
                        contacts: new Set(),
                        callTypes: {},
                        lastActivity: null
                    };
                }
                
                const o = ownerMap[owner];
                o.total++;
                companies.forEach(c => o.companies.add(c));
                contacts.forEach(c => o.contacts.add(c));
                o.callTypes[callType] = (o.callTypes[callType] || 0) + 1;
                
                if (!o.lastActivity || startDate > o.lastActivity) {
                    o.lastActivity = startDate;
                }
            });
        });
        
        // ─── 3. BUILD OPPORTUNITY ANALYSIS ───────────────────────
        const oppMap = {};
        
        activities.forEach(act => {
            const opps = splitField(act['Opportunity(ies)']);
            const contacts = splitField(act['Contact(s)']);
            const owners = splitField(act['Owner(s)']);
            const startDate = act['Start Date'];
            
            opps.forEach(opp => {
                if (!opp) return;
                
                if (!oppMap[opp]) {
                    oppMap[opp] = {
                        total: 0,
                        last30: 0,
                        contacts: new Set(),
                        owners: new Set(),
                        lastActivity: null
                    };
                }
                
                const o = oppMap[opp];
                o.total++;
                if (startDate >= day30) o.last30++;
                contacts.forEach(c => o.contacts.add(c));
                owners.forEach(ow => o.owners.add(ow));
                
                if (!o.lastActivity || startDate > o.lastActivity) {
                    o.lastActivity = startDate;
                }
            });
        });
        
        // ─── 4. BUILD CONTACT ANALYSIS ───────────────────────────
        const contactMap = {};
        
        activities.forEach(act => {
            const contacts = splitField(act['Contact(s)']);
            const owners = splitField(act['Owner(s)']);
            const opps = splitField(act['Opportunity(ies)']);
            const callType = act['Call Type'] || '';
            const startDate = act['Start Date'];
            
            contacts.forEach(contact => {
                if (!contact) return;
                
                if (!contactMap[contact]) {
                    contactMap[contact] = {
                        total: 0,
                        owners: new Set(),
                        opps: new Set(),
                        callTypes: {},
                        lastActivity: null
                    };
                }
                
                const c = contactMap[contact];
                c.total++;
                owners.forEach(o => c.owners.add(o));
                opps.forEach(op => c.opps.add(op));
                c.callTypes[callType] = (c.callTypes[callType] || 0) + 1;
                
                if (!c.lastActivity || startDate > c.lastActivity) {
                    c.lastActivity = startDate;
                }
            });
        });
        
        // ─── 5. CALCULATE RELATIONSHIP HEALTH SCORES ─────────────
        const healthScores = [];
        
        Object.entries(clientMap).forEach(([company, data]) => {
            // Recency Score (25 pts)
            const daysSince = data.lastActivity ? Math.floor((now - data.lastActivity) / (24 * 60 * 60 * 1000)) : 999;
            let recencyScore = 0;
            if (daysSince <= 7) recencyScore = 25;
            else if (daysSince <= 14) recencyScore = 22;
            else if (daysSince <= 30) recencyScore = 18;
            else if (daysSince <= 45) recencyScore = 12;
            else if (daysSince <= 60) recencyScore = 8;
            else if (daysSince <= 90) recencyScore = 4;
            
            // Activity Frequency (20 pts)
            let freqScore = 0;
            if (data.last90 >= 20) freqScore = 20;
            else if (data.last90 >= 15) freqScore = 17;
            else if (data.last90 >= 10) freqScore = 13;
            else if (data.last90 >= 6) freqScore = 9;
            else if (data.last90 >= 3) freqScore = 5;
            else if (data.last90 >= 1) freqScore = 2;
            
            // Meeting Engagement (15 pts)
            let meetingScore = 0;
            if (data.meetings90 >= 8) meetingScore = 15;
            else if (data.meetings90 >= 5) meetingScore = 12;
            else if (data.meetings90 >= 3) meetingScore = 8;
            else if (data.meetings90 >= 1) meetingScore = 4;
            
            // Contact Coverage (10 pts)
            const contactCount = data.contacts.size;
            let contactScore = 0;
            if (contactCount >= 8) contactScore = 10;
            else if (contactCount >= 6) contactScore = 8;
            else if (contactCount >= 4) contactScore = 6;
            else if (contactCount >= 2) contactScore = 3;
            else if (contactCount >= 1) contactScore = 1;
            
            // Owner Coverage (10 pts)
            const ownerCount = data.owners.size;
            let ownerScore = 0;
            if (ownerCount >= 5) ownerScore = 10;
            else if (ownerCount >= 4) ownerScore = 8;
            else if (ownerCount >= 3) ownerScore = 6;
            else if (ownerCount >= 2) ownerScore = 3;
            else if (ownerCount >= 1) ownerScore = 1;
            
            // Opportunity Engagement (10 pts)
            const hasOpps = data.opps.size > 0;
            const hasRecentOppActivity = data.activities.some(a => 
                a['Opportunity(ies)'] && a['Start Date'] >= day30
            );
            let oppScore = 0;
            if (hasOpps && data.opps.size > 1 && hasRecentOppActivity) oppScore = 10;
            else if (hasOpps && hasRecentOppActivity) oppScore = 7;
            else if (hasOpps && !hasRecentOppActivity) oppScore = 3;
            
            const totalScore = recencyScore + freqScore + meetingScore + contactScore + ownerScore + oppScore;
            
            // Status
            let status = 'Critical';
            if (totalScore >= 85) status = 'Healthy';
            else if (totalScore >= 70) status = 'Stable';
            else if (totalScore >= 55) status = 'Watch';
            else if (totalScore >= 40) status = 'Elevated Risk';
            
            // Critical Flags - IMPROVED LOGIC
            const flags = [];
            if (daysSince >= 90) flags.push('No Activity 90+ Days');
            if (contactCount === 0) flags.push('No Contacts');
            else if (contactCount === 1) flags.push('Single Contact');
            if (ownerCount === 0) flags.push('No Owner');
            else if (ownerCount === 1) flags.push('Single Owner');
            if (hasOpps && !hasRecentOppActivity) flags.push('Stale Pursuit');
            if (data.meetings90 === 0 && data.total > 5) flags.push('No Meetings');
            
            healthScores.push({
                company,
                score: totalScore,
                status,
                daysSince,
                owners: ownerCount,
                contacts: contactCount,
                meetings90: data.meetings90,
                activities90: data.last90,
                flags: flags.join('; ') || 'None'
            });
        });
        
        // ─── 6. BUILD WORKSHEET WITH ALL SECTIONS ────────────────
        const aoa = [];
        const BLUE = '4472C4';
        const WHITE = 'FFFFFF';
        const GREEN = 'C6E0B4';
        const YELLOW = 'FFE699';
        const ORANGE = 'F4B084';
        const RED = 'F8CBAD';
        
        function header(text) {
            aoa.push([text]);
        }
        
        function spacer() {
            aoa.push([]);
        }
        
        // SECTION 1: Most Active Clients
        header('MOST ACTIVE CLIENTS');
        aoa.push(['Company', 'Total Activities', 'Last 30d', 'Last 60d', 'Last 90d', 'Owners', 'Contacts', 'Meetings (90d)', 'Last Activity']);
        
        const topClients = Object.entries(clientMap)
            .sort((a, b) => b[1].total - a[1].total)
            .slice(0, 25);
        
        topClients.forEach(([company, data]) => {
            aoa.push([
                company,
                data.total,
                data.last30,
                data.last60,
                data.last90,
                data.owners.size,
                data.contacts.size,
                data.meetings90,
                data.lastActivity
            ]);
        });
        
        spacer();
        spacer();
        
        // SECTION 2: Dormant Clients
        header('CLIENTS NOT CONTACTED RECENTLY');
        aoa.push(['Company', 'Days Since Last Contact', 'Last Activity', 'Last Owner', 'Last Type', 'Open Opps', 'Risk Level']);
        
        const dormant = Object.entries(clientMap)
            .filter(([_, data]) => {
                const days = Math.floor((now - data.lastActivity) / (24 * 60 * 60 * 1000));
                return days >= 30;
            })
            .sort((a, b) => {
                const daysA = Math.floor((now - a[1].lastActivity) / (24 * 60 * 60 * 1000));
                const daysB = Math.floor((now - b[1].lastActivity) / (24 * 60 * 60 * 1000));
                return daysB - daysA;
            })
            .slice(0, 50);
        
        dormant.forEach(([company, data]) => {
            const days = Math.floor((now - data.lastActivity) / (24 * 60 * 60 * 1000));
            let risk = 'Watch';
            if (days >= 180) risk = 'Dormant';
            else if (days >= 90) risk = 'Gap';
            else if (days >= 60) risk = 'At Risk';
            
            aoa.push([
                company,
                days,
                data.lastActivity,
                data.lastOwner,
                data.lastCallType,
                data.opps.size,
                risk
            ]);
        });
        
        spacer();
        spacer();
        
        // SECTION 3: Activity by Owner
        header('ACTIVITY BY OWNER');
        aoa.push(['Owner', 'Total Activities', 'Companies', 'Contacts', 'Avg per Month', 'Last Activity']);
        
        const ownerList = Object.entries(ownerMap)
            .sort((a, b) => b[1].total - a[1].total);
        
        ownerList.forEach(([owner, data]) => {
            const avgPerMonth = Math.round(data.total / 3);
            aoa.push([
                owner,
                data.total,
                data.companies.size,
                data.contacts.size,
                avgPerMonth,
                data.lastActivity
            ]);
        });
        
        spacer();
        spacer();
        
        // SECTION 4: Opportunity Engagement
        header('OPPORTUNITY ENGAGEMENT');
        aoa.push(['Opportunity', 'Total Activities', 'Last 30d', 'Contacts', 'Owners', 'Last Touch']);
        
        const oppList = Object.entries(oppMap)
            .sort((a, b) => b[1].total - a[1].total)
            .slice(0, 50);
        
        oppList.forEach(([opp, data]) => {
            aoa.push([
                opp,
                data.total,
                data.last30,
                data.contacts.size,
                data.owners.size,
                data.lastActivity
            ]);
        });
        
        spacer();
        spacer();
        
        // SECTION 5: Contact Engagement
        header('CONTACT ENGAGEMENT');
        aoa.push(['Contact', 'Interactions', 'Last Contact', 'Related Opps', 'Internal Owners']);
        
        const contactList = Object.entries(contactMap)
            .sort((a, b) => b[1].total - a[1].total)
            .slice(0, 50);
        
        contactList.forEach(([contact, data]) => {
            aoa.push([
                contact,
                data.total,
                data.lastActivity,
                data.opps.size,
                data.owners.size
            ]);
        });
        
        spacer();
        spacer();
        
        // SECTION 6: Relationship Coverage (Risk Analysis)
        header('RELATIONSHIP COVERAGE ANALYSIS');
        aoa.push(['Company', 'Internal Owners', 'Client Contacts', 'Coverage Status']);
        
        const coverage = Object.entries(clientMap)
            .sort((a, b) => (a[1].owners.size + a[1].contacts.size) - (b[1].owners.size + b[1].contacts.size))
            .slice(0, 50);
        
        coverage.forEach(([company, data]) => {
            let status = 'Strong';
            if (data.owners.size === 1 && data.contacts.size === 1) status = 'Critical - Single Threaded';
            else if (data.owners.size === 1) status = 'Risk - Single Owner';
            else if (data.contacts.size === 1) status = 'Risk - Single Contact';
            else if (data.owners.size === 2 && data.contacts.size === 2) status = 'Moderate';
            
            aoa.push([
                company,
                data.owners.size,
                data.contacts.size,
                status
            ]);
        });
        
        spacer();
        spacer();
        
        // SECTION 7: Relationship Health Scores
        header('RELATIONSHIP HEALTH SCORES');
        aoa.push(['Company', 'Score', 'Status', 'Days Since Contact', 'Owners', 'Contacts', 'Meetings (90d)', 'Activities (90d)', 'Risk Flags']);
        
        healthScores.sort((a, b) => b.score - a.score);
        
        healthScores.forEach(h => {
            aoa.push([
                h.company,
                h.score,
                h.status,
                h.daysSince,
                h.owners,
                h.contacts,
                h.meetings90,
                h.activities90,
                h.flags
            ]);
        });
        
        // ─── 7. CREATE AND FORMAT WORKSHEET ──────────────────────
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wsRange = XLSX.utils.decode_range(ws['!ref']);
        
        // Track section header rows for proper formatting
        const sectionHeaders = [];
        const dataHeaders = [];
        
        for (let R = 0; R <= wsRange.e.r; R++) {
            const cell = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
            if (cell && cell.v && typeof cell.v === 'string') {
                if (cell.v.includes('CLIENTS') || cell.v.includes('ACTIVITY') || 
                    cell.v.includes('OPPORTUNITY') || cell.v.includes('CONTACT') || 
                    cell.v.includes('RELATIONSHIP') || cell.v.includes('SCORES')) {
                    sectionHeaders.push(R);
                    // Next row after section header is data header
                    if (R + 1 <= wsRange.e.r) {
                        dataHeaders.push(R + 1);
                    }
                }
            }
        }
        
        // Apply formatting
        for (let R = 0; R <= wsRange.e.r; R++) {
            for (let C = 0; C <= wsRange.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[addr]) continue;
                
                const cellValue = ws[addr].v;
                const isSectionHeader = sectionHeaders.includes(R) && C === 0;
                const isDataHeader = dataHeaders.includes(R);
                
                // Base style
                ws[addr].s = {
                    font: {
                        name: 'Arial',
                        sz: 10,
                        bold: false,
                        color: { rgb: '000000' }
                    },
                    alignment: {
                        vertical: 'center',
                        horizontal: 'left'
                    }
                };
                
                // Section headers
                if (isSectionHeader) {
                    ws[addr].s.font.bold = true;
                    ws[addr].s.font.sz = 12;
                    ws[addr].s.font.color = { rgb: BLUE };
                    ws[addr].s.alignment.horizontal = 'left';
                }
                // Data column headers
                else if (isDataHeader) {
                    ws[addr].s.font.bold = true;
                    ws[addr].s.font.color = { rgb: WHITE };
                    ws[addr].s.fill = {
                        patternType: 'solid',
                        fgColor: { rgb: BLUE }
                    };
                }
                
                // Format dates
                if (ws[addr].v instanceof Date) {
                    ws[addr].t = 'd';
                    ws[addr].z = 'mm/dd/yyyy';
                }
                
                // Conditional formatting for Health Score Status column
                if (typeof cellValue === 'string') {
                    if (cellValue === 'Healthy') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: GREEN }
                        };
                        ws[addr].s.font.bold = true;
                    } else if (cellValue === 'Stable') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: YELLOW }
                        };
                    } else if (cellValue === 'Watch') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: ORANGE }
                        };
                    } else if (cellValue === 'Elevated Risk' || cellValue === 'Critical') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: RED }
                        };
                        ws[addr].s.font.bold = true;
                    }
                    
                    // Risk level formatting for dormant clients
                    if (cellValue === 'Dormant' || cellValue === 'Gap') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: RED }
                        };
                        ws[addr].s.font.bold = true;
                    } else if (cellValue === 'At Risk') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: ORANGE }
                        };
                    }
                    
                    // Coverage status formatting
                    if (cellValue.includes('Critical')) {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: RED }
                        };
                        ws[addr].s.font.bold = true;
                    } else if (cellValue.includes('Risk')) {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: ORANGE }
                        };
                    } else if (cellValue === 'Moderate') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: YELLOW }
                        };
                    } else if (cellValue === 'Strong') {
                        ws[addr].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: GREEN }
                        };
                    }
                }
                
                // Number formatting for scores
                if (typeof cellValue === 'number' && C === 1) {
                    const rowAbove = ws[XLSX.utils.encode_cell({ r: R - 1, c: C })]?.v;
                    if (rowAbove === 'Score') {
                        ws[addr].z = '0';
                        // Score color coding
                        if (cellValue >= 85) {
                            ws[addr].s.fill = {
                                patternType: 'solid',
                                fgColor: { rgb: GREEN }
                            };
                            ws[addr].s.font.bold = true;
                        } else if (cellValue >= 70) {
                            ws[addr].s.fill = {
                                patternType: 'solid',
                                fgColor: { rgb: YELLOW }
                            };
                        } else if (cellValue >= 55) {
                            ws[addr].s.fill = {
                                patternType: 'solid',
                                fgColor: { rgb: ORANGE }
                            };
                        } else {
                            ws[addr].s.fill = {
                                patternType: 'solid',
                                fgColor: { rgb: RED }
                            };
                            ws[addr].s.font.bold = true;
                        }
                    }
                }
            }
        }
        
        // Auto-size columns
        const maxCol = wsRange.e.c;
        const colWidths = [];
        for (let C = 0; C <= maxCol; C++) {
            let maxWidth = 10;
            for (let R = 0; R <= wsRange.e.r; R++) {
                const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
                if (cell && cell.v) {
                    const cellValue = cell.v.toString();
                    maxWidth = Math.max(maxWidth, cellValue.length);
                }
            }
            colWidths.push({ wch: Math.min(maxWidth + 2, 50) });
        }
        ws['!cols'] = colWidths;
        
        // Set row heights for section headers
        if (!ws['!rows']) ws['!rows'] = [];
        sectionHeaders.forEach(R => {
            ws['!rows'][R] = { hpt: 25 };
        });
        
        return ws;
    }
    
    // Main export function
    async function runExport(filters, startDate, endDate) {
        // Create progress overlay
        const overlay = document.createElement('div');
        overlay.id = 'unanet-activities-export-overlay';
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            z-index: 999999;
            display: flex;
            justify-content: center;
            align-items: center;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        `;
        
        const dialog = document.createElement('div');
        dialog.style.cssText = `
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
            min-width: 400px;
            max-width: 500px;
            position: relative;
        `;
        
        // Close button
        const closeBtn = document.createElement('button');
        closeBtn.innerHTML = '×';
        closeBtn.style.cssText = `
            position: absolute;
            top: 10px;
            right: 10px;
            background: none;
            border: none;
            font-size: 28px;
            color: #999;
            cursor: pointer;
            padding: 0;
            width: 30px;
            height: 30px;
            line-height: 30px;
            text-align: center;
            border-radius: 4px;
            transition: all 0.2s;
        `;
        closeBtn.onmouseover = () => {
            closeBtn.style.background = '#f0f0f0';
            closeBtn.style.color = '#333';
        };
        closeBtn.onmouseout = () => {
            closeBtn.style.background = 'none';
            closeBtn.style.color = '#999';
        };
        closeBtn.onclick = () => {
            document.body.removeChild(overlay);
        };
        
        const title = document.createElement('h2');
        title.textContent = 'Unanet Activities Export';
        title.style.cssText = 'margin: 0 0 20px 0; color: #333; font-size: 20px; padding-right: 30px;';
        
        const status = document.createElement('div');
        status.id = 'export-status';
        status.style.cssText = 'margin: 15px 0; color: #666; font-size: 14px; line-height: 1.6; max-height: 200px; overflow-y: auto;';
        
        const progressBar = document.createElement('div');
        progressBar.style.cssText = `
            width: 100%;
            height: 8px;
            background: #e0e0e0;
            border-radius: 4px;
            overflow: hidden;
            margin: 15px 0;
        `;
        
        const progressFill = document.createElement('div');
        progressFill.style.cssText = `
            width: 0%;
            height: 100%;
            background: linear-gradient(90deg, #4CAF50, #45a049);
            border-radius: 4px;
            transition: width 0.3s ease;
        `;
        progressBar.appendChild(progressFill);
        
        dialog.appendChild(closeBtn);
        dialog.appendChild(title);
        dialog.appendChild(status);
        dialog.appendChild(progressBar);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        
        function updateStatus(message, progress = null) {
            status.innerHTML += `<div>${message}</div>`;
            status.scrollTop = status.scrollHeight;
            if (progress !== null) {
                progressFill.style.width = `${progress}%`;
            }
        }
        
        function closeDialog() {
            setTimeout(() => {
                if (document.body.contains(overlay)) {
                    document.body.removeChild(overlay);
                }
            }, 3000);
        }
        
        const API_URL = 'https://services.cosential.com/com/model/activities/callLog.cfc';
        const BATCH_SIZE = 100;
        
        // Helper function to strip HTML and extract text
        function stripHTML(html) {
            if (!html) return '';
            const temp = document.createElement('div');
            temp.innerHTML = html;
            return temp.textContent.trim();
        }
        
        // Helper function to extract names from HTML structure
        function extractNames(html, selector) {
            if (!html) return '';
            const temp = document.createElement('div');
            temp.innerHTML = html;
            const elements = temp.querySelectorAll(selector);
            const names = Array.from(elements).map(el => el.textContent.trim()).filter(n => n);
            return names.join('; ');
        }
        
        // Helper function to parse linked entities
        function parseLinkedField(html) {
            if (!html) return '';
            
            // Extract company names
            let companies = extractNames(html, '.company a');
            if (companies) return companies;
            
            // Extract contact names
            let contacts = extractNames(html, '.contact a');
            if (contacts) return contacts;
            
            // Extract opportunity names
            let opps = extractNames(html, '.opportunity a, .lead a');
            if (opps) return opps;
            
            // Extract personnel names
            let personnel = extractNames(html, '.personnel');
            if (personnel) return personnel;
            
            // Fallback to strip all HTML
            return stripHTML(html);
        }
        
        // Decode status
        function decodeStatus(statusCode) {
            if (statusCode === 0) return 'Not Started';
            if (statusCode === 1) return 'Completed';
            return statusCode;
        }
        
        // Parse date string and return as JavaScript Date object
        function parseDate(dateStr) {
            if (!dateStr) return null;
            try {
                const date = new Date(dateStr);
                if (isNaN(date.getTime())) return null;
                return date;
            } catch (e) {
                return null;
            }
        }
        
        // Fetch activities with pagination
        async function fetchActivities(start = 0) {
            const formData = new URLSearchParams({
                start: start,
                limit: BATCH_SIZE,
                sort: 'startDateDateOnly',
                dir: 'ASC',
                groupBy: 'status',
                groupDir: 'ASC',
                method: 'getGridData',
                xaction: 'read'
            });
            
            const response = await fetch(API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: formData
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            return await response.json();
        }
        
        try {
            updateStatus('🚀 Starting export...', 5);
            
            // Show active filters
            if (startDate || endDate || filters.owner || filters.company || filters.opportunity) {
                updateStatus('🔍 Active filters:');
                if (startDate && endDate) {
                    updateStatus(`   📅 Date: ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
                } else if (startDate) {
                    updateStatus(`   📅 Date: After ${startDate.toLocaleDateString()}`);
                } else if (endDate) {
                    updateStatus(`   📅 Date: Before ${endDate.toLocaleDateString()}`);
                }
                if (filters.owner) updateStatus(`   👤 Owner: ${filters.owner}`);
                if (filters.company) updateStatus(`   🏢 Company: ${filters.company}`);
                if (filters.opportunity) updateStatus(`   💼 Opportunity: ${filters.opportunity}`);
            }
            
            // Fetch all activities
            let allActivities = [];
            let start = 0;
            let totalRecords = 0;
            
            updateStatus('📥 Fetching activities from Unanet...', 10);
            
            // First request to get total count
            const firstBatch = await fetchActivities(0);
            totalRecords = firstBatch.TOTAL;
            updateStatus(`📊 Found ${totalRecords} activities in system`, 15);
            
            // Process first batch
            const columns = firstBatch.DATA.COLUMNS;
            const data = firstBatch.DATA.DATA;
            allActivities.push(...data);
            
            // Fetch remaining batches
            start = BATCH_SIZE;
            while (start < totalRecords) {
                const currentProgress = 15 + (start / totalRecords) * 40;
                updateStatus(`📥 Fetching activities ${start} to ${Math.min(start + BATCH_SIZE, totalRecords)}...`, currentProgress);
                const batch = await fetchActivities(start);
                allActivities.push(...batch.DATA.DATA);
                start += BATCH_SIZE;
            }
            
            updateStatus(`✅ Fetched all ${allActivities.length} activities`, 55);
            
            // Parse activities into structured data
            updateStatus('🔄 Parsing activity data...', 60);
            
            let parsedActivities = allActivities.map(activity => {
                const activityObj = {};
                columns.forEach((col, idx) => {
                    activityObj[col] = activity[idx];
                });
                
                return {
                    'Owner(s)': parseLinkedField(activityObj.CALLERS),
                    'Call Type': activityObj.CALLTYPE || '',
                    'Status': decodeStatus(activityObj.STATUS),
                    'Start Date': parseDate(activityObj.STARTDATEDATEONLY),
                    'End Date': parseDate(activityObj.ENDDATEDATEONLY),
                    'Company(ies)': parseLinkedField(activityObj.LINKEDCOMPANIES),
                    'Subject': activityObj.SUBJECT || '',
                    'Comments': stripHTML(activityObj.COMMENTS),
                    'Attendee(s)': parseLinkedField(activityObj.ATTENDEES),
                    'Lead(s)': parseLinkedField(activityObj.LINKEDLEADS),
                    'Opportunity(ies)': parseLinkedField(activityObj.LINKEDOPPORTUNITIES),
                    'Project(s)': parseLinkedField(activityObj.LINKEDPROJECTS),
                    'Contact(s)': parseLinkedField(activityObj.LINKEDCONTACTS),
                    'Call Disposition': activityObj.CALLDISPOSITION || ''
                };
            });
            
            // Store unfiltered activities for analysis
            const unfilteredActivities = [...parsedActivities];
            
            // Apply filters for Activities tab only
            const beforeFilterCount = parsedActivities.length;
            
            parsedActivities = parsedActivities.filter(activity => {
                // Date filter
                if (startDate && activity['Start Date'] && activity['Start Date'] < startDate) return false;
                if (endDate && activity['Start Date'] && activity['Start Date'] > endDate) return false;
                
                // Owner filter
                if (filters.owner && !activity['Owner(s)'].toLowerCase().includes(filters.owner)) return false;
                
                // Company filter
                if (filters.company && !activity['Company(ies)'].toLowerCase().includes(filters.company)) return false;
                
                // Opportunity filter
                if (filters.opportunity && !activity['Opportunity(ies)'].toLowerCase().includes(filters.opportunity)) return false;
                
                return true;
            });
            
            const filteredCount = beforeFilterCount - parsedActivities.length;
            if (filteredCount > 0) {
                updateStatus(`🔍 Filtered ${filteredCount} activities for display`, 65);
            }
            updateStatus(`📊 ${parsedActivities.length} activities in filtered view`, 68);
            
            if (parsedActivities.length === 0) {
                updateStatus('⚠️ No activities match your filters. Try adjusting your criteria.', 100);
                progressFill.style.background = '#ff9800';
                return;
            }
            
            updateStatus('📋 Sorting by Owner...', 70);
            
            // Sort by Owner
            parsedActivities.sort((a, b) => {
                const ownerA = (a['Owner(s)'] || '').toLowerCase();
                const ownerB = (b['Owner(s)'] || '').toLowerCase();
                return ownerA.localeCompare(ownerB);
            });
            
            // Create Excel file
            updateStatus('📝 Creating Excel file...', 75);
            
            // Load xlsx-js-style (same as ExecIQ)
            if (typeof XLSX === 'undefined') {
                updateStatus('📦 Loading Excel library with styling support...', 78);
                try {
                    await new Promise((resolve, reject) => {
                        const script = document.createElement('script');
                        script.src = SHEETJS;
                        script.onload = () => {
                            let retries = 0;
                            const checkInterval = setInterval(() => {
                                if (window.XLSX && window.XLSX.utils && window.XLSX.utils.aoa_to_sheet) {
                                    clearInterval(checkInterval);
                                    resolve();
                                } else if (++retries > 50) {
                                    clearInterval(checkInterval);
                                    reject(new Error('XLSX not available'));
                                }
                            }, 100);
                        };
                        script.onerror = reject;
                        document.head.appendChild(script);
                    });
                } catch (e) {
                    throw new Error('Unable to load Excel library: ' + e.message);
                }
            }
            
            // NOW build the analysis AFTER XLSX is loaded
            updateStatus('📊 Building relationship intelligence analysis...', 82);
            const analysisSheet = buildActivitiesAnalysis(unfilteredActivities);
            
            updateStatus('✨ Formatting Activities sheet...', 88);
            
            // Create Activities worksheet
            const ws = XLSX.utils.json_to_sheet(parsedActivities);
            
            // Get range
            const range = XLSX.utils.decode_range(ws['!ref']);
            
            // Excel blue color (matching ExecIQ)
            const EXCEL_BLUE = '4472C4';
            
            // Format all cells - Arial 10pt
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cellRef]) continue;
                    
                    // Initialize cell style
                    ws[cellRef].s = {
                        font: {
                            name: 'Arial',
                            sz: 10,
                            bold: R === 0,
                            color: R === 0 ? { rgb: 'FFFFFF' } : { rgb: '000000' }
                        },
                        alignment: {
                            vertical: 'center',
                            horizontal: 'left'
                        }
                    };
                    
                    // Header row (row 0) - Blue background
                    if (R === 0) {
                        ws[cellRef].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: EXCEL_BLUE }
                        };
                    }
                }
            }
            
            // Format date columns
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
                ['D', 'E'].forEach(col => {
                    const cellRef = col + (row + 1);
                    if (ws[cellRef] && ws[cellRef].v instanceof Date) {
                        ws[cellRef].t = 'd';
                        ws[cellRef].z = 'mm/dd/yyyy';
                    }
                });
            }
            
            // Auto-size columns
            const colWidths = Object.keys(parsedActivities[0]).map(key => {
                const maxLength = Math.max(
                    key.length,
                    ...parsedActivities.map(row => {
                        const val = row[key];
                        if (val instanceof Date) return 12;
                        return (val || '').toString().length;
                    })
                );
                return { wch: Math.min(maxLength + 2, 50) };
            });
            ws['!cols'] = colWidths;
            
            // Add autofilter
            ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
            
            // Set row height for header
            if (!ws['!rows']) ws['!rows'] = [];
            ws['!rows'][0] = { hpt: 20 };
            
            // Create workbook
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Activities');
            XLSX.utils.book_append_sheet(wb, analysisSheet, 'Activities Analysis');
            
            updateStatus('💾 Downloading file...', 95);
            
            // Generate filename
            const timestamp = new Date().toISOString().split('T')[0];
            const filename = `Unanet_Activities_Export_${timestamp}.xlsx`;
            
            // Download file
            XLSX.writeFile(wb, filename);
            
            updateStatus(`✅ Export complete! Downloaded: ${filename}`, 100);
            updateStatus(`📊 Total activities exported: ${parsedActivities.length}`, 100);
            updateStatus(`📈 Analysis generated from ${unfilteredActivities.length} activities`, 100);
            updateStatus('🎉 This window will close in 3 seconds...', 100);
            
            closeDialog();
            
        } catch (error) {
            updateStatus(`❌ Error: ${error.message}`, 100);
            progressFill.style.background = '#f44336';
            console.error('Export error:', error);
        }
    }
    
})();
