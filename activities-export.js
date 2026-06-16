// Unanet Activities Export Tool
// Fetches all activities from Unanet CRM and exports to Excel

(async function() {
    // SHEETJS with styling support - same library as ExecIQ
    var SHEETJS = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";

    // ═══════════════════════════════════════════════════════════
    // OPPORTUNITY STAGE CONFIGURATION
    // ═══════════════════════════════════════════════════════════
    //
    // Primary closed detection: ACTIVEIND === 2 from oppActions.cfm
    //
    // CLOSED_STAGE_EXCEPTIONS: stages that are correctly configured as
    // Closed/Won or Closed/Lost in Unanet admin but whose ACTIVEIND is
    // incorrectly set to 1 (open) due to a data quality issue. Add to
    // this list if additional misconfigured stages are identified.
    const CLOSED_STAGE_EXCEPTIONS = new Set([
        'In Construction',
        'In Preconstruction'
    ]);

    const OPP_ACTIONS_URL = '/contact/opportunity/oppActions.cfm';
    const OPP_BODY_BASE   =
        'sort=VCHLEADNUMBER&dir=ASC&action=getOpportunityGridData&json=1' +
        '&selectedCurrency=USD&view=0&SalesCycle=NaN' +
        '&officeId=0&divisionId=0&studioId=0&practiceAreaId=0&territoryId=0' +
        '&stageId=0&priCatId=0&secCatId=0&masterSub=0&staffRoleId=0' +
        '&dateCreated=0&dateModified=0&dateCreatedModified=0&filteredSearch=0&search=' +
        '&visibleColumns=VCHLEADNUMBER&visibleColumns=STAGEID&visibleColumns=VCHPROJECTNAME';

    // Fetches ALL opportunities in parallel and returns a Set of ILEADIDs
    // considered closed via:
    //   1. ACTIVEIND === 2  (primary signal)
    //   2. STAGENAME in CLOSED_STAGE_EXCEPTIONS (safety net for misconfigured stages)
    async function fetchClosedLeadIds(onProgress) {
        const closedLeadIds  = new Set();
        const OPP_CONCURRENCY = 5;
        const OPP_BATCH_SIZE  = 100;
        const OPP_RETRY_DELAY = 500;

        async function fetchOppBatch(start) {
            const body = `start=${start}&limit=${OPP_BATCH_SIZE}&ActiveInd=0&${OPP_BODY_BASE}`;
            const resp = await fetch(OPP_ACTIONS_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body
            });
            if (!resp.ok) throw new Error(`oppActions.cfm returned HTTP ${resp.status}`);
            const text = await resp.text();
            const json = text.replace(/^[\s\S]*?(?={)/, '');
            return JSON.parse(json);
        }

        async function fetchOppBatchWithRetry(start) {
            try { return await fetchOppBatch(start); }
            catch(err) {
                console.warn(`Opp batch at offset ${start} failed, retrying...`, err);
                await new Promise(r => setTimeout(r, OPP_RETRY_DELAY));
                return await fetchOppBatch(start);
            }
        }

        function processOppBatch(data) {
            data.DATA.forEach(opp => {
                if (!opp.ILEADID) return;
                const isClosed = opp.ACTIVEIND === 2 ||
                                 CLOSED_STAGE_EXCEPTIONS.has(opp.STAGENAME || '');
                if (isClosed) closedLeadIds.add(Number(opp.ILEADID));
            });
        }

        // First batch to get total record count
        const firstBatch = await fetchOppBatch(0);
        const totalRecords = firstBatch.ROWCOUNT;
        processOppBatch(firstBatch);
        if (onProgress) onProgress(`📋 Found ${totalRecords} opportunities to scan...`, null);

        // Build offset queue for remaining batches
        const offsets = [];
        for (let s = OPP_BATCH_SIZE; s < totalRecords; s += OPP_BATCH_SIZE) offsets.push(s);

        let completedBatches = 0;
        let nextIndex = 0;

        async function oppWorker() {
            while (true) {
                const myIndex = nextIndex++;
                if (myIndex >= offsets.length) return;
                const offset = offsets[myIndex];
                const batch  = await fetchOppBatchWithRetry(offset);
                processOppBatch(batch);
                completedBatches++;
                if (onProgress) onProgress(
                    `📋 Scanned ${Math.min((completedBatches + 1) * OPP_BATCH_SIZE, totalRecords)} / ${totalRecords} opportunities...`,
                    null
                );
            }
        }

        const workerCount = Math.min(OPP_CONCURRENCY, offsets.length);
        if (workerCount > 0) {
            const workerPromises = [];
            for (let i = 0; i < workerCount; i++) workerPromises.push(oppWorker());
            await Promise.all(workerPromises);
        }

        if (onProgress) onProgress(`✅ Scanned all ${totalRecords} opportunities`, null);
        return closedLeadIds;
    }

    // ═══════════════════════════════════════════════════════════
    // FILTER DIALOG
    // ═══════════════════════════════════════════════════════════

    const filterOverlay = document.createElement('div');
    filterOverlay.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100%; height: 100%;
        background: rgba(0,0,0,0.7); z-index: 999999;
        display: flex; justify-content: center; align-items: center;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    `;

    const filterDialog = document.createElement('div');
    filterDialog.style.cssText = `
        background: white; padding: 30px; border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        min-width: 520px; max-width: 620px; position: relative;
    `;

    const filterTitle = document.createElement('h2');
    filterTitle.textContent = 'Unanet Activities Export - Filters';
    filterTitle.style.cssText = 'margin: 0 0 20px 0; color: #333; font-size: 20px;';

    const filterForm = document.createElement('div');
    filterForm.style.cssText = 'margin: 20px 0;';

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

    // ── Date range ──────────────────────────────────────────────
    const dateRangeSelect = document.createElement('select');
    dateRangeSelect.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';
    dateRangeSelect.innerHTML = `
        <option value="all">All dates</option>
        <option value="last7">Last 7 days</option>
        <option value="last14">Last 14 days</option>
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
        customDateDiv.style.display = dateRangeSelect.value === 'custom' ? 'block' : 'none';
    });

    // ── Text filters ─────────────────────────────────────────────
    const ownerInput = document.createElement('input');
    ownerInput.type = 'text';
    ownerInput.placeholder = 'e.g., John Smith (leave blank for all)';
    ownerInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';

    const companyInput = document.createElement('input');
    companyInput.type = 'text';
    companyInput.placeholder = 'e.g., Acme Corp (leave blank for all)';
    companyInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';

    const opportunityInput = document.createElement('input');
    opportunityInput.type = 'text';
    opportunityInput.placeholder = 'e.g., (25-0044) Security Assessment (leave blank for all)';
    opportunityInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px;';

    // ── Exclude closed opportunities checkbox ────────────────────
    const excludeClosedWrapper = document.createElement('div');
    excludeClosedWrapper.style.cssText = 'margin-bottom: 15px; padding: 12px; background: #f9f9f9; border: 1px solid #e0e0e0; border-radius: 4px;';

    const excludeClosedLabel = document.createElement('label');
    excludeClosedLabel.style.cssText = 'display: flex; align-items: center; gap: 10px; cursor: pointer; font-size: 14px; font-weight: 500; color: #555;';

    const excludeClosedCb = document.createElement('input');
    excludeClosedCb.type = 'checkbox';
    excludeClosedCb.style.cssText = 'width: 16px; height: 16px; cursor: pointer; flex-shrink: 0;';

    const excludeClosedText = document.createElement('span');
    excludeClosedText.innerHTML = 'Exclude Closed Opportunities <span style="font-weight:400;color:#888;font-size:12px;">(fetched at export time)</span>';

    excludeClosedLabel.appendChild(excludeClosedCb);
    excludeClosedLabel.appendChild(excludeClosedText);

    const excludeClosedNote = document.createElement('div');
    excludeClosedNote.style.cssText = 'margin-top: 8px; margin-left: 26px; font-size: 12px; color: #888; line-height: 1.5;';
    excludeClosedNote.textContent = 'Activities linked only to closed opportunities will be excluded from both the Activities and Analysis tabs. Activities with no linked opportunity are always kept.';

    excludeClosedWrapper.appendChild(excludeClosedLabel);
    excludeClosedWrapper.appendChild(excludeClosedNote);

    // ── Build form ───────────────────────────────────────────────
    filterForm.appendChild(createField('Date Range:', dateRangeSelect));
    filterForm.appendChild(customDateDiv);
    filterForm.appendChild(createField('Owner (partial match):', ownerInput));
    filterForm.appendChild(createField('Company (partial match):', companyInput));
    filterForm.appendChild(createField('Opportunity (partial match):', opportunityInput));
    filterForm.appendChild(excludeClosedWrapper);

    // ── Buttons ──────────────────────────────────────────────────
    const buttonDiv = document.createElement('div');
    buttonDiv.style.cssText = 'margin-top: 25px; display: flex; gap: 10px; justify-content: flex-end;';

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText = 'padding: 10px 20px; background: #f0f0f0; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;';
    cancelBtn.onmouseover = () => cancelBtn.style.background = '#e0e0e0';
    cancelBtn.onmouseout  = () => cancelBtn.style.background = '#f0f0f0';

    const exportBtn = document.createElement('button');
    exportBtn.textContent = 'Export Activities';
    exportBtn.style.cssText = 'padding: 10px 20px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;';
    exportBtn.onmouseover = () => exportBtn.style.background = '#45a049';
    exportBtn.onmouseout  = () => exportBtn.style.background = '#4CAF50';

    buttonDiv.appendChild(cancelBtn);
    buttonDiv.appendChild(exportBtn);

    filterDialog.appendChild(filterTitle);
    filterDialog.appendChild(filterForm);
    filterDialog.appendChild(buttonDiv);
    filterOverlay.appendChild(filterDialog);
    document.body.appendChild(filterOverlay);

    cancelBtn.onclick = () => document.body.removeChild(filterOverlay);

    exportBtn.onclick = async () => {
        const filters = {
            dateRange:       dateRangeSelect.value,
            customStartDate: startDateInput.value,
            customEndDate:   endDateInput.value,
            owner:           ownerInput.value.trim().toLowerCase(),
            company:         companyInput.value.trim().toLowerCase(),
            opportunity:     opportunityInput.value.trim().toLowerCase(),
            excludeClosed:   excludeClosedCb.checked
        };

        // Calculate date range
        let startDate = null;
        let endDate   = null;
        const now = new Date();

        switch (filters.dateRange) {
            case 'last7':        startDate = new Date(now - 7   * 86400000); endDate = now; break;
            case 'last14':       startDate = new Date(now - 14  * 86400000); endDate = now; break;
            case 'last30':       startDate = new Date(now - 30  * 86400000); endDate = now; break;
            case 'last60':       startDate = new Date(now - 60  * 86400000); endDate = now; break;
            case 'last90':       startDate = new Date(now - 90  * 86400000); endDate = now; break;
            case 'last6months':  startDate = new Date(now - 180 * 86400000); endDate = now; break;
            case 'last12months': startDate = new Date(now - 365 * 86400000); endDate = now; break;
            case 'currentYear':
                startDate = new Date(now.getFullYear(), 0, 1);
                endDate   = new Date(now.getFullYear(), 11, 31);
                break;
            case 'next90':
                startDate = now;
                endDate   = new Date(now - -90 * 86400000);
                break;
            case 'last90toNext90':
                startDate = new Date(now - 90  * 86400000);
                endDate   = new Date(now - -90 * 86400000);
                break;
            case 'custom':
                if (filters.customStartDate) startDate = new Date(filters.customStartDate);
                if (filters.customEndDate)   endDate   = new Date(filters.customEndDate);
                break;
            default: break;
        }

        document.body.removeChild(filterOverlay);
        await runExport(filters, startDate, endDate);
    };

    // ═══════════════════════════════════════════════════════════
    // ACTIVITIES ANALYSIS BUILDER
    // ═══════════════════════════════════════════════════════════

    function buildActivitiesAnalysis(activities) {
        const now       = new Date();
        const day30     = new Date(now - 30  * 86400000);
        const day60     = new Date(now - 60  * 86400000);
        const day59     = new Date(now - 59  * 86400000);
        const day90     = new Date(now - 90  * 86400000);
        const day3years = new Date(now - 3 * 365 * 86400000);

        function splitField(val) {
            if (!val) return [];
            return val.split(';').map(s => s.trim()).filter(s => s);
        }

        // ── 1. Client analysis ───────────────────────────────────
        const clientMapRaw = {};

        activities.forEach(act => {
            const companies = splitField(act['Company(ies)']);
            const owners    = splitField(act['Owner(s)']);
            const contacts  = splitField(act['Contact(s)']);
            const opps      = splitField(act['Opportunity(ies)']);
            const callType  = act['Call Type'] || '';
            const startDate = act['Start Date'];

            companies.forEach(company => {
                if (!company) return;
                if (!clientMapRaw[company]) {
                    clientMapRaw[company] = {
                        total: 0, last30: 0, last60: 0, last90: 0,
                        owners: new Set(), contacts: new Set(), opps: new Set(),
                        meetings: 0, meetings90: 0, callTypes: {},
                        lastActivity: null, lastOwner: '', lastCallType: '',
                        activities: []
                    };
                }
                const client = clientMapRaw[company];
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
                    client.lastOwner    = owners[0] || '';
                    client.lastCallType = callType;
                }
            });
        });

        // Suppress clients not contacted in 3+ years
        const clientMap = {};
        Object.entries(clientMapRaw).forEach(([company, data]) => {
            if (data.lastActivity && data.lastActivity >= day3years) clientMap[company] = data;
        });

        // ── 2. Owner analysis ────────────────────────────────────
        const ownerMapRaw = {};

        activities.forEach(act => {
            const owners    = splitField(act['Owner(s)']);
            const companies = splitField(act['Company(ies)']);
            const contacts  = splitField(act['Contact(s)']);
            const callType  = act['Call Type'] || '';
            const startDate = act['Start Date'];

            owners.forEach(owner => {
                if (!owner) return;
                if (!ownerMapRaw[owner]) {
                    ownerMapRaw[owner] = {
                        total: 0, companies: new Set(), contacts: new Set(),
                        callTypes: {}, lastActivity: null
                    };
                }
                const o = ownerMapRaw[owner];
                o.total++;
                companies.forEach(c => o.companies.add(c));
                contacts.forEach(c => o.contacts.add(c));
                o.callTypes[callType] = (o.callTypes[callType] || 0) + 1;
                if (!o.lastActivity || startDate > o.lastActivity) o.lastActivity = startDate;
            });
        });

        const ownerMap = {};
        Object.entries(ownerMapRaw).forEach(([owner, data]) => {
            if (data.lastActivity && data.lastActivity >= day3years) ownerMap[owner] = data;
        });

        // ── 3. Opportunity analysis ──────────────────────────────
        const oppMap = {};

        activities.forEach(act => {
            const opps      = splitField(act['Opportunity(ies)']);
            const contacts  = splitField(act['Contact(s)']);
            const owners    = splitField(act['Owner(s)']);
            const startDate = act['Start Date'];

            opps.forEach(opp => {
                if (!opp) return;
                if (!oppMap[opp]) {
                    oppMap[opp] = { total: 0, last30: 0, contacts: new Set(), owners: new Set(), lastActivity: null };
                }
                const o = oppMap[opp];
                o.total++;
                if (startDate >= day30) o.last30++;
                contacts.forEach(c => o.contacts.add(c));
                owners.forEach(ow => o.owners.add(ow));
                if (!o.lastActivity || startDate > o.lastActivity) o.lastActivity = startDate;
            });
        });

        // ── 4. Contact analysis ──────────────────────────────────
        const contactMap = {};

        activities.forEach(act => {
            const contacts  = splitField(act['Contact(s)']);
            const owners    = splitField(act['Owner(s)']);
            const opps      = splitField(act['Opportunity(ies)']);
            const callType  = act['Call Type'] || '';
            const startDate = act['Start Date'];

            contacts.forEach(contact => {
                if (!contact) return;
                if (!contactMap[contact]) {
                    contactMap[contact] = { total: 0, owners: new Set(), opps: new Set(), callTypes: {}, lastActivity: null };
                }
                const c = contactMap[contact];
                c.total++;
                owners.forEach(o => c.owners.add(o));
                opps.forEach(op => c.opps.add(op));
                c.callTypes[callType] = (c.callTypes[callType] || 0) + 1;
                if (!c.lastActivity || startDate > c.lastActivity) c.lastActivity = startDate;
            });
        });

        // ── 5. Relationship health scores ────────────────────────
        const healthScores = [];

        Object.entries(clientMap).forEach(([company, data]) => {
            const daysSince = data.lastActivity
                ? Math.floor((now - data.lastActivity) / 86400000)
                : 999;

            // Recency (25 pts)
            let recencyScore = 0;
            if      (daysSince <= 7)  recencyScore = 25;
            else if (daysSince <= 14) recencyScore = 22;
            else if (daysSince <= 30) recencyScore = 18;
            else if (daysSince <= 45) recencyScore = 12;
            else if (daysSince <= 60) recencyScore = 8;
            else if (daysSince <= 90) recencyScore = 4;

            // Frequency (20 pts)
            let freqScore = 0;
            if      (data.last90 >= 20) freqScore = 20;
            else if (data.last90 >= 15) freqScore = 17;
            else if (data.last90 >= 10) freqScore = 13;
            else if (data.last90 >= 6)  freqScore = 9;
            else if (data.last90 >= 3)  freqScore = 5;
            else if (data.last90 >= 1)  freqScore = 2;

            const contactCount = data.contacts.size;

            // Owner coverage (10 pts)
            const ownerCount = data.owners.size;
            let ownerScore = 0;
            if      (ownerCount >= 5) ownerScore = 10;
            else if (ownerCount >= 4) ownerScore = 8;
            else if (ownerCount >= 3) ownerScore = 6;
            else if (ownerCount >= 2) ownerScore = 3;
            else if (ownerCount >= 1) ownerScore = 1;

            // Opportunity engagement (10 pts)
            const hasOpps = data.opps.size > 0;
            const hasRecentOppActivity = data.activities.some(a =>
                a['Opportunity(ies)'] && a['Start Date'] >= day59
            );
            let oppScore = 0;
            if      (hasOpps && data.opps.size > 1 && hasRecentOppActivity) oppScore = 10;
            else if (hasOpps && hasRecentOppActivity)                        oppScore = 7;
            else if (hasOpps && !hasRecentOppActivity)                       oppScore = 3;

            const totalScore = recencyScore + freqScore + ownerScore + oppScore;

            let status = 'Critical';
            if      (totalScore >= 55) status = 'Healthy';
            else if (totalScore >= 45) status = 'Stable';
            else if (totalScore >= 30) status = 'Watch';
            else if (totalScore >= 15) status = 'Elevated Risk';

            const flags = [];
            if      (daysSince >= 90)                     flags.push('No Activity 90+ Days');
            else if (daysSince >= 30 && daysSince <= 60)  flags.push('Slowing Down');
            if      (ownerCount === 0)                    flags.push('No Owner');
            else if (ownerCount === 1)                    flags.push('Single Owner');
            if      (hasOpps && !hasRecentOppActivity && daysSince >= 59) flags.push('Stale Pursuit');
            if      (daysSince <= 30 && data.last90 <= 2) flags.push('Low Frequency');

            healthScores.push({
                company, score: totalScore, status, daysSince,
                owners: ownerCount, contacts: contactCount,
                meetings90: data.meetings90, activities90: data.last90,
                flags: flags.join('; ') || 'None'
            });
        });

        // ── 6. Build worksheet ───────────────────────────────────
        const aoa = [];
        const BLUE = '4472C4', WHITE = 'FFFFFF';
        const GREEN = 'C6E0B4', YELLOW = 'FFE699', ORANGE = 'F4B084', RED = 'F8CBAD';

        function header(text) { aoa.push([text]); }
        function spacer()     { aoa.push([]);     }

        header('MOST ACTIVE CLIENTS');
        aoa.push(['Company','Total Activities','Last 30d','Last 60d','Last 90d','Owners','Contacts','Meetings (90d)','Last Activity']);
        Object.entries(clientMap).sort((a,b) => b[1].total - a[1].total).slice(0,25).forEach(([company, data]) => {
            aoa.push([company, data.total, data.last30, data.last60, data.last90,
                data.owners.size, data.contacts.size, data.meetings90, data.lastActivity]);
        });

        spacer(); spacer();

        header('CLIENTS NOT CONTACTED RECENTLY');
        aoa.push(['Company','Days Since Last Contact','Last Activity','Last Owner','Last Type','Open Opps','Risk Level']);
        const dormant = Object.entries(clientMap)
            .filter(([_, d]) => Math.floor((now - d.lastActivity) / 86400000) >= 30)
            .sort((a,b) => (now - b[1].lastActivity) - (now - a[1].lastActivity))
            .slice(0, 50);
        dormant.forEach(([company, data]) => {
            const days = Math.floor((now - data.lastActivity) / 86400000);
            let risk = 'Watch';
            if      (days >= 180) risk = 'Dormant';
            else if (days >= 90)  risk = 'Gap';
            else if (days >= 60)  risk = 'At Risk';
            aoa.push([company, days, data.lastActivity, data.lastOwner, data.lastCallType, data.opps.size, risk]);
        });

        spacer(); spacer();

        header('ACTIVITY BY OWNER');
        aoa.push(['Owner','Total Activities','Companies','Contacts','Avg per Month','Last Activity']);
        Object.entries(ownerMap).sort((a,b) => b[1].total - a[1].total).forEach(([owner, data]) => {
            aoa.push([owner, data.total, data.companies.size, data.contacts.size,
                Math.round(data.total / 3), data.lastActivity]);
        });

        spacer(); spacer();

        header('OPPORTUNITY ENGAGEMENT');
        aoa.push(['Opportunity','Total Activities','Last 30d','Contacts','Owners','Last Touch']);
        Object.entries(oppMap).sort((a,b) => b[1].total - a[1].total).slice(0,50).forEach(([opp, data]) => {
            aoa.push([opp, data.total, data.last30, data.contacts.size, data.owners.size, data.lastActivity]);
        });

        spacer(); spacer();

        header('CONTACT ENGAGEMENT');
        aoa.push(['Contact','Interactions','Last Contact','Related Opps','Internal Owners']);
        Object.entries(contactMap).sort((a,b) => b[1].total - a[1].total).slice(0,50).forEach(([contact, data]) => {
            aoa.push([contact, data.total, data.lastActivity, data.opps.size, data.owners.size]);
        });

        spacer(); spacer();

        header('RELATIONSHIP COVERAGE ANALYSIS');
        aoa.push(['Company','Internal Owners','Client Contacts','Coverage Status']);
        Object.entries(clientMap)
            .sort((a,b) => (a[1].owners.size + a[1].contacts.size) - (b[1].owners.size + b[1].contacts.size))
            .slice(0,50)
            .forEach(([company, data]) => {
                let status = 'Strong';
                if      (data.owners.size === 1 && data.contacts.size === 1) status = 'Critical - Single Threaded';
                else if (data.owners.size === 1)                             status = 'Risk - Single Owner';
                else if (data.contacts.size === 1)                           status = 'Risk - Single Contact';
                else if (data.owners.size === 2 && data.contacts.size === 2) status = 'Moderate';
                aoa.push([company, data.owners.size, data.contacts.size, status]);
            });

        spacer(); spacer();

        header('RELATIONSHIP HEALTH SCORES');
        aoa.push(['Company','Score','Status','Days Since Contact','Owners','Contacts','Meetings (90d)','Activities (90d)','Risk Flags']);
        healthScores.sort((a,b) => b.score - a.score).forEach(h => {
            aoa.push([h.company, h.score, h.status, h.daysSince, h.owners, h.contacts,
                h.meetings90, h.activities90, h.flags]);
        });

        // ── 7. Format worksheet ──────────────────────────────────
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wsRange = XLSX.utils.decode_range(ws['!ref']);

        const sectionHeaders = [];
        const dataHeaders    = [];

        for (let R = 0; R <= wsRange.e.r; R++) {
            const cell = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
            if (cell && typeof cell.v === 'string' &&
                /CLIENTS|ACTIVITY|OPPORTUNITY|CONTACT|RELATIONSHIP|SCORES/.test(cell.v)) {
                sectionHeaders.push(R);
                if (R + 1 <= wsRange.e.r) dataHeaders.push(R + 1);
            }
        }

        for (let R = 0; R <= wsRange.e.r; R++) {
            for (let C = 0; C <= wsRange.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[addr]) continue;

                const cellValue       = ws[addr].v;
                const isSectionHeader = sectionHeaders.includes(R) && C === 0;
                const isDataHeader    = dataHeaders.includes(R);

                ws[addr].s = {
                    font: { name: 'Arial', sz: 10, bold: false, color: { rgb: '000000' } },
                    alignment: { vertical: 'center', horizontal: 'left' }
                };

                if (isSectionHeader) {
                    ws[addr].s.font.bold  = true;
                    ws[addr].s.font.sz    = 12;
                    ws[addr].s.font.color = { rgb: BLUE };
                } else if (isDataHeader) {
                    ws[addr].s.font.bold  = true;
                    ws[addr].s.font.color = { rgb: WHITE };
                    ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: BLUE } };
                }

                if (ws[addr].v instanceof Date) { ws[addr].t = 'd'; ws[addr].z = 'mm/dd/yyyy'; }

                if (typeof cellValue === 'string') {
                    if      (cellValue === 'Healthy')       { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: GREEN  } }; ws[addr].s.font.bold = true; }
                    else if (cellValue === 'Stable')        { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: YELLOW } }; }
                    else if (cellValue === 'Watch')         { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: ORANGE } }; }
                    else if (cellValue === 'Elevated Risk' ||
                             cellValue === 'Critical')      { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: RED    } }; ws[addr].s.font.bold = true; }

                    if      (cellValue === 'Dormant' ||
                             cellValue === 'Gap')           { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: RED    } }; ws[addr].s.font.bold = true; }
                    else if (cellValue === 'At Risk')       { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: ORANGE } }; }

                    if      (cellValue.includes('Critical')) { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: RED    } }; ws[addr].s.font.bold = true; }
                    else if (cellValue.includes('Risk'))     { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: ORANGE } }; }
                    else if (cellValue === 'Moderate')       { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: YELLOW } }; }
                    else if (cellValue === 'Strong')         { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: GREEN  } }; }
                }

                if (typeof cellValue === 'number' && C === 1) {
                    const rowAbove = ws[XLSX.utils.encode_cell({ r: R - 1, c: C })]?.v;
                    if (rowAbove === 'Score') {
                        ws[addr].z = '0';
                        if      (cellValue >= 85) { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: GREEN  } }; ws[addr].s.font.bold = true; }
                        else if (cellValue >= 70) { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: YELLOW } }; }
                        else if (cellValue >= 55) { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: ORANGE } }; }
                        else                      { ws[addr].s.fill = { patternType: 'solid', fgColor: { rgb: RED    } }; ws[addr].s.font.bold = true; }
                    }
                }
            }
        }

        const colWidths = [];
        for (let C = 0; C <= wsRange.e.c; C++) {
            let maxWidth = 10;
            for (let R = 0; R <= wsRange.e.r; R++) {
                const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
                if (cell && cell.v) maxWidth = Math.max(maxWidth, cell.v.toString().length);
            }
            colWidths.push({ wch: Math.min(maxWidth + 2, 50) });
        }
        ws['!cols'] = colWidths;

        if (!ws['!rows']) ws['!rows'] = [];
        sectionHeaders.forEach(R => { ws['!rows'][R] = { hpt: 25 }; });

        return ws;
    }

    // ═══════════════════════════════════════════════════════════
    // MAIN EXPORT FUNCTION
    // ═══════════════════════════════════════════════════════════

    async function runExport(filters, startDate, endDate) {
        const overlay = document.createElement('div');
        overlay.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.7); z-index: 999999;
            display: flex; justify-content: center; align-items: center;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        `;

        const dialog = document.createElement('div');
        dialog.style.cssText = `
            background: white; padding: 30px; border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            min-width: 400px; max-width: 500px; position: relative;
        `;

        const closeBtn = document.createElement('button');
        closeBtn.innerHTML = '×';
        closeBtn.style.cssText = `
            position: absolute; top: 10px; right: 10px; background: none; border: none;
            font-size: 28px; color: #999; cursor: pointer; padding: 0;
            width: 30px; height: 30px; line-height: 30px; text-align: center; border-radius: 4px;
        `;
        closeBtn.onmouseover = () => { closeBtn.style.background = '#f0f0f0'; closeBtn.style.color = '#333'; };
        closeBtn.onmouseout  = () => { closeBtn.style.background = 'none';    closeBtn.style.color = '#999'; };
        closeBtn.onclick = () => { if (document.body.contains(overlay)) document.body.removeChild(overlay); };

        const title = document.createElement('h2');
        title.textContent = 'Unanet Activities Export';
        title.style.cssText = 'margin: 0 0 20px 0; color: #333; font-size: 20px; padding-right: 30px;';

        const statusDiv = document.createElement('div');
        statusDiv.style.cssText = 'margin: 15px 0; color: #666; font-size: 14px; line-height: 1.6; max-height: 200px; overflow-y: auto;';

        const progressBar = document.createElement('div');
        progressBar.style.cssText = 'width: 100%; height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; margin: 15px 0;';

        const progressFill = document.createElement('div');
        progressFill.style.cssText = 'width: 0%; height: 100%; background: linear-gradient(90deg,#4CAF50,#45a049); border-radius: 4px; transition: width 0.3s ease;';
        progressBar.appendChild(progressFill);

        dialog.appendChild(closeBtn);
        dialog.appendChild(title);
        dialog.appendChild(statusDiv);
        dialog.appendChild(progressBar);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        function updateStatus(message, progress = null) {
            statusDiv.innerHTML += `<div>${message}</div>`;
            statusDiv.scrollTop = statusDiv.scrollHeight;
            if (progress !== null) progressFill.style.width = `${progress}%`;
        }

        function closeDialog() {
            setTimeout(() => { if (document.body.contains(overlay)) document.body.removeChild(overlay); }, 3000);
        }

        const API_URL     = 'https://services.cosential.com/com/model/activities/callLog.cfc';
        const BATCH_SIZE  = 100;
        const CONCURRENCY = 5;
        const RETRY_DELAY = 500;

        function stripHTML(html) {
            if (!html) return '';
            const temp = document.createElement('div');
            temp.innerHTML = html;
            return temp.textContent.trim();
        }

        function extractNames(html, selector) {
            if (!html) return '';
            const temp = document.createElement('div');
            temp.innerHTML = html;
            return Array.from(temp.querySelectorAll(selector))
                .map(el => el.textContent.trim()).filter(n => n).join('; ');
        }

        function parseLinkedField(html) {
            if (!html) return '';
            let r = extractNames(html, '.company a');           if (r) return r;
            r = extractNames(html, '.contact a');               if (r) return r;
            r = extractNames(html, '.opportunity a, .lead a');  if (r) return r;
            r = extractNames(html, '.personnel');               if (r) return r;
            return stripHTML(html);
        }

        function extractLeadIds(html) {
            if (!html) return [];
            return [...html.matchAll(/LeadID=(\d+)/gi)].map(m => Number(m[1]));
        }

        function decodeStatus(statusCode) {
            if (statusCode === 0) return 'Completed';
            if (statusCode === 1) return 'Not Started';
            if (statusCode === null || statusCode === undefined || statusCode === '') return '';
            return `Status ${statusCode}`;
        }

        function parseDate(dateStr) {
            if (!dateStr) return null;
            try { const d = new Date(dateStr); return isNaN(d.getTime()) ? null : d; }
            catch(e) { return null; }
        }

        async function fetchActivities(start = 0) {
            const resp = await fetch(API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    start, limit: BATCH_SIZE,
                    sort: 'startDateDateOnly', dir: 'ASC',
                    groupBy: 'status', groupDir: 'ASC',
                    method: 'getGridData', xaction: 'read'
                })
            });
            if (!resp.ok) throw new Error(`HTTP error! status: ${resp.status}`);
            return await resp.json();
        }

        async function fetchActivitiesWithRetry(start) {
            try { return await fetchActivities(start); }
            catch(err) {
                console.warn(`Batch at offset ${start} failed, retrying...`, err);
                await new Promise(r => setTimeout(r, RETRY_DELAY));
                return await fetchActivities(start);
            }
        }

        try {
            updateStatus('🚀 Starting export...', 5);

            // Show active filters
            const hasFilters = startDate || endDate || filters.owner ||
                               filters.company || filters.opportunity || filters.excludeClosed;
            if (hasFilters) {
                updateStatus('🔍 Active filters:');
                if (startDate && endDate) updateStatus(`   📅 Date: ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
                else if (startDate)       updateStatus(`   📅 Date: After ${startDate.toLocaleDateString()}`);
                else if (endDate)         updateStatus(`   📅 Date: Before ${endDate.toLocaleDateString()}`);
                if (filters.owner)        updateStatus(`   👤 Owner: ${filters.owner}`);
                if (filters.company)      updateStatus(`   🏢 Company: ${filters.company}`);
                if (filters.opportunity)  updateStatus(`   💼 Opportunity: ${filters.opportunity}`);
                if (filters.excludeClosed) updateStatus('   🚫 Excluding closed opportunities');
            }

            // ── Kick off opp fetch in parallel with activity fetch ──
            // Both fetches run simultaneously; we await the opp result
            // only after all activity batches are done, by which point
            // it should already be resolved.
            let closedLeadIdsPromise = null;
            if (filters.excludeClosed) {
                updateStatus('🔗 Fetching closed opportunities in background...', 8);
                closedLeadIdsPromise = fetchClosedLeadIds((msg) => updateStatus(msg));
            }

            // ── Fetch all activities ──────────────────────────────
            updateStatus('📥 Fetching activities from Unanet...', 10);

            const firstBatch = await fetchActivities(0);

            let totalRecords = 0;
            const cols     = firstBatch.DATA?.COLUMNS || [];
            const firstRow = firstBatch.DATA?.DATA?.[0];
            const trcIdx   = cols.indexOf('TOTALRECORDCOUNT');

            if      (typeof firstBatch.TOTAL          === 'number') totalRecords = firstBatch.TOTAL;
            else if (typeof firstBatch.DATA?.TOTAL    === 'number') totalRecords = firstBatch.DATA.TOTAL;
            else if (trcIdx !== -1 && firstRow && typeof firstRow[trcIdx] === 'number') totalRecords = firstRow[trcIdx];
            else {
                console.error('⚠️ Could not determine totalRecords. Keys:', Object.keys(firstBatch));
                throw new Error('Unable to determine total record count from API response. Check console.');
            }

            updateStatus(`📊 Found ${totalRecords} activities in system`, 15);

            const columns       = firstBatch.DATA.COLUMNS;
            const allActivities = [...firstBatch.DATA.DATA];

            // Parallel batch fetch for remaining activity pages
            const offsets = [];
            for (let s = BATCH_SIZE; s < totalRecords; s += BATCH_SIZE) offsets.push(s);

            const totalBatches = offsets.length;
            let completedBatches = 0;
            let nextIndex = 0;

            if (totalBatches > 0) {
                updateStatus(`📥 Fetching ${totalBatches} remaining batches (${CONCURRENCY} parallel)...`, 15);

                async function worker() {
                    while (true) {
                        const myIndex = nextIndex++;
                        if (myIndex >= offsets.length) return;
                        const offset = offsets[myIndex];
                        const batch  = await fetchActivitiesWithRetry(offset);
                        allActivities.push(...batch.DATA.DATA);
                        completedBatches++;
                        updateStatus(
                            `📥 Batch ${completedBatches}/${totalBatches} complete (offset ${offset})`,
                            15 + (completedBatches / totalBatches) * 40
                        );
                    }
                }

                const workerPromises = [];
                for (let i = 0; i < Math.min(CONCURRENCY, totalBatches); i++) workerPromises.push(worker());
                await Promise.all(workerPromises);
            }

            updateStatus(`✅ Fetched all ${allActivities.length} activities`, 55);

            // ── Await closed LeadIDs (should already be resolved) ──
            let closedLeadIds = new Set();
            if (closedLeadIdsPromise) {
                updateStatus('⏳ Waiting for closed opportunity list...', 57);
                closedLeadIds = await closedLeadIdsPromise;
                updateStatus(`🚫 ${closedLeadIds.size} closed opportunities identified`, 59);
            }

            // ── Parse activities ──────────────────────────────────
            updateStatus('🔄 Parsing activity data...', 60);

            const parsedWithMeta = allActivities.map(rawRow => {
                const activityObj = {};
                columns.forEach((col, idx) => { activityObj[col] = rawRow[idx]; });

                const parsed = {
                    'Owner(s)':         parseLinkedField(activityObj.CALLERS),
                    'Company(ies)':     parseLinkedField(activityObj.LINKEDCOMPANIES),
                    'Subject':          activityObj.SUBJECT || '',
                    'Comments':         stripHTML(activityObj.COMMENTS),
                    'Start Date':       parseDate(activityObj.STARTDATEDATEONLY),
                    'End Date':         parseDate(activityObj.ENDDATEDATEONLY),
                    'Status':           decodeStatus(activityObj.STATUS),
                    'Call Type':        activityObj.CALLTYPE || '',
                    'Call Disposition': activityObj.CALLDISPOSITION || '',
                    'Attendee(s)':      parseLinkedField(activityObj.ATTENDEES),
                    'Lead(s)':          parseLinkedField(activityObj.LINKEDLEADS),
                    'Opportunity(ies)': parseLinkedField(activityObj.LINKEDOPPORTUNITIES),
                    'Project(s)':       parseLinkedField(activityObj.LINKEDPROJECTS),
                    'Contact(s)':       parseLinkedField(activityObj.LINKEDCONTACTS)
                };

                // Extract LeadIDs from raw HTML for closed-opp filtering
                const _leadIds = filters.excludeClosed
                    ? extractLeadIds(activityObj.LINKEDOPPORTUNITIES || '')
                    : [];

                return { parsed, _leadIds };
            });

            // Unfiltered parsed activities for the Analysis tab
            const unfilteredActivities = parsedWithMeta.map(r => r.parsed);

            // ── Apply filters ─────────────────────────────────────
            updateStatus('🔍 Applying filters...', 65);

            const beforeCount = parsedWithMeta.length;

            // Excluded only if activity has linked opps AND ALL of them are closed.
            // Activities with no linked opp are always kept.
            function isExcludedByClosed({ _leadIds }) {
                if (!filters.excludeClosed || _leadIds.length === 0) return false;
                return _leadIds.every(id => closedLeadIds.has(id));
            }

            const filteredWithMeta = parsedWithMeta.filter(item => {
                const { parsed } = item;
                if (startDate && parsed['Start Date'] && parsed['Start Date'] < startDate) return false;
                if (endDate   && parsed['Start Date'] && parsed['Start Date'] > endDate)   return false;
                if (filters.owner       && !parsed['Owner(s)'].toLowerCase().includes(filters.owner))               return false;
                if (filters.company     && !parsed['Company(ies)'].toLowerCase().includes(filters.company))         return false;
                if (filters.opportunity && !parsed['Opportunity(ies)'].toLowerCase().includes(filters.opportunity)) return false;
                if (isExcludedByClosed(item)) return false;
                return true;
            });

            const filteredCount = beforeCount - filteredWithMeta.length;
            if (filteredCount > 0) updateStatus(`🔍 Filtered out ${filteredCount} activities`, 67);

            let parsedActivities = filteredWithMeta.map(r => r.parsed);
            updateStatus(`📊 ${parsedActivities.length} activities in filtered view`, 68);

            if (parsedActivities.length === 0) {
                updateStatus('⚠️ No activities match your filters. Try adjusting your criteria.', 100);
                progressFill.style.background = '#ff9800';
                return;
            }

            // Analysis tab: closed-opp filter only — date/owner/company/opp filters
            // are intentionally excluded so relationship intelligence reflects the
            // full picture minus closed pursuits.
            let analysisActivities = unfilteredActivities;
            if (filters.excludeClosed) {
                analysisActivities = parsedWithMeta
                    .filter(item => !isExcludedByClosed(item))
                    .map(r => r.parsed);
                updateStatus(`📈 Analysis tab: ${analysisActivities.length} activities after closed-opp filter`, 70);
            }

            // ── Sort and build Excel ──────────────────────────────
            updateStatus('📋 Sorting by Owner...', 72);
            parsedActivities.sort((a, b) =>
                (a['Owner(s)'] || '').toLowerCase().localeCompare((b['Owner(s)'] || '').toLowerCase())
            );

            updateStatus('📝 Creating Excel file...', 75);
            if (typeof XLSX === 'undefined') {
                updateStatus('📦 Loading Excel library...', 78);
                await new Promise((resolve, reject) => {
                    const script = document.createElement('script');
                    script.src = SHEETJS;
                    script.onload = () => {
                        let retries = 0;
                        const check = setInterval(() => {
                            if (window.XLSX?.utils?.aoa_to_sheet) { clearInterval(check); resolve(); }
                            else if (++retries > 50) { clearInterval(check); reject(new Error('XLSX not available')); }
                        }, 100);
                    };
                    script.onerror = reject;
                    document.head.appendChild(script);
                });
            }

            updateStatus('📊 Building relationship intelligence analysis...', 82);
            const analysisSheet = buildActivitiesAnalysis(analysisActivities);

            updateStatus('✨ Formatting Activities sheet...', 88);
            const ws    = XLSX.utils.json_to_sheet(parsedActivities);
            const range = XLSX.utils.decode_range(ws['!ref']);
            const EXCEL_BLUE = '4472C4';

            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cellRef]) continue;
                    ws[cellRef].s = {
                        font: { name: 'Arial', sz: 10, bold: R === 0, color: R === 0 ? { rgb: 'FFFFFF' } : { rgb: '000000' } },
                        alignment: { vertical: 'center', horizontal: 'left' }
                    };
                    if (R === 0) ws[cellRef].s.fill = { patternType: 'solid', fgColor: { rgb: EXCEL_BLUE } };
                }
            }

            for (let row = range.s.r + 1; row <= range.e.r; row++) {
                ['E', 'F'].forEach(col => {
                    const cellRef = col + (row + 1);
                    if (ws[cellRef]?.v instanceof Date) { ws[cellRef].t = 'd'; ws[cellRef].z = 'mm/dd/yyyy'; }
                });
            }

            ws['!cols'] = Object.keys(parsedActivities[0]).map(key => ({
                wch: Math.min(Math.max(key.length, ...parsedActivities.map(row => {
                    const val = row[key];
                    if (val instanceof Date) return 12;
                    return (val || '').toString().length;
                })) + 2, 50)
            }));

            ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
            if (!ws['!rows']) ws['!rows'] = [];
            ws['!rows'][0] = { hpt: 20 };

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Activities');
            XLSX.utils.book_append_sheet(wb, analysisSheet, 'Activities Analysis');

            updateStatus('💾 Downloading file...', 95);
            const timestamp = new Date().toISOString().split('T')[0];
            const filename  = `Unanet_Activities_Export_${timestamp}.xlsx`;
            XLSX.writeFile(wb, filename);

            updateStatus(`✅ Export complete! Downloaded: ${filename}`, 100);
            updateStatus(`📊 Activities exported: ${parsedActivities.length}`, 100);
            updateStatus(`📈 Analysis generated from ${analysisActivities.length} activities`, 100);
            updateStatus('🎉 This window will close in 3 seconds...', 100);
            closeDialog();

        } catch(error) {
            updateStatus(`❌ Error: ${error.message}`, 100);
            progressFill.style.background = '#f44336';
            console.error('Export error:', error);
        }
    }

})();
