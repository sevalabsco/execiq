// Unanet Activities Export Tool
// Fetches all activities from Unanet CRM and exports to Excel

(async function() {
    // Create GUI overlay
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
            // Parse "March, 10 2020 00:00:00" format
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
        
        // Fetch all activities
        let allActivities = [];
        let start = 0;
        let totalRecords = 0;
        
        updateStatus('📥 Fetching activities from Unanet...', 10);
        
        // First request to get total count
        const firstBatch = await fetchActivities(0);
        totalRecords = firstBatch.TOTAL;
        updateStatus(`📊 Found ${totalRecords} activities`, 15);
        
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
        
        const parsedActivities = allActivities.map(activity => {
            const activityObj = {};
            columns.forEach((col, idx) => {
                activityObj[col] = activity[idx];
            });
            
            // Return in the requested column order
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
        
        updateStatus('📋 Sorting by Owner...', 70);
        
        // Sort by Owner
        parsedActivities.sort((a, b) => {
            const ownerA = (a['Owner(s)'] || '').toLowerCase();
            const ownerB = (b['Owner(s)'] || '').toLowerCase();
            return ownerA.localeCompare(ownerB);
        });
        
        // Create Excel file
        updateStatus('📝 Creating Excel file...', 75);
        
        // Try loading SheetJS from unpkg (often more CSP-friendly)
        if (typeof XLSX === 'undefined') {
            updateStatus('📦 Loading Excel library...', 80);
            try {
                await new Promise((resolve, reject) => {
                    const script = document.createElement('script');
                    script.src = 'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js';
                    script.onload = resolve;
                    script.onerror = () => {
                        // Fallback to jsdelivr
                        const script2 = document.createElement('script');
                        script2.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
                        script2.onload = resolve;
                        script2.onerror = reject;
                        document.head.appendChild(script2);
                    };
                    document.head.appendChild(script);
                });
            } catch (e) {
                throw new Error('Unable to load Excel library. Please check your browser security settings.');
            }
        }
        
        updateStatus('✨ Formatting Excel file...', 85);
        
        // Create worksheet
        const ws = XLSX.utils.json_to_sheet(parsedActivities);
        
        // Format date columns (D and E are Start Date and End Date in new order)
        const range = XLSX.utils.decode_range(ws['!ref']);
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
        
        // Add autofilter to all columns
        ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Activities');
        
        updateStatus('💾 Downloading file...', 95);
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().split('T')[0];
        const filename = `Unanet_Activities_Export_${timestamp}.xlsx`;
        
        // Download file
        XLSX.writeFile(wb, filename);
        
        updateStatus(`✅ Export complete! Downloaded: ${filename}`, 100);
        updateStatus(`📊 Total activities exported: ${parsedActivities.length}`, 100);
        updateStatus('🎉 This window will close in 3 seconds...', 100);
        
        closeDialog();
        
    } catch (error) {
        updateStatus(`❌ Error: ${error.message}`, 100);
        progressFill.style.background = '#f44336';
        console.error('Export error:', error);
        // Don't auto-close on error so user can read the message
    }
})();
