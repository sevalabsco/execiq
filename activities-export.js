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
            updateStatus('🔄 Parsing and filtering activity data...', 60);
            
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
            
            // Apply filters
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
                updateStatus(`🔍 Filtered out ${filteredCount} activities`, 65);
            }
            updateStatus(`📊 ${parsedActivities.length} activities match your filters`, 70);
            
            if (parsedActivities.length === 0) {
                updateStatus('⚠️ No activities match your filters. Try adjusting your criteria.', 100);
                progressFill.style.background = '#ff9800';
                return;
            }
            
            updateStatus('📋 Sorting by Owner...', 75);
            
            // Sort by Owner
            parsedActivities.sort((a, b) => {
                const ownerA = (a['Owner(s)'] || '').toLowerCase();
                const ownerB = (b['Owner(s)'] || '').toLowerCase();
                return ownerA.localeCompare(ownerB);
            });
            
            // Create Excel file
            updateStatus('📝 Creating Excel file...', 80);
            
            // Load xlsx-js-style (same as ExecIQ)
            if (typeof XLSX === 'undefined') {
                updateStatus('📦 Loading Excel library with styling support...', 85);
                try {
                    await new Promise((resolve, reject) => {
                        const script = document.createElement('script');
                        script.src = SHEETJS;
                        script.onload = () => {
                            let retries = 0;
                            const checkInterval = setInterval(() => {
                                if (window.XLSX) {
                                    clearInterval(checkInterval);
                                    resolve();
                                } else if (++retries > 40) {
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
            
            updateStatus('✨ Formatting Excel file...', 90);
            
            // Create worksheet
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
                            bold: R === 0,  // Bold for header row
                            color: R === 0 ? { rgb: 'FFFFFF' } : { rgb: '000000' }
                        },
                        alignment: {
                            vertical: 'center',
                            horizontal: 'left'
                        }
                    };
                    
                    // Header row (row 0) - Blue background, white text, bold
                    if (R === 0) {
                        ws[cellRef].s.fill = {
                            patternType: 'solid',
                            fgColor: { rgb: EXCEL_BLUE }
                        };
                    }
                }
            }
            
            // Format date columns (D and E are Start Date and End Date)
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
            
            // Set row height for header
            if (!ws['!rows']) ws['!rows'] = [];
            ws['!rows'][0] = { hpt: 20 };
            
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
        }
    }
})();
