/**
 * Borderellen Converter - Results Manager Module
 * Handles all results tab functionality including display, export, and search
 */

// ========== EVENT LISTENERS SETUP ==========

/**
 * Set up event listeners for Results tab buttons
 */
function setupResultsTabListeners() {
    const downloadExcelBtn = document.getElementById('download-excel-btn');
    const exportJsonBtn = document.getElementById('export-json-btn');
    const refreshDataBtn = document.getElementById('refresh-data-btn');
    const searchRecords = document.getElementById('search-records');

    console.log('Setting up Results tab event listeners...');
    console.log('downloadExcelBtn found:', !!downloadExcelBtn);
    console.log('exportJsonBtn found:', !!exportJsonBtn);

    // Remove existing listeners to avoid duplicates
    if (downloadExcelBtn) {
        console.log('Adding click listener to download Excel button');
        downloadExcelBtn.removeEventListener('click', downloadCombinedExcel);
        downloadExcelBtn.addEventListener('click', downloadCombinedExcel);
    } else {
        console.error('download-excel-btn not found!');
    }

    if (exportJsonBtn) {
        console.log('Adding click listener to export JSON button');
        exportJsonBtn.removeEventListener('click', exportCombinedJSON);
        exportJsonBtn.addEventListener('click', exportCombinedJSON);
    } else {
        console.error('export-json-btn not found!');
    }

    if (refreshDataBtn) {
        console.log('Adding click listener to refresh button');
        refreshDataBtn.removeEventListener('click', updateResultsTab);
        refreshDataBtn.addEventListener('click', updateResultsTab);
    } else {
        console.error('refresh-data-btn not found!');
    }

    // Search functionality
    if (searchRecords) {
        console.log('Adding input listener to search field');
        searchRecords.removeEventListener('input', handleSearchInput);
        searchRecords.addEventListener('input', handleSearchInput);
    } else {
        console.error('search-records not found!');
    }
}

/**
 * Handle search input for filtering table rows
 */
function handleSearchInput(e) {
    const searchTerm = e.target.value.toLowerCase();
    const tableRows = document.querySelectorAll('#results-table-body tr');

    tableRows.forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(searchTerm)) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

// ========== MAIN RESULTS TAB UPDATE ==========

/**
 * Update the Filled Broker Template tab with processed data
 */
function updateResultsTab() {
    const processedFiles = window.uploadedFiles?.filter(f =>
        f.parsedData &&
        f.parsedData.length > 0 &&
        f.statusClass === 'status-success'
    ) || [];

    const noDataMessage = document.getElementById('no-data-message');
    const processedDataSection = document.getElementById('processed-data-section');
    const processingOverview = document.getElementById('processing-overview');

    if (processedFiles.length === 0) {
        // Show no data message
        if (noDataMessage) noDataMessage.style.display = 'block';
        if (processedDataSection) processedDataSection.style.display = 'none';
        processingOverview.innerHTML = `
            <div id="no-data-message" style="text-align: center; padding: 40px; color: #888;">
                <p>No processed files available.</p>
                <small>Upload broker files in the Upload tab to see processed data here.</small>
            </div>
        `;
        return;
    }

    // Show processing overview
    displayProcessingOverview(processedFiles);

    // Combine processed data - show first 3 + last 3 records per file for representative sampling
    const combinedData = [];
    processedFiles.forEach(file => {
        const fileRecords = file.parsedData.map(record => ({
            ...record,
            _sourceFile: file.name,
            _sourceId: file.id
        }));

        if (fileRecords.length <= 6) {
            // If 6 or fewer records, show them all
            combinedData.push(...fileRecords);
        } else {
            // Show first 3 + last 3 records from this file
            combinedData.push(
                ...fileRecords.slice(0, 3),
                ...fileRecords.slice(-3)
            );
        }
    });

    // Display the data
    displayCombinedResults(combinedData, processedFiles);

    if (noDataMessage) noDataMessage.style.display = 'none';
    if (processedDataSection) processedDataSection.style.display = 'block';
}

// ========== DISPLAY FUNCTIONS ==========

/**
 * Display processing overview with file summaries
 * @param {Array} processedFiles - Successfully processed files
 */
function displayProcessingOverview(processedFiles) {
    const totalRecords = processedFiles.reduce((sum, f) => sum + f.parsedData.length, 0);
    const totalFiles = processedFiles.length;
    const allFiles = window.uploadedFiles?.length || 0;

    const overview = document.getElementById('processing-overview');
    overview.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 16px;">
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Total Records:</strong> ${totalRecords.toLocaleString()}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Processed Files:</strong> ${totalFiles} of ${allFiles}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Format:</strong> Standard 22-column template
            </div>
        </div>

        <div style="margin-bottom: 16px;">
            <h4 style="margin-bottom: 8px;">File Processing Status:</h4>
            <div style="display: flex; flex-wrap: wrap; gap: 8px;">
                ${processedFiles.map(file => `
                    <div style="display: flex; align-items: center; background: #2a2a2a; padding: 8px 12px; border-radius: 4px;">
                        <span class="status-indicator ${file.statusClass}"></span>
                        <span style="margin-left: 8px;">${file.name}</span>
                        <small style="margin-left: 8px; color: #888;">(${file.parsedData.length} records)</small>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

/**
 * Display combined results table
 * @param {Array} combinedData - All processed records from all files
 * @param {Array} processedFiles - Source files for reference
 */
function displayCombinedResults(combinedData, processedFiles) {
    const tableHead = document.getElementById('results-table-head');
    const tableBody = document.getElementById('results-table-body');
    const recordCountDisplay = document.getElementById('record-count-display');

    if (combinedData.length === 0) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 20px;">No processed data available</td></tr>';
        recordCountDisplay.textContent = '';
        return;
    }

    // Get columns from active template (always show all template columns)
    const templateColumns = window.borderellenTemplate?.columns || [];
    const columns = templateColumns.map(col => col.name);

    // Create table headers
    tableHead.innerHTML = `
        <tr>
            <th>Source File</th>
            ${columns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;

    // Group records by source file for better display organization
    const recordsByFile = {};
    combinedData.forEach(row => {
        if (!recordsByFile[row._sourceFile]) {
            recordsByFile[row._sourceFile] = [];
        }
        recordsByFile[row._sourceFile].push(row);
    });

    // Build table with file separators and record indicators
    let tableHTML = '';
    Object.keys(recordsByFile).forEach(fileName => {
        const fileRecords = recordsByFile[fileName];
        const sourceFile = processedFiles.find(f => f.name === fileName);
        const totalRecords = sourceFile ? sourceFile.parsedData.length : fileRecords.length;
        const showingAll = fileRecords.length === totalRecords;

        // Add file header row
        tableHTML += `<tr style="background: #2a2a2a; border-top: 2px solid #404040;">
            <td colspan="${columns.length + 1}" style="font-weight: bold; padding: 8px;">
                📄 ${fileName}
                <span style="color: #888; font-weight: normal;">(showing ${fileRecords.length} of ${totalRecords.toLocaleString()} records${!showingAll ? ' - first 3 + last 3' : ''})</span>
            </td>
        </tr>`;

        // Add records for this file
        fileRecords.forEach((row, index) => {
            const isGap = !showingAll && index === 3; // Add gap indicator between first 3 and last 3

            if (isGap && totalRecords > 6) {
                tableHTML += `<tr style="background: #1a1a1a;">
                    <td colspan="${columns.length + 1}" style="text-align: center; padding: 4px; color: #666; font-style: italic;">
                        ... ${(totalRecords - 6).toLocaleString()} records omitted ...
                    </td>
                </tr>`;
            }

            tableHTML += `<tr>
                <td style="color: #888; font-size: 11px;">${row._sourceFile}</td>
                ${columns.map(col => {
                    const value = row[col];
                    return `<td>${value !== undefined && value !== null && value !== '' ? value : '<em style="color: #888;">empty</em>'}</td>`;
                }).join('')}
            </tr>`;
        });
    });

    tableBody.innerHTML = tableHTML;

    // Update record count display
    const totalRecordsAll = processedFiles.reduce((sum, f) => sum + f.parsedData.length, 0);
    const displayText = `Showing ${combinedData.length.toLocaleString()} sample records from ${totalRecordsAll.toLocaleString()} total across ${processedFiles.length} files`;

    recordCountDisplay.textContent = displayText;

    // Store FULL data for export (not just the sample)
    const fullCombinedData = [];
    processedFiles.forEach(file => {
        file.parsedData.forEach(record => {
            fullCombinedData.push({
                ...record,
                _sourceFile: file.name,
                _sourceId: file.id
            });
        });
    });
    window.currentCombinedData = fullCombinedData;
}

// ========== EXPORT FUNCTIONS ==========

/**
 * Download combined data as Excel file
 */
async function downloadCombinedExcel() {
    console.log('downloadCombinedExcel called');
    console.log('currentCombinedData:', window.currentCombinedData?.length);
    console.log('appSettings:', window.appSettings);

    if (!window.currentCombinedData || window.currentCombinedData.length === 0) {
        alert('No processed data available for download.');
        return;
    }

    try {
        // Define the 22-column template order
        const columnOrder = [
            'Makelaar', 'Boekingsperiode', 'Polisnr makelaar', 'Verzekerde', 'Branche',
            'Periode van', 'Periode tot', 'Valuta', 'Bruto', 'Provisie%',
            'Provisie', 'Tekencom%', 'Tekencom', 'Netto', 'BAB',
            'Land', 'Aandeel Allianz', 'Tekenjaar', 'Boekdatum tp', 'FactuurDtm',
            'FactuurNr', 'Boekingsreden'
        ];

        // Clean data for export and ensure all 22 columns are present
        const exportData = window.currentCombinedData.map(row => {
            const cleanRow = {};

            // First, add all columns in the correct order
            columnOrder.forEach(columnName => {
                cleanRow[columnName] = row[columnName] || '';
            });

            // Then add any additional columns that aren't internal fields
            Object.keys(row).forEach(key => {
                if (!key.startsWith('_') && !columnOrder.includes(key)) {
                    cleanRow[key] = row[key];
                }
            });

            return cleanRow;
        });

        // Create Excel workbook
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(exportData);

        // Add worksheet
        XLSX.utils.book_append_sheet(wb, ws, 'Processed Data');

        // Generate filename with timestamp
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `Borderellen_Combined_${timestamp}.xlsx`;

        // Download to preferred folder or fallback to browser download
        const success = await window.downloadExcelToPreferredFolder(wb, filename);

        if (success) {
            alert(`Downloaded ${exportData.length} records successfully`);
        }

    } catch (error) {
        console.error('Error downloading Excel:', error);
        alert(`Failed to download Excel file: ${error.message}`);
    }
}

/**
 * Export combined data as JSON file
 */
async function exportCombinedJSON() {
    if (!window.currentCombinedData || window.currentCombinedData.length === 0) {
        alert('No processed data available for export.');
        return;
    }

    try {
        // Clean data for export
        const exportData = window.currentCombinedData.map(row => {
            const cleanRow = {};
            Object.keys(row).forEach(key => {
                if (!key.startsWith('_')) {
                    cleanRow[key] = row[key];
                }
            });
            return cleanRow;
        });

        const dataStr = JSON.stringify(exportData, null, 2);
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `Borderellen_Combined_${timestamp}.json`;

        // Create blob and download to preferred folder or fallback to browser download
        const blob = new Blob([dataStr], { type: 'application/json' });
        const success = await window.downloadToPreferredFolder(blob, filename, 'application/json');

        if (success) {
            alert(`Exported ${exportData.length} records successfully`);
        }

    } catch (error) {
        console.error('Error exporting JSON:', error);
        alert(`Failed to export JSON file: ${error.message}`);
    }
}

// ========== AUTO-NAVIGATION ==========

/**
 * Auto-navigate to appropriate tab on app start
 */
function autoNavigateOnStart() {
    // If there's already an active template, go straight to Upload tab
    if (window.currentTemplateId && window.borderellenTemplate) {
        console.log('Active template detected, auto-navigating to Upload tab');
        // Click on the Upload tab
        const uploadTab = document.querySelector('[data-tab="upload"]');
        if (uploadTab) {
            uploadTab.click();
        }
    }
    // Otherwise, stay on Templates tab (default)
}

/**
 * Auto-select file in Broker Template tab with smart prioritization
 * Focuses on files that need attention first
 */
function autoSelectFileInMappingTab() {
    if (!window.uploadedFiles || window.uploadedFiles.length === 0) {
        return; // No files to select
    }

    // Check if there's already a selection
    const selector = document.getElementById('mapping-file-selector');
    if (selector && selector.value) {
        return; // Already has a selection
    }

    // Priority 1: Files that failed processing (need manual attention)
    let targetFile = window.uploadedFiles.find(f => f.statusClass === 'status-error');

    // Priority 2: Files that have warnings (may need template tweaks)
    if (!targetFile) {
        targetFile = window.uploadedFiles.find(f => f.statusClass === 'status-warning');
    }

    // Priority 3: Unknown broker files (need custom templates)
    if (!targetFile) {
        targetFile = window.uploadedFiles.find(f => f.broker.type === 'unknown');
    }

    // Priority 4: Any uploaded file
    if (!targetFile) {
        targetFile = window.uploadedFiles[0];
    }

    if (targetFile && selector) {
        // Auto-select the prioritized file
        selector.value = targetFile.id;

        // Trigger change event to load file data
        const changeEvent = new Event('change', { bubbles: true });
        selector.dispatchEvent(changeEvent);

        console.log('Auto-selected file for mapping:', targetFile.name, 'Status:', targetFile.statusClass);
    }
}

// ========== GLOBAL EXPORTS ==========

// Make results functions globally accessible
window.setupResultsTabListeners = setupResultsTabListeners;
window.handleSearchInput = handleSearchInput;
window.updateResultsTab = updateResultsTab;
window.displayProcessingOverview = displayProcessingOverview;
window.displayCombinedResults = displayCombinedResults;
window.downloadCombinedExcel = downloadCombinedExcel;
window.exportCombinedJSON = exportCombinedJSON;
window.autoNavigateOnStart = autoNavigateOnStart;
window.autoSelectFileInMappingTab = autoSelectFileInMappingTab;