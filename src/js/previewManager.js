/**
 * Borderellen Converter - Preview Manager Module
 * Handles all data preview functionality including file previews, mapping previews, and export
 */

// ========== DATA FILTERING ==========

/**
 * Filter data for preview based on broker type
 * @param {Array} rawData - Raw data from Excel
 * @param {String} brokerType - Broker type (AON, VGA, BCI, Voogt)
 * @returns {Array} Filtered data
 */
function filterDataForPreview(rawData, brokerType) {
    switch (brokerType) {
        case 'AON':
            return rawData.filter(row => row.PolisNr || row.Verzekerde);
        case 'VGA':
            return rawData.filter(row =>
                row.Soort !== 'Total' &&
                row.Soort !== 'Totaal' &&
                (row.Polisnummer || row['Naam verzekeringnemer'])
            );
        case 'BCI':
        case 'Voogt':
            // These have complex parsing - just return raw data for preview
            return rawData;
        default:
            return rawData;
    }
}

// ========== MAPPING DISPLAY ==========

/**
 * Display mapping summary
 * @param {Object} mapping - Mapping configuration
 */
function displayMappingSummary(mapping) {
    const container = document.getElementById('preview-mapping-summary');
    const mappingCount = Object.keys(mapping).length;
    const fixedCount = Object.values(mapping).filter(v => v.startsWith('FIXED:')).length;
    const columnCount = mappingCount - fixedCount;

    container.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 16px;">
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Total Mappings:</strong> ${mappingCount}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Fixed Values:</strong> ${fixedCount}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Column Mappings:</strong> ${columnCount}
            </div>
        </div>

        <div style="max-height: 200px; overflow-y: auto; background: #2a2a2a; padding: 12px; border-radius: 6px;">
            ${Object.keys(mapping).map(targetField => {
                const source = mapping[targetField];
                const isFixed = source.startsWith('FIXED:');
                const displaySource = isFixed ? source.substring(6) : source;
                const type = isFixed ? 'Fixed' : 'Column';

                return `<div style="display: flex; justify-content: space-between; padding: 4px 0; border-bottom: 1px solid #404040;">
                    <span><strong>${targetField}</strong></span>
                    <span style="color: #888;">${type}: ${displaySource}</span>
                </div>`;
            }).join('')}
        </div>
    `;
}

// ========== TABLE DISPLAY ==========

/**
 * Display preview data table
 * @param {Array} mappedData - Mapped data to display
 */
function displayPreviewTable(mappedData) {
    const tableHead = document.getElementById('preview-table-head');
    const tableBody = document.getElementById('preview-table-body');

    if (mappedData.length === 0) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 20px;">No data to preview</td></tr>';
        return;
    }

    // ALWAYS use active template columns for consistent Results tab display
    const templateColumns = window.borderellenTemplate?.columns || [];
    const columns = templateColumns.length > 0 ? templateColumns.map(col => col.name) :
                   [...new Set(mappedData.flatMap(row => Object.keys(row)))]; // Fallback to data columns if no template

    console.log(`Results tab: Displaying ${columns.length} template columns for ${mappedData.length} records`);

    // Create table headers
    tableHead.innerHTML = `
        <tr>
            ${columns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;

    // Create table rows
    tableBody.innerHTML = mappedData.map(row => `
        <tr>
            ${columns.map(col => `<td>${row[col] || '<em style="color: #888;">empty</em>'}</td>`).join('')}
        </tr>
    `).join('');
}

/**
 * Display preview table with first/last record indicators
 * @param {Array} displayData - Data to display
 * @param {Object} fileData - File data object
 */
function displayPreviewTableWithIndicators(displayData, fileData) {
    const tableHead = document.getElementById('preview-table-head');
    const tableBody = document.getElementById('preview-table-body');

    if (displayData.length === 0) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 20px;">No data to preview</td></tr>';
        return;
    }

    // ALWAYS use active template columns for consistent Results tab display
    const templateColumns = window.borderellenTemplate?.columns || [];
    const columns = templateColumns.length > 0 ? templateColumns.map(col => col.name) :
                   [...new Set(displayData.flatMap(row => Object.keys(row)))]; // Fallback to data columns if no template

    console.log(`Results tab (with indicators): Displaying ${columns.length} template columns`);

    // Create table headers
    tableHead.innerHTML = `
        <tr>
            <th style="width: 40px;">#</th>
            ${columns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;

    // Determine if we have first + last records or just sequential
    const totalRecords = fileData.recordCount || displayData.length;
    const hasGap = displayData.length === 10 && totalRecords > 10;

    // Create table rows with indicators
    tableBody.innerHTML = displayData.map((row, index) => {
        let rowNumber, indicator;

        if (!hasGap) {
            // Sequential data
            rowNumber = index + 1;
            indicator = '';
        } else if (index < 5) {
            // First 5 records
            rowNumber = index + 1;
            indicator = index === 4 ? ' style="border-bottom: 3px solid #00bcd4;"' : '';
        } else {
            // Last 5 records
            rowNumber = totalRecords - (9 - index);
            indicator = index === 5 ? ' style="border-top: 3px solid #00bcd4;"' : '';
        }

        return `
            <tr${indicator}>
                <td style="color: #888; font-weight: bold;">${rowNumber}</td>
                ${columns.map(col => `<td>${row[col] !== undefined && row[col] !== null && row[col] !== '' ? row[col] : '<em style="color: #888;">empty</em>'}</td>`).join('')}
            </tr>
        `;
    }).join('') +
    (hasGap ? '<tr><td colspan="100%" style="text-align: center; padding: 8px; color: #888; font-style: italic;">... ' + (totalRecords - 10) + ' records omitted ...</td></tr>' : '');
}

// ========== STATISTICS DISPLAY ==========

/**
 * Display preview statistics
 * @param {Array} mappedData - Mapped data
 * @param {Object} mapping - Mapping configuration
 * @param {Number} totalRecords - Total records processed
 */
function displayPreviewStats(mappedData, mapping, totalRecords) {
    const container = document.getElementById('preview-stats');

    // Calculate field completion stats
    const fieldStats = {};
    const columns = Object.keys(mapping);

    columns.forEach(col => {
        const filledCount = mappedData.filter(row => row[col] && row[col] !== '').length;
        fieldStats[col] = {
            filled: filledCount,
            empty: mappedData.length - filledCount,
            percentage: mappedData.length > 0 ? Math.round((filledCount / mappedData.length) * 100) : 0
        };
    });

    container.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px; margin-bottom: 16px;">
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Sample Size:</strong> ${mappedData.length} of ${totalRecords} records
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Output Columns:</strong> ${columns.length}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>File:</strong> ${window.currentMappingFile.name}
            </div>
        </div>

        <div style="max-height: 200px; overflow-y: auto; background: #2a2a2a; padding: 12px; border-radius: 6px;">
            <h4 style="margin-bottom: 12px;">Field Completion Rates:</h4>
            ${columns.map(col => {
                const stats = fieldStats[col];
                return `<div style="margin-bottom: 8px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span><strong>${col}</strong></span>
                        <span style="color: #888;">${stats.percentage}% complete</span>
                    </div>
                    <div style="background: #404040; height: 6px; border-radius: 3px; margin-top: 4px;">
                        <div style="background: #00bcd4; height: 100%; width: ${stats.percentage}%; border-radius: 3px;"></div>
                    </div>
                </div>`;
            }).join('')}
        </div>
    `;
}

/**
 * Display unified data quality statistics
 * @param {Array} displayData - Sample data
 * @param {Object} fileData - File data object
 * @param {Number} totalRecords - Total records
 */
function displayUnifiedDataQuality(displayData, fileData, totalRecords) {
    const container = document.getElementById('preview-stats');

    if (displayData.length === 0) {
        container.innerHTML = '<p style="color: #888;">No data available for quality analysis.</p>';
        return;
    }

    // Calculate data quality metrics
    const columns = Object.keys(displayData[0]);
    const qualityIssues = [];
    const fieldStats = {};

    columns.forEach(col => {
        const filledCount = displayData.filter(row =>
            row[col] !== undefined &&
            row[col] !== null &&
            row[col] !== ''
        ).length;

        const emptyCount = displayData.length - filledCount;
        const percentage = Math.round((filledCount / displayData.length) * 100);

        fieldStats[col] = { filled: filledCount, empty: emptyCount, percentage };

        // Flag quality issues (less than 100% completion)
        if (percentage < 100) {
            qualityIssues.push({ field: col, percentage, empty: emptyCount });
        }
    });

    container.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px; margin-bottom: 16px;">
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Sample Size:</strong> ${displayData.length} records
                ${totalRecords > displayData.length ? `<br><small style="color: #888;">of ${totalRecords} total</small>` : ''}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Data Quality:</strong> ${qualityIssues.length === 0 ? 'Excellent' : 'Issues Found'}
                <br><small style="color: ${qualityIssues.length === 0 ? '#4caf50' : '#ff9800'};">
                    ${qualityIssues.length === 0 ? 'All fields populated' : `${qualityIssues.length} fields incomplete`}
                </small>
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Columns:</strong> ${columns.length}
                <br><small style="color: #888;">Ready for processing</small>
            </div>
        </div>

        ${qualityIssues.length > 0 ? `
        <div style="background: #2a2a2a; padding: 12px; border-radius: 6px;">
            <h4 style="margin-bottom: 12px; color: #ff9800;">⚠ Data Quality Issues</h4>
            <p style="color: #888; font-size: 12px; margin-bottom: 12px;">Fields with missing data that may need attention:</p>
            ${qualityIssues.map(issue => `
                <div style="margin-bottom: 8px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span><strong>${issue.field}</strong></span>
                        <span style="color: #ff9800;">${issue.percentage}% complete (${issue.empty} empty)</span>
                    </div>
                    <div style="background: #404040; height: 4px; border-radius: 2px; margin-top: 4px;">
                        <div style="background: ${issue.percentage > 80 ? '#ff9800' : '#f44336'}; height: 100%; width: ${issue.percentage}%; border-radius: 2px;"></div>
                    </div>
                </div>
            `).join('')}
        </div>
        ` : `
        <div style="background: #2a2a2a; padding: 12px; border-radius: 6px;">
            <h4 style="margin-bottom: 12px; color: #4caf50;">✓ Excellent Data Quality</h4>
            <p style="color: #888; font-size: 12px; margin: 0;">All fields are properly populated. Data is ready for processing.</p>
        </div>
        `}
    `;
}

/**
 * Display unified file information (replaces both file info and mapping summary)
 * @param {Object} fileData - File data object
 * @param {Array} displayData - Sample data being shown
 * @param {Number} totalRecords - Total records in file
 */
function displayUnifiedFileInfo(fileData, displayData, totalRecords) {
    const container = document.getElementById('preview-file-info');
    const columnCount = displayData.length > 0 ? Object.keys(displayData[0]).length : 0;

    container.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 16px;">
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>File:</strong> ${fileData.name}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Broker:</strong> ${fileData.broker.name}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Total Records:</strong> ${fileData.recordCount || totalRecords || 'Unknown'}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Columns:</strong> ${columnCount}
            </div>
        </div>

        <div style="background: #2a2a2a; padding: 12px; border-radius: 6px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <h4 style="margin: 0;">Processing Status</h4>
                <span style="background: ${fileData.statusClass === 'status-success' ? '#4caf50' : fileData.statusClass === 'status-warning' ? '#ff9800' : '#f44336'}; padding: 4px 8px; border-radius: 4px; font-size: 12px;">
                    ${fileData.status}
                </span>
            </div>
            ${fileData.broker.type === 'built-in' ?
                '<p style="color: #00bcd4; font-size: 12px; margin: 0;">✓ Automatically processed using built-in broker parser</p>' :
                '<p style="color: #ff9800; font-size: 12px; margin: 0;">⚠ Custom processing or manual template may be required</p>'
            }
        </div>
    `;
}

/**
 * Display file statistics for raw preview
 * @param {Array} displayData - Sample data
 * @param {Object} fileData - File data object
 * @param {Number} totalRecords - Total records
 */
function displayFileStats(displayData, fileData, totalRecords) {
    const container = document.getElementById('preview-stats');

    container.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px;">
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Sample Size:</strong> ${displayData.length} of ${fileData.recordCount || totalRecords} records
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>File Size:</strong> ${fileData.size}
            </div>
            <div style="background: #333; padding: 12px; border-radius: 6px;">
                <strong>Processing Status:</strong> ${fileData.status}
            </div>
        </div>

        ${displayData.length > 0 ? `
        <div style="background: #2a2a2a; padding: 12px; border-radius: 6px; margin-top: 16px;">
            <h4 style="margin-bottom: 12px;">Data Quality:</h4>
            <p style="color: #888; margin-bottom: 8px;">Showing sample of available data from this ${fileData.broker.name} file.</p>
            ${fileData.broker.type === 'built-in' ?
                '<p style="color: #00bcd4; font-size: 12px;">✓ This file format is supported and can be auto-mapped in the Broker Template tab.</p>' :
                '<p style="color: #ff9800; font-size: 12px;">⚠ This file format may need custom template creation.</p>'
            }
        </div>
        ` : ''}
    `;
}

// ========== MAIN PREVIEW FUNCTION ==========

/**
 * Unified preview function for file data (with or without mapping)
 * @param {Object} fileData - File data object
 * @param {Object} mapping - Optional mapping configuration (if null, shows raw data)
 */
async function previewFileData(fileData, mapping = null) {
    try {
        // Show loading state
        window.showPreviewModal();
        document.getElementById('preview-content').innerHTML = '<div style="text-align: center; padding: 40px;"><p>Processing preview...</p></div>';

        // Get sample data - always use the same source for consistency
        let sampleData = [];

        if (fileData.parsedData && fileData.parsedData.length > 0) {
            // Use already parsed data if available (this is the processed broker data)
            // Get first 5 + last 5 records for better data validation
            const allData = fileData.parsedData;
            if (allData.length <= 10) {
                sampleData = allData;
            } else {
                sampleData = [
                    ...allData.slice(0, 5),
                    ...allData.slice(-5)
                ];
            }
        } else {
            // If no parsed data available, parse the raw Excel file with skip rules if available
            const workbook = await ExcelCacheManager.getWorkbook(fileData.file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            let rawData;
            // Use cached analysis from fileData if available, otherwise fall back to global currentPatternAnalysis
            const analysisToUse = fileData.patternAnalysis || window.currentPatternAnalysis;

            if (analysisToUse && analysisToUse.dataSection) {
                // Use GenericBrokerParser to respect skip rules
                console.log('Preview: Using skip rules for raw data extraction');
                const tempConfig = {
                    dataStartMethod: 'skip-rows',
                    skipRows: analysisToUse.dataSection.dataStartIndex || 0,
                    skipColumns: analysisToUse.dataSection.startColumnIndex || 0,
                    headerRow: analysisToUse.dataSection.headerRowIndex,
                    columnMapping: {} // Identity mapping for preview
                };
                const parser = new window.GenericBrokerParser(tempConfig);
                rawData = await parser.parse(workbook, fileData.name);
            } else {
                // Fallback to simple sheet reading
                console.log('Preview: No skip rules available, using simple sheet reading');
                rawData = XLSX.utils.sheet_to_json(worksheet, {
                    cellFormula: false  // Read calculated values instead of formulas
                });
            }

            // Apply broker-specific filtering to match what the parsers would do
            const brokerType = fileData.broker.parser;
            const filteredData = filterDataForPreview(rawData, brokerType);

            // Get first 5 + last 5 records for better data validation
            if (filteredData.length <= 10) {
                sampleData = filteredData;
            } else {
                sampleData = [
                    ...filteredData.slice(0, 5),
                    ...filteredData.slice(-5)
                ];
            }
        }

        let displayData = sampleData;
        let displayMapping = mapping;

        if (!mapping) {
            // No mapping provided - show file preview
            displayMapping = {};

            // For files with parsedData, show the processed data as-is (it's already formatted)
            // For files without parsedData, format the raw Excel data
            if (fileData.parsedData && fileData.parsedData.length > 0) {
                // Already processed data - just use it
                displayData = sampleData;

                // Create identity mapping for display
                if (sampleData.length > 0) {
                    Object.keys(sampleData[0]).forEach(key => {
                        displayMapping[key] = key;
                    });
                }
            } else {
                // Raw Excel data - format dates
                displayData = sampleData.map(row => {
                    const formattedRow = {};
                    Object.keys(row).forEach(key => {
                        let value = row[key];
                        // Format Excel dates in raw preview
                        if (window.isExcelDate && window.isExcelDate(value) && window.isDateField && window.isDateField(key)) {
                            value = window.formatExcelDate(value);
                        }
                        formattedRow[key] = value;
                    });
                    return formattedRow;
                });

                if (sampleData.length > 0) {
                    Object.keys(sampleData[0]).forEach(key => {
                        displayMapping[key] = key;
                    });
                }
            }
        } else {
            // Mapping is provided - check if we have processed data or need to apply mapping
            if (fileData.parsedData && fileData.parsedData.length > 0 &&
                fileData.broker.type === 'built-in') {

                // For built-in brokers, the parsedData IS already correctly formatted
                // Show it directly without any additional mapping
                displayData = sampleData;

                // Create an identity mapping for display (don't show confusing descriptions)
                displayMapping = {};
                if (sampleData.length > 0) {
                    Object.keys(sampleData[0]).forEach(key => {
                        displayMapping[key] = key; // Identity mapping
                    });
                }
            } else {
                // For unknown formats or when no parsed data, apply the manual mapping
                displayData = window.applyMappingToData(sampleData, mapping);
                displayMapping = mapping;
            }
        }

        // Display the preview
        displayFilePreviewResults(displayData, displayMapping, sampleData.length, fileData, !mapping);

    } catch (error) {
        console.error('Error generating file preview:', error);
        document.getElementById('preview-content').innerHTML =
            '<div style="text-align: center; padding: 40px; color: #f44336;"><p>Error generating preview: ' + error.message + '</p></div>';
    }
}

/**
 * Display file preview results (unified layout for all preview types)
 * @param {Array} displayData - Data to display
 * @param {Object} mapping - Mapping configuration (or identity mapping for raw view)
 * @param {Number} totalRecords - Total number of records
 * @param {Object} fileData - Original file data
 * @param {Boolean} isRawPreview - Whether this is a raw file preview (no mapping)
 */
function displayFilePreviewResults(displayData, mapping, totalRecords, fileData, isRawPreview = false) {
    // Unified modal content layout
    document.getElementById('preview-content').innerHTML = `
        <div class="section">
            <h3 class="section-title">File Information</h3>
            <div id="preview-file-info"></div>
        </div>

        <div class="section">
            <h3 class="section-title">Sample Data (First & Last Records)</h3>
            <div style="overflow-x: auto;">
                <table class="data-table" id="preview-table">
                    <thead id="preview-table-head"></thead>
                    <tbody id="preview-table-body"></tbody>
                </table>
            </div>
        </div>

        <div class="section">
            <h3 class="section-title">Data Quality</h3>
            <div id="preview-stats"></div>
        </div>
    `;

    // Always display unified file information
    displayUnifiedFileInfo(fileData, displayData, totalRecords);

    // Display sample data table with first/last indicator
    displayPreviewTableWithIndicators(displayData, fileData);

    // Display unified data quality statistics
    displayUnifiedDataQuality(displayData, fileData, totalRecords);
}

// ========== EXPORT FUNCTIONALITY ==========

/**
 * Export preview results as Excel file
 */
async function exportPreviewAsExcel() {
    if (!window.currentMappingFile || Object.keys(window.currentMapping).length === 0) {
        alert('No preview data available to export.');
        return;
    }

    try {
        // Get the full dataset (not just preview sample)
        let fullData = [];

        if (window.currentMappingFile.parsedData && window.currentMappingFile.parsedData.length > 0) {
            // Use already parsed data if available
            fullData = window.currentMappingFile.parsedData;
        } else {
            // Parse the full file with skip rules if available
            const workbook = await ExcelCacheManager.getWorkbook(window.currentMappingFile.file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            let rawData;
            // Use cached analysis from currentMappingFile if available, otherwise fall back to global currentPatternAnalysis
            const analysisToUse = window.currentMappingFile.patternAnalysis || window.currentPatternAnalysis;

            if (analysisToUse && analysisToUse.dataSection) {
                // Use GenericBrokerParser to respect skip rules
                console.log('Export: Using skip rules for full data extraction');
                const tempConfig = {
                    dataStartMethod: 'skip-rows',
                    skipRows: analysisToUse.dataSection.dataStartIndex || 0,
                    skipColumns: analysisToUse.dataSection.startColumnIndex || 0,
                    headerRow: analysisToUse.dataSection.headerRowIndex,
                    columnMapping: {} // Identity mapping for export
                };
                const parser = new window.GenericBrokerParser(tempConfig);
                rawData = await parser.parse(workbook, window.currentMappingFile.name);
            } else {
                // Fallback to simple sheet reading
                console.log('Export: No skip rules available, using simple sheet reading');
                rawData = XLSX.utils.sheet_to_json(worksheet, {
                    cellFormula: false  // Read calculated values instead of formulas
                });
            }

            // Apply broker-specific filtering
            const brokerType = window.currentMappingFile.broker.parser;
            fullData = filterDataForPreview(rawData, brokerType);
        }

        // Apply mapping to full dataset
        const mappedFullData = window.applyMappingToData(fullData, window.currentMapping);

        // Create Excel file
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(mappedFullData);

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Mapped Data');

        // Generate filename
        const brokerName = window.currentMappingFile.broker.name || 'Unknown';
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `${brokerName}_mapped_${timestamp}.xlsx`;

        // Download to preferred folder or fallback to browser download
        const success = await window.downloadExcelToPreferredFolder(wb, filename);

        if (success) {
            alert(`Exported ${mappedFullData.length} records successfully`);
        }

    } catch (error) {
        console.error('Error exporting preview to Excel:', error);
        alert(`Failed to export to Excel: ${error.message}`);
    }
}

// ========== GLOBAL EXPORTS ==========

// Make preview functions globally accessible
window.filterDataForPreview = filterDataForPreview;
window.displayMappingSummary = displayMappingSummary;
window.displayPreviewTable = displayPreviewTable;
window.displayPreviewTableWithIndicators = displayPreviewTableWithIndicators;
window.displayPreviewStats = displayPreviewStats;
window.displayUnifiedDataQuality = displayUnifiedDataQuality;
window.displayUnifiedFileInfo = displayUnifiedFileInfo;
window.displayFileStats = displayFileStats;
window.previewFileData = previewFileData;
window.displayFilePreviewResults = displayFilePreviewResults;
window.exportPreviewAsExcel = exportPreviewAsExcel;