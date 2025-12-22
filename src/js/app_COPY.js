/**
 * Borderellen Converter - Main Application Controller
 * Coordinates all modules and handles UI interactions
 */

// ========== GLOBAL STATE ==========

// Global settings storage
let appSettings = {
    userName: 'User',
    userEmail: '',
    userSignature: '',
    downloadFolder: '',
    downloadFolderHandle: null
};

// Broker mapping state
let currentMappingFile = null;
let currentMapping = {};
let currentPatternAnalysis = null;

// ========== UI STATE MANAGEMENT ==========

function showMainSection() {
    document.getElementById('template-main-section').style.display = 'block';
    document.getElementById('new-template-section').style.display = 'none';
    document.getElementById('edit-template-section').style.display = 'none';
    document.getElementById('column-config-section').style.display = 'none';
}

function showNewTemplateSection() {
    document.getElementById('template-main-section').style.display = 'none';
    document.getElementById('new-template-section').style.display = 'block';
    document.getElementById('edit-template-section').style.display = 'none';
    document.getElementById('column-config-section').style.display = 'none';
}

function showEditTemplateSection() {
    document.getElementById('template-main-section').style.display = 'none';
    document.getElementById('new-template-section').style.display = 'none';
    document.getElementById('edit-template-section').style.display = 'block';
    document.getElementById('column-config-section').style.display = 'block';

    updateTemplateForm();
    updateTemplateDisplay();
}

// ========== BROKER MAPPING FUNCTIONS ==========

/**
 * Apply automatic mapping for known broker types based on their predefined column mappings
 * @param {string} fileId - ID of the selected file
 */
function applyAutoMappingForBrokerType(fileId) {
    const fileData = window.uploadedFiles?.find(f => f.id == fileId);
    if (!fileData || !fileData.broker || fileData.broker.type !== 'built-in') {
        console.log('File not found or not a built-in broker type, skipping auto-mapping');
        return;
    }

    const brokerType = fileData.broker.parser;
    console.log(`Applying auto-mapping for ${brokerType} broker type`);

    // Clear existing mapping first
    currentMapping = {};

    // Get broker-specific mappings
    const brokerMappings = getBrokerAutoMapping(brokerType, fileData.name);

    // Apply the mappings
    Object.assign(currentMapping, brokerMappings);

    console.log('Auto-mapping applied:', currentMapping);
}

/**
 * Get predefined column mappings for specific broker types
 * @param {string} brokerType - The broker type (AON, VGA, BCI, Voogt)
 * @param {string} filename - Original filename for data extraction
 * @returns {Object} Column mapping object
 */
function getBrokerAutoMapping(brokerType, filename) {
    const mappings = {};

    switch (brokerType) {
        case 'AON':
            // Extract period from filename
            const aonMatch = filename.match(/(\d{2}-\d{4})/);
            const aonPeriod = aonMatch ? aonMatch[1] : '';

            mappings['Makelaar'] = 'FIXED:AON';
            mappings['Boekingsperiode'] = `FIXED:${aonPeriod}`;
            mappings['Valuta'] = 'FIXED:EUR';
            mappings['Polisnr makelaar'] = 'PolisNr';
            mappings['Verzekerde'] = 'Verzekerde';
            mappings['Branche'] = 'Branche';
            mappings['Netto'] = 'Netto';
            mappings['Boekdatum tp'] = 'BoekDtm';
            mappings['FactuurDtm'] = 'FactuurDtm';
            mappings['FactuurNr'] = 'FactuurNr';
            mappings['Tekenjaar'] = 'Tekenjaar';
            mappings['Boekingsreden'] = 'FactuurTekst';
            break;

        case 'VGA':
            // Extract period and code from filename
            const vgaMatch = filename.match(/^VGA (\d{2}-\d{4}) (A\d{3})\.xlsx$/i);
            const vgaPeriod = vgaMatch ? vgaMatch[1] : '';
            const vgaCode = vgaMatch ? vgaMatch[2] : '';

            mappings['Makelaar'] = `FIXED:VGA ${vgaCode}`;
            mappings['Boekingsperiode'] = `FIXED:${vgaPeriod}`;
            mappings['Valuta'] = 'FIXED:EUR';
            mappings['Polisnr makelaar'] = 'Polisnummer';
            mappings['Verzekerde'] = 'Naam verzekeringnemer';
            mappings['FactuurNr'] = 'Factuurnummer';
            mappings['Branche'] = 'Branche';
            mappings['Bruto'] = 'Bruto premie EB';
            mappings['Netto'] = 'Netto Maatschappij EB';
            break;

        case 'BCI':
            // Extract period from filename
            const bciMatch = filename.match(/^BCI (\d{4}-Q[1-4])\.xlsx$/i);
            const bciPeriod = bciMatch ? bciMatch[1] : '';

            mappings['Makelaar'] = 'FIXED:BCI';
            mappings['Boekingsperiode'] = `FIXED:${bciPeriod}`;
            mappings['Valuta'] = 'FIXED:EUR';
            // Note: BCI uses complex two-row parsing, so these are conceptual mappings
            // The actual parsing is handled in BCIParser.parse()
            mappings['Polisnr makelaar'] = 'FIXED:From multi-row pattern (XXXX.XX.XX.XXXX)';
            mappings['Branche'] = 'FIXED:From multi-row pattern (text after policy)';
            mappings['Verzekerde'] = 'FIXED:From row 2, column 1';
            mappings['FactuurNr'] = 'FIXED:From row 2, column 2';
            mappings['Provisie'] = 'FIXED:From row 1, column 4';
            mappings['Netto'] = 'FIXED:From row 1, column 7';
            mappings['Bruto'] = 'FIXED:From row 2, column 5';
            break;

        case 'Voogt':
            // Extract period from filename
            const voogtMatch = filename.match(/^Voogt (\d{2}) (\d{4})\.xlsx$/i);
            const voogtPeriod = voogtMatch ? `${voogtMatch[1]} ${voogtMatch[2]}` : '';

            mappings['Makelaar'] = 'FIXED:Voogt';
            mappings['Boekingsperiode'] = `FIXED:${voogtPeriod}`;
            mappings['Valuta'] = 'FIXED:EUR';
            // Note: Voogt uses position-based parsing, so these are conceptual mappings
            mappings['Polisnr makelaar'] = 'FIXED:Column E (position 4)';
            mappings['Verzekerde'] = 'FIXED:Column J (position 9)';
            mappings['Netto'] = 'FIXED:Column N (position 13)';
            mappings['Provisie'] = 'FIXED:Column P (position 15)';
            mappings['Bruto'] = 'FIXED:Column S (position 18)';
            break;

        default:
            console.log(`No auto-mapping defined for broker type: ${brokerType}`);
    }

    return mappings;
}

function updateMappingFileSelector() {
    const selector = document.getElementById('mapping-file-selector');
    selector.innerHTML = '<option value="">Select a file to map...</option>';

    if (window.uploadedFiles && window.uploadedFiles.length > 0) {
        window.uploadedFiles.forEach(fileData => {
            const option = document.createElement('option');
            option.value = fileData.id;
            option.textContent = `${fileData.name} (${fileData.broker.name})`;
            selector.appendChild(option);
        });

        // Auto-select a file if we're currently on the mapping tab and no file is selected
        const activeTab = document.querySelector('.nav-item.active');
        if (activeTab && activeTab.getAttribute('data-tab') === 'mapping' && !selector.value) {
            setTimeout(() => {
                autoSelectFileInMappingTab();
            }, 50); // Small delay to ensure DOM is updated
        }
    }
}

function updateTemplateDropZones() {
    const container = document.getElementById('template-drop-zones');

    if (!window.currentTemplateId || !window.borderellenTemplate || !window.borderellenTemplate.columns || window.borderellenTemplate.columns.length === 0) {
        container.innerHTML = '<div style="text-align: center; padding: 32px; color: #888;"><p>No template selected. Please select a template in Tab 1.</p></div>';
        return;
    }

    container.innerHTML = '';

    window.borderellenTemplate.columns.forEach((column, index) => {
        const dropZone = document.createElement('div');
        dropZone.className = 'drop-zone';
        dropZone.setAttribute('data-column-id', column.id);
        dropZone.setAttribute('data-column-name', column.name);

        // Check if this field has a mapping
        const mappedSource = currentMapping[column.name];
        if (mappedSource) {
            dropZone.classList.add('filled');

            // Check mapping type and add appropriate class
            let mappingType = 'Column';
            let displayValue = mappedSource;
            let typeIndicator = '';

            if (mappedSource.startsWith('FIXED:')) {
                mappingType = 'Fixed Value';
                displayValue = mappedSource.substring(6);
                dropZone.classList.add('fixed-value');
                typeIndicator = '<span class="mapping-type-indicator">F</span>';
            } else if (mappedSource.startsWith('CALC:')) {
                mappingType = 'Calculation';
                displayValue = mappedSource.substring(5);
                dropZone.classList.add('calculation');
                typeIndicator = '<span class="mapping-type-indicator">C</span>';
            } else {
                dropZone.classList.add('column-mapping');
                typeIndicator = '<span class="mapping-type-indicator">M</span>';
            }

            // Get confidence score and color coding
            let confidenceDisplay = '';
            let confidenceColor = '#888';

            if (mappingType === 'Column' && window.currentMappingConfidences && window.currentMappingConfidences[column.name] !== undefined) {
                const confidence = window.currentMappingConfidences[column.name];

                // Color coding based on confidence
                if (confidence >= 80) {
                    confidenceColor = '#4caf50'; // Green for high confidence
                } else if (confidence >= 50) {
                    confidenceColor = '#ff9800'; // Orange for medium confidence
                } else {
                    confidenceColor = '#f44336'; // Red for low confidence
                }

                confidenceDisplay = ` <span style="color: ${confidenceColor}; font-weight: bold;">(${confidence.toFixed(0)}%)</span>`;
            }

            dropZone.innerHTML = `
                ${typeIndicator}
                <span>${index + 1}. ${column.name} → ${displayValue}${confidenceDisplay}</span>
                <small style="display: block; color: #888; font-size: 11px;">(${mappingType}) Right-click for options</small>
            `;
        } else {
            dropZone.innerHTML = `
                <span>${index + 1}. ${column.name}</span>
                <small style="display: block; color: #888; font-size: 11px;">Drop column or click for fixed value</small>
            `;
        }

        container.appendChild(dropZone);
    });

    // Re-attach drop event listeners
    attachDropZoneListeners();
}

async function loadSourceColumns(fileId) {
    const fileData = window.uploadedFiles?.find(f => f.id == fileId);
    if (!fileData) return;

    currentMappingFile = fileData;
    const container = document.getElementById('source-columns');

    try {
        if (fileData.file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            fileData.file.name.endsWith('.xlsx')) {

            // Show loading message
            container.innerHTML = '<div style="text-align: center; padding: 32px; color: #888;"><p>Analyzing file structure...</p></div>';

            try {
                // Use cached pattern analysis if available, otherwise analyze file
                let patternAnalysis = fileData.patternAnalysis;

                if (!patternAnalysis) {
                    console.log('No cached analysis found, running analysis for mapping tab');
                    patternAnalysis = await DataPatternAnalyzer.analyzeFile(fileData.file);
                    // Cache the analysis result
                    fileData.patternAnalysis = patternAnalysis;
                }

                console.log('Using pattern analysis for mapping tab:', patternAnalysis);

                // Store globally for template saving
                currentPatternAnalysis = patternAnalysis;

                if (patternAnalysis.confidence > 0.3 && patternAnalysis.headerAnalysis?.found) {
                    // High confidence - use detected structure
                    await loadSourceColumnsFromAnalysis(fileData.file, patternAnalysis, container);
                } else {
                    // Low confidence or analysis failed - show manual selection interface
                    await showManualColumnSelection(fileData.file, container);
                }

            } catch (analysisError) {
                console.warn('Pattern analysis failed, showing manual selection:', analysisError);
                await showManualColumnSelection(fileData.file, container);
            }

        } else {
            container.innerHTML = '<div style="text-align: center; padding: 32px; color: #888;"><p>PDF files require manual processing</p></div>';
        }
    } catch (error) {
        console.error('Error loading source columns:', error);
        container.innerHTML = '<div style="text-align: center; padding: 32px; color: #f44336;"><p>Error loading file</p></div>';
    }
}

/**
 * Load source columns using pattern analysis results
 */
async function loadSourceColumnsFromAnalysis(file, patternAnalysis, container) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'binary' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const range = XLSX.utils.decode_range(worksheet['!ref']);

                // Extract headers from detected start cell and column range
                const headerRowIndex = patternAnalysis.suggestedHeaderRow;
                const startCol = patternAnalysis.dataSection.startColumnIndex || range.s.c;
                const endCol = patternAnalysis.dataSection.endColumnIndex || range.e.c;
                const headers = [];

                // Check if this is a multi-row header selection
                const isMultiRowHeader = patternAnalysis.manualSelection && patternAnalysis.manualSelection.headerRows > 1;
                const headerRows = isMultiRowHeader ? patternAnalysis.manualSelection.headerRows : 1;

                console.log(`Extracting headers from start cell ${patternAnalysis.dataSection.startCell || 'A1'}, ${headerRows} header row(s), columns ${XLSX.utils.encode_col(startCol)} to ${XLSX.utils.encode_col(endCol)}`);

                if (isMultiRowHeader) {
                    // Multi-row header: Create separate columns for each row
                    for (let row = 0; row < headerRows; row++) {
                        for (let col = startCol; col <= endCol; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex + row, c: col });
                            const cell = worksheet[cellAddress];
                            let headerText = '';

                            if (cell && cell.v) {
                                headerText = cell.v.toString().trim();
                            } else {
                                headerText = `Row${row + 1}_Col${XLSX.utils.encode_col(col)}`;
                            }

                            headers.push(headerText);
                        }
                    }
                } else {
                    // Single row header (existing behavior)
                    for (let col = startCol; col <= endCol; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
                        const cell = worksheet[cellAddress];
                        let headerText = '';

                        if (cell && cell.v) {
                            headerText = cell.v.toString().trim();
                        } else {
                            headerText = `Column ${XLSX.utils.encode_col(col)}`;
                        }

                        headers.push(headerText);
                    }
                }

                // Auto-detect footer keywords in the last 10 rows
                const autoFooterKeyword = detectFooterKeyword(worksheet, range, startCol, endCol);
                if (autoFooterKeyword) {
                    patternAnalysis.autoFooterKeyword = autoFooterKeyword;
                    console.log(`Auto-detected footer keyword: "${autoFooterKeyword}"`);
                }

                // Show confidence and detected structure info
                const startCell = patternAnalysis.dataSection.startCell || `${XLSX.utils.encode_col(startCol)}${headerRowIndex + 1}`;
                // Build header info with optional footer keyword
                let headerInfo = '';
                if (isMultiRowHeader) {
                    headerInfo = `Header Range: ${patternAnalysis.manualSelection.headerRange} (${headerRows} rows), Processing: ${headerRows} rows per record, Data starts: ${XLSX.utils.encode_col(startCol)}${headerRowIndex + headerRows}`;
                    if (patternAnalysis.manualSelection.footerKeyword) {
                        headerInfo += `<br>Footer Keyword: "${patternAnalysis.manualSelection.footerKeyword}"`;
                    }
                } else {
                    headerInfo = `Start Cell: ${startCell}, Data Range: ${XLSX.utils.encode_col(startCol)}${headerRowIndex + 1}:${XLSX.utils.encode_col(endCol)}${patternAnalysis.suggestedDataEnd + 1}`;
                    // Check if auto-detected footer keyword exists
                    if (patternAnalysis.autoFooterKeyword) {
                        headerInfo += `<br>Auto-detected Footer: "${patternAnalysis.autoFooterKeyword}"`;
                    }
                }

                const structureType = isMultiRowHeader ? 'Manual Header & Footer Selection' : 'Automatic Header & Footer Detection';

                const confidenceInfo = `
                    <div style="background: #e8f4f8; padding: 12px; margin-bottom: 16px; border-radius: 4px; border-left: 4px solid #00bcd4; color: black;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <strong style="color: #00bcd4;">${structureType}</strong><br>
                                ${headerInfo}<br>
                                Confidence: ${Math.round(patternAnalysis.confidence * 100)}%
                            </div>
                            <button class="btn btn-secondary" onclick="showManualHeaderSelection()" style="margin-left: 12px;">Manual select header & footer</button>
                        </div>
                    </div>
                `;

                container.innerHTML = confidenceInfo;
                displaySourceColumns(headers);

                // Trigger auto-mapping for high-confidence detections
                if (patternAnalysis.confidence > 0.7) {
                    console.log('High confidence detection - triggering auto-mapping');
                    setTimeout(() => {
                        generateAutoMappingSuggestions();
                    }, 200); // Small delay to ensure columns are loaded
                }

                resolve();

            } catch (error) {
                console.error('Error extracting columns from analysis:', error);
                reject(error);
            }
        };
        reader.readAsBinaryString(file);
    });
}

/**
 * Show manual grid selection interface when analysis fails
 */
async function showManualColumnSelection(file, container) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'binary' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const range = XLSX.utils.decode_range(worksheet['!ref']);

                // Show manual selection grid (rows 1-50, columns A-AX)
                const maxRow = Math.min(50, range.e.r + 1);
                const maxCol = Math.min(50, range.e.c + 1);

                let gridHtml = `
                    <div style="background: #e8f4f8; padding: 12px; margin-bottom: 16px; border-radius: 4px; border-left: 4px solid #00bcd4; color: black;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <strong style="color: #00bcd4;">If your data table does not start in A1, select the header here so the app will skips rows/columns to start at the correct position. You can even select multiple rows so the app would process sets of N rows as 1 data record.</strong>
                            </div>
                            <button class="btn btn-secondary" onclick="showManualHeaderSelection()" style="margin-left: 12px;">Manual select header & footer</button>
                        </div>
                    </div>
                    <div style="max-height: 400px; overflow: auto; border: 1px solid #ddd;">
                        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                            <thead>
                                <tr style="background: #f5f5f5; position: sticky; top: 0;">
                                    <th style="padding: 4px; border: 1px solid #ddd; width: 40px;">#</th>
                `;

                // Column headers (A, B, C, ...)
                for (let col = 0; col < maxCol; col++) {
                    gridHtml += `<th style="padding: 4px; border: 1px solid #ddd; min-width: 80px;">${XLSX.utils.encode_col(col)}</th>`;
                }
                gridHtml += '</tr></thead><tbody>';

                // Grid rows
                for (let row = 0; row < maxRow; row++) {
                    gridHtml += `<tr data-row="${row}" style="cursor: pointer;" onmouseover="this.style.backgroundColor='#f0f8ff'" onmouseout="this.style.backgroundColor=''">`;
                    gridHtml += `<td style="padding: 4px; border: 1px solid #ddd; background: #f9f9f9; font-weight: bold;">${row + 1}</td>`;

                    for (let col = 0; col < maxCol; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];
                        const cellValue = cell && cell.v ? cell.v.toString() : '';
                        const displayValue = cellValue.length > 10 ? cellValue.substring(0, 10) + '...' : cellValue;

                        gridHtml += `<td style="padding: 4px; border: 1px solid #ddd; cursor: pointer;"
                                        data-row="${row}" data-col="${col}"
                                        onclick="selectHeaderRow(${row})"
                                        title="${cellValue}">${displayValue}</td>`;
                    }
                    gridHtml += '</tr>';
                }

                gridHtml += `
                        </tbody>
                    </table>
                    </div>
                    <div style="margin-top: 12px;">
                        <button id="use-selected-structure" class="btn btn-primary" disabled>Use Selected Structure</button>
                        <span id="selection-status" style="margin-left: 12px; color: #666;">Click on a row to select it as headers</span>
                    </div>
                `;

                container.innerHTML = gridHtml;

                // Add global selection handler
                window.selectedHeaderRow = null;
                window.selectHeaderRow = function(rowIndex) {
                    // Remove previous selection
                    document.querySelectorAll('tr[data-row]').forEach(tr => {
                        tr.style.backgroundColor = '';
                    });

                    // Highlight selected row
                    const selectedRow = document.querySelector(`tr[data-row="${rowIndex}"]`);
                    if (selectedRow) {
                        selectedRow.style.backgroundColor = '#d4edda';
                        window.selectedHeaderRow = rowIndex;

                        document.getElementById('use-selected-structure').disabled = false;
                        document.getElementById('selection-status').textContent = `Header row ${rowIndex + 1} selected`;
                    }
                };

                // Add button handler
                document.getElementById('use-selected-structure').onclick = function() {
                    if (window.selectedHeaderRow !== null) {
                        loadHeadersFromManualSelection(worksheet, window.selectedHeaderRow, range, container);
                    }
                };

                resolve();

            } catch (error) {
                console.error('Error creating manual selection grid:', error);
                reject(error);
            }
        };
        reader.readAsBinaryString(file);
    });
}

/**
 * Show enhanced manual header selection modal with range selection
 */
function showManualHeaderSelection() {
    if (!currentMappingFile) {
        alert('Please select a file first.');
        return;
    }

    // Create modal HTML
    const modalHtml = `
        <div id="manual-header-modal" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); z-index: 1000; display: flex; justify-content: center; align-items: center;">
            <div style="background: #2a2a2a; border-radius: 8px; padding: 24px; max-width: 90vw; max-height: 90vh; overflow: auto; color: white;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                    <h3 style="margin: 0; color: #00bcd4;">Manual Header & Footer Selection</h3>
                    <button onclick="closeManualHeaderSelection()" style="background: none; border: none; color: #888; font-size: 24px; cursor: pointer;">&times;</button>
                </div>
                <div id="manual-selection-instructions" style="background: #333; padding: 12px; border-radius: 4px; margin-bottom: 16px; color: #ccc;">
                    <strong>Instructions:</strong> Click and drag to select your header range. You can select multiple rows and columns to define exactly where your data starts.
                </div>
                <div id="manual-selection-grid" style="border: 1px solid #555; border-radius: 4px; overflow: auto; max-height: 40vh;">
                    <!-- Header grid will be populated here -->
                </div>
                <div style="margin: 16px 0;">
                    <h4 style="margin: 0 0 8px 0; color: #00bcd4;">Footer Detection - Last 10 Rows</h4>
                    <div style="background: #333; padding: 8px; border-radius: 4px; margin-bottom: 8px; color: #ccc; font-size: 12px;">
                        Click on a cell containing a keyword (like "Aantal prolongatie") to stop processing when encountered.
                    </div>
                    <div id="footer-detection-grid" style="border: 1px solid #555; border-radius: 4px; overflow: auto; max-height: 20vh; background: #2a2a2a;">
                        <!-- Footer grid will be populated here -->
                    </div>
                    <div id="footer-keyword-status" style="margin-top: 8px; color: #888; font-size: 12px;">
                        No footer keyword selected
                    </div>
                </div>
                <div style="margin-top: 16px; display: flex; justify-content: space-between; align-items: center;">
                    <div id="selection-status" style="color: #888;">Select a range to continue</div>
                    <div>
                        <button class="btn btn-secondary" onclick="closeManualHeaderSelection()" style="margin-right: 8px;">Cancel</button>
                        <button id="apply-header-selection" class="btn" onclick="applyHeaderSelection()" disabled>Apply Selection</button>
                    </div>
                </div>
            </div>
        </div>
    `;

    // Add modal to page
    document.body.insertAdjacentHTML('beforeend', modalHtml);

    // Load grid content
    loadHeaderSelectionGrid();
}

/**
 * Close manual header selection modal
 */
function closeManualHeaderSelection() {
    const modal = document.getElementById('manual-header-modal');
    if (modal) {
        modal.remove();
    }
    // Reset selection state
    window.headerSelectionState = null;
}

/**
 * Load the grid for header selection
 */
async function loadHeaderSelectionGrid() {
    const gridContainer = document.getElementById('manual-selection-grid');

    try {
        const workbook = await readExcelFile(currentMappingFile.file);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        // Show grid (rows 1-30, columns A-AZ for better visibility)
        const maxRow = Math.min(30, range.e.r + 1);
        const maxCol = Math.min(26, range.e.c + 1);

        let gridHtml = `
            <table id="header-selection-table" style="width: 100%; border-collapse: collapse; font-size: 11px; background: #333;">
                <thead>
                    <tr style="background: #444; position: sticky; top: 0;">
                        <th style="padding: 4px; border: 1px solid #555; width: 40px; color: #888;">#</th>
        `;

        // Column headers (A, B, C, ...)
        for (let col = 0; col < maxCol; col++) {
            gridHtml += `<th style="padding: 4px; border: 1px solid #555; min-width: 60px; color: #888;">${XLSX.utils.encode_col(col)}</th>`;
        }
        gridHtml += '</tr></thead><tbody>';

        // Grid rows
        for (let row = 0; row < maxRow; row++) {
            gridHtml += `<tr data-row="${row}">`;
            gridHtml += `<td style="padding: 4px; border: 1px solid #555; background: #444; font-weight: bold; color: #888;">${row + 1}</td>`;

            for (let col = 0; col < maxCol; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                const cellValue = cell && cell.v ? cell.v.toString() : '';
                const displayValue = cellValue.length > 8 ? cellValue.substring(0, 8) + '...' : cellValue;

                gridHtml += `<td class="header-cell" style="padding: 4px; border: 1px solid #555; cursor: pointer; color: white; background: #333;"
                                data-row="${row}" data-col="${col}"
                                onmousedown="startRangeSelection(${row}, ${col})"
                                onmouseover="updateRangeSelection(${row}, ${col})"
                                onmouseup="endRangeSelection()"
                                title="${cellValue}">${displayValue}</td>`;
            }
            gridHtml += '</tr>';
        }

        gridHtml += '</tbody></table>';
        gridContainer.innerHTML = gridHtml;

        // Initialize selection state
        window.headerSelectionState = {
            isSelecting: false,
            startRow: null,
            startCol: null,
            endRow: null,
            endCol: null,
            footerKeyword: null
        };

        // Load footer detection grid
        loadFooterDetectionGrid(worksheet, range);

        // Prevent text selection during drag
        document.addEventListener('selectstart', preventDefault, { once: false });

    } catch (error) {
        console.error('Error loading header selection grid:', error);
        gridContainer.innerHTML = '<div style="text-align: center; padding: 40px; color: #f44336;">Error loading file</div>';
    }
}

/**
 * Load the footer detection grid with last 10 rows
 */
function loadFooterDetectionGrid(worksheet, range) {
    const footerContainer = document.getElementById('footer-detection-grid');

    try {
        const totalRows = range.e.r + 1;
        const maxCol = Math.min(26, range.e.c + 1); // Show up to column Z
        const startRow = Math.max(0, totalRows - 10); // Last 10 rows

        let footerGridHtml = `
            <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                <thead>
                    <tr style="background: #404040; position: sticky; top: 0;">
                        <th style="padding: 2px 4px; border: 1px solid #555; width: 40px; color: #888; font-size: 10px;">Row</th>
        `;

        // Column headers (A, B, C, ...)
        for (let col = 0; col < maxCol; col++) {
            footerGridHtml += `<th style="padding: 2px 4px; border: 1px solid #555; min-width: 50px; color: #888; font-size: 10px;">${XLSX.utils.encode_col(col)}</th>`;
        }
        footerGridHtml += '</tr></thead><tbody>';

        // Footer rows
        for (let row = startRow; row < totalRows; row++) {
            const excelRowNumber = row + 1; // Convert to 1-based for display
            footerGridHtml += `<tr>`;
            footerGridHtml += `<td style="padding: 2px 4px; border: 1px solid #555; background: #404040; font-weight: bold; color: #888; font-size: 10px;">${excelRowNumber}</td>`;

            for (let col = 0; col < maxCol; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                const cellValue = cell && cell.v ? cell.v.toString() : '';
                const displayValue = cellValue.length > 8 ? cellValue.substring(0, 8) + '...' : cellValue;

                footerGridHtml += `<td class="footer-cell" style="padding: 2px 4px; border: 1px solid #555; cursor: pointer; color: white; background: #2a2a2a; font-size: 10px;"
                                    data-row="${row}" data-col="${col}"
                                    onclick="selectFooterKeyword(${row}, ${col}, '${cellValue.replace(/'/g, "\\'")}')"
                                    title="Click to use '${cellValue}' as footer keyword">${displayValue}</td>`;
            }
            footerGridHtml += '</tr>';
        }

        footerGridHtml += '</tbody></table>';
        footerContainer.innerHTML = footerGridHtml;

    } catch (error) {
        console.error('Error loading footer detection grid:', error);
        footerContainer.innerHTML = '<div style="text-align: center; padding: 20px; color: #f44336; font-size: 12px;">Error loading footer section</div>';
    }
}

/**
 * Select footer keyword from clicked cell
 */
function selectFooterKeyword(row, col, cellValue) {
    const state = window.headerSelectionState;
    const statusEl = document.getElementById('footer-keyword-status');

    // Clear previous footer selection
    document.querySelectorAll('.footer-cell').forEach(cell => {
        cell.style.background = '#2a2a2a';
        cell.style.color = 'white';
    });

    // Highlight selected cell
    const selectedCell = document.querySelector(`.footer-cell[data-row="${row}"][data-col="${col}"]`);
    if (selectedCell) {
        selectedCell.style.background = '#00bcd4';
        selectedCell.style.color = '#000';
    }

    // Store keyword
    state.footerKeyword = cellValue.trim();

    // Update status
    if (state.footerKeyword) {
        statusEl.innerHTML = `<span style="color: #00bcd4;">Footer keyword selected:</span> "${state.footerKeyword}"`;
        statusEl.style.color = '#00bcd4';
    } else {
        statusEl.textContent = 'No footer keyword selected';
        statusEl.style.color = '#888';
    }

    console.log(`Footer keyword selected: "${state.footerKeyword}" from cell ${XLSX.utils.encode_cell({ r: row, c: col })}`);
}

/**
 * Start range selection
 */
function startRangeSelection(row, col) {
    const state = window.headerSelectionState;
    state.isSelecting = true;
    state.startRow = row;
    state.startCol = col;
    state.endRow = row;
    state.endCol = col;

    // Clear previous selection
    clearRangeSelection();

    // Highlight current cell
    updateRangeVisual();
    updateSelectionStatus();
}

/**
 * Update range selection during drag
 */
function updateRangeSelection(row, col) {
    const state = window.headerSelectionState;
    if (!state.isSelecting) return;

    state.endRow = row;
    state.endCol = col;

    updateRangeVisual();
    updateSelectionStatus();
}

/**
 * End range selection
 */
function endRangeSelection() {
    const state = window.headerSelectionState;
    if (!state.isSelecting) return;

    state.isSelecting = false;

    // Enable apply button
    const applyBtn = document.getElementById('apply-header-selection');
    applyBtn.disabled = false;
}

/**
 * Clear range selection visual
 */
function clearRangeSelection() {
    const cells = document.querySelectorAll('.header-cell');
    cells.forEach(cell => {
        cell.style.background = '#333';
    });
}

/**
 * Update visual selection
 */
function updateRangeVisual() {
    const state = window.headerSelectionState;
    if (state.startRow === null) return;

    // Clear previous selection
    clearRangeSelection();

    // Calculate selection bounds
    const minRow = Math.min(state.startRow, state.endRow);
    const maxRow = Math.max(state.startRow, state.endRow);
    const minCol = Math.min(state.startCol, state.endCol);
    const maxCol = Math.max(state.startCol, state.endCol);

    // Highlight selected range
    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const cell = document.querySelector(`.header-cell[data-row="${row}"][data-col="${col}"]`);
            if (cell) {
                cell.style.background = '#00bcd4';
                cell.style.color = '#000';
            }
        }
    }
}

/**
 * Update selection status text
 */
function updateSelectionStatus() {
    const state = window.headerSelectionState;
    const statusEl = document.getElementById('selection-status');

    if (state.startRow === null) {
        statusEl.textContent = 'Select a range to continue';
        return;
    }

    const minRow = Math.min(state.startRow, state.endRow);
    const maxRow = Math.max(state.startRow, state.endRow);
    const minCol = Math.min(state.startCol, state.endCol);
    const maxCol = Math.max(state.startCol, state.endCol);

    const startCell = XLSX.utils.encode_cell({ r: minRow, c: minCol });
    const endCell = XLSX.utils.encode_cell({ r: maxRow, c: maxCol });
    const rowCount = maxRow - minRow + 1;
    const colCount = maxCol - minCol + 1;

    statusEl.textContent = `Selected: ${startCell}:${endCell} (${rowCount} rows × ${colCount} columns)`;
}

/**
 * Apply header selection
 */
function applyHeaderSelection() {
    const state = window.headerSelectionState;

    if (!state || state.startRow === null) {
        alert('Please select a range first.');
        return;
    }

    const minRow = Math.min(state.startRow, state.endRow);
    const maxRow = Math.max(state.startRow, state.endRow);
    const minCol = Math.min(state.startCol, state.endCol);
    const maxCol = Math.max(state.startCol, state.endCol);

    // Store the selection in currentPatternAnalysis for use by other functions
    if (!currentPatternAnalysis) {
        currentPatternAnalysis = {};
    }
    if (!currentPatternAnalysis.dataSection) {
        currentPatternAnalysis.dataSection = {};
    }

    // Update pattern analysis with manual selection
    currentPatternAnalysis.dataSection.startCell = XLSX.utils.encode_cell({ r: minRow, c: minCol });
    currentPatternAnalysis.dataSection.headerRowIndex = minRow;
    currentPatternAnalysis.dataSection.dataStartIndex = maxRow + 1; // Data starts after header
    currentPatternAnalysis.dataSection.startColumnIndex = minCol;
    currentPatternAnalysis.dataSection.endColumnIndex = maxCol;
    currentPatternAnalysis.suggestedHeaderRow = minRow;
    currentPatternAnalysis.confidence = 1.0; // Manual selection = 100% confidence

    // Store additional info for parsingConfig
    currentPatternAnalysis.manualSelection = {
        headerRows: maxRow - minRow + 1,
        headerColumns: maxCol - minCol + 1,
        headerRange: `${XLSX.utils.encode_cell({ r: minRow, c: minCol })}:${XLSX.utils.encode_cell({ r: maxRow, c: maxCol })}`,
        footerKeyword: state.footerKeyword || null
    };

    // Close modal
    closeManualHeaderSelection();

    // Reload source columns with new selection
    const container = document.getElementById('source-columns');
    loadSourceColumnsFromAnalysis(currentMappingFile.file, currentPatternAnalysis, container);

    console.log('Applied manual header selection:', currentPatternAnalysis);
}

/**
 * Auto-detect common footer keywords in the last rows of the worksheet
 */
function detectFooterKeyword(worksheet, range, startCol, endCol) {
    const commonFooterKeywords = [
        'totaal', 'total', 'aantal', 'sum', 'subtotal', 'subtotaal',
        'grand total', 'eindtotaal', 'saldo', 'balance', 'résumé',
        'prolongatie', 'summary', 'samenvatting'
    ];

    try {
        const totalRows = range.e.r + 1;
        const searchRows = Math.max(0, totalRows - 10); // Check last 10 rows

        // Search from bottom up for better accuracy
        for (let row = totalRows - 1; row >= searchRows; row--) {
            for (let col = startCol; col <= endCol; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                    const cellText = cell.v.toString().toLowerCase().trim();

                    // Check for exact matches or partial matches
                    for (const keyword of commonFooterKeywords) {
                        if (cellText.includes(keyword) && cellText.length <= 50) {
                            console.log(`Auto-detected footer keyword "${cell.v}" at ${cellAddress}`);
                            return cell.v.toString().trim();
                        }
                    }
                }
            }
        }
    } catch (error) {
        console.error('Error during auto footer detection:', error);
    }

    return null;
}

/**
 * Context Menu System for Template Fields
 */
class ContextMenu {
    constructor() {
        this.menu = null;
        this.currentTarget = null;
        this.isVisible = false;

        // Bind methods
        this.hide = this.hide.bind(this);
        this.handleClickOutside = this.handleClickOutside.bind(this);
    }

    show(x, y, items, target) {
        this.hide(); // Hide any existing menu

        this.currentTarget = target;
        this.menu = this.createMenu(items);
        document.body.appendChild(this.menu);

        // Position the menu
        this.positionMenu(x, y);

        // Show with animation
        requestAnimationFrame(() => {
            this.menu.classList.add('show');
            this.isVisible = true;
        });

        // Add event listeners
        document.addEventListener('click', this.handleClickOutside);
        document.addEventListener('contextmenu', this.handleClickOutside);
        document.addEventListener('keydown', this.handleKeyDown);
    }

    hide() {
        if (!this.menu) return;

        this.menu.classList.remove('show');
        setTimeout(() => {
            if (this.menu && this.menu.parentNode) {
                this.menu.parentNode.removeChild(this.menu);
            }
            this.menu = null;
            this.currentTarget = null;
            this.isVisible = false;
        }, 150);

        // Remove event listeners
        document.removeEventListener('click', this.handleClickOutside);
        document.removeEventListener('contextmenu', this.handleClickOutside);
        document.removeEventListener('keydown', this.handleKeyDown);
    }

    createMenu(items) {
        const menu = document.createElement('div');
        menu.className = 'context-menu';

        items.forEach(item => {
            if (item.separator) {
                const separator = document.createElement('div');
                separator.className = 'context-menu-separator';
                menu.appendChild(separator);
            } else {
                const menuItem = document.createElement('div');
                menuItem.className = `context-menu-item ${item.disabled ? 'disabled' : ''}`;
                menuItem.innerHTML = `
                    <span class="icon">${item.icon || ''}</span>
                    ${item.text}
                `;

                if (!item.disabled && item.action) {
                    menuItem.addEventListener('click', (e) => {
                        e.stopPropagation();
                        item.action(this.currentTarget);
                        this.hide();
                    });
                }

                menu.appendChild(menuItem);
            }
        });

        return menu;
    }

    positionMenu(x, y) {
        if (!this.menu) return;

        const menuRect = this.menu.getBoundingClientRect();
        const windowWidth = window.innerWidth;
        const windowHeight = window.innerHeight;

        // Adjust X position if menu would go off-screen
        let finalX = x;
        if (x + menuRect.width > windowWidth) {
            finalX = x - menuRect.width;
        }

        // Adjust Y position if menu would go off-screen
        let finalY = y;
        if (y + menuRect.height > windowHeight) {
            finalY = y - menuRect.height;
        }

        this.menu.style.left = Math.max(0, finalX) + 'px';
        this.menu.style.top = Math.max(0, finalY) + 'px';
    }

    handleClickOutside(e) {
        if (!this.menu || this.menu.contains(e.target)) return;
        this.hide();
    }

    handleKeyDown = (e) => {
        if (e.key === 'Escape') {
            this.hide();
        }
    }
}

// Global context menu instance
window.templateContextMenu = new ContextMenu();

/**
 * Prevent text selection during drag
 */
function preventDefault(e) {
    e.preventDefault();
}

/**
 * Load headers from manual selection
 */
function loadHeadersFromManualSelection(worksheet, headerRowIndex, range, container) {
    const headers = [];

    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
        const cell = worksheet[cellAddress];
        if (cell && cell.v) {
            headers.push(cell.v.toString().trim());
        } else {
            headers.push(`Column ${XLSX.utils.encode_col(col)}`);
        }
    }

    const manualInfo = `
        <div style="background: #d1ecf1; padding: 12px; margin-bottom: 16px; border-radius: 4px; border-left: 4px solid #17a2b8;">
            <strong>Manual Selection</strong><br>
            Using row ${headerRowIndex + 1} as headers
        </div>
    `;

    container.innerHTML = manualInfo;
    displaySourceColumns(headers);
}

function displaySourceColumns(headers) {
    const container = document.getElementById('source-columns');

    // Create a container for the column items (append to existing content)
    const columnsContainer = document.createElement('div');

    headers.forEach(header => {
        const columnItem = document.createElement('div');
        columnItem.className = 'column-item';
        columnItem.draggable = true;
        columnItem.setAttribute('data-column-name', header);

        // Guess data type based on column name
        const dataType = window.guessColumnType ? window.guessColumnType(header) : 'text';

        columnItem.innerHTML = `
            <strong>${header}</strong>
            <span style="color: #888;">(${dataType})</span>
        `;

        // Add drag event listeners
        columnItem.addEventListener('dragstart', handleDragStart);
        columnItem.addEventListener('dragend', handleDragEnd);

        columnsContainer.appendChild(columnItem);
    });

    // Append the columns container to the main container
    container.appendChild(columnsContainer);
}

function handleDragStart(e) {
    e.target.classList.add('dragging');
    e.dataTransfer.setData('text/plain', e.target.getAttribute('data-column-name'));
}

function handleDragEnd(e) {
    e.target.classList.remove('dragging');
}

function attachDropZoneListeners() {
    const dropZones = document.querySelectorAll('#template-drop-zones .drop-zone');

    dropZones.forEach(zone => {
        zone.addEventListener('dragover', handleDragOver);
        zone.addEventListener('drop', handleDrop);
        zone.addEventListener('dragleave', handleDragLeave);
        zone.addEventListener('click', handleDropZoneRightClick);
    });
}

function handleDragOver(e) {
    e.preventDefault();
    e.target.closest('.drop-zone').classList.add('drag-over');
}

function handleDragLeave(e) {
    e.target.closest('.drop-zone').classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    const dropZone = e.target.closest('.drop-zone');
    dropZone.classList.remove('drag-over');

    const sourceColumn = e.dataTransfer.getData('text/plain');
    const targetColumn = dropZone.getAttribute('data-column-name');

    if (sourceColumn && targetColumn) {
        // Update mapping
        currentMapping[targetColumn] = sourceColumn;

        // Update visual display using the enhanced template system
        updateTemplateDropZones();

        console.log(`Mapped: ${targetColumn} ← ${sourceColumn}`);
    }
}

function handleDropZoneClick(e) {
    e.preventDefault();
    const dropZone = e.currentTarget;
    const targetColumn = dropZone.getAttribute('data-column-name');

    // Check if it already has a mapping
    const existingMapping = currentMapping[targetColumn];
    let defaultValue = '';

    if (existingMapping && existingMapping.startsWith('FIXED:')) {
        // If it's a fixed value, use current value as default
        defaultValue = existingMapping.substring(6);
    }

    // Prompt user for fixed value
    const fixedValue = prompt(
        `Enter a fixed value for "${targetColumn}" (e.g., EUR, Netherlands, etc.):`,
        defaultValue
    );

    if (fixedValue !== null && fixedValue.trim() !== '') {
        // Store as fixed value with prefix
        currentMapping[targetColumn] = `FIXED:${fixedValue.trim()}`;

        // Update visual display
        updateTemplateDropZones();

        console.log(`Fixed value set: ${targetColumn} = ${fixedValue.trim()}`);
    }
}

function handleDropZoneRightClick(e) {
    e.stopPropagation();

    const dropZone = e.currentTarget;
    const targetColumn = dropZone.getAttribute('data-column-name');
    const existingMapping = currentMapping[targetColumn];

    // Build context menu items based on current state
    const menuItems = [];

    // Remove mapping option (only if mapping exists)
    if (existingMapping) {
        const mappingType = getMappingType(existingMapping);
        menuItems.push({
            icon: '🗑️',
            text: `Remove ${mappingType}`,
            action: (target) => removeMappingAction(target)
        });
        menuItems.push({ separator: true });
    }

    // Add fixed string option
    menuItems.push({
        icon: '📝',
        text: 'Add fixed string',
        action: (target) => addFixedStringAction(target)
    });

    // Add calculation option
    menuItems.push({
        icon: '🧮',
        text: 'Add calculation',
        action: (target) => addCalculationAction(target)
    });

    // Add separator and cancel
    menuItems.push({ separator: true });
    menuItems.push({
        icon: '❌',
        text: 'Cancel',
        action: () => {} // Just closes the menu
    });

    // Show context menu
    window.templateContextMenu.show(e.clientX, e.clientY, menuItems, dropZone);
}

/**
 * Get human-readable mapping type
 */
function getMappingType(mapping) {
    if (mapping.startsWith('FIXED:')) return 'fixed value';
    if (mapping.startsWith('CALC:')) return 'calculation';
    return 'column mapping';
}

/**
 * Remove mapping action
 */
function removeMappingAction(dropZone) {
    const targetColumn = dropZone.getAttribute('data-column-name');

    // Remove mapping
    delete currentMapping[targetColumn];

    // Update visual display
    updateTemplateDropZones();

    console.log(`Removed mapping for ${targetColumn}`);
}

/**
 * Add fixed string action
 */
function addFixedStringAction(dropZone) {
    const targetColumn = dropZone.getAttribute('data-column-name');
    const existingMapping = currentMapping[targetColumn];

    // Show fixed string input modal
    showFixedStringModal(targetColumn, existingMapping);
}

/**
 * Add calculation action
 */
function addCalculationAction(dropZone) {
    const targetColumn = dropZone.getAttribute('data-column-name');
    const existingMapping = currentMapping[targetColumn];

    // Show calculation modal
    showCalculationModal(targetColumn, existingMapping);
}

/**
 * Show fixed string input modal (replaces prompt)
 */
function showFixedStringModal(targetColumn, existingMapping) {
    // Get default value from existing mapping
    let defaultValue = '';
    if (existingMapping && existingMapping.startsWith('FIXED:')) {
        defaultValue = existingMapping.substring(6);
    }

    const modalHtml = `
        <div id="fixed-string-modal" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); z-index: 1000; display: flex; justify-content: center; align-items: center;">
            <div style="background: #2a2a2a; border-radius: 8px; padding: 24px; max-width: 600px; width: 90%; color: white; border: 1px solid #555;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                    <h3 style="margin: 0; color: #00bcd4;">Set Fixed Value</h3>
                    <button onclick="closeFixedStringModal()" style="background: none; border: none; color: #888; font-size: 24px; cursor: pointer;">&times;</button>
                </div>
                <div style="margin-bottom: 16px;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Template Field:</label>
                    <div style="background: #333; padding: 8px 12px; border-radius: 4px; color: #00bcd4; font-weight: bold;">${targetColumn}</div>
                </div>
                <div style="margin-bottom: 16px;">
                    <label for="fixed-value-input" style="display: block; margin-bottom: 8px; color: #ccc;">Fixed Value:</label>
                    <input type="text" id="fixed-value-input" value="${defaultValue}" style="width: 100%; background: #333; border: 1px solid #555; border-radius: 4px; padding: 10px; color: white; font-size: 14px;" placeholder="e.g., EUR, Netherlands, AON..." />
                </div>
                <div style="display: flex; justify-content: flex-end; gap: 8px;">
                    <button class="btn btn-secondary" onclick="closeFixedStringModal()">Cancel</button>
                    <button class="btn" onclick="applyFixedString('${targetColumn}')">Apply</button>
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    document.getElementById('fixed-value-input').focus();
}

/**
 * Close fixed string modal
 */
function closeFixedStringModal() {
    const modal = document.getElementById('fixed-string-modal');
    if (modal) {
        modal.remove();
    }
}

/**
 * Apply fixed string value
 */
function applyFixedString(targetColumn) {
    const input = document.getElementById('fixed-value-input');
    const fixedValue = input.value.trim();

    if (fixedValue !== '') {
        // Store as fixed value with prefix
        currentMapping[targetColumn] = `FIXED:${fixedValue}`;

        // Update visual display
        updateTemplateDropZones();

        console.log(`Fixed value set: ${targetColumn} = ${fixedValue}`);
    }

    closeFixedStringModal();
}

/**
 * Show calculation modal
 */
function showCalculationModal(targetColumn, existingMapping) {
    // Get existing calculation if available
    let existingFormula = '';
    if (existingMapping && existingMapping.startsWith('CALC:')) {
        existingFormula = existingMapping.substring(5);
    }

    // Get available source columns
    const sourceColumns = getAvailableSourceColumns();
    const sourceOptions = sourceColumns.map(col => `<option value="${col}">${col}</option>`).join('');

    const modalHtml = `
        <div id="calculation-modal" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); z-index: 1000; display: flex; justify-content: center; align-items: center;">
            <div style="background: #2a2a2a; border-radius: 8px; padding: 24px; max-width: 700px; width: 90%; color: white; border: 1px solid #555;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                    <h3 style="margin: 0; color: #00bcd4;">Create Calculation</h3>
                    <button onclick="closeCalculationModal()" style="background: none; border: none; color: #888; font-size: 24px; cursor: pointer;">&times;</button>
                </div>

                <div style="margin-bottom: 16px;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Template Field:</label>
                    <div style="background: #333; padding: 8px 12px; border-radius: 4px; color: #00bcd4; font-weight: bold;">${targetColumn}</div>
                </div>

                <div style="margin-bottom: 16px;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Function (Optional):</label>
                    <select id="calc-function" onchange="updateCalculationPreview()" style="background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white; width: 100%;">
                        <option value="">No Function</option>

                        <optgroup label="📝 Text Extraction">
                            <option value="LEFT">LEFT - Extract from start</option>
                            <option value="RIGHT">RIGHT - Extract from end</option>
                            <option value="MID">MID - Extract middle portion</option>
                            <option value="REGEX">REGEX - Pattern extraction</option>
                            <option value="SPLIT">SPLIT - Split by delimiter</option>
                        </optgroup>

                        <optgroup label="✏️ Text Manipulation">
                            <option value="TRIM">TRIM - Remove spaces</option>
                            <option value="UPPER">UPPER - Convert to uppercase</option>
                            <option value="LOWER">LOWER - Convert to lowercase</option>
                            <option value="REPLACE">REPLACE - Replace text</option>
                            <option value="CONCAT">CONCAT - Join text</option>
                        </optgroup>

                        <optgroup label="🧠 Logic Functions">
                            <option value="IF">IF - Conditional logic</option>
                            <option value="AND">AND - All conditions true</option>
                            <option value="OR">OR - Any condition true</option>
                            <option value="NOT">NOT - Negate condition</option>
                        </optgroup>

                        <optgroup label="🔍 Text Analysis">
                            <option value="CONTAINS">CONTAINS - Text contains</option>
                            <option value="STARTSWITH">STARTSWITH - Starts with</option>
                            <option value="ENDSWITH">ENDSWITH - Ends with</option>
                            <option value="LENGTH">LENGTH - Text length</option>
                            <option value="ISEMPTY">ISEMPTY - Check if empty</option>
                        </optgroup>

                        <optgroup label="🔢 Math Functions">
                            <option value="ROUND">ROUND - Round to decimals</option>
                            <option value="ABS">ABS - Absolute value</option>
                            <option value="MIN">MIN - Minimum value</option>
                            <option value="MAX">MAX - Maximum value</option>
                            <option value="CEILING">CEILING - Round up</option>
                            <option value="FLOOR">FLOOR - Round down</option>
                        </optgroup>
                    </select>
                </div>

                <div id="function-params" style="margin-bottom: 16px; display: none;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Function Parameters:</label>
                    <div id="function-params-container">
                        <!-- Dynamic content based on selected function -->
                    </div>
                </div>

                <div style="background: #333; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
                    <h4 style="margin: 0 0 12px 0; color: #00bcd4;">Formula Builder</h4>

                    <div style="display: grid; grid-template-columns: 1fr auto 1fr; gap: 8px; align-items: center; margin-bottom: 12px;">
                        <select id="calc-column-a" style="background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white;">
                            <option value="">Select Column</option>
                            ${sourceOptions}
                        </select>

                        <select id="calc-operator" style="background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white; width: 60px;">
                            <option value="+">+</option>
                            <option value="-">-</option>
                            <option value="*">×</option>
                            <option value="/">/</option>
                        </select>

                        <div style="display: flex; gap: 4px;">
                            <select id="calc-value-type" onchange="toggleCalculationValueType()" style="background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white; width: 80px;">
                                <option value="column">Column</option>
                                <option value="number">Number</option>
                            </select>
                            <select id="calc-column-b" style="background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white; flex: 1;">
                                <option value="">Select Column</option>
                                ${sourceOptions}
                            </select>
                            <input type="number" id="calc-number" step="any" placeholder="1.21" style="background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white; flex: 1; display: none;">
                        </div>
                    </div>
                </div>

                <div style="margin-bottom: 16px;">
                    <label for="calc-formula-preview" style="display: block; margin-bottom: 8px; color: #ccc;">Formula Preview:</label>
                    <div id="calc-formula-preview" style="background: #1a1a1a; border: 1px solid #555; border-radius: 4px; padding: 10px; font-family: 'Courier New', monospace; color: #00bcd4; min-height: 20px;">
                        ${existingFormula || 'Select columns and operator to build formula'}
                    </div>
                </div>

                <div style="margin-bottom: 16px;">
                    <label for="calc-example" style="display: block; margin-bottom: 8px; color: #ccc;">Example Result:</label>
                    <div id="calc-example" style="background: #1a4d1a; border: 1px solid #4caf50; border-radius: 4px; padding: 10px; color: #4caf50; font-size: 12px;">
                        Example: €1000 × 1.21 = €1210
                    </div>
                </div>

                <div style="display: flex; justify-content: flex-end; gap: 8px;">
                    <button class="btn btn-secondary" onclick="closeCalculationModal()">Cancel</button>
                    <button class="btn" onclick="applyCalculation('${targetColumn}')" id="apply-calculation-btn" disabled>Apply Calculation</button>
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);

    // Add event listeners for real-time preview
    document.getElementById('calc-column-a').addEventListener('change', updateCalculationPreview);
    document.getElementById('calc-operator').addEventListener('change', updateCalculationPreview);
    document.getElementById('calc-column-b').addEventListener('change', updateCalculationPreview);
    document.getElementById('calc-number').addEventListener('input', updateCalculationPreview);

    // Set initial values if editing existing calculation
    if (existingFormula) {
        parseExistingFormula(existingFormula);
    }

    updateCalculationPreview();
}

/**
 * Close calculation modal
 */
function closeCalculationModal() {
    const modal = document.getElementById('calculation-modal');
    if (modal) {
        modal.remove();
    }
}

/**
 * Get available source columns for calculations
 */
function getAvailableSourceColumns() {
    const sourceColumnsContainer = document.getElementById('source-columns');
    if (!sourceColumnsContainer) return [];

    const columnElements = sourceColumnsContainer.querySelectorAll('.column-item');
    return Array.from(columnElements).map(el => el.textContent.trim()).filter(col => col);
}

/**
 * Toggle between column and number input
 */
function toggleCalculationValueType() {
    const valueType = document.getElementById('calc-value-type').value;
    const columnSelect = document.getElementById('calc-column-b');
    const numberInput = document.getElementById('calc-number');

    if (valueType === 'number') {
        columnSelect.style.display = 'none';
        numberInput.style.display = 'block';
    } else {
        columnSelect.style.display = 'block';
        numberInput.style.display = 'none';
    }

    updateCalculationPreview();
}

/**
 * Update calculation preview in real-time
 */
function updateCalculationPreview() {
    const functionType = document.getElementById('calc-function').value;
    const columnA = document.getElementById('calc-column-a').value;
    const operator = document.getElementById('calc-operator').value;
    const valueType = document.getElementById('calc-value-type').value;
    const columnB = document.getElementById('calc-column-b').value;
    const numberValue = document.getElementById('calc-number').value;

    const paramsDiv = document.getElementById('function-params');
    const paramsContainer = document.getElementById('function-params-container');
    const previewEl = document.getElementById('calc-formula-preview');
    const exampleEl = document.getElementById('calc-example');
    const applyBtn = document.getElementById('apply-calculation-btn');

    // Show/hide and populate function parameters
    if (functionType) {
        paramsDiv.style.display = 'block';
        paramsContainer.innerHTML = generateFunctionParams(functionType);
    } else {
        paramsDiv.style.display = 'none';
    }

    // Build formula
    let formula = '';
    let isValid = false;

    // Check if it's a text-only function (no math formula needed)
    const textOnlyFunctions = ['LENGTH', 'ISEMPTY', 'TRIM', 'UPPER', 'LOWER'];

    if (functionType && textOnlyFunctions.includes(functionType)) {
        // Text-only functions - just need a column
        if (columnA) {
            formula = buildFunctionFormula(functionType, [columnA]);
            isValid = true;
        }
    } else {
        // Math-based or complex functions
        if (columnA && operator) {
            const operatorSymbol = operator === '*' ? '×' : operator;

            if (valueType === 'number' && numberValue) {
                formula = `${columnA} ${operatorSymbol} ${numberValue}`;
                isValid = true;
            } else if (valueType === 'column' && columnB) {
                formula = `${columnA} ${operatorSymbol} ${columnB}`;
                isValid = true;
            }

            // Apply function wrapper if selected
            if (formula && functionType) {
                formula = buildFunctionFormula(functionType, [formula]);
            }
        }
    }

    // Update preview
    if (formula) {
        previewEl.textContent = formula;

        // Create example calculation
        const exampleA = valueType === 'number' ? 1000 : 1000;
        const exampleB = valueType === 'number' ? parseFloat(numberValue) || 1 : 200;
        let result;

        switch (operator) {
            case '+': result = exampleA + exampleB; break;
            case '-': result = exampleA - exampleB; break;
            case '*': result = exampleA * exampleB; break;
            case '/': result = exampleB !== 0 ? exampleA / exampleB : 'Error'; break;
            default: result = 0;
        }

        if (typeof result === 'number') {
            result = Math.round(result * 100) / 100; // Round to 2 decimals
        }

        const operatorSymbol = operator === '*' ? '×' : operator;
        exampleEl.textContent = `Example: ${exampleA} ${operatorSymbol} ${exampleB} = ${result}`;
        exampleEl.style.backgroundColor = '#1a4d1a';
        exampleEl.style.borderColor = '#4caf50';
        exampleEl.style.color = '#4caf50';
    } else {
        previewEl.textContent = 'Select columns and operator to build formula';
        exampleEl.textContent = 'Example will appear here';
        exampleEl.style.backgroundColor = '#333';
        exampleEl.style.borderColor = '#555';
        exampleEl.style.color = '#888';
    }

    // Enable/disable apply button
    applyBtn.disabled = !isValid;
}

/**
 * Parse existing formula for editing
 */
function parseExistingFormula(formula) {
    // Simple parser for basic formulas like "ColumnA * 1.21" or "ColumnA - ColumnB"
    const match = formula.match(/^([^+\-*/]+)\s*([+\-*/])\s*(.+)$/);

    if (match) {
        const [, columnA, operator, valueB] = match;

        document.getElementById('calc-column-a').value = columnA.trim();
        document.getElementById('calc-operator').value = operator;

        // Check if valueB is a number or column
        const numValue = parseFloat(valueB.trim());
        if (!isNaN(numValue)) {
            document.getElementById('calc-value-type').value = 'number';
            document.getElementById('calc-number').value = valueB.trim();
            toggleCalculationValueType();
        } else {
            document.getElementById('calc-value-type').value = 'column';
            document.getElementById('calc-column-b').value = valueB.trim();
            toggleCalculationValueType();
        }
    }
}

/**
 * Apply calculation to mapping
 */
function applyCalculation(targetColumn) {
    const columnA = document.getElementById('calc-column-a').value;
    const operator = document.getElementById('calc-operator').value;
    const valueType = document.getElementById('calc-value-type').value;
    const columnB = document.getElementById('calc-column-b').value;
    const numberValue = document.getElementById('calc-number').value;

    if (!columnA || !operator) {
        alert('Please complete the formula');
        return;
    }

    // Build formula string for storage
    let formula = '';
    if (valueType === 'number' && numberValue) {
        formula = `${columnA} ${operator} ${numberValue}`;
    } else if (valueType === 'column' && columnB) {
        formula = `${columnA} ${operator} ${columnB}`;
    }

    if (formula) {
        // Store as calculation with prefix
        currentMapping[targetColumn] = `CALC:${formula}`;

        // Update visual display
        updateTemplateDropZones();

        console.log(`Calculation set: ${targetColumn} = ${formula}`);
    }

    closeCalculationModal();
}

/**
 * Generate dynamic parameter inputs based on function type
 */
function generateFunctionParams(functionType) {
    const inputStyle = 'background: #444; border: 1px solid #555; border-radius: 4px; padding: 8px; color: white; width: 100%; margin-bottom: 8px;';

    switch (functionType) {
        case 'ROUND':
            return `<input type="number" id="func-param-1" placeholder="Number of decimal places (e.g., 2)" style="${inputStyle}" value="2">`;

        case 'LEFT':
        case 'RIGHT':
            return `<input type="number" id="func-param-1" placeholder="Number of characters to extract" style="${inputStyle}">`;

        case 'MID':
            return `
                <input type="number" id="func-param-1" placeholder="Start position (1-based)" style="${inputStyle}">
                <input type="number" id="func-param-2" placeholder="Number of characters" style="${inputStyle}">`;

        case 'REGEX':
            return `
                <input type="text" id="func-param-1" placeholder="Regex pattern (e.g., ([A-Z]{2}\\d+))" style="${inputStyle}">
                <input type="number" id="func-param-2" placeholder="Group number (default: 1)" style="${inputStyle}" value="1">`;

        case 'SPLIT':
            return `
                <input type="text" id="func-param-1" placeholder="Delimiter (e.g., ' ', ',', ';')" style="${inputStyle}">
                <input type="number" id="func-param-2" placeholder="Index (1-based)" style="${inputStyle}">`;

        case 'REPLACE':
            return `
                <input type="text" id="func-param-1" placeholder="Text to find" style="${inputStyle}">
                <input type="text" id="func-param-2" placeholder="Replacement text" style="${inputStyle}">`;

        case 'CONCAT':
            return `<input type="text" id="func-param-1" placeholder="Additional text or column name" style="${inputStyle}">`;

        case 'CONTAINS':
        case 'STARTSWITH':
        case 'ENDSWITH':
            return `<input type="text" id="func-param-1" placeholder="Text to search for" style="${inputStyle}">`;

        case 'IF':
            return `
                <select id="func-param-1" style="${inputStyle}">
                    <option value="">Select condition type</option>
                    <option value="equals">Equals (=)</option>
                    <option value="not-equals">Not equals (≠)</option>
                    <option value="contains">Contains</option>
                    <option value="greater">Greater than (>)</option>
                    <option value="less">Less than (<)</option>
                </select>
                <input type="text" id="func-param-2" placeholder="Compare value" style="${inputStyle}">
                <input type="text" id="func-param-3" placeholder="True value" style="${inputStyle}">
                <input type="text" id="func-param-4" placeholder="False value" style="${inputStyle}">`;

        default:
            return '';
    }
}

/**
 * Build function formula with parameters
 */
function buildFunctionFormula(functionType, baseParams) {
    const getParam = (id) => {
        const elem = document.getElementById(id);
        return elem ? elem.value : '';
    };

    const baseExpression = baseParams[0];

    switch (functionType) {
        case 'ROUND':
            const decimals = getParam('func-param-1') || '2';
            return `ROUND(${baseExpression}, ${decimals})`;

        case 'LEFT':
            const leftCount = getParam('func-param-1');
            return leftCount ? `LEFT(${baseExpression}, ${leftCount})` : baseExpression;

        case 'RIGHT':
            const rightCount = getParam('func-param-1');
            return rightCount ? `RIGHT(${baseExpression}, ${rightCount})` : baseExpression;

        case 'MID':
            const start = getParam('func-param-1');
            const length = getParam('func-param-2');
            return (start && length) ? `MID(${baseExpression}, ${start}, ${length})` : baseExpression;

        case 'REGEX':
            const pattern = getParam('func-param-1');
            const group = getParam('func-param-2') || '1';
            return pattern ? `REGEX(${baseExpression}, "${pattern}", ${group})` : baseExpression;

        case 'SPLIT':
            const delimiter = getParam('func-param-1');
            const index = getParam('func-param-2');
            return (delimiter && index) ? `SPLIT(${baseExpression}, "${delimiter}", ${index})` : baseExpression;

        case 'REPLACE':
            const find = getParam('func-param-1');
            const replace = getParam('func-param-2');
            return (find && replace) ? `REPLACE(${baseExpression}, "${find}", "${replace}")` : baseExpression;

        case 'CONCAT':
            const additional = getParam('func-param-1');
            return additional ? `CONCAT(${baseExpression}, "${additional}")` : baseExpression;

        case 'CONTAINS':
            const searchText = getParam('func-param-1');
            return searchText ? `CONTAINS(${baseExpression}, "${searchText}")` : baseExpression;

        case 'STARTSWITH':
            const prefix = getParam('func-param-1');
            return prefix ? `STARTSWITH(${baseExpression}, "${prefix}")` : baseExpression;

        case 'ENDSWITH':
            const suffix = getParam('func-param-1');
            return suffix ? `ENDSWITH(${baseExpression}, "${suffix}")` : baseExpression;

        case 'IF':
            const condition = getParam('func-param-1');
            const compareValue = getParam('func-param-2');
            const trueValue = getParam('func-param-3');
            const falseValue = getParam('func-param-4');

            if (condition && compareValue && trueValue && falseValue) {
                let conditionExpr;
                switch (condition) {
                    case 'equals': conditionExpr = `${baseExpression} = "${compareValue}"`; break;
                    case 'not-equals': conditionExpr = `${baseExpression} ≠ "${compareValue}"`; break;
                    case 'contains': conditionExpr = `CONTAINS(${baseExpression}, "${compareValue}")`; break;
                    case 'greater': conditionExpr = `${baseExpression} > ${compareValue}`; break;
                    case 'less': conditionExpr = `${baseExpression} < ${compareValue}`; break;
                    default: return baseExpression;
                }
                return `IF(${conditionExpr}, "${trueValue}", "${falseValue}")`;
            }
            return baseExpression;

        case 'TRIM':
            return `TRIM(${baseExpression})`;
        case 'UPPER':
            return `UPPER(${baseExpression})`;
        case 'LOWER':
            return `LOWER(${baseExpression})`;
        case 'ABS':
            return `ABS(${baseExpression})`;
        case 'LENGTH':
            return `LENGTH(${baseExpression})`;
        case 'ISEMPTY':
            return `ISEMPTY(${baseExpression})`;
        case 'MIN':
            return `MIN(${baseExpression})`;
        case 'MAX':
            return `MAX(${baseExpression})`;
        case 'CEILING':
            return `CEILING(${baseExpression})`;
        case 'FLOOR':
            return `FLOOR(${baseExpression})`;

        default:
            return baseExpression;
    }
}

function clearMapping() {
    currentMapping = {};
    window.currentMappingConfidences = {}; // Clear confidence scores too
    updateTemplateDropZones();
    hideMappingSummary();
}

/**
 * Update mapping summary display with confidence breakdown
 * @param {Object} confidenceScores - Confidence scores by field name
 */
function updateMappingSummary(confidenceScores) {
    const summaryDiv = document.getElementById('mapping-summary');
    const summaryText = document.getElementById('mapping-summary-text');

    if (!confidenceScores || Object.keys(confidenceScores).length === 0) {
        hideMappingSummary();
        return;
    }

    // Categorize confidence scores
    const scores = Object.values(confidenceScores);
    const high = scores.filter(score => score >= 80).length;
    const medium = scores.filter(score => score >= 50 && score < 80).length;
    const low = scores.filter(score => score < 50).length;
    const total = scores.length;

    // Create summary text with color coding
    let summaryHtml = `<strong>Auto-mapping:</strong> ${total} fields - `;

    const parts = [];
    if (high > 0) parts.push(`<span style="color: #4caf50;">${high} high</span>`);
    if (medium > 0) parts.push(`<span style="color: #ffa726;">${medium} medium</span>`);
    if (low > 0) parts.push(`<span style="color: #f44336;">${low} low</span>`);

    summaryHtml += parts.join(' / ');
    summaryHtml += ' confidence';

    summaryText.innerHTML = summaryHtml;
    summaryDiv.style.display = 'block';
}

/**
 * Hide mapping summary display
 */
function hideMappingSummary() {
    const summaryDiv = document.getElementById('mapping-summary');
    if (summaryDiv) {
        summaryDiv.style.display = 'none';
    }
}

// Auto-mapping functionality moved to AutoMapping class

function generateAutoMappingSuggestions() {
    if (!currentMappingFile || !window.borderellenTemplate || !window.borderellenTemplate.columns) {
        alert('Please select a file and ensure a template is active.');
        return;
    }

    console.log('Starting matrix-based auto-mapping...');

    // Clear existing mapping
    currentMapping = {};

    // Get source columns from the UI
    const sourceColumns = document.querySelectorAll('#source-columns .column-item');
    const sourceColumnNames = Array.from(sourceColumns).map(col => col.getAttribute('data-column-name'));

    if (sourceColumnNames.length === 0) {
        alert('No source columns found. Please select a file first.');
        return;
    }

    try {
        // Use AutoMapping class for suggestions
        const result = AutoMapping.generateMappingSuggestions(
            sourceColumnNames,
            window.borderellenTemplate.columns,
            30 // Confidence threshold
        );

        // Apply the mapping results
        currentMapping = result.mapping;
        window.currentMappingConfidences = result.confidenceScores;

        // Update display
        updateTemplateDropZones();

        // Show summary in UI instead of popup
        updateMappingSummary(result.confidenceScores);

    } catch (error) {
        console.error('Auto-mapping failed:', error);
        alert(`Auto-mapping failed: ${error.message}`);
    }
}

async function saveBrokerTemplate() {
    if (!currentMappingFile || Object.keys(currentMapping).length === 0) {
        alert('Please select a file and create at least one mapping before saving.');
        return;
    }

    const templateName = prompt(
        `Enter a name for this broker template:`,
        `${currentMappingFile.broker.name} Template`
    );

    if (!templateName || templateName.trim() === '') {
        return; // User cancelled or entered empty name
    }

    // Check for duplicate template names
    try {
        const existingMappings = await window.loadAllFileMappings();
        const duplicateName = existingMappings.find(t => t.name === templateName.trim());

        if (duplicateName) {
            const overwrite = confirm(
                `A file mapping named "${templateName}" already exists.\n\n` +
                `Do you want to overwrite it?`
            );
            if (!overwrite) return;
        }
    } catch (error) {
        console.error('Error checking for duplicate templates:', error);
    }

    // Extract suggested keyword from filename
    const suggestedKeyword = window.extractKeywordFromFilename ? window.extractKeywordFromFilename(currentMappingFile.name) : '';
    const defaultKeyword = suggestedKeyword || currentMappingFile.broker.name.split(' ')[0];

    const matchingKeyword = prompt(
        `Enter a keyword for automatic file matching:\n\n` +
        `Files containing this keyword will automatically use this template.\n` +
        `For example, "DeVerkoper" will match "DeVerkoper_04_2025.xlsx"\n\n` +
        `Suggested keyword:`,
        defaultKeyword
    );

    // Allow empty keyword but inform user
    const finalKeyword = matchingKeyword ? matchingKeyword.trim() : '';
    if (!finalKeyword) {
        const proceed = confirm(
            `No keyword specified. This template will only be available for manual selection.\n\n` +
            `Do you want to continue without a keyword?`
        );
        if (!proceed) return;
    }

    try {
        // Create parsing config with auto-detected skip rules
        const parsingConfig = {
            dataStartMethod: 'skip-rows',
            skipRows: 0,
            skipColumns: 0,
            headerRow: null
        };

        // Apply auto-detected skip rules if available
        if (currentPatternAnalysis && currentPatternAnalysis.dataSection) {
            const analysis = currentPatternAnalysis.dataSection;
            parsingConfig.skipRows = analysis.dataStartIndex || 0;
            parsingConfig.skipColumns = analysis.startColumnIndex || 0;
            parsingConfig.headerRow = analysis.headerRowIndex;

            // Add enhanced parsingConfig fields for manual selections
            if (currentPatternAnalysis.manualSelection) {
                const manual = currentPatternAnalysis.manualSelection;
                parsingConfig.headerRows = manual.headerRows;
                parsingConfig.headerColumns = manual.headerColumns;
                parsingConfig.headerRange = manual.headerRange;

                // Add footer keyword detection if provided (manual takes priority)
                if (manual.footerKeyword) {
                    parsingConfig.footerRowKeyword = manual.footerKeyword;
                    console.log(`Footer detection configured: "${manual.footerKeyword}"`);
                } else if (currentPatternAnalysis.autoFooterKeyword) {
                    parsingConfig.footerRowKeyword = currentPatternAnalysis.autoFooterKeyword;
                    console.log(`Auto footer detection configured: "${currentPatternAnalysis.autoFooterKeyword}"`);
                }

                // Configure multi-row data processing if multiple header rows selected
                if (manual.headerRows > 1) {
                    parsingConfig.rowProcessing = {
                        type: 'multi-row',
                        rowsPerRecord: manual.headerRows
                    };
                }

                console.log(`Manual header selection: ${manual.headerRows} rows × ${manual.headerColumns} cols, range: ${manual.headerRange}`);
                if (manual.headerRows > 1) {
                    console.log(`Configured multi-row processing: ${manual.headerRows} rows per record`);
                }
            } else if (currentPatternAnalysis.autoFooterKeyword) {
                // Include auto-detected footer even without manual selection
                parsingConfig.footerRowKeyword = currentPatternAnalysis.autoFooterKeyword;
                console.log(`Auto footer detection configured: "${currentPatternAnalysis.autoFooterKeyword}"`);
            }

            console.log(`Auto-detected skip rules: skip ${parsingConfig.skipRows} rows, skip ${parsingConfig.skipColumns} columns`);
            console.log(`Start cell detected: ${analysis.startCell || 'A1'}`);
        }

        // Create file mapping object (unified format)
        const fileMapping = {
            id: `mapping-${Date.now()}`,
            name: templateName.trim(),
            matchingKeyword: finalKeyword,
            creationMethod: 'drag-drop',
            sourceType: currentMappingFile.broker.parser || currentMappingFile.broker.type,
            sourceName: currentMappingFile.broker.name,
            filePattern: currentMappingFile.name,
            parsingConfig: parsingConfig,
            columnMapping: { ...currentMapping },
            created: new Date().toISOString(),
            lastModified: new Date().toISOString(),
            version: '1.0'
        };

        // Save to unified file mappings store
        const saved = await window.saveFileMapping(fileMapping);
        if (!saved) {
            throw new Error('Failed to save file mapping to database');
        }

        // Also export as JSON file for backup and sharing
        await window.exportFileMappingAsJSON(fileMapping, null, appSettings);

        // Show success message
        alert(`File mapping "${templateName}" saved successfully!\n\nSaved to database and downloaded as JSON file for backup/sharing.`);

        console.log('File mapping saved:', fileMapping);
    } catch (error) {
        console.error('Error saving file mapping:', error);
        alert(`Failed to save file mapping: ${error.message}`);
    }
}

/**
 * Import broker template from JSON file
 */
function importBrokerTemplate() {
    const fileInput = document.getElementById('broker-template-file-input');
    fileInput.click();
}

/**
 * Handle broker template import from file input
 * @param {Event} event - File input change event
 */
async function handleBrokerTemplateImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        // Load and validate the JSON template
        const result = await window.loadFileMappingFromJSON(file);

        if (!result.success) {
            alert(`Failed to import broker template: ${result.error}`);
            return;
        }

        const template = result.mapping;

        // Ask if user wants to save to database or just load temporarily
        const saveToDb = confirm(
            `Import file mapping "${template.name}"?\n\n` +
            `Source Type: ${template.sourceType || template.sourceName}\n` +
            `Mappings: ${Object.keys(template.columnMapping || {}).length} fields\n\n` +
            `Click OK to save to database, or Cancel to load temporarily.`
        );

        if (saveToDb) {
            // The imported mapping is already in the correct format, just save it
            const saved = await window.saveFileMapping(template);
            if (!saved) {
                alert('Failed to save imported file mapping to database.');
                return;
            }
            alert(`File mapping "${template.name}" imported and saved successfully!`);
        } else {
            alert(`File mapping "${template.name}" loaded temporarily. Use "Save File Mapping" to persist it.`);
        }

        // Apply the imported mapping to current session
        currentMapping = { ...template.columnMapping };
        updateTemplateDropZones();

        console.log('Broker template imported:', template);

    } catch (error) {
        console.error('Error importing broker template:', error);
        alert(`Failed to import broker template: ${error.message}`);
    } finally {
        // Reset file input
        event.target.value = '';
    }
}


/**
 * Process file with current template mapping and add to results
 */
async function processFileWithTemplate() {
    if (!currentMappingFile) {
        alert('Please select a file to process.');
        return;
    }

    if (Object.keys(currentMapping).length === 0) {
        alert('Please create at least one mapping before processing.');
        return;
    }

    if (!window.borderellenTemplate || !window.borderellenTemplate.columns) {
        alert('Please ensure a template is active.');
        return;
    }

    try {
        console.log(`Processing file with custom mapping using orchestrator`);

        // Create parsing config with auto-detected skip rules if available
        const parsingConfig = {
            dataStartMethod: 'skip-rows',
            skipRows: 0,
            skipColumns: 0,
            headerRow: null
        };

        if (currentPatternAnalysis && currentPatternAnalysis.dataSection) {
            const analysis = currentPatternAnalysis.dataSection;
            parsingConfig.skipRows = analysis.dataStartIndex || 0;
            parsingConfig.skipColumns = analysis.startColumnIndex || 0;
            parsingConfig.headerRow = analysis.headerRowIndex;

            // Include enhanced fields for manual selections
            if (currentPatternAnalysis.manualSelection) {
                const manual = currentPatternAnalysis.manualSelection;
                parsingConfig.headerRows = manual.headerRows;
                parsingConfig.headerColumns = manual.headerColumns;
                parsingConfig.headerRange = manual.headerRange;

                // Add footer keyword detection if provided (manual takes priority)
                if (manual.footerKeyword) {
                    parsingConfig.footerRowKeyword = manual.footerKeyword;
                } else if (currentPatternAnalysis.autoFooterKeyword) {
                    parsingConfig.footerRowKeyword = currentPatternAnalysis.autoFooterKeyword;
                }

                // Configure multi-row data processing if multiple header rows selected
                if (manual.headerRows > 1) {
                    parsingConfig.rowProcessing = {
                        type: 'multi-row',
                        rowsPerRecord: manual.headerRows
                    };
                }
            } else if (currentPatternAnalysis.autoFooterKeyword) {
                // Include auto-detected footer even without manual selection
                parsingConfig.footerRowKeyword = currentPatternAnalysis.autoFooterKeyword;
            }

            console.log(`Applying auto-detected skip rules: skip ${parsingConfig.skipRows} rows, skip ${parsingConfig.skipColumns} columns`);
        }

        // Use the extended orchestrator with manual mapping override
        const result = await processBrokerFile({
            ...currentMappingFile,
            detectionOverride: {
                type: 'manual-mapping',
                mapping: currentMapping,
                parsingConfig: parsingConfig,
                templateName: 'Custom Template'
            }
        });

        if (result.success) {
            console.log(`Generated ${result.recordCount} processed records`);

            // Update the file with processed results
            currentMappingFile.parsedData = result.data;
            currentMappingFile.recordCount = result.recordCount;
            currentMappingFile.status = 'Processed with Custom Template';
            currentMappingFile.statusClass = 'status-success';
            currentMappingFile.broker = {
                ...currentMappingFile.broker,
                type: 'custom-template',
                name: currentMappingFile.broker.name + ' (Custom)'
            };

            // Update the files display
            updateFilesDisplay();

            alert(`Successfully processed ${result.recordCount} records using custom template mapping!`);

            // Auto-navigate to Results tab to see the processed data
            const resultsTab = document.querySelector('[data-tab="results"]');
            if (resultsTab) {
                resultsTab.click();
            }
        } else {
            throw new Error(result.error || 'Processing failed');
        }

    } catch (error) {
        console.error('Error processing file with template:', error);
        alert(`Failed to process file: ${error.message}`);
    }
}

/**
 * Apply mapping configuration to sample data
 * @param {Array} sampleData - Raw data from Excel
 * @param {Object} mapping - Current mapping configuration
 * @returns {Array} Mapped data
 */
function applyMappingToData(sampleData, mapping) {
    return sampleData.map(row => {
        const mappedRow = {};

        // Apply each mapping rule
        Object.keys(mapping).forEach(targetField => {
            const mappingRule = mapping[targetField];

            if (mappingRule.startsWith('FIXED:')) {
                // Fixed value mapping
                mappedRow[targetField] = mappingRule.substring(6);
            } else if (mappingRule.startsWith('CALC:')) {
                // Calculation mapping
                const formula = mappingRule.substring(5);
                mappedRow[targetField] = executeFormula(formula, row);
            } else {
                // Column mapping - handle undefined/null but preserve falsy values like 0
                let value = row[mappingRule];

                if (value !== undefined && value !== null) {
                    // Check if this looks like an Excel date number and the target field is date-related
                    if (isExcelDate(value) && isDateField(targetField)) {
                        value = formatExcelDate(value);
                    }
                    mappedRow[targetField] = value;
                } else {
                    mappedRow[targetField] = '';
                }
            }
        });

        return mappedRow;
    });
}

/**
 * Check if a value looks like an Excel date serial number
 * @param {any} value - Value to check
 * @returns {boolean} True if looks like Excel date
 */
function isExcelDate(value) {
    // Excel dates are numbers between reasonable bounds
    // Excel epoch starts 1900-01-01 (serial 1) to future dates
    return typeof value === 'number' && value > 0 && value < 100000;
}

/**
 * Check if a field name indicates it should contain date data
 * @param {string} fieldName - Field name to check
 * @returns {boolean} True if field is date-related
 */
function isDateField(fieldName) {
    const dateFieldPatterns = [
        /datum/i,
        /date/i,
        /van$/i,
        /tot$/i,
        /periode/i,
        /dtm$/i
    ];

    return dateFieldPatterns.some(pattern => pattern.test(fieldName));
}

/**
 * Format Excel date serial number to readable date string
 * @param {number} excelDate - Excel serial date number
 * @returns {string} Formatted date string
 */
function formatExcelDate(excelDate) {
    try {
        // Excel epoch starts at 1900-01-01 but has a leap year bug
        // JavaScript Date constructor works with milliseconds
        const excelEpoch = new Date(1900, 0, 1);
        const msPerDay = 24 * 60 * 60 * 1000;

        // Adjust for Excel's leap year bug (it thinks 1900 was a leap year)
        let adjustedDate = excelDate;
        if (excelDate > 59) adjustedDate -= 1;

        const jsDate = new Date(excelEpoch.getTime() + (adjustedDate - 1) * msPerDay);

        // Return formatted date (DD-MM-YYYY)
        return jsDate.toLocaleDateString('nl-NL', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
    } catch (error) {
        console.warn('Failed to convert Excel date:', excelDate, error);
        return excelDate.toString(); // Fallback to original value
    }
}

/**
 * Filter data based on broker-specific rules for preview
 * @param {Array} rawData - Raw Excel data
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
                <strong>File:</strong> ${currentMappingFile.name}
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
 * Show preview modal
 */
function showPreviewModal() {
    document.getElementById('preview-modal').style.display = 'flex';
}

/**
 * Hide preview modal
 */
function hidePreviewModal() {
    document.getElementById('preview-modal').style.display = 'none';
}

/**
 * Unified preview function for file data (with or without mapping)
 * @param {Object} fileData - File data object
 * @param {Object} mapping - Optional mapping configuration (if null, shows raw data)
 */
async function previewFileData(fileData, mapping = null) {
    try {
        // Show loading state
        showPreviewModal();
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
            const workbook = await readExcelFile(fileData.file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            let rawData;
            // Use cached analysis from fileData if available, otherwise fall back to global currentPatternAnalysis
            const analysisToUse = fileData.patternAnalysis || currentPatternAnalysis;

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
                const parser = new GenericBrokerParser(tempConfig);
                rawData = await parser.parse(workbook, fileData.name);
            } else {
                // Fallback to simple sheet reading
                console.log('Preview: No skip rules available, using simple sheet reading');
                rawData = XLSX.utils.sheet_to_json(worksheet);
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
                        if (isExcelDate(value) && isDateField(key)) {
                            value = formatExcelDate(value);
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
                displayData = applyMappingToData(sampleData, mapping);
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

/**
 * Export preview results as Excel file
 */
async function exportPreviewAsExcel() {
    if (!currentMappingFile || Object.keys(currentMapping).length === 0) {
        alert('No preview data available to export.');
        return;
    }

    try {
        // Get the full dataset (not just preview sample)
        let fullData = [];

        if (currentMappingFile.parsedData && currentMappingFile.parsedData.length > 0) {
            // Use already parsed data if available
            fullData = currentMappingFile.parsedData;
        } else {
            // Parse the full file with skip rules if available
            const workbook = await readExcelFile(currentMappingFile.file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            let rawData;
            // Use cached analysis from currentMappingFile if available, otherwise fall back to global currentPatternAnalysis
            const analysisToUse = currentMappingFile.patternAnalysis || currentPatternAnalysis;

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
                const parser = new GenericBrokerParser(tempConfig);
                rawData = await parser.parse(workbook, currentMappingFile.name);
            } else {
                // Fallback to simple sheet reading
                console.log('Export: No skip rules available, using simple sheet reading');
                rawData = XLSX.utils.sheet_to_json(worksheet);
            }

            // Apply broker-specific filtering
            const brokerType = currentMappingFile.broker.parser;
            fullData = filterDataForPreview(rawData, brokerType);
        }

        // Apply mapping to full dataset
        const mappedFullData = applyMappingToData(fullData, currentMapping);

        // Create Excel file
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(mappedFullData);

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Mapped Data');

        // Generate filename
        const brokerName = currentMappingFile.broker.name || 'Unknown';
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `${brokerName}_mapped_${timestamp}.xlsx`;

        // Download to preferred folder or fallback to browser download
        const success = await downloadExcelToPreferredFolder(wb, filename);

        if (success) {
            alert(`Exported ${mappedFullData.length} records successfully`);
        }

    } catch (error) {
        console.error('Error exporting preview to Excel:', error);
        alert(`Failed to export to Excel: ${error.message}`);
    }
}

// ========== DOWNLOAD HELPER FUNCTIONS ==========

/**
 * Download file to preferred folder or fallback to browser download
 * @param {Blob} blob - File content as blob
 * @param {string} filename - Target filename
 * @param {string} mimeType - MIME type for the file
 * @returns {Promise<boolean>} Success status
 */
async function downloadToPreferredFolder(blob, filename, mimeType = 'application/octet-stream') {
    try {
        // Check if File System Access API is available
        if ('showSaveFilePicker' in window) {
            // Determine file type options based on extension
            let fileTypes = [];
            if (filename.endsWith('.json')) {
                fileTypes = [{
                    description: 'JSON files',
                    accept: {'application/json': ['.json']}
                }];
            } else if (filename.endsWith('.eml')) {
                fileTypes = [{
                    description: 'Email files',
                    accept: {'message/rfc822': ['.eml']}
                }];
            } else {
                fileTypes = [{
                    description: 'All files',
                    accept: {'*/*': []}
                }];
            }

            // Use showSaveFilePicker with preferred folder as starting directory
            const options = {
                suggestedName: filename,
                types: fileTypes
            };

            // If we have a folder handle, use it as starting directory suggestion
            if (appSettings.downloadFolderHandle) {
                options.startIn = appSettings.downloadFolderHandle;
            }

            const fileHandle = await window.showSaveFilePicker(options);

            // Write the file
            const writable = await fileHandle.createWritable();
            await writable.write(blob);
            await writable.close();

            return true;
        }
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('User cancelled save dialog');
            return false; // Don't fall back if user cancelled
        }
        console.log('Save picker failed, falling back to browser download:', error);
    }

    // Fallback to standard browser download
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
    return true;
}

/**
 * Download Excel workbook to preferred folder or fallback to browser download
 * @param {Object} workbook - XLSX workbook object
 * @param {string} filename - Target filename
 * @returns {Promise<boolean>} Success status
 */
async function downloadExcelToPreferredFolder(workbook, filename) {
    console.log('downloadExcelToPreferredFolder called with filename:', filename);
    console.log('showSaveFilePicker available:', 'showSaveFilePicker' in window);
    console.log('downloadFolderHandle:', appSettings.downloadFolderHandle);

    try {
        // Check if File System Access API is available
        if ('showSaveFilePicker' in window) {
            // Convert workbook to array buffer first
            const arrayBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

            // Use showSaveFilePicker with preferred folder as starting directory
            const options = {
                suggestedName: filename,
                types: [{
                    description: 'Excel files',
                    accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']}
                }]
            };

            // If we have a folder handle, use it as starting directory suggestion
            if (appSettings.downloadFolderHandle) {
                options.startIn = appSettings.downloadFolderHandle;
            }

            const fileHandle = await window.showSaveFilePicker(options);

            // Write the file
            const writable = await fileHandle.createWritable();
            await writable.write(arrayBuffer);
            await writable.close();

            return true;
        }
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('User cancelled save dialog');
            return false; // Don't fall back if user cancelled
        }
        console.log('Save picker failed, falling back to browser download:', error);
    }

    // Fallback to standard XLSX download
    XLSX.writeFile(workbook, filename);
    return true;
}

// ========== RESULTS TAB FUNCTIONS ==========

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

// ========== RESULTS TAB FUNCTIONS ==========

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

    // Get all unique columns (should be standardized 22-column format)
    const columns = [...new Set(combinedData.flatMap(row =>
        Object.keys(row).filter(key => !key.startsWith('_'))
    ))];

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

/**
 * Download combined data as Excel file
 */
async function downloadCombinedExcel() {
    console.log('downloadCombinedExcel called');
    console.log('currentCombinedData:', window.currentCombinedData?.length);
    console.log('appSettings:', appSettings);

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
        const success = await downloadExcelToPreferredFolder(wb, filename);

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
        const success = await downloadToPreferredFolder(blob, filename, 'application/json');

        if (success) {
            alert(`Exported ${exportData.length} records successfully`);
        }

    } catch (error) {
        console.error('Error exporting JSON:', error);
        alert(`Failed to export JSON file: ${error.message}`);
    }
}

// ========== AUTO-NAVIGATION FUNCTIONS ==========

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
    if (selector.value) {
        return; // Don't override existing selection
    }

    let fileToSelect = null;

    // Priority 1: Parse errors or failed files (highest priority)
    fileToSelect = window.uploadedFiles.find(f =>
        f.statusClass === 'status-error' ||
        f.status.toLowerCase().includes('error') ||
        f.status.toLowerCase().includes('failed')
    );

    // Priority 2: Unknown formats that need templates
    if (!fileToSelect) {
        fileToSelect = window.uploadedFiles.find(f =>
            f.broker.type === 'Unknown' ||
            f.status.includes('Create Template')
        );
    }

    // Priority 3: Files with warnings or no valid records
    if (!fileToSelect) {
        fileToSelect = window.uploadedFiles.find(f =>
            f.statusClass === 'status-warning' ||
            f.status.includes('No Valid Records') ||
            f.recordCount === 0
        );
    }

    // Priority 4: Successfully parsed files that could benefit from mapping validation
    if (!fileToSelect) {
        fileToSelect = window.uploadedFiles.find(f =>
            f.broker.type === 'built-in' &&
            f.statusClass === 'status-success' &&
            f.recordCount > 0
        );
    }

    // Priority 5: Any remaining file (fallback)
    if (!fileToSelect) {
        fileToSelect = window.uploadedFiles[0];
    }

    if (fileToSelect) {
        const priorityReason = fileToSelect.statusClass === 'status-error' ? 'needs troubleshooting' :
                              fileToSelect.broker.type === 'Unknown' ? 'needs template creation' :
                              fileToSelect.statusClass === 'status-warning' ? 'has issues' :
                              'ready for mapping';

        console.log(`Auto-selecting file: ${fileToSelect.name} (${priorityReason})`);

        // Set the selector value and trigger change event
        selector.value = fileToSelect.id;

        // Trigger the change event to load columns and auto-mapping
        const changeEvent = new Event('change', { bubbles: true });
        selector.dispatchEvent(changeEvent);
    }
}

// ========== SETTINGS MANAGEMENT FUNCTIONS ==========

function showSettingsModal() {
    const modal = document.getElementById('settings-modal');
    const userNameInput = document.getElementById('user-name-input');
    const userEmailInput = document.getElementById('user-email-input');
    const userSignatureInput = document.getElementById('user-signature-input');
    const downloadFolderInput = document.getElementById('download-folder-input');

    // Populate current values
    userNameInput.value = appSettings.userName;
    userEmailInput.value = appSettings.userEmail || '';
    userSignatureInput.value = appSettings.userSignature || '';
    downloadFolderInput.value = appSettings.downloadFolder;

    modal.classList.add('show');
}

function hideSettingsModal() {
    const modal = document.getElementById('settings-modal');
    modal.classList.remove('show');
}

async function selectDownloadFolder() {
    try {
        // Use File System Access API if available
        if ('showDirectoryPicker' in window) {
            const dirHandle = await window.showDirectoryPicker();
            appSettings.downloadFolderHandle = dirHandle;
            appSettings.downloadFolder = dirHandle.name;

            document.getElementById('download-folder-input').value = dirHandle.name;

            // Save the folder handle immediately so it persists across sessions
            await saveSettings(appSettings);
            console.log('Saved folder handle to IndexedDB:', dirHandle.name);
        } else {
            // Fallback: inform user about browser limitations
            alert('Your browser does not support folder selection. Downloads will go to your default download folder.');
        }
    } catch (error) {
        console.log('User cancelled folder selection');
    }
}

async function saveSettingsFromModal() {
    const userNameInput = document.getElementById('user-name-input');
    const userEmailInput = document.getElementById('user-email-input');
    const userSignatureInput = document.getElementById('user-signature-input');

    appSettings.userName = userNameInput.value.trim() || 'User';
    appSettings.userEmail = userEmailInput.value.trim() || '';
    appSettings.userSignature = userSignatureInput.value.trim() || '';

    if (await saveSettings(appSettings)) {
        hideSettingsModal();
        alert('Settings saved successfully!');
    } else {
        alert('Error saving settings. Please try again.');
    }
}

// ========== UTILITY FUNCTIONS ==========

/**
 * Update button states based on current selection
 */
function updateButtonStates() {
    const templateSelector = document.getElementById('template-selector');
    const activateTemplateBtn = document.getElementById('activate-template-btn');
    const editTemplateBtn = document.getElementById('edit-template-btn');
    const deleteTemplateBtn = document.getElementById('delete-template-btn');

    const hasSelection = templateSelector && templateSelector.value !== '';
    const hasActiveTemplate = window.currentTemplateId && window.savedTemplates && window.savedTemplates[window.currentTemplateId];
    const selectedId = templateSelector ? templateSelector.value : '';

    // Activate button - enabled only when selection is different from current
    if (activateTemplateBtn) {
        activateTemplateBtn.disabled = !hasSelection || selectedId === window.currentTemplateId;
    }

    // Edit and Delete buttons - enabled when we have an active template
    if (editTemplateBtn) {
        editTemplateBtn.disabled = !hasActiveTemplate;
    }
    if (deleteTemplateBtn) {
        deleteTemplateBtn.disabled = !hasActiveTemplate;
    }
}

// ========== APPLICATION INITIALIZATION ==========

document.addEventListener('DOMContentLoaded', function() {
    // Tab switching functionality
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            // Remove active class from all tab buttons
            tabButtons.forEach(btn => btn.classList.remove('active'));

            // Remove active class from all tab contents and hide them
            tabContents.forEach(content => {
                content.classList.remove('active');
                content.style.display = 'none';
            });

            // Add active class to clicked button
            button.classList.add('active');

            // Add active class to corresponding tab content
            const tabId = button.getAttribute('data-tab');
            const targetTab = document.getElementById(tabId + '-tab');

            if (targetTab) {
                targetTab.classList.add('active');
                targetTab.style.display = 'flex';
                targetTab.style.flexDirection = 'column';
            } else {
                console.error('Tab element not found:', tabId + '-tab');
            }

            // Initialize mapping tab when switched to
            if (tabId === 'mapping') {
                updateTemplateDropZones();
                if (typeof updateMappingFileSelector === 'function') {
                    updateMappingFileSelector();
                }
                attachDropZoneListeners(); // Re-attach listeners after DOM update

                // Initialize keyword management section
                if (typeof window.loadKeywordManagement === 'function') {
                    window.loadKeywordManagement();
                }

                // Auto-select a file with smart prioritization
                setTimeout(() => {
                    autoSelectFileInMappingTab();
                }, 100); // Small delay to ensure DOM is updated
            }

            // Initialize results tab when switched to
            if (tabId === 'results') {
                updateResultsTab();
                // Set up event listeners for results tab (in case they weren't attached during DOMContentLoaded)
                setupResultsTabListeners();
            }

            // Initialize email tab when switched to
            if (tabId === 'email') {
                // Ensure the correct data source is selected and analyzed
                const useResultsData = document.getElementById('use-results-data');
                const emailUploadSection = document.getElementById('email-upload-section');

                if (useResultsData && useResultsData.checked) {
                    emailUploadSection.style.display = 'none';
                    // Auto-analyze current results data if available
                    setTimeout(() => {
                        if (window.currentCombinedData && window.currentCombinedData.length > 0) {
                            analyzeCurrentResultsData();
                        }
                    }, 100);
                }
            }
        });
    });

    // Initialize IndexedDB and load templates
    console.log('Initializing IndexedDB...');
    initIndexedDB().then(async () => {
        console.log('IndexedDB initialized successfully');


        console.log('Loading default template...');
        if (typeof loadDefaultTemplate === 'function') {
            await loadDefaultTemplate();
            console.log('Template loaded successfully');
        } else {
            console.error('loadDefaultTemplate function not found');
        }

        console.log('Loading settings...');
        if (typeof loadSettings === 'function') {
            const settings = await loadSettings();
            appSettings = { ...appSettings, ...settings };
            console.log('Settings loaded successfully:', appSettings);
        } else {
            console.error('loadSettings function not found');
        }

        // Update displays after loading
        if (typeof updateTemplateSelector === 'function') {
            updateTemplateSelector();
            console.log('Template selector updated');
        }
        if (typeof updateActiveTemplateDisplay === 'function') {
            updateActiveTemplateDisplay();
            console.log('Active template display updated');
        }

        // Load email template after settings are loaded
        loadEmailTemplate();
        console.log('Email template loaded from settings');

        console.log('App initialization complete');
        console.log('Current template ID:', window.currentTemplateId);
        console.log('Available templates:', Object.keys(window.savedTemplates || {}));

        // Auto-navigate to Upload tab if template is already active
        autoNavigateOnStart();
    }).catch(error => {
        console.error('Failed to initialize IndexedDB:', error);
        alert('Database initialization failed. Some features may not work properly.');
        // Continue anyway, don't call showMainSection() as it may not exist
    });

    // Template management functionality
    const templateSelector = document.getElementById('template-selector');
    const activateTemplateBtn = document.getElementById('activate-template-btn');
    const editTemplateBtn = document.getElementById('edit-template-btn');
    const newTemplateBtn = document.getElementById('new-template-btn');
    const deleteTemplateBtn = document.getElementById('delete-template-btn');
    const cancelNewBtn = document.getElementById('cancel-new-btn');
    const cancelEditBtn = document.getElementById('cancel-edit-btn');
    const templateUploadZone = document.getElementById('template-upload-zone');
    const templateFileInput = document.getElementById('template-file-input');
    const browseTemplateBtn = document.getElementById('browse-template-btn');
    const createManualBtn = document.getElementById('create-manual-btn');
    const addColumnBtn = document.getElementById('add-column-btn');
    const saveTemplateBtn = document.getElementById('save-template-btn');
    const exportTemplateBtn = document.getElementById('export-template-btn');

    // Template selector event listener
    templateSelector.addEventListener('change', (e) => {
        if (typeof selectTemplate === 'function') {
            selectTemplate(e.target.value);
        }
        updateButtonStates();
    });

    // Template management event listeners
    activateTemplateBtn.addEventListener('click', () => {
        const selectedId = templateSelector.value;
        if (selectedId && typeof selectTemplate === 'function') {
            selectTemplate(selectedId);
        }
    });

    if (editTemplateBtn) {
        editTemplateBtn.addEventListener('click', () => {
            if (!window.currentTemplateId || !window.savedTemplates || !window.savedTemplates[window.currentTemplateId]) {
                alert('Selecteer eerst een template om te bewerken.');
                return;
            }
            if (typeof showEditTemplateSection === 'function') {
                showEditTemplateSection();
            }
        });
    }

    if (newTemplateBtn) {
        newTemplateBtn.addEventListener('click', () => {
            if (typeof createNewTemplate === 'function') {
                createNewTemplate();
            }
        });
    }
    if (deleteTemplateBtn) {
        deleteTemplateBtn.addEventListener('click', () => {
            if (typeof deleteTemplate === 'function') {
                deleteTemplate();
            }
        });
    }

    // Cancel button functionality
    if (cancelNewBtn) {
        cancelNewBtn.addEventListener('click', () => {
            if (typeof showMainSection === 'function') {
                showMainSection();
            }
        });
    }

    if (cancelEditBtn) {
        cancelEditBtn.addEventListener('click', () => {
            if (typeof showMainSection === 'function') {
                showMainSection();
            }
        });
    }

    // Template upload event listeners
    if (templateUploadZone) {
        templateUploadZone.addEventListener('click', () => templateFileInput.click());
    }
    if (browseTemplateBtn) {
        browseTemplateBtn.addEventListener('click', () => templateFileInput.click());
    }
    if (createManualBtn) {
        createManualBtn.addEventListener('click', () => {
            if (typeof createManualTemplate === 'function') {
                createManualTemplate();
            }
        });
    }

    templateFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            if (file.name.endsWith('.json')) {
                window.loadFromJSON(file);
            } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                loadTemplateFromExcel(file);
            } else {
                alert('Please upload an Excel (.xlsx) or JSON (.json) file');
            }
        }
        e.target.value = ''; // Reset input
    });

    // Template upload drag & drop
    templateUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        templateUploadZone.classList.add('dragover');
    });

    templateUploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        templateUploadZone.classList.remove('dragover');
    });

    templateUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        templateUploadZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (file.name.endsWith('.json')) {
                window.loadFromJSON(file);
            } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                loadTemplateFromExcel(file);
            } else {
                alert('Please upload an Excel (.xlsx) or JSON (.json) file');
            }
        }
    });

    // Template management event listeners
    addColumnBtn.addEventListener('click', addTemplateColumn);
    saveTemplateBtn.addEventListener('click', saveTemplate);
    exportTemplateBtn.addEventListener('click', exportTemplate);

    // Initialize file upload system
    console.log('Initializing file upload system...');
    if (typeof initializeFileUpload === 'function') {
        initializeFileUpload();
        console.log('File upload system initialized');
    } else {
        console.error('initializeFileUpload function not found');
    }

    // Mapping tab functionality
    const mappingFileSelector = document.getElementById('mapping-file-selector');
    mappingFileSelector.addEventListener('change', (e) => {
        const fileId = e.target.value;
        if (fileId) {
            loadSourceColumns(fileId);
            // Auto-map for known broker types
            applyAutoMappingForBrokerType(fileId);
            updateTemplateDropZones();
        } else {
            document.getElementById('source-columns').innerHTML = '<div style="text-align: center; padding: 32px; color: #888;"><p>Select a file to see available columns</p></div>';
            currentMapping = {};
            updateTemplateDropZones();
            hideMappingSummary();
        }
    });

    // Mapping tab buttons
    document.getElementById('auto-map-btn').addEventListener('click', generateAutoMappingSuggestions);
    document.getElementById('clear-mapping-btn').addEventListener('click', clearMapping);
    document.getElementById('save-broker-template-btn').addEventListener('click', saveBrokerTemplate);
    document.getElementById('import-broker-template-btn').addEventListener('click', importBrokerTemplate);
    document.getElementById('process-with-template-btn').addEventListener('click', processFileWithTemplate);

    // Keyword management button
    const refreshTemplatesBtn = document.getElementById('refresh-templates-btn');
    if (refreshTemplatesBtn) {
        refreshTemplatesBtn.addEventListener('click', () => {
            if (typeof window.loadKeywordManagement === 'function') {
                window.loadKeywordManagement();
            }
        });
    }

    // Broker template import file input
    const brokerTemplateFileInput = document.getElementById('broker-template-file-input');
    brokerTemplateFileInput.addEventListener('change', handleBrokerTemplateImport);

    // Navigation buttons functionality
    const helpBtn = document.getElementById('help-btn');
    const flowchartBtn = document.getElementById('flowchart-btn');

    // Settings modal functionality
    const settingsBtn = document.getElementById('settings-btn');
    const closeSettingsModal = document.getElementById('close-settings-modal');
    const selectFolderBtn = document.getElementById('select-folder-btn');
    const saveSettingsBtn = document.getElementById('save-settings-btn');
    const cancelSettingsBtn = document.getElementById('cancel-settings-btn');

    // Navigation button event listeners
    helpBtn.addEventListener('click', () => {
        window.open('README.html', '_blank');
    });

    flowchartBtn.addEventListener('click', () => {
        window.open('flowchart.html', '_blank');
    });

    settingsBtn.addEventListener('click', showSettingsModal);
    closeSettingsModal.addEventListener('click', hideSettingsModal);
    selectFolderBtn.addEventListener('click', selectDownloadFolder);
    saveSettingsBtn.addEventListener('click', saveSettingsFromModal);
    cancelSettingsBtn.addEventListener('click', hideSettingsModal);

    // Contact management event listeners
    const contactsBtn = document.getElementById('contacts-btn');
    const closeContactsModal = document.getElementById('close-contacts-modal');
    const closeContactsBtn = document.getElementById('close-contacts-btn');
    const contactForm = document.getElementById('contact-form');
    const cancelContactBtn = document.getElementById('cancel-contact-btn');
    const exportContactsBtn = document.getElementById('export-contacts-btn');
    const importContactsBtn = document.getElementById('import-contacts-btn');
    const contactsFileInput = document.getElementById('contacts-file-input');

    contactsBtn.addEventListener('click', showContactsModal);
    closeContactsModal.addEventListener('click', hideContactsModal);
    closeContactsBtn.addEventListener('click', hideContactsModal);
    contactForm.addEventListener('submit', saveContact);
    cancelContactBtn.addEventListener('click', cancelContactForm);
    exportContactsBtn.addEventListener('click', exportContacts);
    importContactsBtn.addEventListener('click', () => contactsFileInput.click());
    contactsFileInput.addEventListener('change', importContacts);

    // Close modal when clicking outside
    document.getElementById('settings-modal').addEventListener('click', (e) => {
        if (e.target.id === 'settings-modal') {
            hideSettingsModal();
        }
    });

    document.getElementById('contacts-modal').addEventListener('click', (e) => {
        if (e.target.id === 'contacts-modal') {
            hideContactsModal();
        }
    });

    // Preview modal functionality
    const closePreviewModal = document.getElementById('close-preview-modal');
    const closePreviewBtn = document.getElementById('close-preview-btn');
    const exportPreviewBtn = document.getElementById('export-preview-btn');

    closePreviewModal.addEventListener('click', hidePreviewModal);
    closePreviewBtn.addEventListener('click', hidePreviewModal);
    exportPreviewBtn.addEventListener('click', exportPreviewAsExcel);

    // Close preview modal when clicking outside
    document.getElementById('preview-modal').addEventListener('click', (e) => {
        if (e.target.id === 'preview-modal') {
            hidePreviewModal();
        }
    });

    // Initialize template builder

    // Initialize email brokers tab functionality
    initializeEmailBrokersTab();

    // Results tab functionality will be set up when the tab is activated

    // Export functions globally for cross-module access
    window.previewFileData = previewFileData;
    window.applyMappingToData = applyMappingToData;
});

// ========== CUSTOM TEMPLATE BUILDER FUNCTIONS ==========



// ========== BROKER CONTACT MANAGEMENT FUNCTIONS ==========

let currentEditingContactId = null;
let allContacts = [];

/**
 * Show the contact management modal
 */
async function showContactsModal() {
    const modal = document.getElementById('contacts-modal');
    modal.classList.add('show');

    // Load and display contacts
    await loadAndDisplayContacts();

    // Reset form
    resetContactForm();
}

/**
 * Hide the contact management modal
 */
function hideContactsModal() {
    const modal = document.getElementById('contacts-modal');
    modal.classList.remove('show');
    resetContactForm();
}

/**
 * Load contacts from IndexedDB and display them
 */
async function loadAndDisplayContacts() {
    try {
        allContacts = await window.loadAllBrokerContacts();
        displayContactsTable(allContacts);
    } catch (error) {
        console.error('Error loading contacts:', error);
        alert('Error loading contacts: ' + error.message);
    }
}

/**
 * Display contacts in the table
 */
function displayContactsTable(contacts) {
    const noContactsMessage = document.getElementById('no-contacts-message');
    const tableContainer = document.getElementById('contacts-table-container');
    const tableBody = document.getElementById('contacts-table-body');

    if (!contacts || contacts.length === 0) {
        noContactsMessage.style.display = 'block';
        tableContainer.style.display = 'none';
        return;
    }

    noContactsMessage.style.display = 'none';
    tableContainer.style.display = 'block';

    tableBody.innerHTML = contacts.map(contact => `
        <tr>
            <td>${escapeHtml(contact.brokerName)}</td>
            <td>${escapeHtml(contact.firstName)} ${escapeHtml(contact.lastName)}</td>
            <td>${escapeHtml(contact.email)}</td>
            <td>
                <button class="btn btn-secondary" onclick="editContact('${contact.id}')" style="margin-right: 8px; font-size: 12px; padding: 4px 8px;">Edit</button>
                <button class="btn" onclick="deleteContact('${contact.id}')" style="background: #dc3545; border-color: #dc3545; font-size: 12px; padding: 4px 8px;">Delete</button>
            </td>
        </tr>
    `).join('');
}

/**
 * Save contact (create or update)
 */
async function saveContact(event) {
    event.preventDefault();

    const brokerName = document.getElementById('contact-broker-name').value.trim();
    const firstName = document.getElementById('contact-first-name').value.trim();
    const lastName = document.getElementById('contact-last-name').value.trim();
    const email = document.getElementById('contact-email').value.trim();

    if (!brokerName || !firstName || !lastName || !email) {
        alert('Please fill in all required fields.');
        return;
    }

    // Email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        alert('Please enter a valid email address.');
        return;
    }

    const contact = {
        id: currentEditingContactId || `contact-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        brokerName,
        firstName,
        lastName,
        email,
        created: currentEditingContactId ?
            allContacts.find(c => c.id === currentEditingContactId)?.created || new Date().toISOString() :
            new Date().toISOString(),
        lastModified: new Date().toISOString()
    };

    try {
        const success = await window.saveBrokerContact(contact);
        if (success) {
            resetContactForm();
            await loadAndDisplayContacts();
        } else {
            alert('Error saving contact. Please try again.');
        }
    } catch (error) {
        console.error('Error saving contact:', error);
        alert('Error saving contact: ' + error.message);
    }
}

/**
 * Edit contact
 */
function editContact(contactId) {
    const contact = allContacts.find(c => c.id === contactId);
    if (!contact) {
        alert('Contact not found.');
        return;
    }

    currentEditingContactId = contactId;

    document.getElementById('contact-broker-name').value = contact.brokerName;
    document.getElementById('contact-first-name').value = contact.firstName;
    document.getElementById('contact-last-name').value = contact.lastName;
    document.getElementById('contact-email').value = contact.email;

    document.getElementById('form-section-title').textContent = 'Edit Contact';
    document.getElementById('save-contact-btn').textContent = 'Update Contact';

    // Scroll to form for better visibility
    document.getElementById('contact-form-section').scrollIntoView({ behavior: 'smooth' });
}

/**
 * Delete contact
 */
async function deleteContact(contactId) {
    const contact = allContacts.find(c => c.id === contactId);
    if (!contact) {
        alert('Contact not found.');
        return;
    }

    if (!confirm(`Are you sure you want to delete the contact for ${contact.firstName} ${contact.lastName} (${contact.brokerName})?`)) {
        return;
    }

    try {
        const success = await window.deleteBrokerContact(contactId);
        if (success) {
            await loadAndDisplayContacts();
            alert('Contact deleted successfully!');
        } else {
            alert('Error deleting contact. Please try again.');
        }
    } catch (error) {
        console.error('Error deleting contact:', error);
        alert('Error deleting contact: ' + error.message);
    }
}

/**
 * Reset contact form
 */
function resetContactForm() {
    currentEditingContactId = null;
    document.getElementById('contact-form').reset();
    document.getElementById('form-section-title').textContent = 'Add New Contact';
    document.getElementById('save-contact-btn').textContent = 'Save Contact';
}

/**
 * Cancel contact form
 */
function cancelContactForm() {
    resetContactForm();
}

/**
 * Export contacts as JSON
 */
async function exportContacts() {
    try {
        if (!allContacts || allContacts.length === 0) {
            alert('No contacts to export.');
            return;
        }

        await window.exportBrokerContactsAsJSON(allContacts, null, appSettings);
        alert(`Exported ${allContacts.length} contacts successfully!`);
    } catch (error) {
        console.error('Error exporting contacts:', error);
        alert('Error exporting contacts: ' + error.message);
    }
}

/**
 * Import contacts from JSON
 */
async function importContacts(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        const result = await window.loadBrokerContactsFromJSON(file);

        if (result.success) {
            // Save imported contacts to IndexedDB
            let savedCount = 0;
            for (const contact of result.contacts) {
                const success = await window.saveBrokerContact(contact);
                if (success) savedCount++;
            }

            // Refresh contact list
            await loadAndDisplayContacts();

            let message = `Successfully imported ${savedCount} contacts.`;
            if (result.skipped > 0) {
                message += `\n${result.skipped} contacts were skipped due to missing required fields.`;
            }
            alert(message);
        } else {
            alert('Error importing contacts: ' + result.error);
        }
    } catch (error) {
        console.error('Error importing contacts:', error);
        alert('Error importing contacts: ' + error.message);
    }

    // Reset file input
    event.target.value = '';
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ========== EMAIL BROKERS FUNCTIONALITY ==========

let emailUploadedFiles = [];
let missingDataAnalysis = [];
let pendingContactMatching = [];

/**
 * Initialize email brokers tab functionality
 */
function initializeEmailBrokersTab() {
    // Data source radio buttons
    const useResultsData = document.getElementById('use-results-data');
    const useUploadData = document.getElementById('use-upload-data');
    const emailUploadSection = document.getElementById('email-upload-section');

    useResultsData.addEventListener('change', () => {
        if (useResultsData.checked) {
            emailUploadSection.style.display = 'none';
            analyzeCurrentResultsData();
        }
    });

    useUploadData.addEventListener('change', () => {
        if (useUploadData.checked) {
            emailUploadSection.style.display = 'block';
        }
    });

    // File upload functionality
    const emailUploadZone = document.getElementById('email-upload-zone');
    const emailFileInput = document.getElementById('email-file-input');
    const emailBrowseBtn = document.getElementById('email-browse-btn');
    const emailClearFilesBtn = document.getElementById('email-clear-files-btn');

    emailUploadZone.addEventListener('click', () => emailFileInput.click());
    emailBrowseBtn.addEventListener('click', () => emailFileInput.click());
    emailClearFilesBtn.addEventListener('click', clearEmailFiles);

    emailUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        emailUploadZone.style.background = '#444';
    });

    emailUploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        emailUploadZone.style.background = '#333';
    });

    emailUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        emailUploadZone.style.background = '#333';
        handleEmailFilesDrop(e.dataTransfer.files);
    });

    emailFileInput.addEventListener('change', (e) => {
        handleEmailFilesDrop(e.target.files);
    });

    // Analysis and email generation
    const refreshAnalysisBtn = document.getElementById('refresh-analysis-btn');
    const sendEmailsBtn = document.getElementById('send-emails-btn');
    const saveEmailTemplateBtn = document.getElementById('save-email-template-btn');
    const resetEmailTemplateBtn = document.getElementById('reset-email-template-btn');

    refreshAnalysisBtn.addEventListener('click', refreshMissingDataAnalysis);
    sendEmailsBtn.addEventListener('click', generateBrokerEmails);
    saveEmailTemplateBtn.addEventListener('click', saveEmailTemplate);
    resetEmailTemplateBtn.addEventListener('click', resetEmailTemplate);

    // Contact matching modal
    const closeContactMatching = document.getElementById('close-contact-matching');
    const skipContactBtn = document.getElementById('skip-contact-btn');
    const confirmContactBtn = document.getElementById('confirm-contact-btn');

    closeContactMatching.addEventListener('click', closeContactMatchingModal);
    skipContactBtn.addEventListener('click', skipCurrentBrokerContact);
    confirmContactBtn.addEventListener('click', confirmBrokerContact);

    // Filename assignment modal
    const closeFilenameAssignment = document.getElementById('close-filename-assignment');
    const skipFilenameAssignments = document.getElementById('skip-filename-assignments');
    const confirmFilenameAssignments = document.getElementById('confirm-filename-assignments');

    closeFilenameAssignment.addEventListener('click', closeFilenameAssignmentModal);
    skipFilenameAssignments.addEventListener('click', skipFilenameAssignmentProcess);
    confirmFilenameAssignments.addEventListener('click', confirmFilenameAssignmentProcess);

    // Email template will be loaded after settings are loaded from IndexedDB

    // Initialize with current results data if available and "Use current Results data" is selected
    setTimeout(() => {
        if (useResultsData.checked && window.currentCombinedData && window.currentCombinedData.length > 0) {
            analyzeCurrentResultsData();
        }
    }, 100); // Small delay to ensure DOM is ready
}

/**
 * Handle email files drop/upload
 */
async function handleEmailFilesDrop(files) {
    const analysisStatus = document.getElementById('analysis-status');
    const analysisProgress = document.getElementById('analysis-progress');

    analysisStatus.style.display = 'block';
    analysisProgress.innerHTML = '<p style="color: #ffa726;">Processing uploaded files...</p>';

    emailUploadedFiles = [];

    for (let file of files) {
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            alert(`Skipping ${file.name}: Only Excel files (.xlsx, .xls) are supported.`);
            continue;
        }

        try {
            // Read and validate file against active template
            const workbook = await readExcelFile(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(worksheet);

            if (data.length === 0) {
                alert(`File ${file.name} appears to be empty.`);
                continue;
            }

            // Check if columns match active template
            const templateColumns = window.borderellenTemplate?.columns || [];
            const fileColumns = Object.keys(data[0]);
            const templateColumnNames = templateColumns.map(col => col.name);

            // Check for exact column match
            const missingColumns = templateColumnNames.filter(col => !fileColumns.includes(col));
            const extraColumns = fileColumns.filter(col => !templateColumnNames.includes(col));

            if (missingColumns.length > 0 || extraColumns.length > 0) {
                let message = `File ${file.name} does not match the active template:\n`;
                if (missingColumns.length > 0) {
                    message += `Missing columns: ${missingColumns.join(', ')}\n`;
                }
                if (extraColumns.length > 0) {
                    message += `Extra columns: ${extraColumns.join(', ')}\n`;
                }
                message += '\nPlease ensure the file was generated using this application with the same template.';
                alert(message);
                continue;
            }

            emailUploadedFiles.push({
                name: file.name,
                data: data,
                workbook: workbook
            });

        } catch (error) {
            console.error('Error processing file:', file.name, error);
            alert(`Error processing ${file.name}: ${error.message}`);
        }
    }

    if (emailUploadedFiles.length > 0) {
        analysisProgress.innerHTML = `<p style="color: #4caf50;">Successfully loaded ${emailUploadedFiles.length} file(s). Running missing data analysis...</p>`;
        await analyzeMissingData(emailUploadedFiles);
    } else {
        analysisProgress.innerHTML = '<p style="color: #f44336;">No valid files were processed.</p>';
    }
}

/**
 * Clear uploaded files for email analysis
 */
function clearEmailFiles() {
    emailUploadedFiles = [];
    document.getElementById('email-file-input').value = '';
    document.getElementById('analysis-status').style.display = 'none';
    document.getElementById('email-template-section').style.display = 'none';
    document.getElementById('missing-data-results').style.display = 'none';
}

/**
 * Analyze current results data for missing values
 */
async function analyzeCurrentResultsData() {
    if (!window.currentCombinedData || window.currentCombinedData.length === 0) {
        alert('No results data available. Please process some broker files first.');
        return;
    }

    const analysisStatus = document.getElementById('analysis-status');
    const analysisProgress = document.getElementById('analysis-progress');

    analysisStatus.style.display = 'block';
    analysisProgress.innerHTML = '<p style="color: #ffa726;">Analyzing current results data...</p>';

    // Convert current combined data to file-like format for analysis
    // Group by original source file to maintain filename associations
    const groupedByFile = {};

    window.currentCombinedData.forEach(record => {
        const sourceFile = record._sourceFile || 'Unknown File';
        if (!groupedByFile[sourceFile]) {
            groupedByFile[sourceFile] = [];
        }
        groupedByFile[sourceFile].push(record);
    });

    // Convert to file data format
    const fileData = Object.keys(groupedByFile).map(filename => ({
        name: filename,
        data: groupedByFile[filename]
    }));

    await analyzeMissingData(fileData);
}

/**
 * Analyze missing data for all files
 */
async function analyzeMissingData(filesData) {
    try {
        const templateColumns = window.borderellenTemplate?.columns || [];
        const requiredColumns = templateColumns.filter(col => col.required);

        if (requiredColumns.length === 0) {
            alert(`No required columns are defined in the active template.\n\nTemplate has ${templateColumns.length} total columns, but none are marked as "required".\n\nTo fix this:\n1. Go to Template Manager (Tab 1)\n2. Check the "Required" boxes for mandatory columns\n3. Save the template`);
            return;
        }

        console.log(`Email analysis: Found ${requiredColumns.length} required columns out of ${templateColumns.length} total columns`);
        console.log('Required columns:', requiredColumns.map(col => col.name));

        missingDataAnalysis = [];

        // Group data by broker
        const brokerGroups = {};

        for (const fileData of filesData) {
            for (const row of fileData.data) {
                const brokerName = row.Makelaar || row['Broker Name'] || 'Unknown';

                if (!brokerGroups[brokerName]) {
                    brokerGroups[brokerName] = {
                        brokerName: brokerName,
                        rows: [],
                        filenames: [],
                        primaryFilename: fileData.name
                    };
                }

                // Add filename if not already present
                if (!brokerGroups[brokerName].filenames.includes(fileData.name)) {
                    brokerGroups[brokerName].filenames.push(fileData.name);
                }

                brokerGroups[brokerName].rows.push({ ...row, _sourceFilename: fileData.name });
            }
        }

        // Analyze missing data for each broker
        for (const [brokerName, brokerData] of Object.entries(brokerGroups)) {
            // Create a display filename
            const displayFilename = brokerData.filenames.length === 1
                ? brokerData.filenames[0]
                : `${brokerData.filenames.length} files: ${brokerData.filenames.join(', ')}`;

            // Find the best filename for this broker (one containing broker name)
            const brokerSpecificFilename = findBrokerSpecificFilename(brokerName, brokerData.filenames);

            const analysis = {
                brokerName: brokerName,
                filename: displayFilename,
                emailFilename: brokerSpecificFilename || brokerData.primaryFilename,
                primaryFilename: brokerData.primaryFilename,
                allFilenames: brokerData.filenames,
                totalRows: brokerData.rows.length,
                missingColumns: [],
                overallCompletionRate: 0
            };

            let totalRequiredFields = 0;
            let totalFilledFields = 0;

            for (const column of requiredColumns) {
                const columnName = column.name;
                let filledCount = 0;

                for (const row of brokerData.rows) {
                    const value = row[columnName];
                    if (value !== null && value !== undefined && value !== '') {
                        filledCount++;
                    }
                }

                const completionRate = (filledCount / brokerData.rows.length) * 100;

                analysis.missingColumns.push({
                    columnName: columnName,
                    filledCount: filledCount,
                    totalCount: brokerData.rows.length,
                    completionRate: completionRate
                });

                totalRequiredFields += brokerData.rows.length;
                totalFilledFields += filledCount;
            }

            analysis.overallCompletionRate = (totalFilledFields / totalRequiredFields) * 100;

            // Only include brokers that have missing data
            if (analysis.overallCompletionRate < 100) {
                missingDataAnalysis.push(analysis);
            }
        }

        await displayMissingDataResults();

    } catch (error) {
        console.error('Error analyzing missing data:', error);
        document.getElementById('analysis-progress').innerHTML = `<p style="color: #f44336;">Error during analysis: ${error.message}</p>`;
    }
}

/**
 * Display missing data analysis results
 */
async function displayMissingDataResults() {
    const analysisProgress = document.getElementById('analysis-progress');
    const emailTemplateSection = document.getElementById('email-template-section');
    const missingDataResults = document.getElementById('missing-data-results');
    const summaryDiv = document.getElementById('broker-analysis-summary');
    const tableBody = document.getElementById('missing-data-table-body');

    if (missingDataAnalysis.length === 0) {
        analysisProgress.innerHTML = '<p style="color: #4caf50;">✓ All brokers have complete data - no emails needed!</p>';
        emailTemplateSection.style.display = 'none';
        missingDataResults.style.display = 'none';
        return;
    }

    analysisProgress.innerHTML = `<p style="color: #4caf50;">✓ Analysis complete - found ${missingDataAnalysis.length} broker(s) with missing data</p>`;

    // Show summary
    const totalBrokers = missingDataAnalysis.length;
    const avgCompletion = missingDataAnalysis.reduce((sum, broker) => sum + broker.overallCompletionRate, 0) / totalBrokers;

    summaryDiv.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; background: #333; padding: 16px; border-radius: 8px;">
            <div>
                <div style="color: #00bcd4; font-size: 24px; font-weight: bold;">${totalBrokers}</div>
                <div style="color: #888; font-size: 14px;">Brokers with missing data</div>
            </div>
            <div>
                <div style="color: #ffa726; font-size: 24px; font-weight: bold;">${avgCompletion.toFixed(1)}%</div>
                <div style="color: #888; font-size: 14px;">Average completion rate</div>
            </div>
        </div>
    `;

    // Display table
    tableBody.innerHTML = await Promise.all(missingDataAnalysis.map(async (broker) => {
        const contact = await findBrokerContact(broker.brokerName);
        const missingColumnsText = broker.missingColumns
            .map(col => `${col.columnName} (${col.completionRate.toFixed(1)}%)`)
            .join(', ');

        return `
            <tr>
                <td>${escapeHtml(broker.brokerName)}</td>
                <td>${contact ? escapeHtml(`${contact.firstName} ${contact.lastName}`) : '<span style="color: #f44336;">No contact found</span>'}</td>
                <td>${broker.totalRows}</td>
                <td style="max-width: 300px; word-wrap: break-word;">${missingColumnsText}</td>
                <td>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <div style="background: #555; border-radius: 4px; overflow: hidden; flex: 1; height: 8px;">
                            <div style="background: ${broker.overallCompletionRate >= 80 ? '#4caf50' : broker.overallCompletionRate >= 60 ? '#ffa726' : '#f44336'}; width: ${broker.overallCompletionRate}%; height: 100%;"></div>
                        </div>
                        <span style="font-size: 12px; color: #888; min-width: 40px;">${broker.overallCompletionRate.toFixed(1)}%</span>
                    </div>
                </td>
                <td>
                    <button class="btn btn-secondary" onclick="previewEmail('${broker.brokerName}')" style="font-size: 12px; padding: 4px 8px;">Preview Email</button>
                </td>
            </tr>
        `;
    })).then(rows => rows.join(''));

    emailTemplateSection.style.display = 'block';
    missingDataResults.style.display = 'block';
}

/**
 * Find broker contact by name
 */
async function findBrokerContact(brokerName) {
    try {
        const contacts = await window.loadAllBrokerContacts();
        return contacts.find(contact =>
            contact.brokerName.toLowerCase().trim() === brokerName.toLowerCase().trim()
        );
    } catch (error) {
        console.error('Error loading contacts:', error);
        return null;
    }
}

/**
 * Refresh missing data analysis
 */
async function refreshMissingDataAnalysis() {
    const useResultsData = document.getElementById('use-results-data');

    if (useResultsData.checked) {
        await analyzeCurrentResultsData();
    } else if (emailUploadedFiles.length > 0) {
        await analyzeMissingData(emailUploadedFiles);
    } else {
        alert('No data available for analysis. Please select a data source.');
    }
}

/**
 * Validate all placeholders before email generation
 */
async function validateEmailPlaceholders() {
    const subject = document.getElementById('email-subject').value;
    const bodyTemplate = document.getElementById('email-body-template').value;

    const errors = [];
    const brokersNeedingFilenames = [];

    // Check user settings
    if (!appSettings.userEmail || appSettings.userEmail.trim() === '') {
        errors.push('User email address is not configured. Please set it in Settings.');
    }

    // Validate email template placeholders for each broker
    for (const broker of missingDataAnalysis) {
        const contact = broker.manualContact || await findBrokerContact(broker.brokerName);

        if (!contact && bodyTemplate.includes('{contact_last_name}')) {
            errors.push(`No contact found for broker "${broker.brokerName}". Contact is needed for {contact_last_name} placeholder.`);
        }

        if ((!broker.emailFilename || broker.emailFilename.trim() === '') && bodyTemplate.includes('{filename}')) {
            brokersNeedingFilenames.push(broker);
        }
    }

    // If brokers need filenames and user can provide them, show filename assignment modal
    if (brokersNeedingFilenames.length > 0 && errors.length === 0) {
        const proceed = await showFilenameAssignmentModal(brokersNeedingFilenames);
        if (!proceed) {
            errors.push('Filename assignment was cancelled.');
        }
    }

    return errors;
}

/**
 * Show filename assignment modal
 */
async function showFilenameAssignmentModal(brokers) {
    return new Promise((resolve) => {
        const modal = document.getElementById('filename-assignment-modal');
        const container = document.getElementById('filename-assignments-container');

        // Store resolve function for later use
        window.filenameAssignmentResolve = resolve;

        // Create form for each broker
        container.innerHTML = brokers.map(broker => `
            <div class="form-group" style="border: 1px solid #444; padding: 12px; border-radius: 4px; margin-bottom: 12px;">
                <label class="form-label">Filename for "${broker.brokerName}"</label>
                <input type="text" class="form-input filename-input" data-broker="${broker.brokerName}"
                       placeholder="Enter filename (e.g., ${broker.brokerName}_borderel.xlsx)"
                       value="${broker.primaryFilename || ''}" />
                <small style="color: #888; display: block; margin-top: 4px;">
                    Available files: ${broker.allFilenames.join(', ')}
                </small>
            </div>
        `).join('');

        modal.classList.add('show');
    });
}

/**
 * Close filename assignment modal
 */
function closeFilenameAssignmentModal() {
    document.getElementById('filename-assignment-modal').classList.remove('show');
    if (window.filenameAssignmentResolve) {
        window.filenameAssignmentResolve(false);
    }
}

/**
 * Skip filename assignment process
 */
function skipFilenameAssignmentProcess() {
    // Use default names (broker name + "_borderel.xlsx")
    const inputs = document.querySelectorAll('.filename-input');
    inputs.forEach(input => {
        const brokerName = input.dataset.broker;
        const defaultFilename = `${brokerName}_borderel.xlsx`;

        // Find the broker and set default filename
        const broker = missingDataAnalysis.find(b => b.brokerName === brokerName);
        if (broker) {
            broker.emailFilename = defaultFilename;
        }
    });

    closeFilenameAssignmentModal();
    if (window.filenameAssignmentResolve) {
        window.filenameAssignmentResolve(true);
    }
}

/**
 * Confirm filename assignment process
 */
function confirmFilenameAssignmentProcess() {
    const inputs = document.querySelectorAll('.filename-input');
    let allValid = true;

    inputs.forEach(input => {
        const brokerName = input.dataset.broker;
        const filename = input.value.trim();

        if (!filename) {
            allValid = false;
            return;
        }

        // Find the broker and set filename
        const broker = missingDataAnalysis.find(b => b.brokerName === brokerName);
        if (broker) {
            broker.emailFilename = filename;
        }
    });

    if (!allValid) {
        alert('Please provide filenames for all brokers or use "Use Default Names".');
        return;
    }

    closeFilenameAssignmentModal();
    if (window.filenameAssignmentResolve) {
        window.filenameAssignmentResolve(true);
    }
}

/**
 * Generate emails for all brokers with missing data
 */
async function generateBrokerEmails() {
    if (missingDataAnalysis.length === 0) {
        alert('No brokers with missing data found.');
        return;
    }

    // First, validate all placeholders
    const validationErrors = await validateEmailPlaceholders();
    if (validationErrors.length > 0) {
        alert('Please fix the following issues before generating emails:\n\n' + validationErrors.join('\n'));
        return;
    }

    pendingContactMatching = [];

    // Check which brokers need contact selection
    for (const broker of missingDataAnalysis) {
        const contact = await findBrokerContact(broker.brokerName);
        if (!contact) {
            pendingContactMatching.push(broker);
        }
    }

    if (pendingContactMatching.length > 0) {
        await handleContactMatching();
    } else {
        await generateAllEmails();
    }
}

/**
 * Handle contact matching for brokers without exact matches
 */
async function handleContactMatching() {
    if (pendingContactMatching.length === 0) {
        await generateAllEmails();
        return;
    }

    const broker = pendingContactMatching[0];
    await showContactMatchingModal(broker);
}

/**
 * Show contact matching modal
 */
async function showContactMatchingModal(broker) {
    const modal = document.getElementById('contact-matching-modal');
    const messageEl = document.getElementById('contact-matching-message');
    const selector = document.getElementById('broker-contact-selector');

    messageEl.textContent = `No exact contact match found for broker "${broker.brokerName}". Please select a contact or skip this broker.`;

    // Populate contact selector
    const contacts = await window.loadAllBrokerContacts();
    selector.innerHTML = '<option value="">Choose a contact...</option>' +
        contacts.map(contact =>
            `<option value="${contact.id}">${contact.brokerName} - ${contact.firstName} ${contact.lastName}</option>`
        ).join('');

    modal.classList.add('show');
}

/**
 * Close contact matching modal
 */
function closeContactMatchingModal() {
    document.getElementById('contact-matching-modal').classList.remove('show');
}

/**
 * Skip current broker contact
 */
async function skipCurrentBrokerContact() {
    pendingContactMatching.shift(); // Remove first broker
    closeContactMatchingModal();
    await handleContactMatching(); // Continue with next broker
}

/**
 * Confirm broker contact selection
 */
async function confirmBrokerContact() {
    const selector = document.getElementById('broker-contact-selector');
    const selectedContactId = selector.value;

    if (!selectedContactId) {
        alert('Please select a contact or skip this broker.');
        return;
    }

    const broker = pendingContactMatching[0];
    const contacts = await window.loadAllBrokerContacts();
    const selectedContact = contacts.find(c => c.id === selectedContactId);

    if (selectedContact) {
        // Store the manual mapping for this broker
        broker.manualContact = selectedContact;
    }

    pendingContactMatching.shift(); // Remove first broker
    closeContactMatchingModal();
    await handleContactMatching(); // Continue with next broker
}

/**
 * Generate all emails
 */
async function generateAllEmails() {
    const subject = document.getElementById('email-subject').value;
    const bodyTemplate = document.getElementById('email-body-template').value;

    let emailsGenerated = 0;

    for (const broker of missingDataAnalysis) {
        const contact = broker.manualContact || await findBrokerContact(broker.brokerName);

        if (contact) {
            await generateSingleEmail(broker, contact, subject, bodyTemplate);
            emailsGenerated++;
        }
    }

    alert(`Generated ${emailsGenerated} email(s). Check your email client.`);
}

/**
 * Generate single email for a broker
 */
async function generateSingleEmail(broker, contact, subject, bodyTemplate) {
    try {
        // Use template system for all email generation (including table)
        await createAndDownloadEMLFileFromTemplate(
            contact.email,
            subject,
            bodyTemplate,
            contact.lastName,
            broker.emailFilename || broker.filename,
            broker.totalRows,
            broker.missingColumns, // Pass all required columns, not just incomplete ones
            appSettings.userSignature || appSettings.userName || 'User'
        );

        // Small delay between emails to avoid overwhelming the system
        await new Promise(resolve => setTimeout(resolve, 1000));

    } catch (error) {
        console.error('Error generating email for broker:', broker.brokerName, error);
    }
}

/**
 * Preview email for a specific broker
 */
async function previewEmail(brokerName) {
    const broker = missingDataAnalysis.find(b => b.brokerName === brokerName);
    if (!broker) return;

    const contact = broker.manualContact || await findBrokerContact(broker.brokerName);
    if (!contact) {
        alert('No contact found for this broker. Please run the email generation process to select a contact.');
        return;
    }

    const subject = document.getElementById('email-subject').value;
    const bodyTemplate = document.getElementById('email-body-template').value;

    // Use template system for email preview
    await createAndDownloadEMLFileFromTemplate(
        contact.email,
        subject,
        bodyTemplate,
        contact.lastName,
        broker.emailFilename || broker.filename,
        broker.totalRows,
        broker.missingColumns, // Show ALL required columns, not just incomplete ones
        appSettings.userSignature || appSettings.userName || 'User'
    );
}

/**
 * Create and download EML file using email template with placeholders
 */
async function createAndDownloadEMLFileFromTemplate(toEmail, subject, bodyTemplate, contactLastName, filename, totalRows, missingColumns, userSignature) {
    // Process template with placeholders
    const htmlEmailBody = processEmailTemplate(bodyTemplate, {
        contact_email: toEmail,
        contact_last_name: contactLastName,
        filename: filename,
        total_rows: totalRows,
        missing_data_table: createMissingDataTableHTML(missingColumns),
        user_signature: userSignature
    });

    // Create EML file content with proper headers
    const emlContent = `To: ${toEmail}
From: ${appSettings.userEmail || 'noreply@borderellenconverter.nl'}
Subject: ${subject}
Date: ${new Date().toUTCString()}
MIME-Version: 1.0
Content-Type: text/html; charset=UTF-8
Content-Transfer-Encoding: 8bit

${htmlEmailBody}`;

    // Generate filename based on broker name with date and time
    const brokerName = filename.split(/[._]/)[0] || 'Email';
    const now = new Date();
    const timestamp = now.toISOString().slice(0, 19).replace(/[:.]/g, '-');
    const emlFilename = `${brokerName}_OntbrekendeData_${timestamp}.eml`;

    // Create blob and download EML file
    const blob = new Blob([emlContent], { type: 'message/rfc822' });

    // Download EML file using the existing download helper
    const success = await downloadToPreferredFolder(blob, emlFilename, 'message/rfc822');

    if (success) {
        console.log(`EML file created: ${emlFilename}. Please open manually from your downloads folder.`);
    }
}

/**
 * Process email template by replacing placeholders with actual values
 * @param {string} template - Email template with placeholders
 * @param {Object} values - Values to replace placeholders with
 * @returns {string} Processed HTML email content
 */
function processEmailTemplate(template, values) {
    // First, separate signature from main content BEFORE any processing
    let mainTemplate = template;
    let signatureTemplate = '';

    // Find signature placeholder and split there
    if (template.includes('{user_signature}')) {
        const signatureIndex = template.indexOf('{user_signature}');
        // Find the start of the line containing the signature
        const beforeSignature = template.substring(0, signatureIndex);
        const lastNewline = beforeSignature.lastIndexOf('\n');

        mainTemplate = template.substring(0, lastNewline >= 0 ? lastNewline : signatureIndex);
        signatureTemplate = template.substring(lastNewline >= 0 ? lastNewline : signatureIndex);
    }

    // Process main template (everything except signature)
    let processedMainContent = mainTemplate;
    Object.keys(values).forEach(placeholder => {
        if (placeholder === 'user_signature') return; // Skip signature in main content

        const value = values[placeholder];
        const regex = new RegExp(`\\{${placeholder}\\}`, 'g');

        if (placeholder === 'contact_email') {
            // Just insert the email address without special styling
            processedMainContent = processedMainContent.replace(regex, value);
        } else if (placeholder === 'missing_data_table') {
            // Special handling for missing_data_table - add tip box directly after it
            const tableWithTip = value + `<!-- Tip Box -->
<div style="background-color: #e8f4f8; border-left: 4px solid rgb(0, 55, 129); padding: 12px 16px; margin: 8px 0;">
    <p style="font-family: Arial, sans-serif; font-size: 10pt; color: black; margin: 0; font-weight: bold;">
        💡 Tip: Een volledig ingevuld borderellen bestand versnelt de verwerking aanzienlijk!
    </p>
</div>`;
            processedMainContent = processedMainContent.replace(regex, tableWithTip);
        } else {
            // Regular placeholder replacement - all body text should be black
            let styledValue = value;
            if (placeholder === 'contact_last_name' || placeholder === 'filename') {
                styledValue = `<strong style="color: black;">${value}</strong>`;
            } else if (placeholder === 'total_rows') {
                styledValue = `<strong style="color: black;">${value}</strong>`;
            }
            processedMainContent = processedMainContent.replace(regex, styledValue);
        }
    });

    // Process signature separately and keep it as-is (Allianz blue)
    let processedSignature = '';
    if (signatureTemplate.includes('{user_signature}')) {
        processedSignature = signatureTemplate.replace(/{user_signature}/g, createHTMLSignature(values.user_signature));
    }

    // Wrap main content in paragraphs, but preserve signature HTML structure
    const wrappedMainContent = processedMainContent.split('\n').map(line =>
        line.trim() ? `<p style="font-family: Arial, sans-serif; font-size: 10pt; color: black; margin: 0 0 4px 0;">${line.trim()}</p>` : '<br>'
    ).join('');

    return `<div style="font-family: Arial, sans-serif; font-size: 10pt; color: black; line-height: 1.6; max-width: 600px;">
${wrappedMainContent}${processedSignature}
</div>`;
}

/**
 * Create HTML for missing data table
 * @param {Array} missingColumns - Array of missing column data
 * @returns {string} HTML table string
 */
function createMissingDataTableHTML(allRequiredColumns) {
    if (!allRequiredColumns || allRequiredColumns.length === 0) {
        return `
<div style="background-color: #f1f8e9; padding: 15px; border-radius: 6px; border: 1px solid #c8e6c9; text-align: center; margin: 4px 0;">
    <h4 style="font-family: Arial, sans-serif; font-size: 10pt; color: rgb(0, 55, 129); margin: 0 0 8px 0; font-weight: bold;">
        ✅ Uitstekend!
    </h4>
    <p style="font-family: Arial, sans-serif; font-size: 10pt; color: #388e3c; margin: 0; font-weight: bold;">
        Alle verplichte velden zijn volledig ingevuld.
    </p>
</div>`;
    }

    // Calculate totals for the summary row
    let totalFilled = 0;
    let totalFields = 0;

    allRequiredColumns.forEach(col => {
        totalFilled += col.filledCount;
        totalFields += col.totalCount;
    });

    const overallPercentage = totalFields > 0 ? (totalFilled / totalFields * 100).toFixed(1) : '0.0';

    let tableHTML = `<div style="margin: 0 0 4px 0;">
    <table cellpadding="3" cellspacing="0" style="border-collapse: collapse; width: auto; font-family: Arial, sans-serif; font-size: 10pt; border: 1px solid #dee2e6;">
        <thead>
            <tr style="background-color: #e8f4f8;">
                <th style="text-align: left; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Gegevensveld</th>
                <th style="text-align: center; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Gevuld</th>
                <th style="text-align: center; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Totaal</th>
                <th style="text-align: center; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Percentage</th>
            </tr>
        </thead>
        <tbody>`;

    // Show ALL required columns (complete and incomplete)
    allRequiredColumns.forEach((col, index) => {
        const percentage = col.completionRate.toFixed(1);
        const rowBg = index % 2 === 0 ? '#ffffff' : '#f8f9fa';

        // Color-code the percentage based on completion
        const percentageColor = col.completionRate >= 100 ? '#4caf50' : col.completionRate >= 80 ? '#ffa726' : '#f44336';

        tableHTML += `
            <tr style="background-color: ${rowBg};">
                <td style="padding: 3px 3px; color: black; border: 1px solid #dee2e6;">${col.columnName}</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6;">${col.filledCount}</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6;">${col.totalCount}</td>
                <td style="text-align: center; padding: 3px 3px; color: ${percentageColor}; border: 1px solid #dee2e6; font-weight: bold;">${percentage}%</td>
            </tr>`;
    });

    // Add total score row
    const totalRowBg = '#e8f4f8';
    const totalPercentageColor = parseFloat(overallPercentage) >= 100 ? '#4caf50' : parseFloat(overallPercentage) >= 80 ? '#ffa726' : '#f44336';

    tableHTML += `
            <tr style="background-color: ${totalRowBg}; border-top: 2px solid #dee2e6;">
                <td style="padding: 3px 3px; color: black; border: 1px solid #dee2e6; font-weight: bold;">TOTAAL SCORE</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6; font-weight: bold;">${totalFilled}</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6; font-weight: bold;">${totalFields}</td>
                <td style="text-align: center; padding: 3px 3px; color: ${totalPercentageColor}; border: 1px solid #dee2e6; font-weight: bold; font-size: 11pt;">${overallPercentage}%</td>
            </tr>`;

    tableHTML += `
        </tbody>
    </table>
</div>`;

    return tableHTML;
}

/**
 * Create HTML signature with Allianz branding
 * @param {string} userSignature - User's signature text (can contain line breaks)
 * @returns {string} HTML formatted signature
 */
function createHTMLSignature(userSignature) {
    // Parse signature for structured information
    const signatureLines = userSignature.split('\n').filter(line => line.trim());

    if (signatureLines.length === 0) {
        return `<p style="font-family: Arial, sans-serif; font-size: 10pt; color: rgb(0, 55, 129);">${userSignature}</p>`;
    }

    // Try to detect if signature has structured format (Name, Function, Company details)
    const hasStructuredFormat = signatureLines.length >= 2;

    if (hasStructuredFormat && signatureLines.length >= 3) {
        // Structured signature: Name, Function, Company details
        const [name, functionTitle, ...companyDetails] = signatureLines;

        return `
<div style="font-family: Arial, sans-serif; font-size: 10pt; margin-top: 4px;">
    <div style="color: rgb(0, 55, 129); font-weight: bold; font-size: 10pt; margin-bottom: 2px;">
        ${name}
    </div>
    <div style="color: rgb(0, 55, 129); font-size: 10pt; margin-bottom: 2px;">
        ${functionTitle}
    </div>
    <div style="color: rgb(0, 55, 129); font-size: 10pt; margin-bottom: 10px;">
        ${companyDetails.join('<br>')}
    </div>
    ${createSocialMediaLinks()}
</div>`;
    } else {
        // Simple signature - ALWAYS Allianz blue (NEVER black)
        return `
<div style="font-family: Arial, sans-serif; font-size: 10pt; margin-top: 4px;">
    <div style="color: rgb(0, 55, 129); font-weight: bold; font-size: 10pt; margin-bottom: 10px;">
        ${signatureLines[0]}
    </div>
    ${signatureLines.slice(1).map(line =>
        `<div style="color: rgb(0, 55, 129); font-size: 10pt; margin-bottom: 2px;">${line}</div>`
    ).join('')}
    ${createSocialMediaLinks()}
</div>`;
    }
}

/**
 * Create social media links section according to Allianz branding
 * @returns {string} HTML for social media icons and links
 */
function createSocialMediaLinks() {
    // Social media links temporarily disabled - Belgian channels need proper icons
    return '';
}

/**
 * Find the most appropriate filename for a broker
 * @param {string} brokerName - The broker name
 * @param {Array} filenames - Available filenames
 * @returns {string} Best matching filename
 */
function findBrokerSpecificFilename(brokerName, filenames) {
    // First, try to find a filename containing the broker name (case insensitive)
    const brokerNameLower = brokerName.toLowerCase();
    const brokerWords = brokerNameLower.split(/\s+/);

    for (const filename of filenames) {
        const filenameLower = filename.toLowerCase();

        // Check if filename contains the full broker name
        if (filenameLower.includes(brokerNameLower)) {
            return filename;
        }

        // Check if filename contains any significant words from broker name (length > 2)
        for (const word of brokerWords) {
            if (word.length > 2 && filenameLower.includes(word)) {
                return filename;
            }
        }
    }

    // If no match found, return null so primary filename is used
    return null;
}

/**
 * Save email template to settings - Direct IndexedDB save
 */
async function saveEmailTemplate() {
    try {
        const subject = document.getElementById('email-subject').value;
        const body = document.getElementById('email-body-template').value;

        // Direct IndexedDB save without going through complex folder handling
        if (!window.db) await window.initIndexedDB();

        const transaction = window.db.transaction(['settings'], 'readwrite');
        const store = transaction.objectStore('settings');

        // Save email subject
        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'emailSubject', value: subject });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        // Save email body template
        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'emailBodyTemplate', value: body });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        // Update local appSettings
        appSettings.emailSubject = subject;
        appSettings.emailBodyTemplate = body;

        alert('Email template saved successfully!');
        console.log('Email template saved directly to IndexedDB');

    } catch (error) {
        console.error('Error saving email template:', error);
        alert('Error saving email template: ' + error.message);
    }
}

/**
 * Reset email template to default
 */
async function resetEmailTemplate() {
    if (confirm('Are you sure you want to reset the email template to default values?')) {
        const defaultSubject = 'Ontbrekende data in uw borderellen';
        const defaultBody = `{contact_email}

Geachte heer/mevrouw {contact_last_name},

Wij hebben uw borderellen bestand {filename} met {total_rows} boekingen in goede orde ontvangen.
Er ontbreekt echter data, waardoor wij niet in staat zijn deze boekingen snel en accuraat te verwerken.

Gelieve onderstaande gegevens aan te vullen en aan ons te retourneren.
{missing_data_table}

Bedankt voor uw medewerking.

Met vriendelijke groet,

{user_signature}`;

        // Update UI
        document.getElementById('email-subject').value = defaultSubject;
        document.getElementById('email-body-template').value = defaultBody;

        // Save to IndexedDB immediately
        try {
            appSettings.emailSubject = defaultSubject;
            appSettings.emailBodyTemplate = defaultBody;

            const success = await window.saveSettings(appSettings);
            if (success) {
                alert('Email template reset to default and saved successfully!');
            } else {
                alert('Email template was reset but could not be saved. Please try saving manually.');
            }
        } catch (error) {
            console.error('Error saving reset template:', error);
            alert('Email template was reset but could not be saved: ' + error.message);
        }
    }
}

/**
 * Load email template from settings
 */
function loadEmailTemplate() {
    console.log('Loading email template...');
    console.log('appSettings.emailSubject:', appSettings.emailSubject);
    console.log('appSettings.emailBodyTemplate exists:', !!appSettings.emailBodyTemplate);
    console.log('Full appSettings keys:', Object.keys(appSettings));

    const subjectEl = document.getElementById('email-subject');
    const bodyEl = document.getElementById('email-body-template');

    console.log('Elements found - subject:', !!subjectEl, 'body:', !!bodyEl);

    if (appSettings.emailSubject && subjectEl) {
        subjectEl.value = appSettings.emailSubject;
        console.log('✅ Set subject to:', appSettings.emailSubject);
    } else {
        console.log('❌ Did not set subject - emailSubject:', appSettings.emailSubject, 'element:', !!subjectEl);
    }

    if (appSettings.emailBodyTemplate && bodyEl) {
        bodyEl.value = appSettings.emailBodyTemplate;
        console.log('✅ Set body template (first 50 chars):', appSettings.emailBodyTemplate.substring(0, 50));
    } else {
        console.log('❌ Did not set body - emailBodyTemplate exists:', !!appSettings.emailBodyTemplate, 'element:', !!bodyEl);
    }
}

// Make functions globally accessible
window.previewEmail = previewEmail;

// Make functions globally accessible for onclick handlers
window.editContact = editContact;
window.deleteContact = deleteContact;

// Make manual header selection functions globally accessible
window.showManualHeaderSelection = showManualHeaderSelection;
window.closeManualHeaderSelection = closeManualHeaderSelection;
window.startRangeSelection = startRangeSelection;
window.updateRangeSelection = updateRangeSelection;
window.endRangeSelection = endRangeSelection;
window.applyHeaderSelection = applyHeaderSelection;
window.selectFooterKeyword = selectFooterKeyword;

// Calculation engine functions are now in calculationEngine.js

// Make modal functions globally accessible
window.closeFixedStringModal = closeFixedStringModal;
window.applyFixedString = applyFixedString;
window.closeCalculationModal = closeCalculationModal;
window.toggleCalculationValueType = toggleCalculationValueType;
window.updateCalculationPreview = updateCalculationPreview;
window.applyCalculation = applyCalculation;

// Export essential functions to window for cross-module access
window.updateMappingFileSelector = updateMappingFileSelector;
window.showEditTemplateSection = showEditTemplateSection;
window.showMainSection = showMainSection;
window.showNewTemplateSection = showNewTemplateSection;