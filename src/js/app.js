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
window.currentMappingFile = null; // Make this globally accessible
window.currentMapping = {}; // Make this globally accessible
window.currentPatternAnalysis = null; // Make this globally accessible

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
    if (typeof updateActiveTemplateDisplay === 'function') {
        updateActiveTemplateDisplay();
    }
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
    window.currentMapping = {};

    // Get broker-specific mappings
    const brokerMappings = getBrokerAutoMapping(brokerType, fileData.name);

    // Apply the mappings
    Object.assign(window.currentMapping, brokerMappings);

    console.log('Auto-mapping applied:', window.currentMapping);
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


async function loadSourceColumns(fileId) {
    const fileData = window.uploadedFiles?.find(f => f.id == fileId);
    if (!fileData) return;

    window.currentMappingFile = fileData;
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
                window.currentPatternAnalysis = patternAnalysis;

                if (patternAnalysis.confidence > 0.3 && patternAnalysis.headerAnalysis?.found) {
                    // High confidence - use detected structure
                    await loadSourceColumnsFromAnalysis(fileData.file, patternAnalysis, container);
                } else {
                    // Low confidence or analysis failed - show manual selection interface
                    await showManualColumnSelection(fileData, container);
                }

            } catch (analysisError) {
                console.warn('Pattern analysis failed, showing manual selection:', analysisError);
                await showManualColumnSelection(fileData, container);
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
    try {
        const workbook = await ExcelCacheManager.getWorkbook(file);
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

    } catch (error) {
        console.error('Error extracting columns from analysis:', error);
        throw error;
    }
}

/**
 * Show manual grid selection interface when analysis fails
 */
async function showManualColumnSelection(fileData, container) {
    // Store the file for use by the header selection modal
    window.currentMappingFile = fileData;

    // Show a simple interface that directs users to the advanced modal
    container.innerHTML = `
        <div style="background: #e8f4f8; padding: 24px; margin-bottom: 16px; border-radius: 8px; border-left: 4px solid #00bcd4; color: black; text-align: center;">
            <h4 style="color: #00bcd4; margin-bottom: 16px;">Manual Header & Footer Selection Required</h4>
            <p style="margin-bottom: 20px;">The automatic analysis couldn't detect the data structure reliably. Please use the manual selection tool to define your header range and footer keywords.</p>
            <button class="btn btn-primary" onclick="showManualHeaderSelection()" style="padding: 12px 24px; font-size: 16px;">
                Open Manual Header & Footer Selection
            </button>
        </div>
        <div style="background: #f8f9fa; padding: 16px; border-radius: 6px; color: #666;">
            <p style="margin: 0;"><strong>Instructions:</strong></p>
            <ul style="margin: 8px 0 0 20px;">
                <li>Click the button above to open the selection tool</li>
                <li>Drag to select your header range (can be multiple rows/columns)</li>
                <li>Click on footer keywords to stop processing at specific text</li>
                <li>Apply your selection to continue</li>
            </ul>
        </div>
    `;

    return Promise.resolve();
}

/**
 * Show enhanced manual header selection modal with range selection
 */


/**
 * Apply header selection
 */


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
 * Load headers from manual selection
 */






/**
 * Show fixed string input modal (replaces prompt)
 */

function clearMapping() {
    window.currentMapping = {};
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
    if (!window.currentMappingFile || !window.borderellenTemplate || !window.borderellenTemplate.columns) {
        alert('Please select a file and ensure a template is active.');
        return;
    }

    console.log('Starting matrix-based auto-mapping...');

    // Clear existing mapping
    window.currentMapping = {};

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
        window.currentMapping = result.mapping;
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
    if (!window.currentMappingFile || Object.keys(window.currentMapping).length === 0) {
        alert('Please select a file and create at least one mapping before saving.');
        return;
    }

    const templateName = prompt(
        `Enter a name for this broker template:`,
        `${window.currentMappingFile.broker?.name || window.currentMappingFile.name || 'Custom'} Template`
    );

    if (!templateName || templateName.trim() === '') {
        return; // User cancelled or entered empty name
    }

    // Check for duplicate template names
    try {
        const existingMappings = await window.loadAllFileMappings();
        const duplicateName = existingMappings.find(t => t.name === templateName.trim());

        if (duplicateName) {
            alert(
                `A file mapping named "${templateName}" already exists.\n\n` +
                `Template names must be unique to ensure predictable template selection.\n\n` +
                `Please choose a different name.`
            );
            return;
        }
    } catch (error) {
        console.error('Error checking for duplicate templates:', error);
    }

    // Extract suggested keyword from filename
    const suggestedKeyword = window.extractKeywordFromFilename ? window.extractKeywordFromFilename(window.currentMappingFile.name) : '';
    const defaultKeyword = suggestedKeyword || window.currentMappingFile.broker.name.split(' ')[0];

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

    // Check for duplicate keywords (if keyword is provided)
    if (finalKeyword) {
        try {
            const existingMappings = await window.loadAllFileMappings();
            const duplicateKeyword = existingMappings.find(t =>
                t.matchingKeyword &&
                t.matchingKeyword.toLowerCase() === finalKeyword.toLowerCase()
            );

            if (duplicateKeyword) {
                alert(
                    `A template with keyword "${finalKeyword}" already exists (Template: "${duplicateKeyword.name}").\n\n` +
                    `Keywords must be unique to ensure predictable template selection.\n\n` +
                    `Please choose a different keyword.`
                );
                return;
            }
        } catch (error) {
            console.error('Error checking for duplicate keywords:', error);
        }
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
        if (window.currentPatternAnalysis && window.currentPatternAnalysis.dataSection) {
            const analysis = window.currentPatternAnalysis.dataSection;
            parsingConfig.skipRows = analysis.dataStartIndex || 0;
            parsingConfig.skipColumns = analysis.startColumnIndex || 0;
            parsingConfig.headerRow = analysis.headerRowIndex;

            // Add enhanced parsingConfig fields for manual selections
            if (window.currentPatternAnalysis.manualSelection) {
                const manual = window.currentPatternAnalysis.manualSelection;
                parsingConfig.headerRows = manual.headerRows;
                parsingConfig.headerColumns = manual.headerColumns;
                parsingConfig.headerRange = manual.headerRange;

                // Add footer keyword detection if provided (manual takes priority)
                if (manual.footerKeyword) {
                    parsingConfig.footerRowKeyword = manual.footerKeyword;
                    console.log(`Footer detection configured: "${manual.footerKeyword}"`);
                } else if (window.currentPatternAnalysis.autoFooterKeyword) {
                    parsingConfig.footerRowKeyword = window.currentPatternAnalysis.autoFooterKeyword;
                    console.log(`Auto footer detection configured: "${window.currentPatternAnalysis.autoFooterKeyword}"`);
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
            } else if (window.currentPatternAnalysis.autoFooterKeyword) {
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
            sourceType: window.currentMappingFile.broker.parser || window.currentMappingFile.broker.type,
            sourceName: window.currentMappingFile.broker.name,
            filePattern: window.currentMappingFile.name,
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

        // Enable the "Process with Template" button now that template is saved
        const processBtn = document.getElementById('process-with-template-btn');
        if (processBtn) {
            processBtn.disabled = false;
            processBtn.textContent = `Process with Template (${templateName})`;
        }

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
        window.currentMapping = { ...template.columnMapping };
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
    if (!window.currentMappingFile) {
        alert('Please select a file to process.');
        return;
    }

    if (Object.keys(window.currentMapping).length === 0) {
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

        if (window.currentPatternAnalysis && window.currentPatternAnalysis.dataSection) {
            const analysis = window.currentPatternAnalysis.dataSection;
            parsingConfig.skipRows = analysis.dataStartIndex || 0;
            parsingConfig.skipColumns = analysis.startColumnIndex || 0;
            parsingConfig.headerRow = analysis.headerRowIndex;

            // Include enhanced fields for manual selections
            if (window.currentPatternAnalysis.manualSelection) {
                const manual = window.currentPatternAnalysis.manualSelection;
                parsingConfig.headerRows = manual.headerRows;
                parsingConfig.headerColumns = manual.headerColumns;
                parsingConfig.headerRange = manual.headerRange;

                // Add footer keyword detection if provided (manual takes priority)
                if (manual.footerKeyword) {
                    parsingConfig.footerRowKeyword = manual.footerKeyword;
                } else if (window.currentPatternAnalysis.autoFooterKeyword) {
                    parsingConfig.footerRowKeyword = window.currentPatternAnalysis.autoFooterKeyword;
                }

                // Configure multi-row data processing if multiple header rows selected
                if (manual.headerRows > 1) {
                    parsingConfig.rowProcessing = {
                        type: 'multi-row',
                        rowsPerRecord: manual.headerRows
                    };
                }
            } else if (window.currentPatternAnalysis.autoFooterKeyword) {
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
                mapping: window.currentMapping,
                parsingConfig: parsingConfig,
                templateName: 'Custom Template'
            }
        });

        if (result.success) {
            console.log(`Generated ${result.recordCount} processed records`);

            // Update the file with processed results
            window.currentMappingFile.parsedData = result.data;
            window.currentMappingFile.recordCount = result.recordCount;
            window.currentMappingFile.status = 'Processed with Custom Template';
            window.currentMappingFile.statusClass = 'status-success';
            window.currentMappingFile.broker = {
                ...currentMappingFile.broker,
                type: 'custom-template',
                name: window.currentMappingFile.broker.name + ' (Custom)'
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
                console.log(`Executing CALC formula for ${targetField}:`, formula);
                console.log('Available row keys:', Object.keys(row));

                try {
                    if (typeof window.executeFormula !== 'function') {
                        console.error('window.executeFormula function not found!');
                        mappedRow[targetField] = '';
                        return;
                    }
                    const result = window.executeFormula(formula, row);
                    console.log(`CALC result for ${targetField}:`, result);
                    mappedRow[targetField] = result;
                } catch (error) {
                    console.error(`Error executing formula for ${targetField}:`, error);
                    mappedRow[targetField] = '';
                }
            } else {
                // Column mapping - handle undefined/null but preserve falsy values like 0
                let value = row[mappingRule];

                // If direct mapping fails, try normalized header matching (handle line breaks)
                if (value === undefined) {
                    const normalizedMappingRule = mappingRule.replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim();

                    // Find matching key in row data with normalized comparison
                    const matchingKey = Object.keys(row).find(key => {
                        const normalizedKey = key.replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim();
                        return normalizedKey === normalizedMappingRule;
                    });

                    if (matchingKey) {
                        value = row[matchingKey];
                        console.log(`Normalized header mapping: "${mappingRule}" -> "${matchingKey}" = ${value}`);
                    }
                }

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

// ========== UTILITY FUNCTIONS ==========

/**
 * Update button states based on current selection
 */
function updateButtonStates() {
    const editTemplateBtn = document.getElementById('edit-template-btn');
    const deleteTemplateBtn = document.getElementById('delete-template-btn');

    const hasActiveTemplate = window.currentTemplateId && window.savedTemplates && window.savedTemplates[window.currentTemplateId];

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
            window.appSettings = appSettings; // Make appSettings globally accessible
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

    // Template selector event listener - activate template immediately on selection
    templateSelector.addEventListener('change', (e) => {
        const selectedId = e.target.value;
        if (selectedId && typeof selectTemplate === 'function') {
            selectTemplate(selectedId);
            console.log(`Template activated: ${selectedId}`);

            // Show visual feedback that template was activated
            console.log(`✓ Template "${window.savedTemplates[selectedId]?.name}" activated successfully`);

            // Visual feedback in the template display
            const templateNameElement = document.getElementById('active-template-name');
            if (templateNameElement) {
                const originalColor = templateNameElement.style.color;
                templateNameElement.style.color = '#4caf50';
                templateNameElement.textContent = templateNameElement.textContent + ' ✓';
                setTimeout(() => {
                    templateNameElement.style.color = originalColor;
                    if (window.savedTemplates[selectedId]) {
                        templateNameElement.textContent = window.savedTemplates[selectedId].name;
                    }
                }, 2000);
            }
        } else {
            // No template selected
            if (typeof selectTemplate === 'function') {
                selectTemplate('');
            }
        }
        updateButtonStates();
    });

    // Activate template button removed - dropdown provides immediate activation

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
            window.currentMapping = {};
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
    window.saveBrokerTemplate = saveBrokerTemplate;
    window.handleBrokerTemplateImport = handleBrokerTemplateImport;
    window.processFileWithTemplate = processFileWithTemplate;

    // Export navigation functions
    window.showMainSection = showMainSection;
    window.showNewTemplateSection = showNewTemplateSection;
    window.showEditTemplateSection = showEditTemplateSection;
    window.updateMappingFileSelector = updateMappingFileSelector;
});

// ========== ORCHESTRATION FUNCTIONS ==========



/**
 * Process file with existing template - orchestrates the entire processing pipeline
 */
async function processFileWithTemplate() {
    if (!window.currentMappingFile || !currentMapping || Object.keys(currentMapping).length === 0) {
        alert('Please load a file and create mappings first');
        return;
    }

    try {
        console.log('Processing file with template...');
        console.log('File:', window.currentMappingFile);
        console.log('Mapping:', window.currentMapping);
        console.log('Pattern Analysis:', window.currentPatternAnalysis);

        // Use the existing processing logic from preview functionality
        let processedData;
        let detection; // Declare detection variable in function scope

        // Always use template detection for "Process with Template" to ensure we use saved templates
        // Don't use cached parsedData as it might be from old parsing logic
        {
            // Use the same template detection and processing flow as automatic upload
            console.log('Process with Template: Looking for saved template for file:', window.currentMappingFile.file.name);

            // Check if detectBrokerType function exists
            if (typeof window.detectBrokerType !== 'function') {
                throw new Error('detectBrokerType function not available. Please check brokerParsers.js is loaded.');
            }

            // Check if processBrokerFile function exists
            if (typeof window.processBrokerFile !== 'function') {
                throw new Error('processBrokerFile function not available. Please check brokerParsers.js is loaded.');
            }

            // Detect template using the same logic as automatic upload
            console.log('Process with Template: Calling detectBrokerType...');
            detection = await window.detectBrokerType(window.currentMappingFile.file.name);
            console.log('Process with Template: Template detection result:', detection);

            if (detection && detection.type === 'custom' && detection.template) {
                // Use the same broker parser flow as automatic upload
                console.log('Process with Template: Using saved template:', detection.template.name);
                console.log('Process with Template: Calling processBrokerFile...');
                const result = await window.processBrokerFile({
                    file: window.currentMappingFile.file,
                    name: window.currentMappingFile.file.name,
                    id: window.currentMappingFile.id,
                    detectionOverride: detection
                });

                console.log('Process with Template: processBrokerFile result:', result);

                // Check if processing was successful (match fileManager.js structure)
                if (result.success) {
                    processedData = result.data;
                    console.log('Process with Template: Processed data length:', processedData ? processedData.length : 0);

                    if (!processedData || processedData.length === 0) {
                        throw new Error('No data was processed from the template');
                    }
                } else if (result.needsTemplate) {
                    throw new Error('Template processing failed - template needs to be recreated');
                } else {
                    const errorMsg = result.error || 'Unknown error in processBrokerFile';
                    throw new Error(`Template processing failed: ${errorMsg}`);
                }
            } else {
                console.log('Process with Template: No template detected or wrong type:', detection);
                throw new Error('No saved template found for this file. Please save a template first.');
            }

            // Cache the processed data for future use
            window.currentMappingFile.parsedData = processedData;
        }

        console.log('Processed data preview:', processedData.slice(0, 3));

        // Update the current file with processed data (match fileManager.js pattern)
        window.currentMappingFile.parsedData = processedData;
        window.currentMappingFile.recordCount = processedData.length;
        window.currentMappingFile.status = 'Processed with Template';
        window.currentMappingFile.statusClass = 'status-success';

        // Update broker info to indicate template usage
        if (window.currentMappingFile.broker) {
            window.currentMappingFile.broker.name = `${window.currentMappingFile.broker.name} (Template)`;
        }

        // Store processed data globally for results tab
        if (!window.processedFiles) {
            window.processedFiles = [];
        }

        // Remove any existing entry for this file
        const existingIndex = window.processedFiles.findIndex(f => f.fileId === window.currentMappingFile.id);
        if (existingIndex >= 0) {
            window.processedFiles.splice(existingIndex, 1);
        }

        // Add new processed data
        const processedFileData = {
            fileId: window.currentMappingFile.id,
            filename: window.currentMappingFile.file.name,
            brokerType: detection.template.name || 'Custom Template',
            recordCount: processedData.length,
            processedData: processedData,
            mapping: { ...detection.template.columnMapping },
            templateUsed: detection.template.name,
            processedAt: new Date().toISOString()
        };

        window.processedFiles.push(processedFileData);

        // Update the files display to show processed status
        if (typeof updateFilesDisplay === 'function') {
            updateFilesDisplay();
        }

        // Switch to results tab to show results
        const resultsTab = document.querySelector('[data-tab="results"]');
        if (resultsTab) {
            resultsTab.click();

            // Update results display
            if (typeof updateResultsTab === 'function') {
                updateResultsTab();
            }
        }

        alert(`File processed successfully with template! ${processedData.length} records processed.`);

    } catch (error) {
        console.error('Error processing file:', error);
        alert(`Error processing file: ${error.message}`);
    }
}

/**
 * Helper function to get broker parser by type
 */
function getBrokerParser(brokerType) {
    const parsers = {
        'AON': window.AONParser,
        'VGA': window.VGAParser,
        'BCI': window.BCIParser,
        'Voogt': window.VoogtParser
    };

    return parsers[brokerType] || null;
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}