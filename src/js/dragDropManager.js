/**
 * Borderellen Converter - Drag & Drop Manager Module
 * Handles all drag and drop interactions for column mapping
 */

// ========== DRAG EVENTS ==========

/**
 * Handle drag start event
 */
function handleDragStart(e) {
    e.target.classList.add('dragging');
    e.dataTransfer.setData('text/plain', e.target.getAttribute('data-column-name'));
}

/**
 * Handle drag end event
 */
function handleDragEnd(e) {
    e.target.classList.remove('dragging');
}

// ========== DROP ZONE MANAGEMENT ==========

/**
 * Update template drop zones with current mapping state
 */
function updateTemplateDropZones() {
    const container = document.getElementById('template-drop-zones');

    if (!window.currentTemplateId || !window.borderellenTemplate || !window.borderellenTemplate.columns || window.borderellenTemplate.columns.length === 0) {
        container.innerHTML = '<div style="text-align: center; padding: 32px; color: #888;"><p>No template selected. Please select a template in Tab 1.</p></div>';

        // Disable "Process with Template" button when no template
        const processBtn = document.getElementById('process-with-template-btn');
        if (processBtn) {
            processBtn.disabled = true;
            processBtn.textContent = 'Save Template First';
        }
        return;
    }


    container.innerHTML = '';

    window.borderellenTemplate.columns.forEach((column, index) => {
        const dropZone = document.createElement('div');
        dropZone.className = 'drop-zone';
        dropZone.setAttribute('data-column-id', column.id);
        dropZone.setAttribute('data-column-name', column.name);

        // Check if this field has a mapping
        const mappedSource = window.currentMapping && window.currentMapping[column.name];
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

    // Check if there's a saved template for the current file and update button state
    updateProcessButtonState();

    // Re-attach drop event listeners
    attachDropZoneListeners();
}

/**
 * Update "Process with Template" button state based on saved template availability
 */
async function updateProcessButtonState() {
    const processBtn = document.getElementById('process-with-template-btn');
    if (!processBtn) {
        console.log('Process button not found');
        return;
    }

    console.log('Updating process button state...');
    console.log('Current mapping file:', window.currentMappingFile);

    // Check if we have a current file and can detect a saved template for it
    if (window.currentMappingFile && window.currentMappingFile.file) {
        try {
            console.log('Detecting template for file:', window.currentMappingFile.file.name);
            const detection = await window.detectBrokerType(window.currentMappingFile.file.name);
            console.log('Template detection result:', detection);

            if (detection && detection.type === 'custom' && detection.template) {
                // Template exists - enable button
                console.log('Template found, enabling button:', detection.template.name);
                processBtn.disabled = false;
                processBtn.textContent = `Process with Template (${detection.template.name})`;
            } else {
                // No template found - disable button
                console.log('No template found, disabling button');
                processBtn.disabled = true;
                processBtn.textContent = 'Save Template First';
            }
        } catch (error) {
            // Error detecting template - disable button
            console.log('Error detecting template, disabling button:', error);
            processBtn.disabled = true;
            processBtn.textContent = 'Save Template First';
        }
    } else {
        // No file loaded - disable button
        console.log('No file loaded, disabling button');
        processBtn.disabled = true;
        processBtn.textContent = 'Save Template First';
    }
}

/**
 * Attach event listeners to drop zones
 */
function attachDropZoneListeners() {
    const dropZones = document.querySelectorAll('#template-drop-zones .drop-zone');

    dropZones.forEach(zone => {
        zone.addEventListener('dragover', handleDragOver);
        zone.addEventListener('drop', handleDrop);
        zone.addEventListener('dragleave', handleDragLeave);
        zone.addEventListener('click', handleDropZoneRightClick);
    });
}

// ========== DROP EVENTS ==========

/**
 * Handle drag over event
 */
function handleDragOver(e) {
    e.preventDefault();
    e.target.closest('.drop-zone').classList.add('drag-over');
}

/**
 * Handle drag leave event
 */
function handleDragLeave(e) {
    e.target.closest('.drop-zone').classList.remove('drag-over');
}

/**
 * Handle drop event
 */
function handleDrop(e) {
    e.preventDefault();
    const dropZone = e.target.closest('.drop-zone');
    dropZone.classList.remove('drag-over');

    const sourceColumn = e.dataTransfer.getData('text/plain');
    const targetColumn = dropZone.getAttribute('data-column-name');

    if (sourceColumn && targetColumn) {
        // Update mapping
        window.currentMapping[targetColumn] = sourceColumn;

        // Update visual display using the enhanced template system
        updateTemplateDropZones();

        console.log(`Mapped: ${targetColumn} ← ${sourceColumn}`);
    }
}

// ========== DROP ZONE INTERACTIONS ==========

/**
 * Handle drop zone click for fixed value entry
 */
function handleDropZoneClick(e) {
    e.preventDefault();
    const dropZone = e.currentTarget;
    const targetColumn = dropZone.getAttribute('data-column-name');

    // Check if it already has a mapping
    const existingMapping = window.currentMapping && window.currentMapping[targetColumn];
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
        window.currentMapping[targetColumn] = `FIXED:${fixedValue.trim()}`;

        // Update visual display
        updateTemplateDropZones();

        console.log(`Fixed value set: ${targetColumn} = ${fixedValue.trim()}`);
    }
}

/**
 * Handle drop zone right-click for context menu
 */
function handleDropZoneRightClick(e) {
    e.stopPropagation();

    const dropZone = e.currentTarget;
    const targetColumn = dropZone.getAttribute('data-column-name');
    const existingMapping = window.currentMapping[targetColumn];

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

// ========== MAPPING UTILITIES ==========

/**
 * Get human-readable mapping type
 */
function getMappingType(mapping) {
    if (mapping.startsWith('FIXED:')) return 'fixed value';
    if (mapping.startsWith('CALC:')) return 'calculation';
    return 'column mapping';
}

// ========== ACTION HANDLERS ==========

/**
 * Remove mapping action
 */
function removeMappingAction(dropZone) {
    const targetColumn = dropZone.getAttribute('data-column-name');

    // Remove mapping
    delete window.currentMapping[targetColumn];

    // Update visual display
    updateTemplateDropZones();

    console.log(`Removed mapping for ${targetColumn}`);
}

/**
 * Add fixed string action
 */
function addFixedStringAction(dropZone) {
    const targetColumn = dropZone.getAttribute('data-column-name');
    const existingMapping = window.currentMapping[targetColumn];

    // Show fixed string input modal
    window.showFixedStringModal(targetColumn, existingMapping);
}

/**
 * Add calculation action
 */
function addCalculationAction(dropZone) {
    const targetColumn = dropZone.getAttribute('data-column-name');
    const existingMapping = window.currentMapping[targetColumn];

    // Show calculation modal
    window.showCalculationModal(targetColumn, existingMapping);
}

// ========== COLUMN LOADING FUNCTIONS ==========

/**
 * Load source columns for mapping based on file ID
 * @param {string} fileId - ID of the selected file
 */
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
                    patternAnalysis = await window.DataPatternAnalyzer.analyzeFile(fileData.file);
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
                const autoFooterKeyword = window.detectFooterKeyword ? window.detectFooterKeyword(worksheet, range, startCol, endCol) : null;
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
                            <button class="btn btn-secondary" onclick="window.showManualHeaderSelection()" style="margin-left: 12px;">Manual select header & footer</button>
                        </div>
                    </div>
                `;

                container.innerHTML = confidenceInfo;
                displaySourceColumns(headers);

                // Trigger auto-mapping for high-confidence detections
                if (patternAnalysis.confidence > 0.7) {
                    console.log('High confidence detection - triggering auto-mapping');
                    setTimeout(() => {
                        window.generateAutoMappingSuggestions();
                    }, 200); // Small delay to ensure columns are loaded
                }

    } catch (error) {
        console.error('Error extracting columns from analysis:', error);
        throw error;
    }
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

/**
 * Display source columns in the UI
 */
function displaySourceColumns(headers) {
    const container = document.getElementById('source-columns');

    // Create a container for the column items (append to existing content)
    const columnsContainer = document.createElement('div');

    // Always add filename as the first virtual column
    const filenameColumn = document.createElement('div');
    filenameColumn.className = 'column-item filename-column';
    filenameColumn.draggable = true;
    filenameColumn.setAttribute('data-column-name', 'Filename');

    filenameColumn.innerHTML = `
        <strong>📄 Filename</strong>
        <span style="color: #888;">(text)</span>
    `;

    // Add drag event listeners
    filenameColumn.addEventListener('dragstart', handleDragStart);
    filenameColumn.addEventListener('dragend', handleDragEnd);

    columnsContainer.appendChild(filenameColumn);

    // Add regular source columns
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

// ========== GLOBAL EXPORTS ==========

// Make drag and drop functions globally accessible
window.handleDragStart = handleDragStart;
window.handleDragEnd = handleDragEnd;
window.updateTemplateDropZones = updateTemplateDropZones;
window.attachDropZoneListeners = attachDropZoneListeners;
window.handleDragOver = handleDragOver;
window.handleDragLeave = handleDragLeave;
window.handleDrop = handleDrop;
window.handleDropZoneClick = handleDropZoneClick;
window.handleDropZoneRightClick = handleDropZoneRightClick;
window.getMappingType = getMappingType;
window.removeMappingAction = removeMappingAction;
window.addFixedStringAction = addFixedStringAction;
window.addCalculationAction = addCalculationAction;

// Make column loading functions globally accessible
window.loadSourceColumns = loadSourceColumns;
window.loadSourceColumnsFromAnalysis = loadSourceColumnsFromAnalysis;
window.loadHeadersFromManualSelection = loadHeadersFromManualSelection;
window.displaySourceColumns = displaySourceColumns;