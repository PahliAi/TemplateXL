/**
 * Borderellen Converter - Header Selection Manager Module
 * Handles manual header and footer selection functionality
 */

// ========== HEADER SELECTION STATE ==========

// Global header selection state (also accessible via window.headerSelectionState)
let headerSelectionState = null;

// ========== MAIN HEADER SELECTION MODAL ==========

/**
 * Show manual header selection modal
 */
function showManualHeaderSelection() {
    if (!window.currentMappingFile) {
        alert('Please select a file first.');
        return;
    }

    // Create modal HTML
    const modalHtml = `
        <div id="manual-header-modal" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); z-index: 1000; display: flex; justify-content: center; align-items: center;">
            <div style="background: #2a2a2a; border-radius: 8px; padding: 24px; max-width: 95vw; max-height: 95vh; overflow: hidden; color: white; display: flex; flex-direction: column;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                    <h3 style="margin: 0; color: #00bcd4;">Manual Header & Footer Selection</h3>
                    <button onclick="closeManualHeaderSelection()" style="background: none; border: none; color: #888; font-size: 24px; cursor: pointer;">&times;</button>
                </div>
                <div id="manual-selection-instructions" style="background: #333; padding: 12px; border-radius: 4px; margin-bottom: 16px; color: #ccc;">
                    <strong>Instructions:</strong> Click and drag to select your header range, or use the cell coordinate inputs below for precise selection. The grid auto-scrolls when selecting near edges.
                </div>

                <!-- Cell coordinate inputs -->
                <div style="background: #333; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
                    <div style="display: flex; align-items: center; gap: 16px; flex-wrap: wrap;">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <label style="color: #ccc; min-width: 60px;">Start Cell:</label>
                            <input type="text" id="start-cell-input" placeholder="A1" style="background: #2a2a2a; border: 1px solid #555; color: white; padding: 4px 8px; border-radius: 4px; width: 60px; text-transform: uppercase;" />
                        </div>
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <label style="color: #ccc; min-width: 60px;">End Cell:</label>
                            <input type="text" id="end-cell-input" placeholder="Z5" style="background: #2a2a2a; border: 1px solid #555; color: white; padding: 4px 8px; border-radius: 4px; width: 60px; text-transform: uppercase;" />
                        </div>
                        <button onclick="selectCellRange()" style="background: #00bcd4; color: #000; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer;">Select Range</button>
                        <button onclick="scrollToSelection()" style="background: #444; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer;">Scroll to Selection</button>
                    </div>
                </div>

                <div id="manual-selection-grid" style="border: 1px solid #555; border-radius: 4px; overflow: auto; flex: 1; max-height: 35vh; position: relative;">
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
                    <div id="footer-keyword-editor" style="margin-top: 8px; display: none;">
                        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
                            <label style="color: #ccc; font-size: 12px; min-width: 80px;">Edit keyword:</label>
                            <input type="text" id="footer-keyword-input" style="background: #2a2a2a; border: 1px solid #555; color: white; padding: 4px 8px; border-radius: 4px; flex: 1; font-size: 12px;" placeholder="Enter footer keyword..." onkeypress="if(event.key==='Enter') updateFooterKeyword()" />
                            <button onclick="updateFooterKeyword()" style="background: #00bcd4; color: #000; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">Update</button>
                            <button onclick="clearFooterKeyword()" style="background: #f44336; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">Clear</button>
                        </div>
                        <div style="color: #888; font-size: 11px;">
                            Tip: Make keywords generic (e.g., "Total bookings" instead of "Total bookings for Q1") for better template reusability.
                        </div>
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
    headerSelectionState = null;
    window.headerSelectionState = null;
}

// ========== GRID LOADING ==========

/**
 * Load the grid for header selection
 */
async function loadHeaderSelectionGrid() {
    const gridContainer = document.getElementById('manual-selection-grid');

    try {
        const workbook = await ExcelCacheManager.getWorkbook(window.currentMappingFile.file);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        // Show grid (rows 1-30, all available columns)
        const maxRow = Math.min(30, range.e.r + 1);
        const maxCol = range.e.c + 1; // Show all columns, no artificial limit

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

        // Grid rows (worksheet is now compacted with sequential row numbering)
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
        headerSelectionState = {
            isSelecting: false,
            startRow: null,
            startCol: null,
            endRow: null,
            endCol: null,
            footerKeyword: null
        };
        window.headerSelectionState = headerSelectionState;

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
        const maxCol = range.e.c + 1; // Show all available columns
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

// ========== FOOTER KEYWORD SELECTION ==========

/**
 * Select footer keyword from clicked cell
 */
function selectFooterKeyword(row, col, cellValue) {
    const state = headerSelectionState;
    const statusEl = document.getElementById('footer-keyword-status');
    const editorEl = document.getElementById('footer-keyword-editor');
    const inputEl = document.getElementById('footer-keyword-input');

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

    // Store keyword and populate input
    state.footerKeyword = cellValue.trim();
    inputEl.value = state.footerKeyword;

    // Update status and show editor
    if (state.footerKeyword) {
        statusEl.innerHTML = `<span style="color: #00bcd4;">Footer keyword selected:</span> "${state.footerKeyword}"`;
        statusEl.style.color = '#00bcd4';
        editorEl.style.display = 'block';
    } else {
        statusEl.textContent = 'No footer keyword selected';
        statusEl.style.color = '#888';
        editorEl.style.display = 'none';
    }

    console.log(`Footer keyword selected: "${state.footerKeyword}" from cell ${XLSX.utils.encode_cell({ r: row, c: col })}`);
}

/**
 * Update footer keyword from input field
 */
function updateFooterKeyword() {
    const state = headerSelectionState;
    const inputEl = document.getElementById('footer-keyword-input');
    const statusEl = document.getElementById('footer-keyword-status');

    const newKeyword = inputEl.value.trim();

    if (!newKeyword) {
        alert('Please enter a footer keyword');
        return;
    }

    // Update keyword in state
    state.footerKeyword = newKeyword;

    // Update status display
    statusEl.innerHTML = `<span style="color: #00bcd4;">Footer keyword updated:</span> "${state.footerKeyword}"`;
    statusEl.style.color = '#00bcd4';

    console.log(`Footer keyword updated to: "${state.footerKeyword}"`);
}

/**
 * Clear footer keyword selection
 */
function clearFooterKeyword() {
    const state = headerSelectionState;
    const statusEl = document.getElementById('footer-keyword-status');
    const editorEl = document.getElementById('footer-keyword-editor');
    const inputEl = document.getElementById('footer-keyword-input');

    // Clear keyword from state
    state.footerKeyword = null;
    inputEl.value = '';

    // Clear visual selection
    document.querySelectorAll('.footer-cell').forEach(cell => {
        cell.style.background = '#2a2a2a';
        cell.style.color = 'white';
    });

    // Update status and hide editor
    statusEl.textContent = 'No footer keyword selected';
    statusEl.style.color = '#888';
    editorEl.style.display = 'none';

    console.log('Footer keyword cleared');
}

// ========== CELL COORDINATE INPUT ==========

/**
 * Select range using cell coordinate inputs
 */
function selectCellRange() {
    const startCellInput = document.getElementById('start-cell-input');
    const endCellInput = document.getElementById('end-cell-input');

    const startCell = startCellInput.value.trim().toUpperCase();
    const endCell = endCellInput.value.trim().toUpperCase();

    if (!startCell || !endCell) {
        alert('Please enter both start and end cell coordinates (e.g., A1, Z5)');
        return;
    }

    try {
        const startCoord = XLSX.utils.decode_cell(startCell);
        const endCoord = XLSX.utils.decode_cell(endCell);

        // Validate coordinates are within worksheet bounds
        const gridContainer = document.getElementById('manual-selection-grid');
        const table = gridContainer.querySelector('table');
        if (!table) return;

        // Update selection state
        const state = headerSelectionState;
        state.startRow = startCoord.r;
        state.startCol = startCoord.c;
        state.endRow = endCoord.r;
        state.endCol = endCoord.c;
        state.isSelecting = false;

        // Update visual selection and status
        updateRangeVisual();
        updateSelectionStatus();
        updateCellInputs();

        // Enable apply button
        const applyBtn = document.getElementById('apply-header-selection');
        applyBtn.disabled = false;

        // Scroll to selection
        scrollToSelection();

    } catch (error) {
        alert('Invalid cell coordinates. Please use format like A1, B5, AA10, etc.');
    }
}

/**
 * Scroll to current selection in the grid
 */
function scrollToSelection() {
    const state = headerSelectionState;
    if (!state || state.startRow === null) return;

    const gridContainer = document.getElementById('manual-selection-grid');
    const minRow = Math.min(state.startRow, state.endRow);
    const minCol = Math.min(state.startCol, state.endCol);

    // Find the target cell
    const targetCell = gridContainer.querySelector(`.header-cell[data-row="${minRow}"][data-col="${minCol}"]`);
    if (targetCell) {
        // Scroll to make the cell visible
        targetCell.scrollIntoView({
            behavior: 'smooth',
            block: 'center',
            inline: 'center'
        });
    }
}

/**
 * Update cell input values based on current selection
 */
function updateCellInputs() {
    const state = headerSelectionState;
    if (!state || state.startRow === null) return;

    const startCellInput = document.getElementById('start-cell-input');
    const endCellInput = document.getElementById('end-cell-input');

    const minRow = Math.min(state.startRow, state.endRow);
    const maxRow = Math.max(state.startRow, state.endRow);
    const minCol = Math.min(state.startCol, state.endCol);
    const maxCol = Math.max(state.startCol, state.endCol);

    startCellInput.value = XLSX.utils.encode_cell({ r: minRow, c: minCol });
    endCellInput.value = XLSX.utils.encode_cell({ r: maxRow, c: maxCol });
}

// ========== RANGE SELECTION ==========

/**
 * Start range selection
 */
function startRangeSelection(row, col) {
    const state = headerSelectionState;
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
    updateCellInputs();
}

/**
 * Update range selection during drag
 */
function updateRangeSelection(row, col) {
    const state = headerSelectionState;
    if (!state.isSelecting) return;

    state.endRow = row;
    state.endCol = col;

    updateRangeVisual();
    updateSelectionStatus();
    updateCellInputs();

    // Auto-scroll when near edges
    autoScrollGrid(row, col);
}

/**
 * End range selection
 */
function endRangeSelection() {
    const state = headerSelectionState;
    if (!state.isSelecting) return;

    state.isSelecting = false;

    // Update cell inputs with final selection
    updateCellInputs();

    // Enable apply button
    const applyBtn = document.getElementById('apply-header-selection');
    applyBtn.disabled = false;
}

/**
 * Auto-scroll grid when selecting near edges
 */
function autoScrollGrid(row, col) {
    const gridContainer = document.getElementById('manual-selection-grid');
    if (!gridContainer) return;

    const currentCell = gridContainer.querySelector(`.header-cell[data-row="${row}"][data-col="${col}"]`);
    if (!currentCell) return;

    const containerRect = gridContainer.getBoundingClientRect();
    const cellRect = currentCell.getBoundingClientRect();

    const scrollMargin = 50; // Pixels from edge to trigger scroll
    const scrollStep = 30; // Pixels to scroll

    // Check if cell is near edges and scroll accordingly
    if (cellRect.left < containerRect.left + scrollMargin) {
        gridContainer.scrollLeft -= scrollStep;
    } else if (cellRect.right > containerRect.right - scrollMargin) {
        gridContainer.scrollLeft += scrollStep;
    }

    if (cellRect.top < containerRect.top + scrollMargin) {
        gridContainer.scrollTop -= scrollStep;
    } else if (cellRect.bottom > containerRect.bottom - scrollMargin) {
        gridContainer.scrollTop += scrollStep;
    }
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
    const state = headerSelectionState;
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
    const state = headerSelectionState;
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

// ========== APPLY SELECTION ==========

/**
 * Apply header selection
 */
async function applyHeaderSelection() {
    const state = headerSelectionState;

    if (!state || state.startRow === null) {
        alert('Please select a range first.');
        return;
    }

    const minRow = Math.min(state.startRow, state.endRow);
    const maxRow = Math.max(state.startRow, state.endRow);
    const minCol = Math.min(state.startCol, state.endCol);
    const maxCol = Math.max(state.startCol, state.endCol);


    // Store the selection in currentPatternAnalysis for use by other functions
    if (!window.currentPatternAnalysis) {
        window.currentPatternAnalysis = {};
    }
    if (!window.currentPatternAnalysis.dataSection) {
        window.currentPatternAnalysis.dataSection = {};
    }

    // Update pattern analysis with manual selection
    window.currentPatternAnalysis.dataSection.startCell = XLSX.utils.encode_cell({ r: minRow, c: minCol });
    window.currentPatternAnalysis.dataSection.headerRowIndex = minRow;
    window.currentPatternAnalysis.dataSection.dataStartIndex = maxRow + 1; // Data starts after header
    window.currentPatternAnalysis.dataSection.startColumnIndex = minCol;
    window.currentPatternAnalysis.dataSection.endColumnIndex = maxCol;
    window.currentPatternAnalysis.suggestedHeaderRow = minRow;
    window.currentPatternAnalysis.confidence = 1.0; // Manual selection = 100% confidence

    // Store additional info for parsingConfig
    window.currentPatternAnalysis.manualSelection = {
        headerRows: maxRow - minRow + 1,
        headerColumns: maxCol - minCol + 1,
        headerRange: `${XLSX.utils.encode_cell({ r: minRow, c: minCol })}:${XLSX.utils.encode_cell({ r: maxRow, c: maxCol })}`,
        footerKeyword: state.footerKeyword || null
    };

    // Close modal
    closeManualHeaderSelection();

    // Reload source columns using the main app's loadSourceColumnsFromAnalysis function
    // This properly handles multi-row headers (no code duplication)
    const container = document.getElementById('source-columns');
    if (container) {
        await loadSourceColumnsFromAnalysis(window.currentMappingFile.file, window.currentPatternAnalysis, container);
    }

    console.log('Applied header selection:', {
        headerRange: window.currentPatternAnalysis.manualSelection.headerRange,
        footerKeyword: state.footerKeyword
    });
}

// ========== HEADER LOADING ==========
// Note: loadHeadersFromSelection() has been removed to eliminate code duplication.
// We now use loadSourceColumnsFromAnalysis() from app.js which properly handles multi-row headers.

/**
 * Load headers from manual selection (legacy function for compatibility)
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
 * Display source columns in the mapping interface
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
    filenameColumn.addEventListener('dragstart', window.handleDragStart);
    filenameColumn.addEventListener('dragend', window.handleDragEnd);

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
        columnItem.addEventListener('dragstart', window.handleDragStart);
        columnItem.addEventListener('dragend', window.handleDragEnd);

        columnsContainer.appendChild(columnItem);
    });

    // Append the columns container to the main container
    container.appendChild(columnsContainer);
}

// ========== UTILITY FUNCTIONS ==========

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
 * Prevent default event (used for preventing text selection during drag)
 */
function preventDefault(e) {
    e.preventDefault();
}

// ========== GLOBAL EXPORTS ==========

// Make header selection functions globally accessible
window.showManualHeaderSelection = showManualHeaderSelection;
window.closeManualHeaderSelection = closeManualHeaderSelection;
window.loadHeaderSelectionGrid = loadHeaderSelectionGrid;
window.loadFooterDetectionGrid = loadFooterDetectionGrid;
window.selectFooterKeyword = selectFooterKeyword;
window.updateFooterKeyword = updateFooterKeyword;
window.clearFooterKeyword = clearFooterKeyword;
window.selectCellRange = selectCellRange;
window.scrollToSelection = scrollToSelection;
window.updateCellInputs = updateCellInputs;
window.autoScrollGrid = autoScrollGrid;
window.startRangeSelection = startRangeSelection;
window.updateRangeSelection = updateRangeSelection;
window.endRangeSelection = endRangeSelection;
window.clearRangeSelection = clearRangeSelection;
window.updateRangeVisual = updateRangeVisual;
window.updateSelectionStatus = updateSelectionStatus;
window.applyHeaderSelection = applyHeaderSelection;
window.loadHeadersFromManualSelection = loadHeadersFromManualSelection;
window.displaySourceColumns = displaySourceColumns;
window.detectFooterKeyword = detectFooterKeyword;
window.preventDefault = preventDefault;