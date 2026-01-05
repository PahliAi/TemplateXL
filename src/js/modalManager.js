/**
 * Borderellen Converter - Modal Management Module
 * Handles all modal dialogs including fixed string, calculation, preview, settings, and contacts modals
 */

// ========== MODAL UTILITIES ==========

/**
 * Show any modal by setting display style
 */
function showModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.add('show');
    }
}

/**
 * Hide any modal by removing show class
 */
function hideModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.remove('show');
    }
}

// ========== FIXED STRING MODAL ==========

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
        window.currentMapping[targetColumn] = `FIXED:${fixedValue}`;

        // Update visual display
        window.updateTemplateDropZones();

        console.log(`Fixed value set: ${targetColumn} = ${fixedValue}`);
    }

    closeFixedStringModal();
}

// ========== CALCULATION MODAL ==========

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
            <div style="background: #2a2a2a; border-radius: 8px; padding: 24px; max-width: 800px; width: 90%; color: white; border: 1px solid #555; max-height: 90vh; overflow-y: auto;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                    <h3 style="margin: 0; color: #00bcd4;">Create Formula</h3>
                    <button onclick="closeCalculationModal()" style="background: none; border: none; color: #888; font-size: 24px; cursor: pointer;">&times;</button>
                </div>

                <div style="margin-bottom: 16px;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Target Field:</label>
                    <div style="background: #333; padding: 8px 12px; border-radius: 4px; color: #00bcd4; font-weight: bold;">${targetColumn}</div>
                </div>

                <div style="margin-bottom: 16px;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Available Source Columns:</label>
                    <div id="column-buttons" style="display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 8px;">
                        ${sourceColumns.map(col =>
                            `<button type="button" onclick="insertColumnName('${col}')" style="background: #444; border: 1px solid #666; border-radius: 4px; padding: 4px 8px; color: #00bcd4; font-size: 11px; cursor: pointer; white-space: nowrap;">${col}</button>`
                        ).join('')}
                    </div>
                    <div style="font-size: 11px; color: #888; margin-bottom: 8px;">
                        Click column names above to insert into formula. Examples: Filename, PolisNr, Bruto
                    </div>
                </div>

                <div style="margin-bottom: 8px; padding: 8px; background: #1a4d1a; border: 1px solid #4caf50; border-radius: 4px;">
                    <div style="color: #4caf50; font-weight: bold; font-size: 12px; margin-bottom: 4px;">💡 Pro Tip: Excel Copy-Paste!</div>
                    <div style="color: #ccc; font-size: 11px;">
                        Create your formula in Excel first, then copy and paste it here! Just replace Excel column references (A1, B2) with our column names (Filename, Bruto).
                    </div>
                </div>

                <div style="margin-bottom: 16px;">
                    <label for="calc-formula-input" style="display: block; margin-bottom: 8px; color: #ccc;">
                        Formula Input:
                    </label>
                    <textarea id="calc-formula-input"
                              placeholder="Enter formula here... Examples:
• Math: Bruto * 1.21
• Text: LEFT(Filename, FIND('_', Filename)-1)
• Logic: IF(CONTAINS(Branche, 'AUTO'), 'Automotive', Branche)
• Copy Excel formulas and replace column references!"
                              style="background: #1a1a1a; border: 1px solid #555; border-radius: 4px; padding: 12px; color: white; width: 100%; min-height: 80px; font-family: 'Courier New', monospace; font-size: 13px; resize: vertical; box-sizing: border-box;">${existingFormula}</textarea>
                </div>

                <div style="display: flex; gap: 8px; margin-bottom: 16px;">
                    <button type="button" onclick="testFormula()" class="btn" style="background: #4caf50; border: none; color: white; padding: 8px 16px; border-radius: 4px; cursor: pointer;">
                        Test Formula
                    </button>
                    <div style="font-size: 11px; color: #888; padding: 8px 0;">
                        Test your formula against the first 5 rows of source data to verify it works correctly.
                    </div>
                </div>

                <div id="test-results" style="display: none; margin-bottom: 16px;">
                    <label style="display: block; margin-bottom: 8px; color: #ccc;">Test Results:</label>
                    <div id="test-results-content" style="background: #1a1a1a; border: 1px solid #555; border-radius: 4px; padding: 12px; max-height: 200px; overflow-y: auto;">
                        <!-- Test results will appear here -->
                    </div>
                </div>

                <div style="display: flex; justify-content: flex-end; gap: 8px;">
                    <button class="btn btn-secondary" onclick="closeCalculationModal()" style="background: #666; border: none; color: white; padding: 8px 16px; border-radius: 4px; cursor: pointer;">Cancel</button>
                    <button class="btn" onclick="applyCalculation('${targetColumn}')" id="apply-calculation-btn" style="background: #00bcd4; border: none; color: white; padding: 8px 16px; border-radius: 4px; cursor: pointer;">Apply Formula</button>
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);

    // Focus on the formula input
    const formulaInput = document.getElementById('calc-formula-input');
    if (formulaInput) {
        formulaInput.focus();
        if (existingFormula) {
            formulaInput.setSelectionRange(formulaInput.value.length, formulaInput.value.length);
        }
    }
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
 * Insert column name into formula textarea at cursor position
 */
function insertColumnName(columnName) {
    const formulaInput = document.getElementById('calc-formula-input');
    if (!formulaInput) return;

    const startPos = formulaInput.selectionStart;
    const endPos = formulaInput.selectionEnd;
    const currentValue = formulaInput.value;

    const beforeCursor = currentValue.substring(0, startPos);
    const afterCursor = currentValue.substring(endPos);

    const newValue = beforeCursor + columnName + afterCursor;
    formulaInput.value = newValue;

    // Position cursor after the inserted column name
    const newCursorPos = startPos + columnName.length;
    formulaInput.focus();
    formulaInput.setSelectionRange(newCursorPos, newCursorPos);
}

/**
 * Test formula against first 5 rows of source data
 */
async function testFormula() {
    const formulaInput = document.getElementById('calc-formula-input');
    const testResultsDiv = document.getElementById('test-results');
    const testResultsContent = document.getElementById('test-results-content');

    if (!formulaInput || !testResultsDiv || !testResultsContent) return;

    const formula = formulaInput.value.trim();
    if (!formula) {
        alert('Please enter a formula to test');
        return;
    }

    // Show loading
    testResultsContent.innerHTML = '<div style="color: #888; padding: 8px;">Testing formula...</div>';
    testResultsDiv.style.display = 'block';

    try {
        // Get sample data from current file
        const sampleData = await getTestDataForFormula();
        if (!sampleData || sampleData.length === 0) {
            testResultsContent.innerHTML = `
                <div style="color: #f44336; padding: 8px;">
                    No test data available. Please make sure you have uploaded and analyzed a file first.
                </div>
            `;
            return;
        }

        // Create a temporary mapping to test this formula
        const tempMapping = { 'TestField': `CALC:${formula}` };

        // Use the existing applyMappingToData function to test the formula
        const testResults = window.applyMappingToData(sampleData.slice(0, 5), tempMapping);

        // Display results in a simple two-column format
        let html = `
            <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                <thead>
                    <tr style="border-bottom: 1px solid #555;">
                        <th style="text-align: left; padding: 4px; color: #ccc;">Source Data (relevant columns)</th>
                        <th style="text-align: left; padding: 4px; color: #ccc;">Formula Result</th>
                    </tr>
                </thead>
                <tbody>
        `;

        testResults.forEach((result, i) => {
            const originalRow = sampleData[i];

            // Show only columns referenced in the formula
            const relevantData = {};
            Object.keys(originalRow).forEach(key => {
                if (formula.includes(key)) {
                    relevantData[key] = originalRow[key];
                }
            });

            const sourcePreview = Object.keys(relevantData).length > 0
                ? Object.entries(relevantData).map(([k, v]) => `${k}: "${v}"`).join(', ')
                : 'No matching columns';

            html += `
                <tr style="border-bottom: 1px solid #333;">
                    <td style="padding: 4px; color: #ccc; max-width: 300px; overflow: hidden; text-overflow: ellipsis;">${sourcePreview}</td>
                    <td style="padding: 4px; color: #4caf50; font-weight: bold;">"${result.TestField}"</td>
                </tr>
            `;
        });

        html += '</tbody></table>';
        testResultsContent.innerHTML = html;

    } catch (error) {
        testResultsContent.innerHTML = `
            <div style="color: #f44336; padding: 8px;">
                Formula Error: ${error.message}
            </div>
        `;
    }
}

/**
 * Get test data for formula validation
 */
async function getTestDataForFormula() {
    console.log('getTestDataForFormula: currentMappingFile exists:', !!window.currentMappingFile);
    console.log('getTestDataForFormula: file exists:', !!window.currentMappingFile?.file);
    console.log('getTestDataForFormula: file name:', window.currentMappingFile?.file?.name);
    console.log('getTestDataForFormula: currentPatternAnalysis exists:', !!window.currentPatternAnalysis);
    console.log('getTestDataForFormula: currentPatternAnalysis.filename:', window.currentPatternAnalysis?.filename);

    // Validate that we have file context - this is required for template creation workflow
    if (!window.currentMappingFile || !window.currentMappingFile.file) {
        throw new Error('No file loaded. Please select a file from the File Mapping tab first.');
    }

    // Always re-extract test data using current pattern analysis to ensure fresh, normalized headers
    // Don't use cached parsedData as it may have outdated column names or non-normalized headers
    console.log('Re-extracting test data with current pattern analysis for accurate formula testing');

    console.log('Reading file for test data:', window.currentMappingFile.file.name);

    // Extract data from the loaded file
    try {
        const workbook = await ExcelCacheManager.getWorkbook(window.currentMappingFile.file);

        // Use pattern analysis if available, otherwise simple extraction
        if (window.currentPatternAnalysis && window.GenericParser && typeof window.GenericParser.parseWithAnalysis === 'function') {
            console.log('Using GenericParser with pattern analysis');

            // FIX: Ensure filename is set if missing
            if (!window.currentPatternAnalysis.filename && window.currentMappingFile?.file?.name) {
                console.warn('currentPatternAnalysis missing filename, adding it now:', window.currentMappingFile.file.name);
                window.currentPatternAnalysis.filename = window.currentMappingFile.file.name;
            }

            const sampleData = await window.GenericParser.parseWithAnalysis(workbook, window.currentPatternAnalysis);
            return sampleData.slice(0, 5);
        } else {
            console.log('Using simple sheet_to_json extraction');
            // Simple extraction with filename added
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let rawData = XLSX.utils.sheet_to_json(worksheet, {
                cellFormula: false  // Read calculated values instead of formulas
            });

            // Ensure we have some data to work with
            if (!rawData || rawData.length === 0) {
                throw new Error('No data found in Excel file');
            }

            // Always add filename column for testing
            rawData = rawData.map(row => ({
                ...row,
                Filename: window.currentMappingFile.file.name
            }));

            return rawData.slice(0, 5);
        }
    } catch (error) {
        throw new Error(`Failed to read file for testing: ${error.message}`);
    }
}

/**
 * Get available source columns for calculations
 */
function getAvailableSourceColumns() {
    const sourceColumnsContainer = document.getElementById('source-columns');
    if (!sourceColumnsContainer) return [];

    const columnElements = sourceColumnsContainer.querySelectorAll('.column-item');
    return Array.from(columnElements).map(el => el.getAttribute('data-column-name')).filter(col => col);
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

        // Only create parameter inputs if they don't exist or if function type changed
        const currentFunctionType = paramsContainer.getAttribute('data-function-type');
        if (currentFunctionType !== functionType) {
            // Preserve existing values before recreating inputs
            const existingValues = {};
            for (let i = 1; i <= 4; i++) {
                const input = paramsContainer.querySelector(`#func-param-${i}`);
                if (input) {
                    existingValues[`param${i}`] = input.value;
                }
            }

            // Recreate inputs
            paramsContainer.innerHTML = window.generateFunctionParams(functionType);
            paramsContainer.setAttribute('data-function-type', functionType);

            // Restore existing values
            for (let i = 1; i <= 4; i++) {
                const input = paramsContainer.querySelector(`#func-param-${i}`);
                if (input && existingValues[`param${i}`]) {
                    input.value = existingValues[`param${i}`];
                }
            }
        }
    } else {
        paramsDiv.style.display = 'none';
    }

    // Build formula
    let formula = '';
    let isValid = false;

    // Check if it's a function that can work standalone (no math formula needed)
    const standaloneFunctions = [
        'LENGTH', 'ISEMPTY', 'TRIM', 'UPPER', 'LOWER',           // Text manipulation
        'LEFT', 'RIGHT', 'MID', 'REGEX', 'SPLIT',               // Text extraction
        'REPLACE', 'CONCAT',                                     // Text modification
        'CONTAINS', 'STARTSWITH', 'ENDSWITH',                   // Text analysis
        'ABS', 'CEILING', 'FLOOR',                              // Math functions
        'IF', 'AND', 'OR', 'NOT'                               // Logic functions
    ];

    if (functionType && standaloneFunctions.includes(functionType)) {
        // Text-only functions - just need a column
        if (columnA) {
            const modal = document.getElementById('calculation-modal');
            const param1 = modal?.querySelector('#func-param-1')?.value || '';
            const param2 = modal?.querySelector('#func-param-2')?.value || '';
            const param3 = modal?.querySelector('#func-param-3')?.value || '';
            const param4 = modal?.querySelector('#func-param-4')?.value || '';
            formula = buildFunctionFormulaWithParams(functionType, columnA, param1, param2, param3, param4);
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
                const modal = document.getElementById('calculation-modal');
                const param1 = modal?.querySelector('#func-param-1')?.value || '';
                const param2 = modal?.querySelector('#func-param-2')?.value || '';
                const param3 = modal?.querySelector('#func-param-3')?.value || '';
                const param4 = modal?.querySelector('#func-param-4')?.value || '';
                formula = buildFunctionFormulaWithParams(functionType, formula, param1, param2, param3, param4);
            }
        }
    }

    // Update preview
    if (formula) {
        previewEl.textContent = formula;

        // Create example based on formula type
        let exampleText = '';
        if (functionType && standaloneFunctions.includes(functionType)) {
            // Function-only examples
            switch (functionType) {
                case 'UPPER': exampleText = 'Example: "john smith" → "JOHN SMITH"'; break;
                case 'LOWER': exampleText = 'Example: "JOHN SMITH" → "john smith"'; break;
                case 'TRIM': exampleText = 'Example: "  text  " → "text"'; break;
                case 'LENGTH': exampleText = 'Example: "Hello World" → 11'; break;
                case 'LEFT': exampleText = 'Example: "Hello" → "He" (first 2 chars)'; break;
                case 'RIGHT': exampleText = 'Example: "Hello" → "lo" (last 2 chars)'; break;
                case 'MID': exampleText = 'Example: "Hello" → "ll" (middle portion)'; break;
                case 'CONTAINS': exampleText = 'Example: "Hello World" contains "World" → true'; break;
                case 'STARTSWITH': exampleText = 'Example: "Hello" starts with "He" → true'; break;
                case 'ENDSWITH': exampleText = 'Example: "Hello" ends with "lo" → true'; break;
                case 'ABS': exampleText = 'Example: -123 → 123'; break;
                case 'CEILING': exampleText = 'Example: 4.3 → 5'; break;
                case 'FLOOR': exampleText = 'Example: 4.9 → 4'; break;
                default: exampleText = `Example: ${functionType}(${columnA || 'column'}) result`;
            }
        } else if (operator) {
            // Math calculation examples
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
            exampleText = `Example: ${exampleA} ${operatorSymbol} ${exampleB} = ${result}`;
        }

        exampleEl.textContent = exampleText;
        exampleEl.style.backgroundColor = '#1a4d1a';
        exampleEl.style.borderColor = '#4caf50';
        exampleEl.style.color = '#4caf50';
    } else {
        previewEl.textContent = 'Select a function or column + operator to build formula';
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
    const formulaInput = document.getElementById('calc-formula-input');

    if (!formulaInput) {
        alert('Formula input not found');
        return;
    }

    const formula = formulaInput.value.trim();

    if (!formula) {
        alert('Please enter a formula');
        return;
    }

    // Validate formula syntax by testing it
    try {
        // Try to execute the formula against test data to validate syntax
        if (window.executeFormula) {
            // Create a dummy row with some sample data for validation
            const testRow = {
                'Filename': 'Test_File.xlsx',
                'TestColumn': 'TestValue',
                'TestNumber': 100
            };
            window.executeFormula(formula, testRow);
        }
    } catch (error) {
        alert(`Formula Error: ${error.message}\n\nPlease check your formula syntax and try again.`);
        return;
    }

    // Store as calculation with CALC: prefix
    window.currentMapping[targetColumn] = `CALC:${formula}`;

    // Update visual display
    if (window.updateTemplateDropZones) {
        window.updateTemplateDropZones();
    }

    console.log(`Formula applied: ${targetColumn} = CALC:${formula}`);

    closeCalculationModal();
}

// ========== PREVIEW MODAL ==========

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

// ========== SETTINGS MODAL ==========

/**
 * Show settings modal
 */
function showSettingsModal() {
    const modal = document.getElementById('settings-modal');
    const userNameInput = document.getElementById('user-name-input');
    const userEmailInput = document.getElementById('user-email-input');
    const userSignatureInput = document.getElementById('user-signature-input');
    const downloadFolderInput = document.getElementById('download-folder-input');

    // Populate current values
    userNameInput.value = window.appSettings.userName;
    userEmailInput.value = window.appSettings.userEmail || '';
    userSignatureInput.value = window.appSettings.userSignature || '';
    downloadFolderInput.value = window.appSettings.downloadFolder;

    // Set up tab click handlers (do this each time modal opens to ensure handlers are attached)
    document.querySelectorAll('.settings-tab').forEach(tab => {
        if (tab && tab.dataset) {
            // Remove existing listeners by cloning
            const newTab = tab.cloneNode(true);
            tab.parentNode.replaceChild(newTab, tab);

            // Add fresh click listener
            newTab.addEventListener('click', () => {
                switchSettingsTab(newTab.dataset.tab);
            });
        }
    });

    // Reset to General tab
    switchSettingsTab('general');

    modal.classList.add('show');
}

/**
 * Switch between settings tabs
 */
function switchSettingsTab(tabName) {
    // Update tab buttons
    document.querySelectorAll('.settings-tab').forEach(tab => {
        const isActive = tab.dataset.tab === tabName;
        tab.classList.toggle('active', isActive);
        tab.style.color = isActive ? '#00bcd4' : '#888';
        tab.style.borderBottomColor = isActive ? '#00bcd4' : 'transparent';
    });

    // Update tab content
    document.querySelectorAll('.settings-tab-content').forEach(content => {
        content.style.display = content.id === `${tabName}-settings-tab` ? 'block' : 'none';
    });

    // Load file mappings when switching to mappings tab
    if (tabName === 'mappings') {
        loadFileMappingsList();
    }
}

/**
 * Load and display file mappings list
 */
async function loadFileMappingsList() {
    const listContainer = document.getElementById('file-mappings-list');

    try {
        const mappings = await window.loadAllFileMappings();

        if (mappings.length === 0) {
            listContainer.innerHTML = '<div style="padding: 24px; text-align: center; color: #888;">No file mappings saved yet.</div>';
            return;
        }

        listContainer.innerHTML = mappings.map(mapping => `
            <div class="file-mapping-item" style="padding: 12px 16px; border-bottom: 1px solid #333; display: flex; align-items: center; gap: 12px;">
                <input type="checkbox" class="mapping-checkbox" data-mapping-id="${mapping.id}" style="width: 18px; height: 18px; cursor: pointer;">
                <div style="flex: 1;">
                    <div style="font-weight: 500; color: #fff; margin-bottom: 4px;">${mapping.name}</div>
                    <div style="font-size: 12px; color: #888;">
                        Keyword: ${mapping.matchingKeyword || 'None'} |
                        ${Object.keys(mapping.columnMapping || {}).length} mappings |
                        Created: ${new Date(mapping.created).toLocaleDateString()}
                    </div>
                </div>
                <button class="btn btn-secondary" onclick="deleteSingleMappingFromList('${mapping.id}')" style="padding: 6px 12px; font-size: 12px; background: #d32f2f; border-color: #d32f2f;">Delete</button>
            </div>
        `).join('');

        // Add checkbox change listeners
        document.querySelectorAll('.mapping-checkbox').forEach(checkbox => {
            checkbox.addEventListener('change', updateMappingSelectionUI);
        });

        updateMappingSelectionUI();
    } catch (error) {
        console.error('Error loading file mappings:', error);
        listContainer.innerHTML = '<div style="padding: 24px; text-align: center; color: #f44336;">Error loading file mappings.</div>';
    }
}

/**
 * Update UI based on mapping selection
 */
function updateMappingSelectionUI() {
    const checkboxes = document.querySelectorAll('.mapping-checkbox');
    const selectedCount = Array.from(checkboxes).filter(cb => cb.checked).length;

    document.getElementById('selected-mappings-count').textContent = `${selectedCount} selected`;
    document.getElementById('export-selected-mappings-btn').disabled = selectedCount === 0;
    document.getElementById('delete-selected-mappings-btn').disabled = selectedCount === 0;
}

/**
 * Select all file mappings
 */
function selectAllMappings() {
    document.querySelectorAll('.mapping-checkbox').forEach(checkbox => {
        checkbox.checked = true;
    });
    updateMappingSelectionUI();
}

/**
 * Deselect all file mappings
 */
function deselectAllMappings() {
    document.querySelectorAll('.mapping-checkbox').forEach(checkbox => {
        checkbox.checked = false;
    });
    updateMappingSelectionUI();
}

/**
 * Export selected file mappings
 */
async function exportSelectedMappings() {
    const selectedIds = Array.from(document.querySelectorAll('.mapping-checkbox:checked'))
        .map(cb => cb.dataset.mappingId);

    if (selectedIds.length === 0) {
        alert('Please select at least one mapping to export.');
        return;
    }

    try {
        const allMappings = await window.loadAllFileMappings();
        const selectedMappings = allMappings.filter(m => selectedIds.includes(m.id));

        // Create export object
        const exportData = {
            exportDate: new Date().toISOString(),
            version: '1.0',
            mappings: selectedMappings
        };

        // Create JSON file
        const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `file-mappings-export-${new Date().toISOString().split('T')[0]}.json`;
        a.click();
        URL.revokeObjectURL(url);

        alert(`Exported ${selectedMappings.length} file mapping(s) successfully!`);
    } catch (error) {
        console.error('Error exporting mappings:', error);
        alert('Error exporting mappings. Please try again.');
    }
}

/**
 * Delete selected file mappings
 */
async function deleteSelectedMappings() {
    const selectedIds = Array.from(document.querySelectorAll('.mapping-checkbox:checked'))
        .map(cb => cb.dataset.mappingId);

    if (selectedIds.length === 0) {
        alert('Please select at least one mapping to delete.');
        return;
    }

    const confirmDelete = confirm(
        `Are you sure you want to delete ${selectedIds.length} file mapping(s)?\n\n` +
        `This action cannot be undone.`
    );

    if (!confirmDelete) return;

    try {
        for (const id of selectedIds) {
            await window.deleteFileMapping(id);
        }

        alert(`Deleted ${selectedIds.length} file mapping(s) successfully!`);
        loadFileMappingsList(); // Reload the list
    } catch (error) {
        console.error('Error deleting mappings:', error);
        alert('Error deleting mappings. Please try again.');
    }
}

/**
 * Delete a single file mapping (UI wrapper)
 */
async function deleteSingleMappingFromList(mappingId) {
    const confirmDelete = confirm('Are you sure you want to delete this file mapping?\n\nThis action cannot be undone.');

    if (!confirmDelete) return;

    try {
        await window.deleteFileMapping(mappingId);
        alert('File mapping deleted successfully!');
        loadFileMappingsList(); // Reload the list
    } catch (error) {
        console.error('Error deleting mapping:', error);
        alert('Error deleting mapping. Please try again.');
    }
}

/**
 * Hide settings modal
 */
function hideSettingsModal() {
    const modal = document.getElementById('settings-modal');
    modal.classList.remove('show');
}

/**
 * Select download folder using File System Access API
 */
async function selectDownloadFolder() {
    try {
        // Use File System Access API if available
        if ('showDirectoryPicker' in window) {
            const dirHandle = await window.showDirectoryPicker();
            window.appSettings.downloadFolderHandle = dirHandle;
            window.appSettings.downloadFolder = dirHandle.name;

            document.getElementById('download-folder-input').value = dirHandle.name;

            // Save the folder handle immediately so it persists across sessions
            await window.saveSettings(window.appSettings);
            console.log('Saved folder handle to IndexedDB:', dirHandle.name);
        } else {
            // Fallback: inform user about browser limitations
            alert('Your browser does not support folder selection. Downloads will go to your default download folder.');
        }
    } catch (error) {
        console.log('User cancelled folder selection');
    }
}

/**
 * Save settings from modal
 */
async function saveSettingsFromModal() {
    const userNameInput = document.getElementById('user-name-input');
    const userEmailInput = document.getElementById('user-email-input');
    const userSignatureInput = document.getElementById('user-signature-input');

    window.appSettings.userName = userNameInput.value.trim() || 'User';
    window.appSettings.userEmail = userEmailInput.value.trim() || '';
    window.appSettings.userSignature = userSignatureInput.value.trim() || '';

    if (await window.saveSettings(window.appSettings)) {
        hideSettingsModal();
        alert('Settings saved successfully!');
    } else {
        alert('Error saving settings. Please try again.');
    }
}

// ========== CONTACTS MODAL ==========

/**
 * Show contacts modal
 */
async function showContactsModal() {
    const modal = document.getElementById('contacts-modal');
    modal.classList.add('show');

    // Load and display contacts
    await window.loadAndDisplayContacts();

    // Reset form
    window.resetContactForm();
}

/**
 * Hide contacts modal
 */
function hideContactsModal() {
    const modal = document.getElementById('contacts-modal');
    modal.classList.remove('show');
    window.resetContactForm();
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
 * Build function formula with parameters passed directly
 */
function buildFunctionFormulaWithParams(functionType, baseExpression, param1, param2, param3, param4) {
    switch (functionType) {
        case 'ROUND':
            const decimals = param1 || '2';
            return `ROUND(${baseExpression}, ${decimals})`;

        case 'LEFT':
            if (!param1 || isNaN(parseInt(param1))) {
                throw new Error('LEFT function requires a numeric parameter (number of characters)');
            }
            return `LEFT(${baseExpression}, ${param1})`;

        case 'RIGHT':
            if (!param1 || isNaN(parseInt(param1))) {
                throw new Error('RIGHT function requires a numeric parameter (number of characters)');
            }
            return `RIGHT(${baseExpression}, ${param1})`;

        case 'MID':
            if (!param1 || isNaN(parseInt(param1)) || !param2 || isNaN(parseInt(param2))) {
                throw new Error('MID function requires two numeric parameters (start position and length)');
            }
            return `MID(${baseExpression}, ${param1}, ${param2})`;

        case 'REGEX':
            if (!param1) {
                throw new Error('REGEX function requires a pattern parameter');
            }
            const group = param2 || '1';
            return `REGEX(${baseExpression}, "${param1}", ${group})`;

        case 'SPLIT':
            if (!param1 || !param2 || isNaN(parseInt(param2))) {
                throw new Error('SPLIT function requires a delimiter and numeric index parameter');
            }
            return `SPLIT(${baseExpression}, "${param1}", ${param2})`;

        case 'REPLACE':
            return (param1 && param2) ? `REPLACE(${baseExpression}, "${param1}", "${param2}")` : baseExpression;

        case 'CONCAT':
            return param1 ? `CONCAT(${baseExpression}, "${param1}")` : baseExpression;

        case 'CONTAINS':
            return param1 ? `CONTAINS(${baseExpression}, "${param1}")` : baseExpression;

        case 'STARTSWITH':
            return param1 ? `STARTSWITH(${baseExpression}, "${param1}")` : baseExpression;

        case 'ENDSWITH':
            return param1 ? `ENDSWITH(${baseExpression}, "${param1}")` : baseExpression;

        case 'IF':
            if (param1 && param2 && param3 && param4) {
                let conditionExpr;
                switch (param1) {
                    case 'equals': conditionExpr = `${baseExpression} = "${param2}"`; break;
                    case 'not-equals': conditionExpr = `${baseExpression} ≠ "${param2}"`; break;
                    case 'contains': conditionExpr = `CONTAINS(${baseExpression}, "${param2}")`; break;
                    case 'greater': conditionExpr = `${baseExpression} > ${param2}`; break;
                    case 'less': conditionExpr = `${baseExpression} < ${param2}`; break;
                    default: return baseExpression;
                }
                return `IF(${conditionExpr}, "${param3}", "${param4}")`;
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

/**
 * Build function formula with parameters (legacy DOM-based version)
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

// ========== GLOBAL EXPORTS ==========

// Make modal functions globally accessible
window.showModal = showModal;
window.hideModal = hideModal;

// Fixed string modal
window.showFixedStringModal = showFixedStringModal;
window.closeFixedStringModal = closeFixedStringModal;
window.applyFixedString = applyFixedString;

// Calculation modal
window.showCalculationModal = showCalculationModal;
window.closeCalculationModal = closeCalculationModal;
window.insertColumnName = insertColumnName;
window.testFormula = testFormula;
window.getTestDataForFormula = getTestDataForFormula;
window.getAvailableSourceColumns = getAvailableSourceColumns;
window.applyCalculation = applyCalculation;

// Preview modal
window.showPreviewModal = showPreviewModal;
window.hidePreviewModal = hidePreviewModal;

// Settings modal
window.showSettingsModal = showSettingsModal;
window.hideSettingsModal = hideSettingsModal;
window.selectDownloadFolder = selectDownloadFolder;
window.saveSettingsFromModal = saveSettingsFromModal;
window.switchSettingsTab = switchSettingsTab;
window.loadFileMappingsList = loadFileMappingsList;
window.selectAllMappings = selectAllMappings;
window.deselectAllMappings = deselectAllMappings;
window.exportSelectedMappings = exportSelectedMappings;
window.deleteSelectedMappings = deleteSelectedMappings;
window.deleteSingleMappingFromList = deleteSingleMappingFromList;

// Contacts modal
window.showContactsModal = showContactsModal;
window.hideContactsModal = hideContactsModal;

// Function helpers
window.generateFunctionParams = generateFunctionParams;
window.buildFunctionFormula = buildFunctionFormula;