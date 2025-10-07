/**
 * Borderellen Converter - Template Manager
 * Handles template creation, editing, and management
 */

// Global template state - exposed to window for cross-module access
window.savedTemplates = {}; // Store multiple templates by ID
window.currentTemplateId = null; // Currently active template
window.borderellenTemplate = {
    id: null,
    name: '',
    version: '1.0',
    description: '',
    matchingKeyword: '',
    columns: [],
    originalFileName: ''
};
let templateColumnIdCounter = 0;

/**
 * Helper function to detect column data type based on name
 */
function guessColumnType(columnName) {
    const name = columnName.toLowerCase();

    // Date columns
    if (name.includes('datum') || name.includes('date') || name.includes('dtm') ||
        name.includes('periode') || name.includes('van') || name.includes('tot')) {
        return 'date';
    }

    // Number columns
    if (name.includes('%') || name.includes('bedrag') || name.includes('bruto') ||
        name.includes('netto') || name.includes('provisie') || name.includes('com') ||
        name.includes('aandeel') || name.includes('jaar') || name.includes('nr')) {
        return 'number';
    }

    // Default to text
    return 'text';
}

/**
 * Process Excel file to extract template columns
 */
async function processTemplateExcel(file) {
    try {
        const workbook = await ExcelCacheManager.getWorkbook(file);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

                // Get the range of the worksheet
                const range = XLSX.utils.decode_range(worksheet['!ref']);

                // Read first row to get column headers
                const columns = [];
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
                    const cell = worksheet[cellAddress];

                    if (cell && cell.v) {
                        const columnName = cell.v.toString().trim();
                        if (columnName) {
                            columns.push({
                                name: columnName,
                                type: guessColumnType(columnName),
                                required: false, // Default to not required
                                description: `Auto-imported from ${file.name}`
                            });
                        }
                    }
                }

        return {
            columns: columns,
            fileName: file.name,
            sheetName: firstSheetName
        };

    } catch (error) {
        throw error;
    }
}

/**
 * Load template from Excel file
 */
async function loadTemplateFromExcel(file) {
    try {
        const result = await processTemplateExcel(file);

        if (result.columns.length === 0) {
            alert('No column headers found in the Excel file. Make sure the first row contains column names.');
            return;
        }

        // Create new template ID
        const templateId = 'template_' + Date.now();
        const baseName = file.name.replace(/\.[^/.]+$/, ""); // Remove extension

        // Reset template and load new columns
        borderellenTemplate = {
            id: templateId,
            name: `${baseName} Template`,
            description: `Template automatically created from ${file.name} (${result.columns.length} columns)`,
            version: '1.0',
            columns: [],
            originalFileName: file.name,
            createdDate: new Date().toISOString()
        };

        templateColumnIdCounter = 0;

        // Add columns with IDs
        result.columns.forEach(col => {
            borderellenTemplate.columns.push({
                id: ++templateColumnIdCounter,
                ...col
            });
        });

        // Save template automatically (but not JSON yet)
        await saveTemplateToStorage(templateId, borderellenTemplate);
        savedTemplates[templateId] = { ...borderellenTemplate };
        currentTemplateId = templateId;

        // Update displays and go to edit mode
        updateTemplateSelector();
        updateActiveTemplateDisplay();
        showEditTemplateSection();

        alert(`Template created successfully!\n\nFound ${result.columns.length} columns from ${file.name}\nSheet: ${result.sheetName}\n\nClick "Save Changes" to export the JSON file.\n\nYou can now modify column types and descriptions as needed.`);

    } catch (error) {
        alert(`Error loading template from Excel: ${error.message}`);
        console.error('Template loading error:', error);
    }
}

/**
 * Load templates from storage and initialize
 */
async function loadDefaultTemplate() {
    const result = await loadTemplatesFromStorage();

    window.savedTemplates = result.templates;

    if (result.currentTemplateId && window.savedTemplates[result.currentTemplateId]) {
        window.borderellenTemplate = { ...window.savedTemplates[result.currentTemplateId] };
        window.currentTemplateId = result.currentTemplateId;

        // Find highest ID to continue counter
        templateColumnIdCounter = Math.max(...window.borderellenTemplate.columns.map(col => col.id), 0);

        updateTemplateForm();
        updateTemplateDisplay();
    }

    updateTemplateSelector();
    updateActiveTemplateDisplay();

    // Update button states after templates are loaded
    if (typeof updateButtonStates === 'function') {
        updateButtonStates();
    }

    showMainSection();
}

/**
 * Update template selector dropdown
 */
function updateTemplateSelector() {
    const selector = document.getElementById('template-selector');
    selector.innerHTML = '<option value="">Choose a template...</option>';

    Object.entries(window.savedTemplates).forEach(([id, template]) => {
        const option = document.createElement('option');
        option.value = id;
        option.textContent = `${template.name} (${template.columns.length} columns)`;
        selector.appendChild(option);
    });

    // Set the selector to the current active template
    if (window.currentTemplateId) {
        selector.value = window.currentTemplateId;
    }

    // Update button states
    if (typeof updateButtonStates === 'function') {
        updateButtonStates();
    }
}

/**
 * Select and activate a template
 */
function selectTemplate(templateId) {
    if (templateId && window.savedTemplates[templateId]) {
        window.borderellenTemplate = { ...window.savedTemplates[templateId] };
        window.currentTemplateId = templateId;

        // Find highest ID to continue counter
        templateColumnIdCounter = Math.max(...window.borderellenTemplate.columns.map(col => col.id), 0);

        updateTemplateForm();
        updateTemplateDisplay();
        updateActiveTemplateDisplay();
        updateTemplateSelector();

        // Save to IndexedDB for persistence
        if (typeof window.saveCurrentTemplateId === 'function') {
            window.saveCurrentTemplateId(templateId);
        }

        // Update button states after selection
        if (typeof updateButtonStates === 'function') {
            updateButtonStates();
        }
    } else {
        // No template selected
        window.borderellenTemplate = {
            id: null,
            name: '',
            version: '1.0',
            description: '',
            columns: [],
            originalFileName: ''
        };
        window.currentTemplateId = null;
        templateColumnIdCounter = 0;

        // Clear from IndexedDB for persistence
        if (typeof window.saveCurrentTemplateId === 'function') {
            window.saveCurrentTemplateId(null);
        }

        updateTemplateForm();
        updateTemplateDisplay();
        updateActiveTemplateDisplay();
        updateTemplateSelector();

        // Update button states after clearing selection
        if (typeof updateButtonStates === 'function') {
            updateButtonStates();
        }
    }
}

/**
 * Start new template creation process
 */
function createNewTemplate() {
    if (window.currentTemplateId && hasUnsavedChanges()) {
        if (!confirm('You have unsaved changes. Create new template anyway?')) {
            return;
        }
    }

    // Show the new template section with upload options
    if (typeof showNewTemplateSection === 'function') {
        showNewTemplateSection();
    }
}

/**
 * Create new manual template (called from the "Create Manual Template" button)
 */
function createManualTemplate() {
    const templateId = 'template_' + Date.now();
    window.borderellenTemplate = {
        id: templateId,
        name: 'New Template',
        description: 'Manually created template',
        version: '1.0',
        columns: [{
            id: 1,
            name: 'Column 1',
            type: 'text',
            required: false,
            description: ''
        }],
        originalFileName: '',
        createdDate: new Date().toISOString()
    };

    templateColumnIdCounter = 1;
    window.currentTemplateId = templateId;

    saveTemplateToStorage(templateId, window.borderellenTemplate);
    window.savedTemplates[templateId] = { ...window.borderellenTemplate };

    updateTemplateSelector();
    updateActiveTemplateDisplay();
    showEditTemplateSection();
}

/**
 * Delete current template
 */
async function deleteTemplate() {
    if (!window.currentTemplateId || !window.savedTemplates[window.currentTemplateId]) return;

    const templateName = window.savedTemplates[window.currentTemplateId].name;
    if (!confirm(`Delete template "${templateName}"? This cannot be undone.`)) {
        return;
    }

    // Delete from IndexedDB
    await deleteTemplateFromStorage(window.currentTemplateId);

    // Delete from memory
    delete window.savedTemplates[window.currentTemplateId];

    // Select first available template or none
    const remainingIds = Object.keys(window.savedTemplates);
    if (remainingIds.length > 0) {
        selectTemplate(remainingIds[0]);
    } else {
        selectTemplate('');
    }

    updateTemplateSelector();
    updateActiveTemplateDisplay();

    // Return to main section after deletion
    if (typeof showMainSection === 'function') {
        showMainSection();
    }
}

/**
 * Check for unsaved changes
 */
function hasUnsavedChanges() {
    if (!window.currentTemplateId || !window.savedTemplates[window.currentTemplateId]) return false;

    const saved = window.savedTemplates[window.currentTemplateId];
    return JSON.stringify(saved) !== JSON.stringify(window.borderellenTemplate);
}

/**
 * Update form fields with current template data
 */
function updateTemplateForm() {
    const templateNameInput = document.getElementById('template-name-input');
    const templateDescInput = document.getElementById('template-desc-input');
    const templateKeywordInput = document.getElementById('template-keyword-input');

    if (templateNameInput) templateNameInput.value = window.borderellenTemplate.name || '';
    if (templateDescInput) templateDescInput.value = window.borderellenTemplate.description || '';
    if (templateKeywordInput) templateKeywordInput.value = window.borderellenTemplate.matchingKeyword || '';
}

/**
 * Update active template display
 */
function updateActiveTemplateDisplay() {
    const nameElement = document.getElementById('active-template-name');
    const infoElement = document.getElementById('active-template-info');

    if (window.currentTemplateId && window.savedTemplates[window.currentTemplateId]) {
        const template = window.savedTemplates[window.currentTemplateId];
        nameElement.textContent = template.name;
        infoElement.textContent = `${template.columns.length} columns • ${template.description || 'No description'}`;
        nameElement.style.color = '#00bcd4';
    } else {
        nameElement.textContent = 'No Template Selected';
        infoElement.textContent = 'Please select a template from the dropdown below';
        nameElement.style.color = '#888';
    }
}

/**
 * Update template column display table
 */
function updateTemplateDisplay() {
    const templateTableBody = document.getElementById('template-table-body');
    templateTableBody.innerHTML = '';

    window.borderellenTemplate.columns.forEach((column, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>
                <input type="text" class="form-input" value="${column.name}"
                       style="min-width: 150px;"
                       onchange="updateTemplateColumn(${column.id}, 'name', this.value)">
            </td>
            <td>
                <select class="form-input" style="width: 100px;"
                        onchange="updateTemplateColumn(${column.id}, 'type', this.value)">
                    <option value="text" ${column.type === 'text' ? 'selected' : ''}>Text</option>
                    <option value="number" ${column.type === 'number' ? 'selected' : ''}>Number</option>
                    <option value="date" ${column.type === 'date' ? 'selected' : ''}>Date</option>
                </select>
            </td>
            <td>
                <input type="checkbox" ${column.required ? 'checked' : ''}
                       onchange="updateTemplateColumn(${column.id}, 'required', this.checked)">
            </td>
            <td>
                <input type="text" class="form-input" value="${column.description || ''}"
                       placeholder="Optional description"
                       onchange="updateTemplateColumn(${column.id}, 'description', this.value)">
            </td>
            <td>
                <button class="btn btn-secondary" style="padding: 4px 8px;"
                        onclick="removeTemplateColumn(${column.id})">Remove</button>
            </td>
        `;
        templateTableBody.appendChild(row);
    });
}

/**
 * Update template column property
 */
function updateTemplateColumn(columnId, field, value) {
    const column = borderellenTemplate.columns.find(col => col.id === columnId);
    if (column) {
        column[field] = value;
    }
}

/**
 * Add new column to template
 */
function addTemplateColumn() {
    const newColumn = {
        id: ++templateColumnIdCounter,
        name: 'New Column',
        type: 'text',
        required: false,
        description: ''
    };

    borderellenTemplate.columns.push(newColumn);
    updateTemplateDisplay();
}

/**
 * Remove column from template
 */
function removeTemplateColumn(columnId) {
    if (borderellenTemplate.columns.length <= 1) {
        alert('Template must have at least one column.');
        return;
    }

    if (confirm('Remove this column from the template?')) {
        borderellenTemplate.columns = borderellenTemplate.columns.filter(col => col.id !== columnId);
        updateTemplateDisplay();
    }
}

/**
 * Update template name
 */
function updateTemplateName(value) {
    borderellenTemplate.name = value;
}

/**
 * Update template description
 */
function updateTemplateDescription(value) {
    borderellenTemplate.description = value;
}

/**
 * Update template matching keyword
 */
function updateTemplateKeyword(value) {
    borderellenTemplate.matchingKeyword = value;
}

/**
 * Activate selected template
 */
function activateTemplate() {
    const selector = document.getElementById('template-selector');
    const selectedId = selector.value;

    if (selectedId && savedTemplates[selectedId]) {
        selectTemplate(selectedId);
        // Activate template button removed - using dropdown for immediate activation
        updateActiveTemplateDisplay();
        updateButtonStates();

        // Update mapping tab if it's currently active
        if (document.getElementById('mapping-tab').classList.contains('active')) {
            updateTemplateDropZones();
        }
    }
}

/**
 * Update button states based on selection
 */
function updateButtonStates() {
    const hasSelection = document.getElementById('template-selector').value !== '';
    const hasActiveTemplate = currentTemplateId && savedTemplates[currentTemplateId];

    // Activate button - enabled only when selection is different from current
    const selectedId = document.getElementById('template-selector').value;
    // Activate template button removed - using dropdown for immediate activation

    // Other buttons - enabled when we have an active template
    document.getElementById('edit-template-btn').disabled = !hasActiveTemplate;
    document.getElementById('delete-template-btn').disabled = !hasActiveTemplate;
}

/**
 * Save template changes
 */
async function saveTemplate() {
    if (!currentTemplateId) {
        alert('No template selected. Please create or select a template first.');
        return;
    }

    // Update template name and description from form
    const templateName = document.getElementById('template-name-input');
    const templateDesc = document.getElementById('template-desc-input');

    if (templateName) borderellenTemplate.name = templateName.value;
    if (templateDesc) borderellenTemplate.description = templateDesc.value;

    // Save to storage
    if (await saveTemplateToStorage(currentTemplateId, borderellenTemplate)) {
        savedTemplates[currentTemplateId] = { ...borderellenTemplate };

        updateTemplateSelector();
        updateActiveTemplateDisplay();

        // Export JSON file with same name as original Excel file or template name
        let jsonFileName;
        if (borderellenTemplate.originalFileName) {
            // Use original Excel filename
            const baseName = borderellenTemplate.originalFileName.replace(/\.[^/.]+$/, "");
            jsonFileName = baseName + '.json';
        } else {
            // Use template name
            jsonFileName = borderellenTemplate.name.replace(/[^a-z0-9]/gi, '_') + '.json';
        }

        // Get app settings for export
        const settings = await loadSettings();
        await exportTemplateAsJSON(borderellenTemplate, jsonFileName, { downloadFolderHandle: settings.downloadFolderHandle });

        alert(`Template saved successfully!\n\nJSON file "${jsonFileName}" has been downloaded.`);
    } else {
        alert('Error saving template');
    }
}

/**
 * Export current template as JSON
 */
async function exportTemplate() {
    if (!window.currentTemplateId || !window.borderellenTemplate.name) {
        alert('No template selected to export.');
        return;
    }

    const settings = await loadSettings();
    const filename = `${window.borderellenTemplate.name.replace(/[^a-z0-9]/gi, '_')}.json`;
    await exportTemplateAsJSON(window.borderellenTemplate, filename, { downloadFolderHandle: settings.downloadFolderHandle });
}

/**
 * Load template from JSON file
 */
async function loadFromJSON(file) {
    const result = await loadTemplateFromJSON(file);

    if (!result.success) {
        alert(`Error importing JSON template: ${result.error}`);
        return;
    }

    const imported = result.template;

    // Create new template ID to avoid conflicts
    const templateId = 'template_' + Date.now();
    const baseName = file.name.replace(/\.[^/.]+$/, ""); // Remove extension

    // Assign new IDs to columns to avoid conflicts
    templateColumnIdCounter = 0;
    imported.columns.forEach(col => {
        col.id = ++templateColumnIdCounter;
    });

    // Create new template object
    window.borderellenTemplate = {
        id: templateId,
        name: imported.name || `${baseName} Template`,
        description: imported.description || `Template imported from ${file.name}`,
        version: imported.version || '1.0',
        columns: imported.columns,
        originalFileName: imported.originalFileName || file.name,
        createdDate: new Date().toISOString(),
        importedFrom: file.name
    };

    // Save template
    await saveTemplateToStorage(templateId, window.borderellenTemplate);
    window.savedTemplates[templateId] = { ...window.borderellenTemplate };
    window.currentTemplateId = templateId;

    // Update displays and go to edit mode
    updateTemplateSelector();
    updateActiveTemplateDisplay();
    showEditTemplateSection();

    alert(`Template imported successfully!\n\nLoaded ${window.borderellenTemplate.columns.length} columns from ${file.name}\n\nYou can now modify the template and save to export updated JSON.`);
}

// Export essential functions to window for cross-module access
window.guessColumnType = guessColumnType;
window.loadFromJSON = loadFromJSON;
window.loadDefaultTemplate = loadDefaultTemplate;
window.updateTemplateSelector = updateTemplateSelector;
window.updateActiveTemplateDisplay = updateActiveTemplateDisplay;
window.selectTemplate = selectTemplate;
window.createNewTemplate = createNewTemplate;
window.createManualTemplate = createManualTemplate;
window.saveTemplate = saveTemplate;
window.exportTemplate = exportTemplate;
window.deleteTemplate = deleteTemplate;