/**
 * Borderellen Converter - File Manager
 * Handles file upload, processing, and management
 */

// Global file storage - accessible from other modules
window.uploadedFiles = [];
let fileIdCounter = 0;

// Color mapping for different broker types
const brokerColors = {
    'AON': '#004d55',
    'VGA': '#4d3300',
    'BCI': '#1a4d00',
    'Voogt': '#4d0033',
    'custom': '#2d5f2f',
    'Unknown': '#4d1a00',
    'Error': '#f44336',
    'PDF': '#666'
};

/**
 * Format file size in human readable format
 * @param {number} bytes - File size in bytes
 * @returns {string} Formatted file size
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Process Excel file to get basic record count (fallback for unknown formats)
 * @param {File} file - Excel file to process
 * @returns {Promise<number>} Number of records
 */
async function processExcelFile(file) {
    try {
        const workbook = await ExcelCacheManager.getWorkbook(file);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        return jsonData.length;
    } catch (error) {
        console.error('Error processing Excel file:', error);
        return 0;
    }
}

/**
 * Add file to upload table with proper broker parsing
 * @param {File} file - File to add
 */
async function addFileToTable(file) {
    const fileId = ++fileIdCounter;

    // Process file using enhanced broker detection and parsing
    let recordCount = '-';
    let status = 'Processing...';
    let statusClass = 'status-warning';
    let parsedData = [];
    let brokerInfo = null;
    let patternAnalysis = null;

    if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        file.name.endsWith('.xlsx')) {
        try {
            // FLOW 1: First check if this is a built-in parser (no analysis needed)
            const detection = await detectBrokerType(file.name);
            console.log('Initial detection for', file.name, ':', detection);

            let patternAnalysis = null;
            let tempFileData = { file: file, name: file.name };

            if (detection.type === 'built-in') {
                console.log('FLOW 1: Using built-in parser:', detection.parser);
            } else if (detection.type === 'custom') {
                console.log('FLOW 2: Using keyword-matched custom template:', detection.name);
                // Template already found via keyword matching in detectBrokerType
            } else {
                // FLOW 3: Unknown format - run pattern analysis for potential template creation
                console.log('FLOW 3: Running pattern analysis for unknown format:', file.name);
                patternAnalysis = await DataPatternAnalyzer.analyzeFile(file);
                console.log('Pattern analysis completed:', patternAnalysis);

                tempFileData.patternAnalysis = patternAnalysis;
            }

            // Template matching is now handled directly in detectBrokerType via keywords

            // Use enhanced broker processing that checks custom templates too
            const result = await processBrokerFile(tempFileData);

            if (result.success) {
                parsedData = result.data;
                // Use actual parsed data length, not just the reported recordCount
                recordCount = parsedData ? parsedData.length : 0;
                brokerInfo = {
                    ...result.brokerInfo,
                    color: brokerColors[result.brokerInfo.type] || brokerColors[result.brokerInfo.parser] || brokerColors['Unknown']
                };
                status = recordCount > 0 ? 'Parsed Successfully' : 'No Valid Records';
                statusClass = recordCount > 0 ? 'status-success' : 'status-warning';

                // Mark if auto-applied template was used
                if (detection.type === 'custom') {
                    status = `Auto-processed with ${detection.name}`;
                    brokerInfo.name = detection.name;
                }
            } else if (result.needsTemplate) {
                // Check if this was supposed to be a template-linked file
                if (detection.type === 'custom') {
                    console.error(`Template processing failed for ${detection.name}:`, result.error);
                    brokerInfo = { type: 'Error', name: 'Template Error', color: brokerColors['Error'] };
                    status = `Template Error: ${detection.name}`;
                    statusClass = 'status-error';
                    recordCount = 0;
                } else {
                    // Genuinely unknown format that needs template creation
                    brokerInfo = { type: 'Unknown', name: 'Unknown Format', color: brokerColors['Unknown'] };
                    recordCount = await processExcelFile(file);
                    status = 'Create Template';
                    statusClass = 'status-warning';
                }
            } else {
                brokerInfo = { type: 'Error', name: 'Parse Error', color: brokerColors['Error'] };
                recordCount = 0;
                status = `Parse Error: ${result.error}`;
                statusClass = 'status-error';
            }
        } catch (error) {
            console.error('Error processing file:', error);
            brokerInfo = { type: 'Error', name: 'Error', color: brokerColors['Error'] };
            status = `Error: ${error.message}`;
            statusClass = 'status-error';
            recordCount = 0;
        }
    } else if (file.type === 'application/pdf') {
        brokerInfo = { type: 'PDF', name: 'PDF File', color: brokerColors['PDF'] };
        status = 'PDF - Manual Processing Required';
        statusClass = 'status-warning';
    }

    // Fallback broker info if not set
    if (!brokerInfo) {
        brokerInfo = { type: 'Unknown', name: 'Unknown', color: brokerColors['Unknown'] };
    }

    const fileData = {
        id: fileId,
        file: file,
        name: file.name,
        broker: brokerInfo,
        size: formatFileSize(file.size),
        status: status,
        statusClass: statusClass,
        recordCount: recordCount,
        parsedData: parsedData, // Store the parsed data
        selectedTemplateId: null, // Will be set if template was auto-applied
        patternAnalysis: null // Will store the analysis result for reuse
    };

    // Store pattern analysis for reuse if Excel file was processed
    if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        file.name.endsWith('.xlsx')) {
        // Pattern analysis was computed above, store it in fileData
        fileData.patternAnalysis = patternAnalysis;
    }

    // Template associations are now handled via keyword matching in detectBrokerType

    window.uploadedFiles.push(fileData);
    updateFilesDisplay();
}

/**
 * Update files display table
 */
function updateFilesDisplay() {
    const noFilesMessage = document.getElementById('no-files-message');
    const filesTableContainer = document.getElementById('files-table-container');
    const filesTableBody = document.getElementById('files-table-body');

    if (window.uploadedFiles.length === 0) {
        noFilesMessage.style.display = 'block';
        filesTableContainer.style.display = 'none';
        return;
    }

    noFilesMessage.style.display = 'none';
    filesTableContainer.style.display = 'block';

    filesTableBody.innerHTML = '';

    window.uploadedFiles.forEach(fileData => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${fileData.name}</td>
            <td>
                <span style="background: ${fileData.broker.color}; padding: 4px 8px; border-radius: 4px; font-size: 12px;">
                    ${fileData.broker.name}
                </span>
            </td>
            <td>${fileData.size}</td>
            <td>
                <span class="status-indicator ${fileData.statusClass}"></span>
                ${fileData.status}
            </td>
            <td>${fileData.recordCount}</td>
            <td>
                ${fileData.broker.type === 'Unknown' ?
                    `<select class="form-input" style="max-width: 150px; font-size: 12px;" onchange="selectTemplateForFile(${fileData.id}, this.value)" id="template-selector-${fileData.id}">
                        <option value="">Select Template...</option>
                    </select>` :
                    `<button class="btn btn-secondary" style="padding: 4px 8px; font-size: 12px;" onclick="previewFile(${fileData.id})">Preview</button>`
                }
                <button class="btn btn-secondary" style="padding: 4px 8px; font-size: 12px; margin-left: 4px;" onclick="removeFile(${fileData.id})">Remove</button>
            </td>
        `;
        filesTableBody.appendChild(row);
    });

    // Update mapping file selector
    updateMappingFileSelector();

    // Populate template dropdowns for unknown files
    populateTemplateDropdowns();
}

/**
 * Populate template selection dropdowns for unknown files
 */
async function populateTemplateDropdowns() {
    try {
        // Load unified file mappings
        const allMappings = await window.loadAllFileMappings();
        const allTemplates = allMappings.map(m => ({ ...m, source: m.creationMethod || 'unknown' }));

        // Update each template selector
        window.uploadedFiles.forEach(fileData => {
            if (fileData.broker.type === 'Unknown') {
                const selector = document.getElementById(`template-selector-${fileData.id}`);
                if (selector) {
                    // Clear existing options except the first
                    while (selector.options.length > 1) {
                        selector.remove(1);
                    }

                    // Add template options
                    allTemplates.forEach(template => {
                        const option = document.createElement('option');
                        option.value = `${template.source}:${template.id}`;
                        option.textContent = `${template.name} (${template.source})`;
                        selector.appendChild(option);
                    });

                    // Set selected value if file has associated template
                    if (fileData.selectedTemplateId) {
                        selector.value = fileData.selectedTemplateId;
                    }

                    // Add option to create new template
                    const createOption = document.createElement('option');
                    createOption.value = 'create-new';
                    createOption.textContent = '+ Create New Template';
                    selector.appendChild(createOption);
                }
            }
        });

    } catch (error) {
        console.error('Error loading templates for dropdowns:', error);
    }
}

/**
 * Handle template selection for a file
 * @param {number} fileId - File ID
 * @param {string} templateValue - Selected template value
 */
async function selectTemplateForFile(fileId, templateValue) {
    const fileData = window.uploadedFiles.find(f => f.id === fileId);
    if (!fileData) {
        alert('File not found.');
        return;
    }

    if (templateValue === 'create-new') {
        // Navigate to Broker Mapping tab to create new template
        createTemplateForFile(fileId);
        return;
    }

    if (!templateValue) {
        // No template selected
        return;
    }

    try {
        const [source, templateId] = templateValue.split(':');
        let template = null;

        // Load the selected template from unified store
        const allMappings = await window.loadAllFileMappings();
        template = allMappings.find(t => t.id === templateId);

        if (!template) {
            alert('Selected template not found.');
            return;
        }

        // Process the file with the selected template using the orchestrator
        const result = await processBrokerFile({
            ...fileData,
            // Override detection to use custom template parser that respects parsingConfig
            detectionOverride: {
                type: 'custom',
                parser: 'GenericBrokerParser',
                template: template,
                name: template.name
            }
        });

        if (result.success) {
            // Update file with processed results
            fileData.parsedData = result.data;
            fileData.recordCount = result.recordCount;
            fileData.status = `Processed with ${template.name}`;
            fileData.statusClass = 'status-success';
            fileData.broker = {
                ...fileData.broker,
                type: 'custom-template',
                name: template.name
            };

            // Suggest adding keyword to template if it doesn't have one
            fileData.selectedTemplateId = templateValue;
            await suggestKeywordForTemplate(fileData.name, template);

            // Update display
            updateFilesDisplay();

            alert(`Successfully processed ${result.recordCount} records using template "${template.name}"!`);
        } else {
            alert(`Failed to process with template: ${result.error}`);
        }

    } catch (error) {
        console.error('Error applying template to file:', error);
        alert(`Error applying template: ${error.message}`);
    }
}

/**
 * Suggest adding a keyword to template for automatic matching if it doesn't have one
 * @param {string} filename - Original filename that was processed
 * @param {Object} template - Template object
 */
async function suggestKeywordForTemplate(filename, template) {
    try {
        // Skip if template already has a keyword
        if (template.matchingKeyword && template.matchingKeyword.trim() !== '') {
            console.log(`Template "${template.name}" already has keyword: "${template.matchingKeyword}"`);
            return;
        }

        // Extract potential keyword from filename
        const suggestedKeyword = extractKeywordFromFilename(filename);

        if (suggestedKeyword) {
            const userConfirm = confirm(
                `Template "${template.name}" doesn't have a matching keyword. ` +
                `Would you like to add "${suggestedKeyword}" as the keyword so similar files are automatically processed? ` +
                `(You can always change this later)`
            );

            if (userConfirm) {
                // Update the template with the keyword
                template.matchingKeyword = suggestedKeyword;
                await CustomBrokerTemplateManager.saveTemplate(template);
                console.log(`Added keyword "${suggestedKeyword}" to template "${template.name}"`);

                alert(`Keyword "${suggestedKeyword}" added to template "${template.name}". Similar files will now be processed automatically!`);
            }
        }
    } catch (error) {
        console.error('Error suggesting keyword for template:', error);
    }
}

/**
 * Extract a potential keyword from filename for template matching
 * @param {string} filename - Filename to analyze
 * @returns {string} Suggested keyword or empty string
 */
function extractKeywordFromFilename(filename) {
    // Remove file extension
    const baseName = filename.replace(/\.(xlsx|xls|pdf)$/i, '');

    // Split by common separators and take the first meaningful part
    const parts = baseName.split(/[_\-\s\.]+/);

    // Look for the first part that's not a number or date-like pattern
    for (const part of parts) {
        const cleaned = part.trim();

        // Skip if empty, purely numeric, or looks like a date
        if (!cleaned || /^\d+$/.test(cleaned) || /^\d{2,4}$/.test(cleaned)) {
            continue;
        }

        // Return the first meaningful word (at least 3 characters)
        if (cleaned.length >= 3) {
            return cleaned;
        }
    }

    return '';
}


/**
 * Remove file from upload list
 * @param {number} fileId - File ID to remove
 */
function removeFile(fileId) {
    window.uploadedFiles = window.uploadedFiles.filter(f => f.id !== fileId);
    updateFilesDisplay();
}

/**
 * Preview file data using the main preview modal
 * @param {number} fileId - File ID to preview
 */
function previewFile(fileId) {
    const fileData = window.uploadedFiles.find(f => f.id === fileId);
    if (!fileData) {
        alert('File not found.');
        return;
    }

    // Use the unified preview function
    if (typeof window.previewFileData === 'function') {
        window.previewFileData(fileData);
    } else {
        // Fallback to simple alert if main preview not available
        alert(`Preview for ${fileData.name}\nBroker: ${fileData.broker.name}\nRecords: ${fileData.recordCount}\n\nStatus: ${fileData.status}`);
    }
}

/**
 * Create template for unknown file formats
 * @param {number} fileId - File ID to create template for
 */
function createTemplateForFile(fileId) {
    const fileData = window.uploadedFiles.find(f => f.id === fileId);
    if (fileData) {
        // Switch to File Mapping tab and auto-select the file
        document.querySelector('[data-tab="mapping"]').click();

        // Auto-select the file in the mapping tab
        setTimeout(() => {
            const selector = document.getElementById('mapping-file-selector');
            if (selector) {
                selector.value = fileId;
                // Trigger the change event to load the file for mapping
                selector.dispatchEvent(new Event('change'));
            }
        }, 100);
    } else {
        alert(`File not found. Please re-upload the file and try again.`);
    }
}

/**
 * Clear all uploaded files
 */
function clearAllFiles() {
    if (window.uploadedFiles.length === 0) return;

    if (confirm(`Remove all ${window.uploadedFiles.length} uploaded files?`)) {
        window.uploadedFiles = [];
        updateFilesDisplay();
    }
}

/**
 * Handle file uploads (drag & drop or file input)
 * @param {FileList} files - Files to process
 */
async function handleFiles(files) {
    const filesToProcess = [];

    // First validate all files
    Array.from(files).forEach(file => {
        const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                          'application/vnd.ms-excel', 'application/pdf'];
        const validExtensions = ['.xlsx', '.xls', '.pdf'];

        if (validTypes.includes(file.type) || validExtensions.some(ext => file.name.toLowerCase().endsWith(ext))) {
            // Check if file already exists
            if (!window.uploadedFiles.some(f => f.name === file.name && f.file.size === file.size)) {
                filesToProcess.push(file);
            } else {
                alert(`File "${file.name}" is already uploaded.`);
            }
        } else {
            alert(`File "${file.name}" is not a supported format. Please upload Excel (.xlsx) or PDF files.`);
        }
    });

    // Process files sequentially and wait for all to complete
    for (const file of filesToProcess) {
        await addFileToTable(file);
    }

    // Auto-navigate based on processing results (only if files were actually processed)
    if (filesToProcess.length > 0) {
        setTimeout(() => {
            checkAndAutoNavigate();
        }, 500); // Small delay to ensure UI updates are complete
    }
}

/**
 * Check processing results and auto-navigate to appropriate tab
 */
function checkAndAutoNavigate() {
    if (!window.uploadedFiles || window.uploadedFiles.length === 0) {
        return; // No files to analyze
    }

    // Analyze processing results
    const successfulFiles = window.uploadedFiles.filter(f =>
        f.statusClass === 'status-success' && f.recordCount > 0
    );
    const failedFiles = window.uploadedFiles.filter(f =>
        f.statusClass === 'status-error' ||
        f.statusClass === 'status-warning' ||
        f.recordCount === 0
    );

    // Check for unknown files that could benefit from mapping (like perfect D4 tables)
    const unknownFiles = window.uploadedFiles.filter(f =>
        f.broker && f.broker.name === 'Unknown Format'
    );

    console.log(`Auto-navigation check: ${successfulFiles.length} successful, ${failedFiles.length} failed/problematic, ${unknownFiles.length} unknown format`);

    if (unknownFiles.length > 0) {
        // Unknown files detected - go to Broker Mapping tab for template creation/mapping
        console.log('Unknown files detected - navigating to Broker Mapping tab for template creation');
        const mappingTab = document.querySelector('[data-tab="mapping"]');
        if (mappingTab) {
            mappingTab.click();
            // The existing autoSelectFileInMappingTab() will prioritize unknown files
        }
    } else if (failedFiles.length > 0) {
        // Some files need attention - go to Broker Mapping tab
        console.log('Some files need attention - navigating to Broker Mapping tab');
        const mappingTab = document.querySelector('[data-tab="mapping"]');
        if (mappingTab) {
            mappingTab.click();
            // The existing autoSelectFileInMappingTab() will prioritize problematic files
        }
    } else if (failedFiles.length === 0 && successfulFiles.length > 0) {
        // All files processed successfully - go to Results tab
        console.log('All files successful - navigating to Results tab');
        const resultsTab = document.querySelector('[data-tab="results"]');
        if (resultsTab) {
            resultsTab.click();
        }
    }
    // If no successful files and no failed files, stay on Upload tab (e.g., only PDFs)
}

/**
 * Set up file upload event listeners
 */
function initializeFileUpload() {
    // File upload functionality
    const uploadZone = document.getElementById('upload-zone');
    const fileInput = document.getElementById('file-input');
    const browseBtn = document.getElementById('browse-btn');
    const clearAllBtn = document.getElementById('clear-all-btn');

    // Upload zone click
    uploadZone.addEventListener('click', () => {
        fileInput.click();
    });

    // Browse button
    browseBtn.addEventListener('click', () => {
        fileInput.click();
    });

    // Clear all button
    clearAllBtn.addEventListener('click', clearAllFiles);

    // File input change
    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        e.target.value = ''; // Reset input so same file can be selected again
    });

    // Drag and drop events
    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files);
    });

    // Prevent default drag behavior on body
    document.body.addEventListener('dragover', (e) => e.preventDefault());
    document.body.addEventListener('drop', (e) => e.preventDefault());
}

/**
 * Load and display templates with their keywords for management
 */
async function loadKeywordManagement() {
    try {
        // Use unified file mappings store only
        const allMappings = await window.loadAllFileMappings();
        const allTemplates = allMappings.map(m => ({ ...m, source: m.creationMethod || 'unknown' }));

        const tableBody = document.getElementById('keyword-templates-table-body');
        const noTemplatesMsg = document.getElementById('no-templates-message');
        const templateCountDisplay = document.getElementById('template-count-display');

        if (allTemplates.length === 0) {
            tableBody.innerHTML = '';
            noTemplatesMsg.style.display = 'block';
            templateCountDisplay.textContent = 'No file mappings found';
            return;
        }

        noTemplatesMsg.style.display = 'none';
        templateCountDisplay.textContent = `${allTemplates.length} file mappings found`;

        tableBody.innerHTML = allTemplates.map(template => {
            const keyword = template.matchingKeyword || '';
            const hasKeyword = keyword.trim() !== '';
            const formattedDate = template.lastModified || template.created || 'Unknown';

            return `
                <tr>
                    <td>${template.name}</td>
                    <td>
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <input type="text"
                                   class="form-input"
                                   style="max-width: 150px; font-size: 14px;"
                                   value="${keyword}"
                                   placeholder="Enter keyword..."
                                   id="keyword-input-${template.id}"
                                   ${hasKeyword ? '' : 'style="border-color: #ff9800;"'}>
                            ${hasKeyword ?
                                '<span style="color: #4caf50; font-size: 12px;">✓</span>' :
                                '<span style="color: #ff9800; font-size: 12px;">!</span>'
                            }
                        </div>
                    </td>
                    <td>
                        <span class="template-type-badge template-type-${template.source}">
                            ${template.source === 'custom' ? 'Template Builder' : 'Drag & Drop'}
                        </span>
                    </td>
                    <td style="font-size: 12px; color: #888;">
                        ${new Date(formattedDate).toLocaleDateString()}
                    </td>
                    <td>
                        <div style="display: flex; gap: 4px; flex-wrap: wrap;">
                            <button class="btn btn-secondary"
                                    style="padding: 4px 8px; font-size: 12px;"
                                    onclick="saveTemplateKeyword('${template.id}', '${template.source || 'unknown'}')">
                                Save Keyword
                            </button>
                            <button class="btn btn-secondary edit-mapping-btn"
                                    style="padding: 4px 8px; font-size: 12px; background-color: #007bff; border-color: #007bff;"
                                    data-template-id="${template.id || ''}"
                                    data-template-name="${encodeURIComponent(template.name || 'Unnamed Template')}">
                                Edit Mapping
                            </button>
                            <button class="btn btn-secondary delete-template-btn"
                                    style="padding: 4px 8px; font-size: 12px; background-color: #d32f2f; border-color: #d32f2f;"
                                    data-template-id="${template.id || ''}"
                                    data-template-source="${template.source || 'unknown'}"
                                    data-template-name="${encodeURIComponent(template.name || 'Unnamed Template')}">
                                Delete
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');

        // Remove any existing event listeners first
        document.querySelectorAll('.delete-template-btn, .edit-mapping-btn').forEach(btn => {
            const newBtn = btn.cloneNode(true);
            btn.parentNode.replaceChild(newBtn, btn);
        });

        // Add fresh event listeners for edit buttons
        document.querySelectorAll('.edit-mapping-btn').forEach(button => {
            button.addEventListener('click', function(event) {
                event.preventDefault();
                event.stopImmediatePropagation();

                const templateId = this.dataset.templateId;
                const templateName = decodeURIComponent(this.dataset.templateName || 'Unnamed Template');

                if (!templateId) {
                    alert('Error: Template ID is missing. Cannot edit.');
                    return;
                }

                window.editFileMappingFromTable(templateId, templateName);
            }, { once: true });
        });

        // Add fresh event listeners for delete buttons
        document.querySelectorAll('.delete-template-btn').forEach(button => {
            button.addEventListener('click', function(event) {
                event.preventDefault();
                event.stopImmediatePropagation();

                const templateId = this.dataset.templateId;
                const templateSource = this.dataset.templateSource;
                const templateName = decodeURIComponent(this.dataset.templateName || 'Unnamed Template');

                if (!templateId) {
                    alert('Error: Template ID is missing. Cannot delete.');
                    return;
                }

                window.deleteFileMappingWithUI(templateId, templateSource, templateName);
            }, { once: true });
        });

    } catch (error) {
        console.error('Error loading keyword management:', error);
        document.getElementById('template-count-display').textContent = 'Error loading file mappings';
    }
}

/**
 * Save keyword for a specific template
 * @param {string} templateId - Template ID
 * @param {string} source - Template source ('custom' or 'broker')
 */
async function saveTemplateKeyword(templateId, source) {
    try {
        const keywordInput = document.getElementById(`keyword-input-${templateId}`);
        const newKeyword = keywordInput.value.trim();

        // Load the template from unified store
        const allMappings = await window.loadAllFileMappings();
        const template = allMappings.find(t => t.id === templateId);

        if (!template) {
            alert('Template not found');
            return;
        }

        // Check for duplicate keywords (if keyword is provided and changed)
        if (newKeyword && newKeyword !== template.matchingKeyword) {
            const duplicateKeyword = allMappings.find(t =>
                t.id !== templateId && // Exclude current template
                t.matchingKeyword &&
                t.matchingKeyword.toLowerCase() === newKeyword.toLowerCase()
            );

            if (duplicateKeyword) {
                alert(
                    `A template with keyword "${newKeyword}" already exists (Template: "${duplicateKeyword.name}").\n\n` +
                    `Keywords must be unique to ensure predictable template selection.\n\n` +
                    `Please choose a different keyword.`
                );
                return;
            }
        }

        // Update keyword
        template.matchingKeyword = newKeyword;
        template.lastModified = new Date().toISOString();

        // Save to unified store
        await window.saveFileMapping(template);

        // Refresh display
        await loadKeywordManagement();

        alert(`Keyword "${newKeyword}" saved for template "${template.name}"!`);
    } catch (error) {
        console.error('Error saving template keyword:', error);
        alert(`Error saving keyword: ${error.message}`);
    }
}


/**
 * Delete a file mapping with UI confirmation and refresh
 * @param {string} templateId - File mapping ID to delete
 * @param {string} source - Creation method ('custom' or 'broker')
 * @param {string} templateName - File mapping name for confirmation
 */
async function deleteFileMappingWithUI(templateId, source, templateName) {
    // Prevent multiple simultaneous calls
    if (deleteFileMappingWithUI.isDeleting) {
        return;
    }

    try {
        deleteFileMappingWithUI.isDeleting = true;

        // Validate inputs
        if (!templateId || templateId === 'undefined') {
            alert('Error: Invalid template ID. Cannot delete.');
            return;
        }

        if (!templateName || templateName === 'undefined') {
            alert('Error: Invalid template name. Cannot delete.');
            return;
        }

        // Simple confirmation - it's just an IndexedDB entry, not critical data
        const confirmed = confirm(
            `Delete file mapping "${templateName}"?\n\n` +
            `This will remove the mapping from browser storage.\n` +
            `The exported JSON files remain on your computer.`
        );

        if (!confirmed) return;

        // Delete from unified store only
        const success = await window.deleteFileMapping(templateId);

        if (success) {
            alert(`File mapping "${templateName}" deleted.`);
            await loadKeywordManagement();
        } else {
            alert(`Failed to delete "${templateName}". Check console for errors.`);
        }

    } catch (error) {
        console.error('Error deleting file mapping:', error);
        alert(`❌ Error deleting file mapping: ${error.message}`);
    } finally {
        deleteFileMappingWithUI.isDeleting = false;
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
            if (window.appSettings && window.appSettings.downloadFolderHandle) {
                options.startIn = window.appSettings.downloadFolderHandle;
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
    console.log('downloadFolderHandle:', window.appSettings?.downloadFolderHandle);

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
            if (window.appSettings && window.appSettings.downloadFolderHandle) {
                options.startIn = window.appSettings.downloadFolderHandle;
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

// Export essential functions to window for cross-module access
window.initializeFileUpload = initializeFileUpload;
window.loadKeywordManagement = loadKeywordManagement;
window.saveTemplateKeyword = saveTemplateKeyword;
window.deleteFileMappingWithUI = deleteFileMappingWithUI;
window.extractKeywordFromFilename = extractKeywordFromFilename;
window.downloadToPreferredFolder = downloadToPreferredFolder;
window.downloadExcelToPreferredFolder = downloadExcelToPreferredFolder;