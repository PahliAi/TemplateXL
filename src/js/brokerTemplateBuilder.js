/**
 * Borderellen Converter - Broker Template Builder
 * Visual interface for creating custom broker templates with parsing rules
 */

class BrokerTemplateBuilder {
    constructor(containerId) {
        this.container = document.getElementById(containerId);
        this.currentFile = null;
        this.analysisResults = null;
        this.currentConfig = this.getDefaultConfig();
        this.previewData = null;
        this.isVisible = false;

        this.initialize();
    }

    /**
     * Gets default configuration for new templates
     * @returns {Object} Default configuration
     */
    getDefaultConfig() {
        return {
            templateName: '',
            matchingKeyword: '',
            patternDescription: '',
            filenamePattern: '',
            dataStartMethod: 'auto-detect',
            skipRows: 0,
            headerRowOffset: -1,
            rowProcessing: {
                type: 'single',
                multiRowConfig: null
            },
            rowFilters: [],
            dataValidation: [],
            columnMapping: {}
        };
    }

    /**
     * Initialize the template builder interface
     */
    initialize() {
        if (!this.container) {
            console.error('Template builder container not found');
            return;
        }

        this.createInterface();
        this.attachEventListeners();
    }

    /**
     * Creates the main interface structure
     */
    createInterface() {
        this.container.innerHTML = `
            <div class="template-builder-overlay" style="display: none;">
                <div class="template-builder-modal">
                    <div class="template-builder-header">
                        <h2>Create Custom Broker Template</h2>
                        <button class="template-builder-close" id="close-template-builder">×</button>
                    </div>

                    <div class="template-builder-content">
                        <!-- Step 1: File Upload & Analysis -->
                        <div class="builder-step" id="step-upload">
                            <h3>Step 1: Upload Sample File</h3>
                            <p>Upload a sample Excel file to analyze its structure and create parsing rules.</p>

                            <div class="upload-zone-small" id="sample-file-upload">
                                <div class="upload-icon">📁</div>
                                <div class="upload-text">Drop sample Excel file here</div>
                            </div>
                            <input type="file" id="sample-file-input" accept=".xlsx,.xls" style="display: none;">

                            <div id="analysis-results" style="display: none;">
                                <h4>Analysis Results</h4>
                                <div class="analysis-summary"></div>
                                <div class="pattern-visualization"></div>
                            </div>
                        </div>

                        <!-- Step 2: Template Configuration -->
                        <div class="builder-step" id="step-config" style="display: none;">
                            <h3>Step 2: Configure Template</h3>

                            <div class="config-section">
                                <div class="form-group">
                                    <label class="form-label">Template Name</label>
                                    <input type="text" class="form-input" id="template-name-input"
                                           placeholder="Enter a descriptive name for this template">
                                </div>

                                <div class="form-group">
                                    <label class="form-label">Matching Keyword</label>
                                    <input type="text" class="form-input" id="template-keyword-input"
                                           placeholder="Enter keyword to match files (e.g., 'Verkoper')">
                                    <small class="form-help">Files containing this keyword will automatically use this template</small>
                                </div>

                                <div class="form-group">
                                    <label class="form-label">Filename Pattern (regex)</label>
                                    <input type="text" class="form-input" id="filename-pattern-input"
                                           placeholder="e.g., ^CustomBroker_\\\\d{2}-\\\\d{4}\\\\.xlsx$">
                                    <small class="form-help">Legacy pattern for backward compatibility (optional)</small>
                                </div>
                            </div>

                            <div class="parsing-config-section">
                                <h4>Parsing Configuration</h4>

                                <div class="form-group">
                                    <label class="form-label">Data Start Method</label>
                                    <select class="form-input" id="data-start-method">
                                        <option value="auto-detect">Auto-detect (recommended)</option>
                                        <option value="skip-rows">Skip first N rows</option>
                                        <option value="find-date">Find first date column</option>
                                        <option value="find-pattern">Find by pattern</option>
                                    </select>
                                </div>

                                <div class="form-group" id="skip-rows-config" style="display: none;">
                                    <label class="form-label">Number of rows to skip</label>
                                    <input type="number" class="form-input" id="skip-rows-input" min="0" max="50" value="0">
                                </div>

                                <div class="form-group">
                                    <label class="form-label">Row Processing</label>
                                    <select class="form-input" id="row-processing-type">
                                        <option value="single">Single row per record</option>
                                        <option value="multi-row">Multi-row processing (advanced)</option>
                                    </select>
                                </div>
                            </div>

                            <div class="row-filters-section">
                                <h4>Row Filtering Rules</h4>
                                <div id="row-filters-container">
                                    <p class="no-filters-message">No filters configured. Add filters to skip unwanted rows.</p>
                                </div>
                                <button class="btn btn-secondary" id="add-row-filter">Add Filter</button>
                            </div>
                        </div>

                        <!-- Step 3: Column Mapping -->
                        <div class="builder-step" id="step-mapping" style="display: none;">
                            <h3>Step 3: Map Columns</h3>

                            <div class="mapping-preview">
                                <div class="detected-columns">
                                    <h4>Detected Columns</h4>
                                    <div id="detected-columns-list"></div>
                                </div>

                                <div class="template-fields">
                                    <h4>Template Fields</h4>
                                    <div id="template-fields-list"></div>
                                </div>
                            </div>

                            <div class="auto-mapping-section">
                                <button class="btn" id="auto-map-columns">Auto-Map Similar Names</button>
                                <button class="btn btn-secondary" id="clear-all-mappings">Clear All</button>
                            </div>
                        </div>

                        <!-- Step 4: Test & Preview -->
                        <div class="builder-step" id="step-test" style="display: none;">
                            <h3>Step 4: Test Template</h3>

                            <div class="test-results">
                                <div class="test-summary"></div>
                                <div class="preview-table-container">
                                    <h4>Preview (First 5 Records)</h4>
                                    <div id="preview-table"></div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="template-builder-footer">
                        <div class="step-navigation">
                            <button class="btn btn-secondary" id="prev-step" disabled>Previous</button>
                            <button class="btn" id="next-step">Next</button>
                            <button class="btn" id="save-template" style="display: none;">Save Template</button>
                        </div>
                        <div class="step-indicator">
                            <span class="step-dot active" data-step="1">1</span>
                            <span class="step-dot" data-step="2">2</span>
                            <span class="step-dot" data-step="3">3</span>
                            <span class="step-dot" data-step="4">4</span>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    /**
     * Attach event listeners to interface elements
     */
    attachEventListeners() {
        // Close button
        const closeBtn = document.getElementById('close-template-builder');
        closeBtn.addEventListener('click', () => this.hide());

        // File upload
        const uploadZone = document.getElementById('sample-file-upload');
        const fileInput = document.getElementById('sample-file-input');

        uploadZone.addEventListener('click', () => fileInput.click());
        uploadZone.addEventListener('dragover', this.handleDragOver.bind(this));
        uploadZone.addEventListener('drop', this.handleFileDrop.bind(this));
        fileInput.addEventListener('change', this.handleFileSelect.bind(this));

        // Configuration changes
        document.getElementById('data-start-method').addEventListener('change', this.handleDataStartMethodChange.bind(this));
        document.getElementById('template-name-input').addEventListener('input', this.handleConfigChange.bind(this));
        document.getElementById('filename-pattern-input').addEventListener('input', this.handleConfigChange.bind(this));

        // Navigation
        document.getElementById('next-step').addEventListener('click', this.nextStep.bind(this));
        document.getElementById('prev-step').addEventListener('click', this.prevStep.bind(this));
        document.getElementById('save-template').addEventListener('click', this.saveTemplate.bind(this));

        // Row filters
        document.getElementById('add-row-filter').addEventListener('click', this.addRowFilter.bind(this));

        // Column mapping
        document.getElementById('auto-map-columns').addEventListener('click', this.autoMapColumns.bind(this));
        document.getElementById('clear-all-mappings').addEventListener('click', this.clearAllMappings.bind(this));

        this.currentStep = 1;
    }

    /**
     * Shows the template builder interface
     */
    show() {
        const overlay = this.container.querySelector('.template-builder-overlay');
        if (overlay) {
            overlay.style.display = 'flex';
            this.isVisible = true;
        } else {
            console.error('Template builder overlay not found - interface may not be initialized');
        }
    }

    /**
     * Hides the template builder interface
     */
    hide() {
        const overlay = this.container.querySelector('.template-builder-overlay');
        overlay.style.display = 'none';
        this.isVisible = false;
        this.reset();
    }

    /**
     * Resets the builder to initial state
     */
    reset() {
        this.currentFile = null;
        this.analysisResults = null;
        this.currentConfig = this.getDefaultConfig();
        this.previewData = null;
        this.goToStep(1);
    }

    /**
     * Handle file drag over
     */
    handleDragOver(e) {
        e.preventDefault();
        e.currentTarget.classList.add('dragover');
    }

    /**
     * Handle file drop
     */
    async handleFileDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            await this.processFile(files[0]);
        }
    }

    /**
     * Handle file selection
     */
    async handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            await this.processFile(file);
        }
    }

    /**
     * Check if the interface is ready for processing
     */
    isInterfaceReady() {
        const overlay = this.container.querySelector('.template-builder-overlay');
        if (!overlay || overlay.style.display === 'none') return false;

        const container = document.getElementById('analysis-results');
        if (!container) return false;

        const summaryDiv = container.querySelector('.analysis-summary');
        const visualDiv = container.querySelector('.pattern-visualization');

        return !!(summaryDiv && visualDiv);
    }

    /**
     * Process uploaded file and analyze structure
     */
    async processFile(file) {
        try {
            console.log('Processing file:', file.name);
            this.currentFile = file;

            // Wait for interface to be ready if needed
            let retries = 0;
            while (!this.isInterfaceReady() && retries < 10) {
                await new Promise(resolve => setTimeout(resolve, 100));
                retries++;
            }

            if (!this.isInterfaceReady()) {
                console.error('Interface not ready after waiting');
                return;
            }

            // Show loading state
            const resultsContainer = document.getElementById('analysis-results');
            resultsContainer.style.display = 'block';

            // Show loading in summary div instead of overwriting entire container
            const summaryDiv = resultsContainer.querySelector('.analysis-summary');
            if (summaryDiv) {
                summaryDiv.innerHTML = '<p>Analyzing file structure...</p>';
            }

            // Analyze file structure
            this.analysisResults = await DataPatternAnalyzer.analyzeFile(file);
            console.log('Analysis results:', this.analysisResults);

            // Display analysis results
            this.displayAnalysisResults();

            // Auto-populate some config based on filename
            this.suggestConfigFromFilename(file.name);

            // Enable next step
            document.getElementById('next-step').disabled = false;

        } catch (error) {
            console.error('Error analyzing file:', error);
            const resultsContainer = document.getElementById('analysis-results');
            resultsContainer.innerHTML = `<p style="color: #f44336;">Error analyzing file: ${error.message}</p>`;
        }
    }

    /**
     * Display analysis results in the interface
     */
    displayAnalysisResults() {
        const container = document.getElementById('analysis-results');
        if (!container) {
            console.error('Analysis results container not found');
            return;
        }

        // Make sure the container is visible
        container.style.display = 'block';

        const summaryDiv = container.querySelector('.analysis-summary');
        const visualDiv = container.querySelector('.pattern-visualization');

        if (!summaryDiv || !visualDiv) {
            console.error('Analysis results child containers not found', {
                summaryDiv: !!summaryDiv,
                visualDiv: !!visualDiv,
                containerHTML: container.innerHTML.substring(0, 200)
            });
            return;
        }

        // Summary
        summaryDiv.innerHTML = `
            <div class="analysis-item">
                <strong>Filled Row Pattern:</strong> ${this.analysisResults.filledRowPattern}
            </div>
            <div class="analysis-item">
                <strong>Detected Data Section:</strong> Rows ${this.analysisResults.suggestedDataStart + 1} - ${this.analysisResults.suggestedDataEnd + 1}
            </div>
            <div class="analysis-item">
                <strong>Confidence:</strong> ${Math.round(this.analysisResults.confidence * 100)}%
            </div>
        `;

        // Pattern visualization
        const pattern = this.analysisResults.filledRowPattern;
        let visualization = '<div class="pattern-viz">';
        for (let i = 0; i < pattern.length; i++) {
            const char = pattern[i];
            const isDataRow = i >= this.analysisResults.suggestedDataStart && i <= this.analysisResults.suggestedDataEnd;
            const className = char === 'F' ? (isDataRow ? 'filled-data' : 'filled') : 'empty';
            visualization += `<span class="pattern-char ${className}" title="Row ${i + 1}">${char}</span>`;
        }
        visualization += '</div>';

        visualDiv.innerHTML = `
            <h5>Row Pattern Visualization</h5>
            ${visualization}
            <div class="pattern-legend">
                <span class="legend-item"><span class="pattern-char filled-data">F</span> Data rows</span>
                <span class="legend-item"><span class="pattern-char filled">F</span> Other filled rows</span>
                <span class="legend-item"><span class="pattern-char empty">E</span> Empty rows</span>
            </div>
        `;
    }

    /**
     * Suggest configuration based on filename
     */
    suggestConfigFromFilename(filename) {
        // Auto-populate template name
        const baseName = filename.replace(/\.(xlsx?|csv)$/i, '');
        document.getElementById('template-name-input').value = `${baseName} Template`;

        // Suggest filename pattern
        const escaped = filename.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const pattern = escaped.replace(/\d+/g, '\\d+').replace(/[A-Za-z]+/g, '[A-Za-z]+');
        document.getElementById('filename-pattern-input').value = `^${pattern}$`;

        this.currentConfig.templateName = `${baseName} Template`;
        this.currentConfig.filenamePattern = `^${pattern}$`;
    }

    /**
     * Handle data start method change
     */
    handleDataStartMethodChange(e) {
        const method = e.target.value;
        const skipConfig = document.getElementById('skip-rows-config');

        if (method === 'skip-rows') {
            skipConfig.style.display = 'block';
        } else {
            skipConfig.style.display = 'none';
        }

        this.currentConfig.dataStartMethod = method;
    }

    /**
     * Handle configuration changes
     */
    handleConfigChange() {
        this.currentConfig.templateName = document.getElementById('template-name-input').value;
        this.currentConfig.matchingKeyword = document.getElementById('template-keyword-input').value;
        this.currentConfig.filenamePattern = document.getElementById('filename-pattern-input').value;
        this.currentConfig.skipRows = parseInt(document.getElementById('skip-rows-input').value) || 0;
    }

    /**
     * Navigate to next step
     */
    nextStep() {
        if (this.currentStep < 4) {
            this.goToStep(this.currentStep + 1);
        }
    }

    /**
     * Navigate to previous step
     */
    prevStep() {
        if (this.currentStep > 1) {
            this.goToStep(this.currentStep - 1);
        }
    }

    /**
     * Go to specific step
     */
    goToStep(step) {
        // Hide all steps
        document.querySelectorAll('.builder-step').forEach(el => {
            el.style.display = 'none';
        });

        // Show target step
        document.getElementById(`step-${this.getStepName(step)}`).style.display = 'block';

        // Update navigation
        document.getElementById('prev-step').disabled = step === 1;

        if (step === 4) {
            document.getElementById('next-step').style.display = 'none';
            document.getElementById('save-template').style.display = 'block';
        } else {
            document.getElementById('next-step').style.display = 'block';
            document.getElementById('save-template').style.display = 'none';
        }

        // Update step indicators
        document.querySelectorAll('.step-dot').forEach((dot, index) => {
            dot.classList.toggle('active', index < step);
        });

        this.currentStep = step;

        // Initialize step-specific content
        if (step === 3 && this.analysisResults) {
            this.initializeColumnMapping();
        } else if (step === 4) {
            this.generatePreview();
        }
    }

    /**
     * Get step name from number
     */
    getStepName(step) {
        const names = ['', 'upload', 'config', 'mapping', 'test'];
        return names[step] || 'upload';
    }

    /**
     * Initialize column mapping interface
     */
    initializeColumnMapping() {
        if (!this.analysisResults?.headerAnalysis?.suggestedColumns) {
            return;
        }

        const detectedList = document.getElementById('detected-columns-list');
        const templateList = document.getElementById('template-fields-list');

        // Show detected columns
        detectedList.innerHTML = this.analysisResults.headerAnalysis.suggestedColumns
            .map(col => `<div class="column-item" data-column="${col}">${col}</div>`)
            .join('');

        // Show template fields (from current borderellenTemplate)
        if (typeof borderellenTemplate !== 'undefined') {
            templateList.innerHTML = borderellenTemplate.columns
                .map(col => `
                    <div class="template-field" data-field="${col.name}">
                        <span>${col.name}</span>
                        <select class="mapping-select" data-target="${col.name}">
                            <option value="">Select source column...</option>
                            ${this.analysisResults.headerAnalysis.suggestedColumns
                                .map(srcCol => `<option value="${srcCol}">${srcCol}</option>`)
                                .join('')}
                        </select>
                    </div>
                `)
                .join('');
        }
    }

    /**
     * Auto-map columns based on name similarity
     */
    autoMapColumns() {
        // Simple auto-mapping logic
        document.querySelectorAll('.mapping-select').forEach(select => {
            const targetField = select.dataset.target.toLowerCase();
            const options = Array.from(select.options);

            const match = options.find(option => {
                const sourceField = option.value.toLowerCase();
                return sourceField.includes(targetField) || targetField.includes(sourceField);
            });

            if (match) {
                select.value = match.value;
                this.currentConfig.columnMapping[select.dataset.target] = {
                    type: 'column',
                    source: match.value
                };
            }
        });
    }

    /**
     * Clear all column mappings
     */
    clearAllMappings() {
        document.querySelectorAll('.mapping-select').forEach(select => {
            select.value = '';
        });
        this.currentConfig.columnMapping = {};
    }

    /**
     * Add row filter
     */
    addRowFilter() {
        // Simple implementation - could be expanded
        const filter = {
            type: 'require-any',
            fields: ['Column1'],
            description: 'Skip empty rows'
        };

        this.currentConfig.rowFilters.push(filter);
        this.updateRowFiltersDisplay();
    }

    /**
     * Update row filters display
     */
    updateRowFiltersDisplay() {
        const container = document.getElementById('row-filters-container');

        if (this.currentConfig.rowFilters.length === 0) {
            container.innerHTML = '<p class="no-filters-message">No filters configured. Add filters to skip unwanted rows.</p>';
        } else {
            container.innerHTML = this.currentConfig.rowFilters
                .map((filter, index) => `
                    <div class="filter-item">
                        <span>${filter.description}</span>
                        <button class="btn-small" onclick="this.removeRowFilter(${index})">Remove</button>
                    </div>
                `)
                .join('');
        }
    }

    /**
     * Generate preview of template processing
     */
    async generatePreview() {
        if (!this.currentFile || !this.analysisResults) {
            return;
        }

        try {
            // Create temporary template
            const template = CustomBrokerTemplateManager.createTemplateFromAnalysis(
                this.analysisResults,
                this.currentConfig
            );

            // Test the template
            const testResults = await CustomBrokerTemplateManager.testTemplate(template, this.currentFile);

            this.displayTestResults(testResults);
        } catch (error) {
            console.error('Error generating preview:', error);
            document.querySelector('.test-summary').innerHTML =
                `<p style="color: #f44336;">Error generating preview: ${error.message}</p>`;
        }
    }

    /**
     * Display test results
     */
    displayTestResults(results) {
        const summaryDiv = document.querySelector('.test-summary');
        const previewDiv = document.getElementById('preview-table');

        if (results.success) {
            summaryDiv.innerHTML = `
                <div class="test-success">
                    <strong>✓ Template test successful!</strong>
                    <p>Processed ${results.recordCount} records from sample file.</p>
                </div>
            `;

            // Show preview table
            if (results.sampleData.length > 0) {
                const headers = Object.keys(results.sampleData[0]);
                let tableHTML = '<table class="preview-table"><thead><tr>';
                headers.forEach(header => {
                    tableHTML += `<th>${header}</th>`;
                });
                tableHTML += '</tr></thead><tbody>';

                results.sampleData.forEach(row => {
                    tableHTML += '<tr>';
                    headers.forEach(header => {
                        tableHTML += `<td>${row[header] || ''}</td>`;
                    });
                    tableHTML += '</tr>';
                });
                tableHTML += '</tbody></table>';

                previewDiv.innerHTML = tableHTML;
            }

            this.previewData = results;
        } else {
            summaryDiv.innerHTML = `
                <div class="test-error">
                    <strong>✗ Template test failed</strong>
                    <p>${results.error}</p>
                </div>
            `;
            previewDiv.innerHTML = '';
        }
    }

    /**
     * Save the configured template
     */
    async saveTemplate() {
        try {
            // Collect final configuration
            this.handleConfigChange();

            // Collect column mappings
            document.querySelectorAll('.mapping-select').forEach(select => {
                if (select.value) {
                    this.currentConfig.columnMapping[select.dataset.target] = {
                        type: 'column',
                        source: select.value
                    };
                }
            });

            // Create final template
            const template = CustomBrokerTemplateManager.createTemplateFromAnalysis(
                this.analysisResults,
                {
                    ...this.currentConfig,
                    sampleFile: this.currentFile.name,
                    testRecordCount: this.previewData?.recordCount || 0
                }
            );

            // Save template
            await CustomBrokerTemplateManager.saveTemplate(template);

            alert('Template saved successfully! JSON backup has been downloaded.');
            this.hide();

            // Refresh file selector in mapping tab
            if (typeof updateMappingFileSelector === 'function') {
                updateMappingFileSelector();
            }

        } catch (error) {
            console.error('Error saving template:', error);
            alert(`Error saving template: ${error.message}`);
        }
    }
}