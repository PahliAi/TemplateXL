/**
 * Borderellen Converter - Generic Broker Parser
 * Configurable parsing engine that can handle any Excel format based on JSON configuration
 */

class GenericBrokerParser {
    constructor(config) {
        this.config = this.validateConfig(config);
        console.log('Generic parser initialized with config:', this.config);
    }

    /**
     * Extracts the calculated value from a cell, handling formulas properly
     * @param {Object} cell - XLSX cell object
     * @returns {*} The calculated value or null
     */
    getCellValue(cell) {
        if (!cell) return null;

        // Debug logging for formulas only
        if (cell.f) {
            console.log('Formula cell found:', {
                formula: cell.f,
                value: cell.v,
                formatted: cell.w,
                type: cell.t
            });
        }

        // Priority 1: If cell has formula, try to get the calculated value
        if (cell.f) {
            console.log('Formula cell found:', {
                formula: cell.f,
                value: cell.v,
                formatted: cell.w,
                type: cell.t
            });

            if (cell.v !== undefined && cell.v !== null) {
                console.log('Using cell.v for formula:', cell.v);
                return cell.v;
            }

            if (cell.w !== undefined && cell.w !== null) {
                console.log('Using cell.w for formula:', cell.w);
                const numValue = parseFloat(cell.w);
                if (!isNaN(numValue)) {
                    return numValue;
                }
                return cell.w;
            }

            console.warn('Formula cell has no calculated value:', cell);
            return null;
        }

        // Priority 2: Regular cell value
        return cell.v !== null ? cell.v : null;
    }

    /**
     * Validates and sets default values for parser configuration
     * @param {Object} config - Parser configuration
     * @returns {Object} Validated configuration
     */
    validateConfig(config) {
        const defaultConfig = {
            id: config.id || `generic-${Date.now()}`,
            name: config.name || 'Generic Parser',
            dataStartMethod: 'auto-detect',
            skipRows: 0,
            skipColumns: 0,
            endColumn: null,
            headerRow: null,
            rowProcessing: {
                type: 'single',
                multiRowConfig: null
            },
            rowFilters: [],
            dataValidation: [],
            columnMapping: {},
            filenamePattern: null
        };

        return { ...defaultConfig, ...config };
    }

    /**
     * Parses a workbook using the configured rules
     * @param {Object} workbook - XLSX workbook object
     * @param {String} filename - Original filename for pattern extraction
     * @returns {Array} Parsed data array
     */
    async parse(workbook, filename) {
        console.log(`Generic parser processing: ${filename}`);

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        // Step 1: Determine data start location (row and column boundaries)
        const dataStartInfo = await this.determineDataStart(worksheet);
        console.log(`Data start info:`, dataStartInfo);

        // Step 2: Get header information using detected boundaries
        const headers = this.extractHeaders(worksheet, dataStartInfo);
        console.log(`Headers:`, headers);

        // Step 3: Extract raw data based on processing type
        let rawData;
        const processingType = this.config.rowProcessing?.type || 'single';
        const isMultiRowHeaders = (this.config.headerRows || 1) > 1;

        // Use multi-row processing if either explicitly set OR if multi-row headers detected
        if (processingType === 'multi-row' || isMultiRowHeaders) {
            console.log(`Using multi-row data processing: ${this.config.headerRows || 1} rows per record`);
            rawData = this.extractSingleRowData(worksheet, dataStartInfo, headers, filename);
        } else {
            rawData = this.extractSingleRowData(worksheet, dataStartInfo, headers, filename);
        }

        console.log(`Extracted ${rawData.length} raw records`);

        // Step 4: Apply row filters
        const filteredData = this.applyRowFilters(rawData);
        console.log(`After filtering: ${filteredData.length} records`);

        // Step 5: Apply column mapping (skip if using string-based template mappings)
        let mappedData;
        if (this.hasStringBasedMappings()) {
            // String-based mappings (CALC:, FIXED:) should be handled by external applyMappingToData
            console.log('String-based mappings detected, skipping internal mapping (will be handled externally)');
            mappedData = filteredData;
        } else {
            // Structured mappings can be handled internally
            mappedData = this.applyColumnMapping(filteredData, filename);
        }
        console.log(`Final mapped data: ${mappedData.length} records`);

        // Step 6: Apply data validation
        const validatedData = this.applyDataValidation(mappedData);
        console.log(`After validation: ${validatedData.length} records`);

        return validatedData;
    }

    /**
     * Determines where data starts based on configuration
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {Object} Data start information with row and column ranges
     */
    async determineDataStart(worksheet) {
        switch (this.config.dataStartMethod) {
            case 'skip-rows':
                return {
                    dataStartRow: this.config.skipRows || 0,
                    headerRow: this.config.headerRow,
                    startColumn: this.config.skipColumns || 0,
                    endColumn: this.config.endColumn || null,
                    method: 'manual'
                };

            case 'auto-detect':
                const analysis = DataPatternAnalyzer.analyzeSheet(worksheet);
                console.log('Auto-detect analysis:', analysis);

                // Apply auto-detected skip rules
                const startColumn = analysis.dataSection?.startColumnIndex || 0;
                const endColumn = analysis.dataSection?.endColumnIndex || null;
                const headerRow = analysis.dataSection?.headerRowIndex;
                const dataStartRow = analysis.dataSection?.dataStartIndex || analysis.suggestedDataStart;

                console.log(`Auto-detected: skip ${dataStartRow} rows, skip ${startColumn} columns, data range: ${startColumn}-${endColumn}`);

                return {
                    dataStartRow: dataStartRow,
                    headerRow: headerRow,
                    startColumn: startColumn,
                    endColumn: endColumn,
                    method: 'auto-detect',
                    startCell: analysis.dataSection?.startCell
                };

            case 'find-pattern':
                const patternRow = this.findPatternStart(worksheet);
                return {
                    dataStartRow: patternRow,
                    headerRow: this.config.headerRow,
                    startColumn: this.config.skipColumns || 0,
                    endColumn: this.config.endColumn || null,
                    method: 'find-pattern'
                };

            case 'find-date':
                const dateRow = this.findDateStart(worksheet);
                return {
                    dataStartRow: dateRow,
                    headerRow: this.config.headerRow,
                    startColumn: this.config.skipColumns || 0,
                    endColumn: this.config.endColumn || null,
                    method: 'find-date'
                };

            default:
                return {
                    dataStartRow: 0,
                    headerRow: this.config.headerRow,
                    startColumn: 0,
                    endColumn: null,
                    method: 'default'
                };
        }
    }

    /**
     * Finds data start by searching for a specific pattern
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {Number} Row index where pattern is found
     */
    findPatternStart(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const pattern = this.config.patternConfig?.pattern;

        if (!pattern) return 0;

        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 0 });
            const cell = worksheet[cellAddress];

            if (cell && cell.v && new RegExp(pattern).test(cell.v.toString())) {
                return row;
            }
        }

        return 0;
    }

    /**
     * Finds data start by looking for first date in specified column
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {Number} Row index where first date is found
     */
    findDateStart(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const dateColumn = this.config.dateConfig?.column || 0;

        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: dateColumn });
            const cell = worksheet[cellAddress];

            if (cell && cell.v && DataPatternAnalyzer.isLikelyDate(cell.v)) {
                return row;
            }
        }

        return 0;
    }

    /**
     * Extracts headers from the worksheet using detected column boundaries
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} dataStartInfo - Data start information with column ranges
     * @returns {Array} Array of header names
     */
    extractHeaders(worksheet, dataStartInfo) {
        const headerRow = dataStartInfo.headerRow !== null ?
            dataStartInfo.headerRow :
            dataStartInfo.dataStartRow - 1;

        if (headerRow < 0) {
            // No headers, use column positions within detected range
            const startCol = dataStartInfo.startColumn || 0;
            const endCol = dataStartInfo.endColumn || XLSX.utils.decode_range(worksheet['!ref']).e.c;
            const columnCount = endCol - startCol + 1;
            return Array.from({ length: columnCount }, (_, i) => `Column${startCol + i + 1}`);
        }

        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const headers = [];

        // Use detected column boundaries instead of full range
        const startCol = dataStartInfo.startColumn || range.s.c;
        const endCol = dataStartInfo.endColumn || range.e.c;

        // Check for multi-row headers
        const headerRows = this.config.headerRows || 1;
        const isMultiRowHeader = headerRows > 1;

        console.log(`Extracting headers from ${headerRows} row(s) starting at row ${headerRow}, columns ${startCol}-${endCol}`);

        if (isMultiRowHeader) {
            // Multi-row header: Create separate columns for each row
            for (let row = 0; row < headerRows; row++) {
                for (let col = startCol; col <= endCol; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: headerRow + row, c: col });
                    const cell = worksheet[cellAddress];
                    let headerText = '';

                    if (cell && cell.v) {
                        // Clean header text: remove newlines and extra whitespace
                        headerText = cell.v.toString().replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim();
                    } else {
                        headerText = `Row${row + 1}_Col${col + 1}`;
                    }

                    headers.push(headerText);
                }
            }
        } else {
            // Single row header (existing behavior)
            for (let col = startCol; col <= endCol; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
                const cell = worksheet[cellAddress];
                // Clean header text: remove newlines and extra whitespace
                const headerText = cell && cell.v ?
                    cell.v.toString().replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim() :
                    `Column${col + 1}`;
                headers.push(headerText);
            }
        }

        console.log(`Extracted headers:`, headers);
        return headers;
    }

    /**
     * Extracts data using single-row processing within detected table boundaries
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} dataStartInfo - Data start information with column ranges
     * @param {Array} headers - Column headers
     * @returns {Array} Array of row objects
     */
    extractSingleRowData(worksheet, dataStartInfo, headers, filename) {
        const data = [];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        // Use detected boundaries
        const headerRows = this.config.headerRows || 1;
        const baseDataStart = dataStartInfo.dataStartRow;
        const startCol = dataStartInfo.startColumn || range.s.c;
        const endCol = dataStartInfo.endColumn || range.e.c;
        const isMultiRowData = headerRows > 1;

        console.log(`Extracting data from rows ${baseDataStart}-${range.e.r}, columns ${startCol}-${endCol}, multi-row processing: ${isMultiRowData ? headerRows + ' rows per record' : 'single row per record'}`);

        if (isMultiRowData) {
            // Multi-row data processing: N rows = 1 record
            for (let row = baseDataStart; row <= range.e.r; row += headerRows) {
                // Check for footer keyword in any row of this multi-row record
                let foundFooterKeyword = false;
                if (this.config.footerRowKeyword) {
                    for (let rowOffset = 0; rowOffset < headerRows && !foundFooterKeyword; rowOffset++) {
                        const currentRow = row + rowOffset;
                        if (currentRow > range.e.r) break;

                        for (let col = startCol; col <= endCol; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: col });
                            const cell = worksheet[cellAddress];
                            if (cell && cell.v && cell.v.toString().includes(this.config.footerRowKeyword)) {
                                console.log(`Footer keyword "${this.config.footerRowKeyword}" found at ${cellAddress}, stopping multi-row processing`);
                                foundFooterKeyword = true;
                                break;
                            }
                        }
                    }
                }

                if (foundFooterKeyword) {
                    console.log(`Stopped processing at row ${row} due to footer keyword`);
                    break;
                }

                const recordData = {};
                let hasData = false;
                let headerIndex = 0;

                // Process each row in the multi-row record
                for (let rowOffset = 0; rowOffset < headerRows; rowOffset++) {
                    const currentRow = row + rowOffset;
                    if (currentRow > range.e.r) break;

                    // Process each column in this row
                    for (let col = startCol; col <= endCol; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: col });
                        const cell = worksheet[cellAddress];
                        const value = this.getCellValue(cell);

                        // Map to header using sequential index across all rows
                        const headerName = headers[headerIndex] || `Row${rowOffset + 1}_Col${col + 1}`;
                        recordData[headerName] = value;
                        headerIndex++;

                        if (value !== null && value !== '') {
                            hasData = true;
                        }
                    }
                }

                if (hasData) {
                    recordData.__rowIndex = row;
                    recordData.__multiRowRecord = true;
                    recordData.__recordRows = headerRows;
                    recordData.Filename = filename;
                    data.push(recordData);
                }
            }
        } else {
            // Single row data processing (existing behavior)
            for (let row = baseDataStart; row <= range.e.r; row++) {
                // Check for footer keyword in this row
                if (this.config.footerRowKeyword) {
                    let foundFooterKeyword = false;
                    for (let col = startCol; col <= endCol; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v && cell.v.toString().includes(this.config.footerRowKeyword)) {
                            console.log(`Footer keyword "${this.config.footerRowKeyword}" found at ${cellAddress}, stopping single-row processing`);
                            foundFooterKeyword = true;
                            break;
                        }
                    }

                    if (foundFooterKeyword) {
                        console.log(`Stopped processing at row ${row} due to footer keyword`);
                        break;
                    }
                }

                const rowData = {};
                let hasData = false;

                // Only read columns within detected boundaries
                for (let col = startCol; col <= endCol; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    const cell = worksheet[cellAddress];
                    const value = this.getCellValue(cell);

                    // Map to header using relative position within detected range
                    const headerIndex = col - startCol;
                    const headerName = headers[headerIndex] || `Column${col + 1}`;
                    rowData[headerName] = value;

                    if (value !== null && value !== '') {
                        hasData = true;
                    }
                }

                if (hasData) {
                    rowData.__rowIndex = row;
                    rowData.Filename = filename;
                    data.push(rowData);
                }
            }
        }

        return data;
    }

    /**
     * Extracts data using multi-row processing (like BCI format)
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} dataStartInfo - Data start information with column ranges
     * @param {Object} range - Sheet range
     * @returns {Array} Array of combined row objects
     */
    extractMultiRowData(worksheet, dataStartInfo, range) {
        const data = [];
        const multiConfig = this.config.rowProcessing.multiRowConfig;

        if (!multiConfig) {
            console.warn('Multi-row processing requested but no configuration provided');
            return [];
        }

        const rowCount = multiConfig.rowCount || 2;
        const rowStep = multiConfig.rowStep || rowCount;

        for (let row = dataStartInfo.dataStartRow; row <= range.e.r; row += rowStep) {
            // Check stop condition if configured
            if (multiConfig.stopCondition && this.checkStopCondition(worksheet, row, multiConfig.stopCondition)) {
                break;
            }

            const combinedData = {};
            let hasValidData = false;

            // Process each row in the group
            for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
                const currentRow = row + rowOffset;
                if (currentRow > range.e.r) break;

                const rowConfig = multiConfig.fieldMapping[`row${rowOffset + 1}`];
                if (!rowConfig) continue;

                // Extract pattern-based fields (like BCI policy number parsing)
                if (rowConfig.pattern) {
                    const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: 0 });
                    const cell = worksheet[cellAddress];

                    if (cell && cell.v) {
                        const match = cell.v.toString().match(new RegExp(rowConfig.pattern));
                        if (match && rowConfig.groups) {
                            rowConfig.groups.forEach((groupName, index) => {
                                if (groupName && match[index + 1]) {
                                    combinedData[groupName] = match[index + 1];
                                    hasValidData = true;
                                }
                            });
                        }
                    }
                }

                // Extract column-based fields
                if (rowConfig.columns) {
                    rowConfig.columns.forEach((fieldName, colOffset) => {
                        if (fieldName) {
                            const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: colOffset });
                            const cell = worksheet[cellAddress];
                            const value = this.getCellValue(cell);

                            combinedData[fieldName] = value;
                            if (value !== null && value !== '') {
                                hasValidData = true;
                            }
                        }
                    });
                }
            }

            if (hasValidData) {
                combinedData.__rowIndex = row;
                data.push(combinedData);
            }
        }

        return data;
    }

    /**
     * Checks if stop condition is met for multi-row processing
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Number} row - Current row index
     * @param {Object} stopCondition - Stop condition configuration
     * @returns {Boolean} True if processing should stop
     */
    checkStopCondition(worksheet, row, stopCondition) {
        if (stopCondition.type === 'pattern-mismatch') {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 0 });
            const cell = worksheet[cellAddress];

            if (!cell || !cell.v) return true;

            const pattern = new RegExp(stopCondition.pattern);
            return !pattern.test(cell.v.toString());
        }

        return false;
    }

    /**
     * Applies configured row filters to the data
     * @param {Array} data - Raw data array
     * @returns {Array} Filtered data array
     */
    applyRowFilters(data) {
        let filteredData = [...data];

        this.config.rowFilters.forEach(filter => {
            switch (filter.type) {
                case 'require-any':
                    filteredData = filteredData.filter(row => {
                        return filter.fields.some(field => {
                            const value = row[field];
                            return value !== null && value !== undefined && value !== '';
                        });
                    });
                    break;

                case 'require-all':
                    filteredData = filteredData.filter(row => {
                        return filter.fields.every(field => {
                            const value = row[field];
                            return value !== null && value !== undefined && value !== '';
                        });
                    });
                    break;

                case 'exclude-values':
                    filteredData = filteredData.filter(row => {
                        const value = row[filter.field];
                        return value === null || value === undefined ||
                               !filter.values.includes(value.toString());
                    });
                    break;

                case 'pattern-match':
                    filteredData = filteredData.filter(row => {
                        const value = row[filter.field];
                        if (!value) return false;
                        return new RegExp(filter.pattern).test(value.toString());
                    });
                    break;
            }
        });

        return filteredData;
    }

    /**
     * Check if column mapping uses string-based format (CALC:, FIXED:) vs structured format
     * @returns {boolean} True if string-based mappings are detected
     */
    hasStringBasedMappings() {
        if (!this.config.columnMapping || typeof this.config.columnMapping !== 'object') {
            return false;
        }

        // Check if any mapping values are strings (indicating CALC:, FIXED:, or direct column names)
        const mappingValues = Object.values(this.config.columnMapping);
        return mappingValues.some(mapping => typeof mapping === 'string');
    }

    /**
     * Applies column mapping to transform source data to target format
     * @param {Array} data - Filtered data array
     * @param {String} filename - Original filename for pattern extraction
     * @returns {Array} Mapped data array
     */
    applyColumnMapping(data, filename) {
        // Safety check for undefined/null/empty columnMapping
        if (!this.config.columnMapping || typeof this.config.columnMapping !== 'object' || Object.keys(this.config.columnMapping).length === 0) {
            console.log('No column mapping configured, returning data as-is');
            return data;
        }

        return data.map(row => {
            const mappedRow = {};

            Object.entries(this.config.columnMapping).forEach(([targetField, mapping]) => {
                switch (mapping.type) {
                    case 'fixed':
                        mappedRow[targetField] = mapping.value;
                        break;

                    case 'column':
                        mappedRow[targetField] = row[mapping.source] || '';
                        break;

                    case 'multi-row-field':
                        mappedRow[targetField] = row[mapping.source] || '';
                        break;

                    case 'filename-extract':
                        const match = filename.match(new RegExp(mapping.pattern));
                        mappedRow[targetField] = match && match[mapping.group] ? match[mapping.group] : '';
                        break;

                    case 'transform':
                        const sourceValue = row[mapping.source] || '';
                        mappedRow[targetField] = this.applyTransform(sourceValue, mapping.transform);
                        break;

                    default:
                        mappedRow[targetField] = '';
                }
            });

            return mappedRow;
        });
    }

    /**
     * Applies data transformations (currency, date formatting, etc.)
     * @param {*} value - Source value
     * @param {String} transformType - Type of transformation
     * @returns {*} Transformed value
     */
    applyTransform(value, transformType) {
        if (!value) return value;

        switch (transformType) {
            case 'currency':
                // Clean currency symbols and convert to number
                const cleanValue = value.toString().replace(/[€$£,\s]/g, '');
                return parseFloat(cleanValue) || 0;

            case 'date':
                return DataPatternAnalyzer.isLikelyDate(value) ? new Date(value) : value;

            case 'uppercase':
                return value.toString().toUpperCase();

            case 'lowercase':
                return value.toString().toLowerCase();

            case 'trim':
                return value.toString().trim();

            default:
                return value;
        }
    }

    /**
     * Applies data validation rules
     * @param {Array} data - Mapped data array
     * @returns {Array} Validated data array
     */
    applyDataValidation(data) {
        const validatedData = [];

        data.forEach((row, index) => {
            let isValid = true;
            const validationErrors = [];

            this.config.dataValidation.forEach(validation => {
                const value = row[validation.field];

                switch (validation.type) {
                    case 'pattern-match':
                        if (value && !new RegExp(validation.pattern).test(value.toString())) {
                            isValid = false;
                            validationErrors.push(`${validation.field}: Does not match pattern ${validation.pattern}`);
                        }
                        break;

                    case 'required':
                        if (!value || value === '') {
                            isValid = false;
                            validationErrors.push(`${validation.field}: Required field is empty`);
                        }
                        break;

                    case 'numeric':
                        if (value && isNaN(Number(value))) {
                            isValid = false;
                            validationErrors.push(`${validation.field}: Must be numeric`);
                        }
                        break;
                }
            });

            if (isValid) {
                validatedData.push(row);
            } else {
                console.warn(`Row ${index + 1} validation failed:`, validationErrors);
            }
        });

        return validatedData;
    }

    /**
     * Static method to parse workbook using pattern analysis
     * @param {Object} workbook - XLSX workbook object
     * @param {Object} patternAnalysis - Pattern analysis results
     * @returns {Array} Parsed data array
     */
    static async parseWithAnalysis(workbook, patternAnalysis) {
        console.log('GenericParser.parseWithAnalysis called with:', patternAnalysis);

        // Create parser configuration from pattern analysis
        const config = {
            id: 'pattern-analysis-parser',
            name: 'Pattern Analysis Parser',
            dataStartMethod: 'skip-rows',
            skipRows: patternAnalysis.dataSection?.dataStartIndex || 0,
            skipColumns: patternAnalysis.dataSection?.startColumnIndex || 0,
            endColumn: patternAnalysis.dataSection?.endColumnIndex || null,
            headerRow: patternAnalysis.suggestedHeaderRow,
            headerRows: patternAnalysis.manualSelection?.headerRows || 1,
            footerRowKeyword: patternAnalysis.manualSelection?.footerKeyword || patternAnalysis.autoFooterKeyword,
            rowProcessing: patternAnalysis.manualSelection?.headerRows > 1 ? {
                type: 'multi-row',
                rowsPerRecord: patternAnalysis.manualSelection.headerRows
            } : {
                type: 'single'
            },
            rowFilters: [],
            dataValidation: [],
            columnMapping: {} // Will be applied externally
        };

        console.log('Created parser config:', config);

        // Create parser instance and parse
        const parser = new GenericBrokerParser(config);
        const filename = patternAnalysis.filename || 'pattern-analysis-file';
        return await parser.parse(workbook, filename);
    }

    /**
     * Static method to create parser from JSON template
     * @param {Object} template - JSON template object
     * @returns {GenericBrokerParser} Parser instance
     */
    static fromTemplate(template) {
        const config = {
            id: template.id,
            name: template.name,
            ...template.parsingConfig,
            columnMapping: template.columnMapping,
            filenamePattern: template.filePattern
        };

        return new GenericBrokerParser(config);
    }
}

// Export the class for global access
window.GenericParser = GenericBrokerParser;