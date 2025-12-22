/**
 * Borderellen Converter - Worksheet Cleaner
 * Preprocesses Excel worksheets by removing empty rows/columns before parsing
 * Used exclusively for Generic Parser (not for built-in brokers)
 */

class WorksheetCleaner {
    /**
     * Apply cleanup rules to raw worksheet BEFORE parsing
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} preprocessRules - Cleanup configuration
     * @returns {Object} Cleaned worksheet
     */
    static cleanWorksheet(worksheet, preprocessRules) {
        if (!preprocessRules) {
            console.log('No preprocess rules provided, skipping worksheet cleanup');
            return worksheet;
        }

        console.log('Starting worksheet cleanup with rules:', preprocessRules);

        const range = XLSX.utils.decode_range(worksheet['!ref']);
        console.log(`Original worksheet range: ${worksheet['!ref']} (${range.e.r + 1} rows, ${range.e.c + 1} cols)`);

        const skipInfo = {
            skipRows: [],
            skipColumns: []
        };

        // 1. Find empty rows to remove
        if (preprocessRules.removeEmptyRows?.enabled) {
            skipInfo.skipRows = this.findEmptyRows(worksheet, range, preprocessRules.removeEmptyRows);
            console.log(`Found ${skipInfo.skipRows.length} empty rows to remove:`, skipInfo.skipRows);
        }

        // 2. Find empty columns to remove
        if (preprocessRules.removeEmptyColumns?.enabled) {
            skipInfo.skipColumns = this.findEmptyColumns(worksheet, range, preprocessRules.removeEmptyColumns);
            console.log(`Found ${skipInfo.skipColumns.length} empty columns to remove:`, skipInfo.skipColumns);
        }

        // 3. Find noise rows (headers/footers with specific patterns)
        if (preprocessRules.removeHeaderNoise?.enabled) {
            const noiseRows = this.findNoiseRows(worksheet, range, preprocessRules.removeHeaderNoise);
            console.log(`Found ${noiseRows.length} noise rows to remove:`, noiseRows);
            skipInfo.skipRows = [...new Set([...skipInfo.skipRows, ...noiseRows])];
        }

        // 4. Rebuild worksheet without skipped rows/columns
        if (skipInfo.skipRows.length === 0 && skipInfo.skipColumns.length === 0) {
            console.log('No rows or columns to remove, returning original worksheet');
            return worksheet;
        }

        const cleanedWorksheet = this.rebuildWorksheet(worksheet, range, skipInfo);
        console.log(`Cleaned worksheet range: ${cleanedWorksheet['!ref']}`);

        return cleanedWorksheet;
    }

    /**
     * Find empty rows based on rules
     * @param {Object} worksheet - XLSX worksheet
     * @param {Object} range - Worksheet range
     * @param {Object} rules - removeEmptyRows configuration
     * @returns {Array} Array of row indices to skip
     */
    static findEmptyRows(worksheet, range, rules) {
        const emptyRows = [];

        for (let row = range.s.r; row <= range.e.r; row++) {
            if (rules.mode === 'all-empty') {
                // Check if ALL cells in row are empty
                if (this.isRowEmpty(worksheet, row, range.s.c, range.e.c)) {
                    emptyRows.push(row);
                }

            } else if (rules.mode === 'specific-columns-empty') {
                // Check if specific columns are ALL empty
                const checkColumns = rules.columns || [];
                if (checkColumns.length === 0) {
                    console.warn('specific-columns-empty mode requires columns array');
                    continue;
                }

                let isEmpty = true;
                for (let colLetter of checkColumns) {
                    const col = XLSX.utils.decode_col(colLetter);
                    const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                    const cell = worksheet[cellAddress];
                    if (this.isCellFilled(cell)) {
                        isEmpty = false;
                        break;
                    }
                }
                if (isEmpty) {
                    emptyRows.push(row);
                }
            }
        }

        return emptyRows;
    }

    /**
     * Check if a row is completely empty
     * @param {Object} worksheet - XLSX worksheet
     * @param {Number} row - Row index
     * @param {Number} startCol - Start column index
     * @param {Number} endCol - End column index
     * @returns {Boolean}
     */
    static isRowEmpty(worksheet, row, startCol, endCol) {
        for (let col = startCol; col <= endCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            const cell = worksheet[cellAddress];
            if (this.isCellFilled(cell)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Check if a cell is filled (not empty/null/undefined)
     * @param {Object} cell - XLSX cell object
     * @returns {Boolean}
     */
    static isCellFilled(cell) {
        if (!cell) return false;
        if (cell.v === null || cell.v === undefined) return false;
        if (typeof cell.v === 'string' && cell.v.trim() === '') return false;
        return true;
    }

    /**
     * Find empty columns based on rules
     * @param {Object} worksheet - XLSX worksheet
     * @param {Object} range - Worksheet range
     * @param {Object} rules - removeEmptyColumns configuration
     * @returns {Array} Array of column indices to skip
     */
    static findEmptyColumns(worksheet, range, rules) {
        const emptyColumns = [];
        const threshold = rules.threshold || 1.0; // Default: 100% empty

        for (let col = range.s.c; col <= range.e.c; col++) {
            const totalRows = range.e.r - range.s.r + 1;
            let emptyCount = 0;

            for (let row = range.s.r; row <= range.e.r; row++) {
                const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                const cell = worksheet[cellAddress];
                if (!this.isCellFilled(cell)) {
                    emptyCount++;
                }
            }

            const emptyPercentage = emptyCount / totalRows;
            if (emptyPercentage >= threshold) {
                emptyColumns.push(col);
            }
        }

        return emptyColumns;
    }

    /**
     * Find noise rows containing specific patterns
     * @param {Object} worksheet - XLSX worksheet
     * @param {Object} range - Worksheet range
     * @param {Object} rules - removeHeaderNoise configuration
     * @returns {Array} Array of row indices to skip
     */
    static findNoiseRows(worksheet, range, rules) {
        const noiseRows = [];
        const patterns = rules.patterns || [];

        if (patterns.length === 0) {
            return noiseRows;
        }

        for (let row = range.s.r; row <= range.e.r; row++) {
            // Check all cells in row for noise patterns
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                    const cellValue = cell.v.toString().toLowerCase();

                    // Check if cell contains any noise pattern
                    for (let pattern of patterns) {
                        const patternLower = pattern.toLowerCase();
                        if (cellValue.includes(patternLower)) {
                            noiseRows.push(row);
                            break; // Row found, no need to check more cells
                        }
                    }

                    if (noiseRows.includes(row)) break; // Move to next row
                }
            }
        }

        return [...new Set(noiseRows)]; // Remove duplicates
    }

    /**
     * Rebuild worksheet without skipped rows and columns
     * @param {Object} originalWorksheet - Original XLSX worksheet
     * @param {Object} range - Original range
     * @param {Object} skipInfo - {skipRows: [], skipColumns: []}
     * @returns {Object} New cleaned worksheet
     */
    static rebuildWorksheet(originalWorksheet, range, skipInfo) {
        const newWorksheet = {};
        let newRowIndex = 0;

        // Build column mapping (old col index -> new col index)
        const columnMapping = new Map();
        let newColIndex = 0;
        for (let col = range.s.c; col <= range.e.c; col++) {
            if (!skipInfo.skipColumns.includes(col)) {
                columnMapping.set(col, newColIndex);
                newColIndex++;
            }
        }

        // Copy cells to new worksheet (skip removed rows/columns)
        for (let row = range.s.r; row <= range.e.r; row++) {
            if (skipInfo.skipRows.includes(row)) {
                continue; // Skip this row
            }

            for (let col = range.s.c; col <= range.e.c; col++) {
                if (skipInfo.skipColumns.includes(col)) {
                    continue; // Skip this column
                }

                const oldAddress = XLSX.utils.encode_cell({r: row, c: col});
                const cell = originalWorksheet[oldAddress];

                if (cell) {
                    const newAddress = XLSX.utils.encode_cell({
                        r: newRowIndex,
                        c: columnMapping.get(col)
                    });

                    // Deep copy cell to avoid reference issues
                    newWorksheet[newAddress] = {
                        v: cell.v,
                        t: cell.t,
                        w: cell.w,
                        f: cell.f,
                        z: cell.z,
                        s: cell.s
                    };
                }
            }

            newRowIndex++;
        }

        // Update worksheet range
        const newEndRow = newRowIndex - 1;
        const newEndCol = columnMapping.size - 1;

        if (newEndRow < 0 || newEndCol < 0) {
            console.error('Worksheet cleanup resulted in empty worksheet!');
            // Return original worksheet to prevent errors
            return originalWorksheet;
        }

        newWorksheet['!ref'] = XLSX.utils.encode_range({
            s: {r: 0, c: 0},
            e: {r: newEndRow, c: newEndCol}
        });

        // Copy worksheet properties if they exist
        if (originalWorksheet['!cols']) {
            // Filter columns based on mapping
            newWorksheet['!cols'] = originalWorksheet['!cols']
                .filter((col, idx) => !skipInfo.skipColumns.includes(idx));
        }

        console.log(`Rebuilt worksheet: ${range.e.r + 1} → ${newEndRow + 1} rows, ${range.e.c + 1} → ${newEndCol + 1} cols`);

        return newWorksheet;
    }

    /**
     * Get statistics about what was cleaned
     * @param {Object} originalWorksheet - Original worksheet
     * @param {Object} cleanedWorksheet - Cleaned worksheet
     * @returns {Object} Cleanup statistics
     */
    static getCleanupStats(originalWorksheet, cleanedWorksheet) {
        const originalRange = XLSX.utils.decode_range(originalWorksheet['!ref']);
        const cleanedRange = XLSX.utils.decode_range(cleanedWorksheet['!ref']);

        return {
            originalRows: originalRange.e.r + 1,
            originalCols: originalRange.e.c + 1,
            cleanedRows: cleanedRange.e.r + 1,
            cleanedCols: cleanedRange.e.c + 1,
            rowsRemoved: (originalRange.e.r + 1) - (cleanedRange.e.r + 1),
            colsRemoved: (originalRange.e.c + 1) - (cleanedRange.e.c + 1)
        };
    }
}

// Make available globally
window.WorksheetCleaner = WorksheetCleaner;
