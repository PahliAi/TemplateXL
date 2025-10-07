/**
 * Centralized Excel Cache Manager
 *
 * Single entry point for all Excel file operations to ensure consistency
 * and prevent coordinate mismatches between analysis and execution phases.
 *
 * Features:
 * - WeakMap cache keyed by File objects for automatic garbage collection
 * - Consistent compacted worksheet state across all consumers
 * - Multi-file support with isolated cache entries
 * - Performance optimization through single-read caching
 */
class ExcelCacheManager {
    // WeakMap automatically garbage collects when File objects are released
    static cache = new WeakMap();

    /**
     * Reads and caches an Excel file with consistent processing
     * @param {File} file - Excel file to read
     * @returns {Promise<Object>} XLSX workbook object (compacted)
     */
    static async getWorkbook(file) {
        // Check if already cached
        if (this.cache.has(file)) {
            return this.cache.get(file);
        }

        try {
            const workbook = await this._readExcelFile(file);

            // Cache the processed workbook
            this.cache.set(file, workbook);

            return workbook;
        } catch (error) {
            console.error('ExcelCacheManager: Failed to read Excel file:', error);
            throw error;
        }
    }

    /**
     * Gets a specific worksheet from cached workbook
     * @param {File} file - Excel file
     * @param {string} sheetName - Name of sheet to retrieve (defaults to first sheet)
     * @returns {Promise<Object>} XLSX worksheet object
     */
    static async getSheet(file, sheetName = null) {
        const workbook = await this.getWorkbook(file);

        if (!sheetName) {
            sheetName = workbook.SheetNames[0];
        }

        if (!workbook.Sheets[sheetName]) {
            throw new Error(`Sheet "${sheetName}" not found in workbook`);
        }

        return workbook.Sheets[sheetName];
    }

    /**
     * Gets all sheet names from cached workbook
     * @param {File} file - Excel file
     * @returns {Promise<string[]>} Array of sheet names
     */
    static async getSheetNames(file) {
        const workbook = await this.getWorkbook(file);
        return workbook.SheetNames;
    }

    /**
     * Clears cache entry for specific file
     * @param {File} file - File to remove from cache
     */
    static clearCache(file) {
        this.cache.delete(file);
    }

    /**
     * Gets cache status (for debugging)
     * @param {File} file - File to check
     * @returns {boolean} True if file is cached
     */
    static isCached(file) {
        return this.cache.has(file);
    }

    /**
     * Internal method to read Excel file with consistent processing
     * This replaces the scattered XLSX.read() calls throughout the codebase
     * @param {File} file - Excel file to read
     * @returns {Promise<Object>} XLSX workbook object
     * @private
     */
    static _readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    // Read workbook with consistent options
                    const workbook = XLSX.read(e.target.result, {
                        type: 'binary',
                        cellFormula: false  // Read calculated values instead of formulas
                    });

                    // Apply consistent compaction to all worksheets
                    workbook.SheetNames.forEach(sheetName => {
                        const originalSheet = workbook.Sheets[sheetName];
                        const compactedSheet = this._removeEmptyRowsFromWorksheet(originalSheet);

                        // Log compaction details for debugging
                        if (originalSheet['!ref'] && compactedSheet['!ref']) {
                            const originalRange = XLSX.utils.decode_range(originalSheet['!ref']);
                            const compactedRange = XLSX.utils.decode_range(compactedSheet['!ref']);
                            const originalRows = originalRange.e.r + 1;
                            const compactedRows = compactedRange.e.r + 1;
                            const removedRows = originalRows - compactedRows;

                            if (removedRows > 0) {
                                console.log(`[ExcelCacheManager] Compacted worksheet "${sheetName}": ${originalRows} rows → ${compactedRows} rows (removed ${removedRows} empty rows)`);
                            }
                        }

                        workbook.Sheets[sheetName] = compactedSheet;
                    });

                    resolve(workbook);
                } catch (error) {
                    reject(new Error(`Failed to parse Excel file: ${error.message}`));
                }
            };

            reader.onerror = () => {
                reject(new Error('Failed to read file'));
            };

            reader.readAsBinaryString(file);
        });
    }

    /**
     * Removes empty rows from worksheet (moved from brokerParsers.js for consistency)
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {Object} Worksheet with empty rows removed
     * @private
     */
    static _removeEmptyRowsFromWorksheet(worksheet) {
        if (!worksheet || !worksheet['!ref']) {
            return worksheet;
        }

        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const newWorksheet = {};
        let newRowIndex = range.s.r;
        let hasContent = false;

        // Copy worksheet properties
        Object.keys(worksheet).forEach(key => {
            if (key.startsWith('!')) {
                newWorksheet[key] = worksheet[key];
            }
        });

        // Process each row
        for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
            let rowHasContent = false;

            // Check if row has any non-empty content
            for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                const cell = worksheet[cellAddress];

                if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
                    rowHasContent = true;
                    break;
                }
            }

            // If row has content, copy it to new position
            if (rowHasContent) {
                for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
                    const oldAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                    const newAddress = XLSX.utils.encode_cell({ r: newRowIndex, c: colIndex });

                    if (worksheet[oldAddress]) {
                        newWorksheet[newAddress] = worksheet[oldAddress];
                    }
                }
                newRowIndex++;
                hasContent = true;
            }
        }

        // Update range if we have content
        if (hasContent) {
            newWorksheet['!ref'] = XLSX.utils.encode_range({
                s: { r: range.s.r, c: range.s.c },
                e: { r: newRowIndex - 1, c: range.e.c }
            });
        } else {
            // Empty worksheet
            newWorksheet['!ref'] = XLSX.utils.encode_range({
                s: { r: 0, c: 0 },
                e: { r: 0, c: 0 }
            });
        }

        return newWorksheet;
    }
}

// Export globally for cross-module access
window.ExcelCacheManager = ExcelCacheManager;