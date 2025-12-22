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
     * @param {string|Object} sheetSelector - Sheet name, pattern object {pattern: 'regex'}, or null for first sheet
     * @returns {Promise<Object>} XLSX worksheet object
     */
    static async getSheet(file, sheetSelector = null) {
        const workbook = await this.getWorkbook(file);

        let sheetName = null;

        if (!sheetSelector) {
            // Default: first sheet
            sheetName = workbook.SheetNames[0];
        } else if (typeof sheetSelector === 'string') {
            // Direct sheet name
            sheetName = sheetSelector;
        } else if (typeof sheetSelector === 'object' && sheetSelector.pattern) {
            // Pattern matching
            sheetName = SheetSelectorManager.findSheetByPattern(workbook.SheetNames, sheetSelector.pattern);
            if (!sheetName) {
                console.warn(`No sheet found matching pattern: ${sheetSelector.pattern}, using first sheet`);
                sheetName = workbook.SheetNames[0];
            } else {
                console.log(`Sheet selected by pattern "${sheetSelector.pattern}": ${sheetName}`);
            }
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
                    // Use 'array' type for better compatibility with both .xlsx and .xls files
                    const workbook = XLSX.read(e.target.result, {
                        type: 'array',
                        cellFormula: false  // Read calculated values instead of formulas
                    });

                    // Debug: Log workbook info
                    console.log(`[ExcelCacheManager] Loaded workbook with ${workbook.SheetNames.length} sheets:`, workbook.SheetNames);

                    // Apply consistent compaction to all worksheets
                    workbook.SheetNames.forEach(sheetName => {
                        const originalSheet = workbook.Sheets[sheetName];

                        // [FIX] If worksheet has no !ref, try to generate it from cell addresses
                        if (!originalSheet['!ref']) {
                            console.warn(`[ExcelCacheManager] Worksheet "${sheetName}" has no !ref, attempting to generate...`);

                            // Debug: Log all keys in worksheet
                            const allKeys = Object.keys(originalSheet);
                            console.log(`[ExcelCacheManager] Worksheet "${sheetName}" has ${allKeys.length} keys:`, allKeys.slice(0, 20));

                            const generatedRef = this._generateRangeFromCells(originalSheet);
                            if (generatedRef) {
                                originalSheet['!ref'] = generatedRef;
                                console.log(`[ExcelCacheManager] Generated !ref for "${sheetName}": ${generatedRef}`);
                            } else {
                                console.error(`[ExcelCacheManager] Could not generate !ref for "${sheetName}" - sheet appears empty or has no valid cell addresses`);
                                // Create a minimal default range to prevent errors
                                originalSheet['!ref'] = 'A1:A1';
                                console.warn(`[ExcelCacheManager] Using fallback range A1:A1 for "${sheetName}"`);
                            }
                        }

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

            // Use readAsArrayBuffer for better compatibility with both .xlsx and .xls files
            reader.readAsArrayBuffer(file);
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

    /**
     * Generate !ref range from cell addresses (for worksheets missing !ref property)
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {string|null} Range string (e.g., "A1:Z100") or null if no cells found
     * @private
     */
    static _generateRangeFromCells(worksheet) {
        if (!worksheet) return null;

        let minRow = Infinity;
        let maxRow = -Infinity;
        let minCol = Infinity;
        let maxCol = -Infinity;
        let foundCells = false;

        // Iterate through all worksheet properties looking for cell addresses
        for (let key in worksheet) {
            // Skip metadata properties (start with !)
            if (key.startsWith('!')) continue;

            // Try to decode cell address
            try {
                const addr = XLSX.utils.decode_cell(key);
                minRow = Math.min(minRow, addr.r);
                maxRow = Math.max(maxRow, addr.r);
                minCol = Math.min(minCol, addr.c);
                maxCol = Math.max(maxCol, addr.c);
                foundCells = true;
            } catch (e) {
                // Not a valid cell address, skip
                continue;
            }
        }

        if (!foundCells) {
            return null;
        }

        // Generate range string
        return XLSX.utils.encode_range({
            s: { r: minRow, c: minCol },
            e: { r: maxRow, c: maxCol }
        });
    }

    /**
     * Clears cache entry for specific file
     * @param {File} file - File to remove from cache
     */
    static clearFile(file) {
        this.cache.delete(file);
        console.log('[ExcelCacheManager] Cleared cache for file:', file.name);
    }

    /**
     * Clears entire cache
     */
    static clearAll() {
        const count = this.cache.size;
        this.cache.clear();
        console.log(`[ExcelCacheManager] Cleared entire cache (${count} files removed)`);
    }

    /**
     * Get cache statistics
     */
    static getCacheStats() {
        return {
            size: this.cache.size,
            files: Array.from(this.cache.keys()).map(f => f.name)
        };
    }
}

// Export globally for cross-module access
window.ExcelCacheManager = ExcelCacheManager;