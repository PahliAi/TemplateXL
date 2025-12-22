/**
 * Borderellen Converter - Data Filter Engine
 * Applies row-level filters to parsed data (after parsing, before mapping)
 * Used exclusively for Generic Parser (not for built-in brokers)
 */

class DataFilterEngine {
    /**
     * Apply all data filters to parsed data
     * @param {Array} data - Parsed data array
     * @param {Array} dataFilters - Array of filter configurations
     * @returns {Array} Filtered data
     */
    static applyFilters(data, dataFilters) {
        if (!dataFilters || dataFilters.length === 0) {
            console.log('No data filters to apply');
            return data;
        }

        console.log(`Applying ${dataFilters.length} data filters to ${data.length} rows`);

        let filteredData = [...data];
        let filterStats = [];

        dataFilters.forEach((filter, index) => {
            if (!filter.enabled && filter.enabled !== undefined) {
                console.log(`Filter ${index + 1} disabled, skipping`);
                return;
            }

            const beforeCount = filteredData.length;
            filteredData = this.applySingleFilter(filteredData, filter);
            const afterCount = filteredData.length;
            const removed = beforeCount - afterCount;

            filterStats.push({
                filterType: filter.type,
                field: filter.field,
                beforeCount,
                afterCount,
                removed
            });

            const filterDesc = filter.field ? `${filter.type} on "${filter.field}"` : filter.type;
            console.log(`Filter ${index + 1} (${filterDesc}): ${beforeCount} → ${afterCount} rows (${removed} removed)`);

            // WARNING if filter removes all data
            if (afterCount === 0 && beforeCount > 0) {
                console.warn(`⚠️ Filter ${index + 1} removed ALL ${beforeCount} rows! Check filter: ${JSON.stringify(filter)}`);
            }
        });

        console.log(`Final result: ${data.length} → ${filteredData.length} rows (${data.length - filteredData.length} total removed)`);

        if (filteredData.length === 0 && data.length > 0) {
            console.error('❌ All data was filtered out! Review your data filters.');
        }

        return filteredData;
    }

    /**
     * Apply a single filter to data
     * @param {Array} data - Data array
     * @param {Object} filter - Filter configuration
     * @returns {Array} Filtered data
     */
    static applySingleFilter(data, filter) {
        switch (filter.type) {
            case 'require-any':
                return this.filterRequireAny(data, filter);

            case 'require-all':
                return this.filterRequireAll(data, filter);

            case 'numeric-compare':
                return this.filterNumericCompare(data, filter);

            case 'exclude-values':
                return this.filterExcludeValues(data, filter);

            case 'include-values':
                return this.filterIncludeValues(data, filter);

            case 'text-contains':
                return this.filterTextContains(data, filter);

            case 'text-not-contains':
                return this.filterTextNotContains(data, filter);

            case 'empty':
                return this.filterEmpty(data, filter);

            case 'not-empty':
                return this.filterNotEmpty(data, filter);

            case 'date-compare':
                return this.filterDateCompare(data, filter);

            case 'date-valid':
                return this.filterDateValid(data, filter);

            case 'regex':
                return this.filterRegex(data, filter);

            default:
                console.warn(`Unknown filter type: ${filter.type}`);
                return data;
        }
    }

    /**
     * Filter: At least one of the specified fields must be filled
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'require-any', fields: ['field1', 'field2']}
     * @returns {Array}
     */
    static filterRequireAny(data, filter) {
        return data.filter(row => {
            return filter.fields.some(field => {
                const value = row[field];
                return this.isValueFilled(value);
            });
        });
    }

    /**
     * Filter: All specified fields must be filled
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'require-all', fields: ['field1', 'field2']}
     * @returns {Array}
     */
    static filterRequireAll(data, filter) {
        return data.filter(row => {
            return filter.fields.every(field => {
                const value = row[field];
                return this.isValueFilled(value);
            });
        });
    }

    /**
     * Filter: Numeric comparison (>, <, >=, <=, ==, !=)
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'numeric-compare', field: 'Bruto', operator: '>', value: 250}
     * @returns {Array}
     */
    static filterNumericCompare(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return false;

            const numValue = this.parseNumber(value);
            if (numValue === null) return false;

            const compareValue = parseFloat(filter.value);
            if (isNaN(compareValue)) {
                console.warn(`Invalid comparison value: ${filter.value}`);
                return true; // Keep row if filter is invalid
            }

            switch (filter.operator) {
                case '>': return numValue > compareValue;
                case '<': return numValue < compareValue;
                case '>=': return numValue >= compareValue;
                case '<=': return numValue <= compareValue;
                case '==': return numValue === compareValue;
                case '!=': return numValue !== compareValue;
                default:
                    console.warn(`Unknown operator: ${filter.operator}`);
                    return true;
            }
        });
    }

    /**
     * Filter: Exclude rows with specific values (blacklist)
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'exclude-values', field: 'Soort', values: ['Total', 'Totaal']}
     * @returns {Array}
     */
    static filterExcludeValues(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return true;

            const valueStr = value.toString().trim();
            return !filter.values.some(excludeVal => excludeVal.toString().trim() === valueStr);
        });
    }

    /**
     * Filter: Include only rows with specific values (whitelist)
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'include-values', field: 'Branche', values: ['Vervoer', 'Transport']}
     * @returns {Array}
     */
    static filterIncludeValues(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return false;

            const valueStr = value.toString().trim();
            return filter.values.some(includeVal => includeVal.toString().trim() === valueStr);
        });
    }

    /**
     * Filter: Text contains substring (case-insensitive)
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'text-contains', field: 'Verzekerde', value: 'BV'}
     * @returns {Array}
     */
    static filterTextContains(data, filter) {
        const searchValue = filter.value.toString().toLowerCase();

        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return false;

            const valueStr = value.toString().toLowerCase();
            return valueStr.includes(searchValue);
        });
    }

    /**
     * Filter: Text does NOT contain substring (case-insensitive)
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'text-not-contains', field: 'Verzekerde', value: 'test'}
     * @returns {Array}
     */
    static filterTextNotContains(data, filter) {
        const searchValue = filter.value.toString().toLowerCase();

        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return true;

            const valueStr = value.toString().toLowerCase();
            return !valueStr.includes(searchValue);
        });
    }

    /**
     * Filter: Field is empty
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'empty', field: 'Notes'}
     * @returns {Array}
     */
    static filterEmpty(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            return !this.isValueFilled(value);
        });
    }

    /**
     * Filter: Field is NOT empty
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'not-empty', field: 'PolisNr'}
     * @returns {Array}
     */
    static filterNotEmpty(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            return this.isValueFilled(value);
        });
    }

    /**
     * Filter: Date comparison (before/after/between)
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'date-compare', field: 'Boekdatum tp', operator: 'after', value: '2024-01-01'}
     * @returns {Array}
     */
    static filterDateCompare(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return false;

            const dateValue = this.parseDate(value);
            if (!dateValue) return false;

            const compareDate = this.parseDate(filter.value);
            if (!compareDate) {
                console.warn(`Invalid comparison date: ${filter.value}`);
                return true;
            }

            switch (filter.operator) {
                case 'before':
                    return dateValue < compareDate;
                case 'after':
                    return dateValue > compareDate;
                case 'equals':
                    return dateValue.getTime() === compareDate.getTime();
                case 'between':
                    const endDate = this.parseDate(filter.endValue);
                    if (!endDate) return true;
                    return dateValue >= compareDate && dateValue <= endDate;
                default:
                    console.warn(`Unknown date operator: ${filter.operator}`);
                    return true;
            }
        });
    }

    /**
     * Filter: Field contains a valid date
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'date-valid', field: 'Oorspr_BoekDtm'}
     * @returns {Array}
     */
    static filterDateValid(data, filter) {
        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return false;

            const dateValue = this.parseDate(value);
            return dateValue !== null; // Keep row if date is valid
        });
    }

    /**
     * Filter: Regex pattern match
     * @param {Array} data - Data array
     * @param {Object} filter - {type: 'regex', field: 'PolisNr', pattern: '^\\d{4}\\.\\d{2}'}
     * @returns {Array}
     */
    static filterRegex(data, filter) {
        let regex;
        try {
            regex = new RegExp(filter.pattern, filter.flags || 'i');
        } catch (error) {
            console.error(`Invalid regex pattern: ${filter.pattern}`, error);
            return data; // Keep all data if regex is invalid
        }

        return data.filter(row => {
            const value = row[filter.field];
            if (value === null || value === undefined) return false;

            const valueStr = value.toString();
            return regex.test(valueStr);
        });
    }

    /**
     * Helper: Check if value is filled (not empty/null/undefined)
     * @param {*} value - Value to check
     * @returns {Boolean}
     */
    static isValueFilled(value) {
        if (value === null || value === undefined) return false;
        if (typeof value === 'string' && value.trim() === '') return false;
        return true;
    }

    /**
     * Helper: Parse number from various formats
     * @param {*} value - Value to parse
     * @returns {Number|null}
     */
    static parseNumber(value) {
        if (typeof value === 'number') return value;

        if (typeof value === 'string') {
            // Remove common thousand separators and replace decimal comma with dot
            const cleaned = value.replace(/[.,\s]/g, (match, offset, str) => {
                // Keep last comma/dot as decimal separator
                const lastCommaIndex = str.lastIndexOf(',');
                const lastDotIndex = str.lastIndexOf('.');
                const lastSeparatorIndex = Math.max(lastCommaIndex, lastDotIndex);

                if (offset === lastSeparatorIndex) {
                    return '.'; // Decimal separator
                }
                return ''; // Remove thousand separator
            });

            const num = parseFloat(cleaned);
            return isNaN(num) ? null : num;
        }

        return null;
    }

    /**
     * Helper: Parse date from various formats
     * @param {*} value - Value to parse
     * @returns {Date|null}
     */
    static parseDate(value) {
        if (value instanceof Date) return value;

        if (typeof value === 'number') {
            // Excel date serial number
            const excelEpoch = new Date(1899, 11, 30);
            const date = new Date(excelEpoch.getTime() + value * 86400000);
            return isNaN(date.getTime()) ? null : date;
        }

        if (typeof value === 'string') {
            // Try parsing common date formats
            const date = new Date(value);
            if (!isNaN(date.getTime())) return date;

            // Try DD-MM-YYYY format
            const parts = value.split(/[-/]/);
            if (parts.length === 3) {
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1;
                const year = parseInt(parts[2], 10);
                const date = new Date(year, month, day);
                if (!isNaN(date.getTime())) return date;
            }
        }

        return null;
    }

    /**
     * Get filter statistics
     * @param {Array} originalData - Original data
     * @param {Array} filteredData - Filtered data
     * @returns {Object} Statistics
     */
    static getFilterStats(originalData, filteredData) {
        const removed = originalData.length - filteredData.length;
        const percentage = originalData.length > 0
            ? ((removed / originalData.length) * 100).toFixed(1)
            : 0;

        return {
            originalCount: originalData.length,
            filteredCount: filteredData.length,
            removedCount: removed,
            removedPercentage: percentage
        };
    }
}

// Make available globally
window.DataFilterEngine = DataFilterEngine;
