/**
 * Borderellen Converter - Data Pattern Analyzer
 * Analyzes Excel sheets to detect data patterns and suggest parsing configuration
 */

class DataPatternAnalyzer {
    /**
     * Analyzes an Excel worksheet to detect data patterns
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {Object} Analysis results with suggestions
     */
    static analyzeSheet(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        console.log(`Analyzing sheet with range: ${worksheet['!ref']}`);

        // Use new two-dimensional density analysis
        const densityAnalysis = this.analyzeCellDensity(worksheet, range);
        console.log(`Density analysis:`, densityAnalysis);

        // Find data section using both row and column density
        const dataSection = this.findDataSectionByDensity(densityAnalysis, worksheet);
        console.log(`Data section detected:`, dataSection);

        // Analyze header row (use detected header row, not start - 1)
        const headerRowIndex = dataSection.headerRowIndex !== null ? dataSection.headerRowIndex : Math.max(0, dataSection.start - 1);
        const headerAnalysis = this.analyzeHeaderRow(worksheet, headerRowIndex);

        // Analyze data quality
        const qualityAnalysis = this.analyzeDataQuality(worksheet, dataSection, range);

        return {
            dataSection: dataSection,
            suggestedHeaderRow: headerRowIndex,
            suggestedDataStart: dataSection.dataStartIndex || dataSection.start,
            suggestedDataEnd: dataSection.end,
            confidence: dataSection.confidence || 0.8,
            headerAnalysis: headerAnalysis,
            qualityAnalysis: qualityAnalysis,
            suggestions: this.generateSuggestions(dataSection, headerAnalysis, qualityAnalysis)
        };
    }

    /**
     * Analyzes a file (uses centralized Excel cache for consistency)
     * @param {File} file - File object to analyze
     * @returns {Promise<Object>} Analysis results
     */
    static async analyzeFile(file) {
        try {
            // Use centralized cache to ensure same compacted state as execution
            const workbook = await ExcelCacheManager.getWorkbook(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const analysis = this.analyzeSheet(worksheet);

            analysis.filename = file.name;
            analysis.sheetName = workbook.SheetNames[0];
            analysis.analyzedAt = new Date().toISOString();

            return analysis;
        } catch (error) {
            throw new Error(`Failed to analyze file ${file.name}: ${error.message}`);
        }
    }

    /**
     * Analyzes cell density in both row and column dimensions
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} range - Decoded range object
     * @returns {Object} Density analysis results
     */
    static analyzeCellDensity(worksheet, range) {
        const maxRowsToAnalyze = range.e.r + 1;  // Analyze ALL rows in worksheet
        const maxColsToAnalyze = range.e.c + 1;  // Analyze ALL columns in worksheet

        console.log(`Analyzing full worksheet range: ${range.e.r + 1} rows x ${range.e.c + 1} columns`);

        const rowDensities = [];
        const columnDensities = [];

        // Analyze row densities (count filled cells per row)
        for (let row = range.s.r; row < maxRowsToAnalyze; row++) {
            let filledCells = 0;
            let totalCells = 0;

            for (let col = range.s.c; col < maxColsToAnalyze; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                totalCells++;
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    filledCells++;
                }
            }

            rowDensities.push({
                row: row,
                filledCells: filledCells,
                totalCells: totalCells,
                density: totalCells > 0 ? filledCells / totalCells : 0,
                isEmpty: filledCells === 0
            });
        }

        // Analyze column densities (count filled cells per column)
        for (let col = range.s.c; col < maxColsToAnalyze; col++) {
            let filledCells = 0;
            let totalCells = 0;

            for (let row = range.s.r; row < maxRowsToAnalyze; row++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                totalCells++;
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    filledCells++;
                }
            }

            columnDensities.push({
                column: col,
                filledCells: filledCells,
                totalCells: totalCells,
                density: totalCells > 0 ? filledCells / totalCells : 0,
                isEmpty: filledCells === 0
            });
        }

        return {
            rowDensities: rowDensities,
            columnDensities: columnDensities,
            maxRowsAnalyzed: maxRowsToAnalyze,
            maxColsAnalyzed: maxColsToAnalyze
        };
    }

    /**
     * Find data section using two-dimensional density analysis
     * @param {Object} densityAnalysis - Results from analyzeCellDensity
     * @returns {Object} Data section with start, end, length, and confidence
     */
    static findDataSectionByDensity(densityAnalysis, worksheet) {
        const { rowDensities, columnDensities } = densityAnalysis;

        // Step 1: Find rows with highest density (potential data rows)
        const rowsByDensity = [...rowDensities]
            .filter(r => !r.isEmpty)
            .sort((a, b) => b.density - a.density);

        // Step 2: Find columns with highest density (potential data columns)
        const colsByDensity = [...columnDensities]
            .filter(c => !c.isEmpty)
            .sort((a, b) => b.density - a.density);

        if (rowsByDensity.length === 0 || colsByDensity.length === 0) {
            return { start: 0, end: 0, length: 0, confidence: 0, reason: 'No data found' };
        }

        // Step 3: Identify the primary data density threshold
        const topRowDensity = rowsByDensity[0].density;
        const topColDensity = colsByDensity[0].density;

        // Use 70% of top density as threshold for "high density" rows/columns
        const rowThreshold = Math.max(0.5, topRowDensity * 0.7);
        const colThreshold = Math.max(0.3, topColDensity * 0.7);

        // Step 4: Find consecutive high-density rows
        const highDensityRows = rowDensities
            .map((r, idx) => ({ ...r, index: idx }))
            .filter(r => r.density >= rowThreshold);

        // Step 5: Find all consecutive sequences of high-density rows
        let allSequences = this.findAllConsecutiveSequences(highDensityRows);

        // Step 6: Find START CELL by intersecting row and column density
        let headerRowIndex = null;
        let startColumnIndex = null;
        let endColumnIndex = null;
        let dataStartIndex = null;
        let dataEndIndex = null;
        let totalDataLength = 0;

        // Find the row with highest density (likely header row)
        const highestDensityRow = rowsByDensity[0];
        if (highestDensityRow && highestDensityRow.density > 0.7) {
            headerRowIndex = highestDensityRow.row;
            console.log(`Header row identified at row ${headerRowIndex} (density: ${highestDensityRow.density.toFixed(2)})`);

            // Find the column with highest density (likely first data column)
            const highestDensityColumn = colsByDensity[0];
            if (highestDensityColumn && highestDensityColumn.density > 0.2) {
                startColumnIndex = highestDensityColumn.column;
                console.log(`Start column identified at column ${XLSX.utils.encode_col(startColumnIndex)} (density: ${highestDensityColumn.density.toFixed(2)})`);

                // Find longest consecutive string of non-empty header cells
                // This is the ACTUAL table boundary, not limited by density analysis

                const range = XLSX.utils.decode_range(worksheet['!ref']);
                console.log(`Scanning full header row ${headerRowIndex} from column ${XLSX.utils.encode_col(startColumnIndex)} to ${XLSX.utils.encode_col(range.e.c)}`);

                // Find the longest consecutive sequence starting from startColumnIndex
                const headerCells = [];
                endColumnIndex = startColumnIndex;

                for (let col = startColumnIndex; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
                    const cell = worksheet[cellAddress];
                    const hasContent = cell && cell.v !== null && cell.v !== undefined && cell.v !== '';

                    if (hasContent) {
                        headerCells.push({ col, content: cell.v });
                        endColumnIndex = col;  // Extend table boundary
                        console.log(`Header ${cellAddress}: "${cell.v}"`);
                    } else {
                        // Allow small gaps (1-2 empty cells) in header sequence
                        let gapSize = 0;
                        let nextCol = col + 1;

                        // Look ahead for content after small gap
                        while (nextCol <= range.e.c && gapSize < 2) {
                            const nextAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: nextCol });
                            const nextCell = worksheet[nextAddress];
                            const nextHasContent = nextCell && nextCell.v !== null && nextCell.v !== undefined && nextCell.v !== '';

                            if (nextHasContent) {
                                // Found content after gap - include the gap and continue
                                console.log(`Gap of ${gapSize + 1} cells found, but continuing due to content at ${nextAddress}`);

                                // Include empty cells in the gap
                                for (let gapCol = col; gapCol < nextCol; gapCol++) {
                                    headerCells.push({ col: gapCol, content: '' });
                                }

                                col = nextCol - 1; // Loop will increment, so we'll process nextCol next
                                break;
                            }

                            gapSize++;
                            nextCol++;
                        }

                        // If no content found after reasonable gap, end the table here
                        if (gapSize >= 2 || nextCol > range.e.c) {
                            console.log(`End of header sequence at column ${XLSX.utils.encode_col(col - 1)} (gap too large or end of sheet)`);
                            break;
                        }
                    }
                }

                console.log(`Found ${headerCells.length} header positions, table range: ${XLSX.utils.encode_col(startColumnIndex)} to ${XLSX.utils.encode_col(endColumnIndex)}`);
                console.log('Complete header sequence:', headerCells.map(h => `${XLSX.utils.encode_col(h.col)}:"${h.content}"`));

                console.log(`Data column range: ${XLSX.utils.encode_col(startColumnIndex)} to ${XLSX.utils.encode_col(endColumnIndex)}`);

                // Find last row with any data in the detected column range
                let lastDataRow = headerRowIndex;
                for (let row = headerRowIndex + 1; row <= range.e.r; row++) {
                    let hasDataInRow = false;
                    for (let col = startColumnIndex; col <= endColumnIndex; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                            hasDataInRow = true;
                            break;
                        }
                    }
                    if (hasDataInRow) {
                        lastDataRow = row;
                    }
                }

                console.log(`Data extends from header row ${headerRowIndex} to last data row ${lastDataRow}`);

                // Analyze density between header and last data row to detect summary rows
                dataStartIndex = headerRowIndex + 1;
                dataEndIndex = lastDataRow;

                // Calculate row densities for the data section and detect summary rows
                const dataRowDensities = [];
                const summaryRows = [];

                for (let row = dataStartIndex; row <= dataEndIndex; row++) {
                    let filledCells = 0;
                    let totalCells = endColumnIndex - startColumnIndex + 1;
                    let rowText = '';

                    for (let col = startColumnIndex; col <= endColumnIndex; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                            filledCells++;
                            rowText += cell.v.toString().toLowerCase() + ' ';
                        }
                    }

                    const density = totalCells > 0 ? filledCells / totalCells : 0;
                    dataRowDensities.push({ row, density, filledCells, rowText: rowText.trim() });

                    // Detect summary rows: high density + keywords
                    const isSummaryRow = (density > 0.7) && (
                        rowText.includes('totaal') ||
                        rowText.includes('total') ||
                        rowText.includes('sum') ||
                        rowText.includes('subtotal') ||
                        rowText.includes('grand total') ||
                        rowText.includes('eindtotaal')
                    );

                    if (isSummaryRow) {
                        summaryRows.push(row);
                        console.log(`Summary row detected at ${row}: "${rowText}" (density: ${density.toFixed(2)})`);
                    }
                }

                // Exclude summary rows from data section
                if (summaryRows.length > 0) {
                    const firstSummaryRow = Math.min(...summaryRows);
                    if (firstSummaryRow < dataEndIndex) {
                        dataEndIndex = firstSummaryRow - 1;
                        console.log(`Adjusted data end to row ${dataEndIndex} (excluding summary rows)`);
                    }
                }

                totalDataLength = Math.max(1, dataEndIndex - dataStartIndex + 1);
                console.log(`Final data section: rows ${dataStartIndex} to ${dataEndIndex} (${totalDataLength} data rows)`);
            } else {
                console.log(`No clear start column found (highest density: ${colsByDensity[0]?.density.toFixed(2) || 'none'})`);
                startColumnIndex = 0; // Fallback to column A
                endColumnIndex = Math.min(20, columnDensities.length - 1); // Default range
            }
        }

        // Step 7: Calculate confidence based on detected structure (both row AND column)
        let confidence = 0;
        let sectionStart = headerRowIndex || 0;
        let sectionEnd = dataEndIndex || headerRowIndex || 0;
        let avgDensity = 0;

        if (headerRowIndex !== null && startColumnIndex !== null) {
            // We found both header row AND start column - high confidence
            const headerDensity = rowDensities[headerRowIndex].density;
            const columnDensity = columnDensities[startColumnIndex].density;

            // Calculate average density of data section
            if (dataStartIndex !== null && dataEndIndex !== null) {
                let densitySum = headerDensity;
                for (let i = dataStartIndex; i <= dataEndIndex; i++) {
                    densitySum += rowDensities[i].density;
                }
                avgDensity = densitySum / totalDataLength;

                // Very high confidence for clear start cell + data pattern
                confidence = Math.min(1.0,
                    headerDensity * // Header quality (perfect header = high confidence)
                    columnDensity * // Column quality (high column density = high confidence)
                    (totalDataLength / Math.max(1, rowDensities.length * 0.1)) * // Data length score (very lenient)
                    1.2 // Boost for having both dimensions
                );
            } else {
                // Only header found, but we have start column too
                avgDensity = headerDensity;
                confidence = headerDensity * columnDensity * 0.9; // High confidence for perfect start cell
            }

            sectionStart = headerRowIndex;
            sectionEnd = dataEndIndex || headerRowIndex;
        } else if (headerRowIndex !== null) {
            // Fall back to row-only analysis
            const headerDensity = rowDensities[headerRowIndex].density;
            avgDensity = headerDensity;
            confidence = headerDensity * 0.6; // Lower confidence without column info
            sectionStart = headerRowIndex;
            sectionEnd = dataEndIndex || headerRowIndex;
        }

        const startCell = headerRowIndex !== null && startColumnIndex !== null
            ? `${XLSX.utils.encode_col(startColumnIndex)}${headerRowIndex + 1}`
            : 'Not detected';

        console.log(`Final analysis: Start cell ${startCell}, Header at ${headerRowIndex}, Data ${dataStartIndex}-${dataEndIndex}, Confidence: ${confidence.toFixed(3)}`);

        return {
            start: sectionStart,
            end: sectionEnd,
            length: totalDataLength || 1,
            avgDensity: avgDensity,
            confidence: confidence,
            headerRowIndex: headerRowIndex,
            dataStartIndex: dataStartIndex || headerRowIndex,
            startColumnIndex: startColumnIndex,
            endColumnIndex: endColumnIndex,
            startCell: startCell,
            pattern: this.createPatternFromSequence(rowDensities, sectionStart, sectionEnd)
        };
    }

    /**
     * Find all consecutive sequences of rows/columns
     * @param {Array} items - Array of items with index property
     * @returns {Array} Array of sequence objects with start, length, avgDensity
     */
    static findAllConsecutiveSequences(items) {
        if (items.length === 0) return [];

        const sequences = [];
        let currentSequence = {
            start: items[0].index,
            length: 1,
            densitySum: items[0].density
        };

        for (let i = 1; i < items.length; i++) {
            if (items[i].index === items[i-1].index + 1) {
                // Consecutive - extend current sequence
                currentSequence.length++;
                currentSequence.densitySum += items[i].density;
            } else {
                // Not consecutive - save current and start new sequence
                sequences.push({
                    start: currentSequence.start,
                    length: currentSequence.length,
                    avgDensity: currentSequence.densitySum / currentSequence.length
                });

                currentSequence = {
                    start: items[i].index,
                    length: 1,
                    densitySum: items[i].density
                };
            }
        }

        // Add final sequence
        sequences.push({
            start: currentSequence.start,
            length: currentSequence.length,
            avgDensity: currentSequence.densitySum / currentSequence.length
        });

        return sequences;
    }

    /**
     * Find longest consecutive sequence of rows/columns
     * @param {Array} items - Array of items with index property
     * @returns {Object} Sequence info with start, length, avgDensity
     */
    static findLongestConsecutiveSequence(items) {
        if (items.length === 0) return { start: 0, length: 0, avgDensity: 0 };

        let maxLength = 1;
        let maxStart = items[0].index;
        let maxDensitySum = items[0].density;

        let currentLength = 1;
        let currentStart = items[0].index;
        let currentDensitySum = items[0].density;

        for (let i = 1; i < items.length; i++) {
            if (items[i].index === items[i-1].index + 1) {
                // Consecutive
                currentLength++;
                currentDensitySum += items[i].density;
            } else {
                // Check if current sequence is better
                if (currentLength > maxLength ||
                    (currentLength === maxLength && currentDensitySum > maxDensitySum)) {
                    maxLength = currentLength;
                    maxStart = currentStart;
                    maxDensitySum = currentDensitySum;
                }

                // Start new sequence
                currentLength = 1;
                currentStart = items[i].index;
                currentDensitySum = items[i].density;
            }
        }

        // Check final sequence
        if (currentLength > maxLength ||
            (currentLength === maxLength && currentDensitySum > maxDensitySum)) {
            maxLength = currentLength;
            maxStart = currentStart;
            maxDensitySum = currentDensitySum;
        }

        return {
            start: maxStart,
            length: maxLength,
            avgDensity: maxLength > 0 ? maxDensitySum / maxLength : 0
        };
    }

    /**
     * Calculate average column density within specific rows
     * @param {Object} densityAnalysis - Full density analysis
     * @param {Array} rowIndices - Row indices to analyze
     * @returns {Number} Average column density in those rows
     */
    static calculateColumnDensityInRows(densityAnalysis, rowIndices) {
        const { columnDensities } = densityAnalysis;

        let totalDensity = 0;
        let validColumns = 0;

        for (const colData of columnDensities) {
            if (colData.density > 0) {
                totalDensity += colData.density;
                validColumns++;
            }
        }

        return validColumns > 0 ? totalDensity / validColumns : 0;
    }


    /**
     * Create pattern string for a specific sequence
     * @param {Array} rowDensities - All row densities
     * @param {Number} start - Start index
     * @param {Number} end - End index
     * @returns {String} Pattern for the sequence
     */
    static createPatternFromSequence(rowDensities, start, end) {
        const fillThreshold = 0.5;
        return rowDensities
            .slice(start, end + 1)
            .map(r => r.density > fillThreshold ? 'F' : 'E')
            .join('');
    }

    /**
     * Creates pattern string of filled (F) and empty (E) rows
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} range - Decoded range object
     * @returns {String} Pattern like "FFFEEFEEFFFFFFFFFEEFFFEF"
     */
    static getFilledRowPattern(worksheet, range) {
        const pattern = [];
        const fillThreshold = 0.5; // Row is "filled" if >50% of cells have data

        for (let row = range.s.r; row <= range.e.r; row++) {
            let filledCells = 0;
            let totalCells = 0;

            // Check each column in this row
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                totalCells++;
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    filledCells++;
                }
            }

            // Determine if row is filled or empty
            const fillRatio = totalCells > 0 ? filledCells / totalCells : 0;
            pattern.push(fillRatio > fillThreshold ? 'F' : 'E');
        }

        return pattern.join('');
    }

    /**
     * Finds the longest consecutive sequence of 'F' characters
     * @param {String} pattern - Pattern string like "FFFEEFEEFFFFFFFFFEEFFFEF"
     * @returns {Object} Object with start, end, length of longest sequence
     */
    static findLongestFilledSequence(pattern) {
        let maxLength = 0;
        let maxStart = 0;
        let maxEnd = 0;
        let currentLength = 0;
        let currentStart = 0;

        for (let i = 0; i < pattern.length; i++) {
            if (pattern[i] === 'F') {
                if (currentLength === 0) {
                    currentStart = i;
                }
                currentLength++;
            } else {
                if (currentLength > maxLength) {
                    maxLength = currentLength;
                    maxStart = currentStart;
                    maxEnd = currentStart + currentLength - 1;
                }
                currentLength = 0;
            }
        }

        // Check if the last sequence was the longest
        if (currentLength > maxLength) {
            maxLength = currentLength;
            maxStart = currentStart;
            maxEnd = currentStart + currentLength - 1;
        }

        return {
            start: maxStart,
            end: maxEnd,
            length: maxLength,
            pattern: pattern.slice(maxStart, maxEnd + 1)
        };
    }

    /**
     * Analyzes the header row to understand column structure
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Number} headerRowIndex - Row index to analyze as header
     * @returns {Object} Header analysis results
     */
    static analyzeHeaderRow(worksheet, headerRowIndex) {
        if (headerRowIndex < 0) {
            return { found: false, reason: 'No header row detected' };
        }

        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const headers = [];
        let filledHeaders = 0;

        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
            const cell = worksheet[cellAddress];
            const headerValue = cell && cell.v ? cell.v.toString().trim() : '';

            headers.push(headerValue);
            if (headerValue) filledHeaders++;
        }

        return {
            found: filledHeaders > 0,
            row: headerRowIndex,
            headers: headers,
            filledCount: filledHeaders,
            totalCount: headers.length,
            fillRatio: headers.length > 0 ? filledHeaders / headers.length : 0,
            suggestedColumns: headers.filter(h => h && h.length > 0)
        };
    }

    /**
     * Analyzes data quality in the detected data section
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} dataSection - Detected data section
     * @param {Object} range - Sheet range
     * @returns {Object} Quality analysis results
     */
    static analyzeDataQuality(worksheet, dataSection, range) {
        const sampleSize = Math.min(10, dataSection.length); // Analyze first 10 rows
        const columnStats = new Map();

        for (let row = dataSection.start; row < dataSection.start + sampleSize; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                const value = cell && cell.v !== null ? cell.v : null;

                if (!columnStats.has(col)) {
                    columnStats.set(col, {
                        filledCount: 0,
                        emptyCount: 0,
                        dataTypes: new Set(),
                        sampleValues: []
                    });
                }

                const stats = columnStats.get(col);
                if (value !== null && value !== '') {
                    stats.filledCount++;
                    stats.dataTypes.add(typeof value);
                    if (stats.sampleValues.length < 3) {
                        stats.sampleValues.push(value);
                    }
                } else {
                    stats.emptyCount++;
                }
            }
        }

        // Convert to array with column analysis
        const columnAnalysis = [];
        for (let col = range.s.c; col <= range.e.c; col++) {
            const stats = columnStats.get(col) || { filledCount: 0, emptyCount: sampleSize };
            columnAnalysis.push({
                column: col,
                fillRatio: stats.filledCount / sampleSize,
                dataTypes: stats.dataTypes ? Array.from(stats.dataTypes) : [],
                sampleValues: stats.sampleValues || [],
                likelyDataColumn: stats.filledCount > 0
            });
        }

        return {
            sampleSize: sampleSize,
            columnAnalysis: columnAnalysis,
            likelyDataColumns: columnAnalysis.filter(c => c.likelyDataColumn).length,
            averageFillRatio: columnAnalysis.reduce((sum, c) => sum + c.fillRatio, 0) / columnAnalysis.length
        };
    }

    /**
     * Generates parsing suggestions based on analysis
     * @param {Object} dataSection - Detected data section
     * @param {Object} headerAnalysis - Header analysis
     * @param {Object} qualityAnalysis - Quality analysis
     * @returns {Object} Parsing suggestions
     */
    static generateSuggestions(dataSection, headerAnalysis, qualityAnalysis) {
        const suggestions = {
            dataStartMethod: 'auto-detect',
            confidence: 'high',
            recommendations: [],
            warnings: []
        };

        // Determine confidence level
        if (dataSection.length < 5) {
            suggestions.confidence = 'low';
            suggestions.warnings.push('Very short data section detected. Manual review recommended.');
        } else if (dataSection.length < 10) {
            suggestions.confidence = 'medium';
            suggestions.warnings.push('Short data section. Verify detection accuracy.');
        }

        // Header recommendations
        if (headerAnalysis.found && headerAnalysis.fillRatio > 0.7) {
            suggestions.recommendations.push(`Strong header row detected at row ${headerAnalysis.row + 1}`);
            suggestions.useHeaderRow = true;
            suggestions.headerRow = headerAnalysis.row;
        } else if (headerAnalysis.found) {
            suggestions.recommendations.push(`Partial header row detected at row ${headerAnalysis.row + 1}. Manual verification recommended.`);
            suggestions.useHeaderRow = true;
            suggestions.headerRow = headerAnalysis.row;
        } else {
            suggestions.recommendations.push('No clear header row detected. Consider position-based parsing.');
            suggestions.useHeaderRow = false;
        }

        // Data processing recommendations
        if (qualityAnalysis.averageFillRatio > 0.8) {
            suggestions.recommendations.push('High data quality detected. Standard single-row processing recommended.');
            suggestions.rowProcessing = 'single';
        } else if (qualityAnalysis.averageFillRatio > 0.4) {
            suggestions.recommendations.push('Mixed data quality. Review for possible multi-row patterns.');
            suggestions.rowProcessing = 'single';
            suggestions.warnings.push('Consider filtering empty or incomplete rows.');
        } else {
            suggestions.recommendations.push('Low data density. Multi-row or complex parsing may be needed.');
            suggestions.rowProcessing = 'review';
        }

        // Skip rows suggestion
        if (dataSection.start > 0) {
            suggestions.recommendations.push(`Skip first ${dataSection.start} rows to reach data section.`);
            suggestions.skipRows = dataSection.start;
        }

        return suggestions;
    }

    /**
     * Detects potential date columns by analyzing data patterns
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} dataSection - Data section to analyze
     * @returns {Array} Array of column indices that likely contain dates
     */
    static detectDateColumns(worksheet, dataSection) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const dateColumns = [];
        const sampleSize = Math.min(5, dataSection.length);

        for (let col = range.s.c; col <= range.e.c; col++) {
            let dateCount = 0;
            let totalCount = 0;

            for (let row = dataSection.start; row < dataSection.start + sampleSize; row++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                    totalCount++;
                    if (this.isLikelyDate(cell.v)) {
                        dateCount++;
                    }
                }
            }

            if (totalCount > 0 && dateCount / totalCount > 0.7) {
                dateColumns.push({
                    column: col,
                    confidence: dateCount / totalCount,
                    sampleSize: totalCount
                });
            }
        }

        return dateColumns;
    }

    /**
     * Checks if a value is likely a date
     * @param {*} value - Value to check
     * @returns {Boolean} True if likely a date
     */
    static isLikelyDate(value) {
        if (!value) return false;

        // Check if it's already a Date object
        if (value instanceof Date) return true;

        // Check if it's a number that could be an Excel date serial
        if (typeof value === 'number' && value > 25567 && value < 65000) {
            return true; // Excel date range roughly 1970-2078
        }

        // Try parsing as string date
        if (typeof value === 'string') {
            const dateRegex = /^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$/;
            if (dateRegex.test(value)) return true;

            const parsedDate = Date.parse(value);
            return !isNaN(parsedDate);
        }

        return false;
    }

    /**
     * Detects potential multi-row patterns (like BCI format)
     * @param {Object} worksheet - XLSX worksheet object
     * @param {Object} dataSection - Data section to analyze
     * @returns {Object|null} Multi-row pattern if detected
     */
    static detectMultiRowPattern(worksheet, dataSection) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sampleRows = Math.min(6, dataSection.length);

        // Check for alternating patterns (like BCI's 2-row structure)
        const rowPatterns = [];

        for (let row = dataSection.start; row < dataSection.start + sampleRows; row++) {
            const rowData = [];
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                rowData.push(cell && cell.v ? 'F' : 'E');
            }
            rowPatterns.push(rowData.join(''));
        }

        // Look for repeating 2-row patterns
        if (rowPatterns.length >= 4) {
            const isAlternating = rowPatterns[0] === rowPatterns[2] &&
                                 rowPatterns[1] === rowPatterns[3] &&
                                 rowPatterns[0] !== rowPatterns[1];

            if (isAlternating) {
                return {
                    type: 'alternating-2-row',
                    pattern1: rowPatterns[0],
                    pattern2: rowPatterns[1],
                    confidence: 0.8
                };
            }
        }

        return null;
    }
}