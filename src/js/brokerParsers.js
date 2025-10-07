/**
 * Borderellen Converter - Broker Parsers
 * Handles parsing of different broker Excel formats according to specifications
 */

class AONParser {
    static detect(filename) {
        return /^AON B550 (\d{2}-\d{4})\.xlsx$/i.test(filename);
    }

    static parse(workbook, filename) {
        console.log('AON Parser: Processing file', filename);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet, {
            cellFormula: false  // Read calculated values instead of formulas
        });

        // Extract period from filename
        const match = filename.match(/(\d{2}-\d{4})/);
        const period = match ? match[1] : '';

        console.log(`AON Parser: Found ${data.length} rows, period: ${period}`);

        return data
            .filter(row => row.PolisNr || row.Verzekerde) // Skip if both empty
            .map(row => ({
                'Makelaar': 'AON',
                'Boekingsperiode': period,
                'Polisnr makelaar': row.PolisNr || '',
                'Verzekerde': row.Verzekerde || '',
                'Branche': row.Branche || '',
                'Periode van': row.PeriodeVan || '',
                'Periode tot': row.PeriodeTot || '',
                'Valuta': 'EUR',
                'Bruto': row.Bruto || '',
                'Provisie%': row.ProvisiePercentage || '',
                'Provisie': row.Provisie || '',
                'Tekencom%': row.TekencomPercentage || '',
                'Tekencom': row.Tekencom || '',
                'Netto': row.Netto || '',
                'BAB': row.BAB || '',
                'Land': row.Land || '',
                'Aandeel Allianz': row.AandeelAllianz || '',
                'Tekenjaar': row.Tekenjaar || '',
                'Boekdatum tp': row.BoekDtm || '',
                'FactuurDtm': row.FactuurDtm || '',
                'FactuurNr': row.FactuurNr || '',
                'Boekingsreden': row.FactuurTekst || ''
            }));
    }
}

class VGAParser {
    static detect(filename) {
        return /^VGA (\d{2}-\d{4}) (A\d{3})\.xlsx$/i.test(filename);
    }

    static parse(workbook, filename) {
        console.log('VGA Parser: Processing file', filename);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet, {
            cellFormula: false  // Read calculated values instead of formulas
        });

        const matches = filename.match(/^VGA (\d{2}-\d{4}) (A\d{3})\.xlsx$/i);
        const period = matches ? matches[1] : '';
        const code = matches ? matches[2] : '';

        console.log(`VGA Parser: Found ${data.length} rows, period: ${period}, code: ${code}`);

        return data
            .filter(row => row.Soort !== 'Total' && row.Soort !== 'Totaal') // Exclude summary rows
            .filter(row => row.Polisnummer || row['Naam verzekeringnemer']) // Skip if both empty
            .map(row => ({
                'Makelaar': `VGA ${code}`,
                'Boekingsperiode': period,
                'Polisnr makelaar': row.Polisnummer || '',
                'Verzekerde': row['Naam verzekeringnemer'] || '',
                'Branche': row.Branche || '',
                'Periode van': row.PeriodeVan || '',
                'Periode tot': row.PeriodeTot || '',
                'Valuta': 'EUR',
                'Bruto': row['Bruto premie EB'] || '',
                'Provisie%': row.ProvisiePercentage || '',
                'Provisie': row.Provisie || '',
                'Tekencom%': row.TekencomPercentage || '',
                'Tekencom': row.Tekencom || '',
                'Netto': row['Netto Maatschappij EB'] || '',
                'BAB': row.BAB || '',
                'Land': row.Land || '',
                'Aandeel Allianz': row.AandeelAllianz || '',
                'Tekenjaar': row.Tekenjaar || '',
                'Boekdatum tp': row.BoekDtm || '',
                'FactuurDtm': row.FactuurDtm || '',
                'FactuurNr': row.Factuurnummer || '',
                'Boekingsreden': row.FactuurTekst || ''
            }));
    }
}

class BCIParser {
    static detect(filename) {
        return /^BCI (\d{4})-Q([1-4])\.xlsx$/i.test(filename);
    }

    static parse(workbook, filename) {
        console.log('BCI Parser: Processing file', filename);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        const matches = filename.match(/^BCI (\d{4}-Q[1-4])\.xlsx$/i);
        const period = matches ? matches[1] : '';
        const result = [];

        console.log(`BCI Parser: Processing ${period}, scanning for data start`);

        // Find the actual data start by looking for the policy pattern
        let dataStartRow = -1;
        for (let scanRow = 0; scanRow <= range.e.r; scanRow++) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: scanRow, c: 0 })];
            if (cell && cell.v) {
                const cellValue = cell.v.toString().trim();
                // Look for policy pattern: XXXX.XX.XX.XXXX with optional suffix
                if (/^\d{4}\.\w{1,3}\.\d{2}\.\d{4}(\w*)/.test(cellValue)) {
                    dataStartRow = scanRow;
                    console.log(`BCI Parser: Found data start at row ${scanRow}: "${cellValue}"`);
                    break;
                }
            }
        }

        if (dataStartRow === -1) {
            console.log('BCI Parser: No policy data found in file');
            return result;
        }

        // Process data starting from found row
        for (let row = dataStartRow; row <= range.e.r; row += 2) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
            if (!cell || !cell.v) break;

            const firstCol = cell.v.toString();
            console.log(`BCI Parser: Processing row ${row}: "${firstCol}"`);

            // Check pattern: XXXX.XX.XX.XXXX[suffix] BRANCHE
            const match = firstCol.match(/^(\d{4}\.\w{1,3}\.\d{2}\.\d{4}\w*)\s+(.+)$/);
            if (!match) {
                console.log(`BCI Parser: Pattern not matched, stopping at row ${row}`);
                break; // End of borderel data
            }

            const [, polisnr, branche] = match;

            // Get data from current and next row
            const row1Data = this.getRowData(worksheet, row);
            const row2Data = this.getRowData(worksheet, row + 1);

            result.push({
                'Makelaar': 'BCI',
                'Boekingsperiode': period,
                'Polisnr makelaar': polisnr,
                'Verzekerde': row2Data[0] || '',
                'Branche': branche.trim(),
                'Periode van': '',
                'Periode tot': '',
                'Valuta': 'EUR',
                'Bruto': row2Data[4] || '',
                'Provisie%': '',
                'Provisie': row1Data[3] || '',
                'Tekencom%': '',
                'Tekencom': '',
                'Netto': row1Data[6] || '',
                'BAB': '',
                'Land': '',
                'Aandeel Allianz': '',
                'Tekenjaar': '',
                'Boekdatum tp': '',
                'FactuurDtm': '',
                'FactuurNr': row2Data[1] || '',
                'Boekingsreden': ''
            });
        }

        console.log(`BCI Parser: Processed ${result.length} borderel records`);
        console.log('BCI Parser: First few records:', result.slice(0, 2));
        return result;
    }

    static getRowData(worksheet, rowIndex) {
        const data = [];
        for (let col = 0; col < 10; col++) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: col })];
            data.push(cell ? cell.v : null);
        }
        return data;
    }
}

class VoogtParser {
    static detect(filename) {
        return /^Voogt (\d{2}) (\d{4})\.xlsx$/i.test(filename);
    }

    static parse(workbook, filename) {
        console.log('Voogt Parser: Processing file', filename);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // Use defval to handle merged cells better and get raw data
        const data = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: null,  // Use null for empty cells
            raw: false,    // Convert numbers to strings to avoid Excel date issues
            cellFormula: false  // Read calculated values instead of formulas
        });

        const matches = filename.match(/^Voogt (\d{2}) (\d{4})\.xlsx$/i);
        const period = matches ? `${matches[1]} ${matches[2]}` : '';

        console.log(`Voogt Parser: Found ${data.length} rows, period: ${period}`);
        console.log('Voogt Parser: Worksheet range:', worksheet['!ref']);

        // Debug: Show rows around the data area (15-25) to see the actual structure
        console.log('Voogt Parser: Rows 15-25 for debugging:');
        for (let i = 15; i < Math.min(25, data.length); i++) {
            console.log(`Row ${i} (Excel row ${i + 1}):`, data[i]);
        }

        // Look for header row first (contains "Datum", "Polisnummer", etc.)
        let headerRow = -1;
        for (let i = 0; i < Math.min(10, data.length); i++) {
            const row = data[i];
            if (row && row.length > 0) {
                // Check if this row contains typical Voogt headers
                const rowText = row.join('').toLowerCase();
                if (rowText.includes('datum') && rowText.includes('polisnummer')) {
                    headerRow = i;
                    console.log(`Voogt Parser: Found header row at ${i}:`, row);
                    break;
                }
            }
        }

        // Find first data row (should be right after header row)
        let startRow = -1;
        const searchStart = headerRow >= 0 ? headerRow + 1 : 0;

        for (let i = searchStart; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;

            const cellValue = row[0]; // Column A (Datum)
            console.log(`Voogt Parser: Checking row ${i}, column A value:`, cellValue, `Type: ${typeof cellValue}`);

            // Look for date patterns more broadly
            if (cellValue && this.isVoogtDate(cellValue)) {
                startRow = i;
                console.log(`Voogt Parser: Found data start at row ${i} (Excel row ${i + 1}): ${cellValue}`);
                break;
            }
        }

        if (startRow === -1) {
            console.log('Voogt Parser: No date column found, trying policy number detection...');

            // Alternative: look for rows with policy numbers (should be in column C based on your image)
            for (let i = searchStart; i < data.length; i++) {
                const row = data[i];
                if (row && row.length > 2) {
                    // Look for policy number pattern (AB followed by numbers)
                    const policyNum = row[2]; // Column C
                    if (policyNum && typeof policyNum === 'string' && /^AB\d+/.test(policyNum)) {
                        startRow = i;
                        console.log(`Voogt Parser: Found data start at row ${i} using policy number: ${policyNum}`);
                        break;
                    }
                }
            }
        }

        if (startRow === -1) {
            console.log('Voogt Parser: Still no data found');
            return [];
        }

        const result = [];
        let processedCount = 0;

        console.log(`Voogt Parser: Starting processing from row ${startRow} (Excel row ${startRow + 1})`);
        console.log(`Voogt Parser: Will process up to row ${data.length - 1} (Excel row ${data.length})`);

        for (let i = startRow; i < data.length; i++) {
            const row = data[i];
            processedCount++;
            console.log(`Voogt Parser: [${processedCount}] Processing row ${i} (Excel row ${i + 1}):`, row);

            // Check if we've hit a summary row (contains "Totaal" or similar)
            if (row && row.length > 0) {
                const rowText = row.join('').toLowerCase();
                if (rowText.includes('totaal') || rowText.includes('total')) {
                    console.log(`Voogt Parser: Found summary row at ${i} (Excel row ${i + 1}), stopping`);
                    break;
                }
            }

            // More flexible end-of-data detection
            if (!row || row.length === 0) {
                console.log(`Voogt Parser: Empty row at ${i} (Excel row ${i + 1}), stopping`);
                break;
            }

            // Check if this row has the minimum required data
            const hasDate = row[0] && this.isVoogtDate(row[0]);
            const hasPolicy = row[5] && row[5].toString().trim() !== '';
            const hasDescription = row[10] && row[10].toString().trim() !== '';
            const hasAmount = row[14] && row[14].toString().trim() !== '';

            console.log(`Voogt Parser: Row ${i} data check - Date: "${row[0]}" (${hasDate}), Policy: "${row[5]}" (${hasPolicy}), Description: "${row[10]}" (${hasDescription})`);
            console.log(`Voogt Parser: Row ${i} amounts - Netto:${row[14]}, Provisie:${row[16]}, Bruto:${row[19]}`);

            // Row is valid if it has date AND (policy OR description OR amount)
            if (!hasDate || (!hasPolicy && !hasDescription && !hasAmount)) {
                console.log(`Voogt Parser: No essential data at row ${i} (Excel row ${i + 1}), stopping`);
                break;
            }

            const record = {
                'Makelaar': 'Voogt',
                'Boekingsperiode': period,
                'Polisnr makelaar': (row[5] || '').toString().trim(),
                'Verzekerde': (row[10] || '').toString().trim(),
                'Branche': (row[11] || '').toString().trim(),
                'Periode van': '',
                'Periode tot': '',
                'Valuta': 'EUR',
                'Bruto': this.parseNumber(row[19]) || '',
                'Provisie%': '',
                'Provisie': this.parseNumber(row[16]) || '',
                'Tekencom%': '',
                'Tekencom': '',
                'Netto': this.parseNumber(row[14]) || '',
                'BAB': '',
                'Land': '',
                'Aandeel Allianz': '',
                'Tekenjaar': '',
                'Boekdatum tp': (row[0] || '').toString().trim(),
                'FactuurDtm': '',
                'FactuurNr': '',
                'Boekingsreden': ''
            };

            console.log(`Voogt Parser: Created record for Excel row ${i + 1}:`, record);
            result.push(record);
        }

        console.log(`Voogt Parser: Processed ${result.length} records`);
        return result;
    }

    static isDate(value) {
        if (!value) return false;

        // Check if it's an Excel serial date (number between reasonable bounds)
        if (typeof value === 'number' && value > 1 && value < 100000) {
            console.log(`Voogt Parser: Detected Excel serial date: ${value}`);
            return true;
        }

        // Check if it's a parseable date string
        if (typeof value === 'string') {
            const parsed = Date.parse(value);
            if (!isNaN(parsed)) {
                console.log(`Voogt Parser: Detected date string: ${value}`);
                return true;
            }
        }

        // Check if it's a Date object
        if (value instanceof Date && !isNaN(value.getTime())) {
            console.log(`Voogt Parser: Detected Date object: ${value}`);
            return true;
        }

        return false;
    }

    static isVoogtDate(value) {
        if (!value) return false;

        // Check for DD-MM-YYYY pattern (like "01-01-2024")
        if (typeof value === 'string') {
            const datePattern = /^\d{2}-\d{2}-\d{4}$/;
            if (datePattern.test(value)) {
                console.log(`Voogt Parser: Detected DD-MM-YYYY date: ${value}`);
                return true;
            }
        }

        // Check if it's an Excel serial date
        if (typeof value === 'number' && value > 1 && value < 100000) {
            console.log(`Voogt Parser: Detected Excel serial date: ${value}`);
            return true;
        }

        // Fallback to standard date detection
        return this.isDate(value);
    }

    static parseNumber(value) {
        if (!value) return '';

        // Handle string numbers (European format: 1.234,56 or simple: 1234,56)
        if (typeof value === 'string') {
            const trimmed = value.trim();

            // Handle dash/hyphen as zero or empty
            if (trimmed === '-' || trimmed === '') return '';

            // Check if it contains both dots and commas (European format like 1.234,56)
            if (trimmed.includes('.') && trimmed.includes(',')) {
                // Remove thousand separators (dots) and convert decimal comma to dot
                const cleaned = trimmed.replace(/\./g, '').replace(',', '.');
                const parsed = parseFloat(cleaned);
                return isNaN(parsed) ? '' : parsed;
            }
            // Check if it only contains comma (simple format like 1234,56)
            else if (trimmed.includes(',') && !trimmed.includes('.')) {
                // Just convert decimal comma to dot
                const cleaned = trimmed.replace(',', '.');
                const parsed = parseFloat(cleaned);
                return isNaN(parsed) ? '' : parsed;
            }
            // Check if it only contains dots (could be thousand separators or decimal)
            else if (trimmed.includes('.') && !trimmed.includes(',')) {
                // If there's only one dot and it's near the end (decimal), keep it
                const dotIndex = trimmed.lastIndexOf('.');
                if (dotIndex > trimmed.length - 4) {
                    // Likely decimal separator
                    const parsed = parseFloat(trimmed);
                    return isNaN(parsed) ? '' : parsed;
                } else {
                    // Likely thousand separator, remove it
                    const cleaned = trimmed.replace(/\./g, '');
                    const parsed = parseFloat(cleaned);
                    return isNaN(parsed) ? '' : parsed;
                }
            }
            // No separators, just parse as number
            else {
                const parsed = parseFloat(trimmed);
                return isNaN(parsed) ? '' : parsed;
            }
        }

        // Handle numeric values
        if (typeof value === 'number') {
            return value;
        }

        return '';
    }
}

/**
 * Unified broker detection using keyword-based template matching
 * @param {String} filename - Filename to detect broker type
 * @returns {Promise<Object>} Broker detection result
 */
async function detectBrokerType(filename) {
    console.log(`Detecting broker type for: ${filename}`);

    // 1. Check built-in parsers first (highest priority)
    if (AONParser.detect(filename)) {
        return { type: 'built-in', parser: 'AON', name: 'AON B550' };
    }
    if (VGAParser.detect(filename)) {
        return { type: 'built-in', parser: 'VGA', name: 'VGA' };
    }
    if (BCIParser.detect(filename)) {
        return { type: 'built-in', parser: 'BCI', name: 'BCI' };
    }
    if (VoogtParser.detect(filename)) {
        return { type: 'built-in', parser: 'Voogt', name: 'Voogt' };
    }

    // 2. Check unified file mappings by keyword
    const fileMapping = await window.loadFileMappingByKeyword(filename);
    if (fileMapping) {
        console.log(`Found file mapping: ${fileMapping.name} (keyword: "${fileMapping.matchingKeyword}", method: ${fileMapping.creationMethod})`);
        return {
            type: 'custom',
            parser: 'GenericBrokerParser',
            template: fileMapping,
            name: fileMapping.name
        };
    }


    // 3. Unknown format - suggest creating new template
    return { type: 'unknown', filename, name: 'Unknown Format' };
}

/**
 * Main broker processing function
 * @param {Object} fileData - File data object with file, name, and broker info
 * @returns {Promise} Promise resolving to parse result
 */
async function processBrokerFile(fileData) {
    try {
        const workbook = await ExcelCacheManager.getWorkbook(fileData.file);
        let parsedData = [];

        // Detect broker type (including custom templates) or use override
        const detection = fileData.detectionOverride || await detectBrokerType(fileData.name);
        console.log(`${fileData.detectionOverride ? 'Override' : 'Detected'} broker:`, detection);

        switch (detection.type) {
            case 'built-in':
                parsedData = await processBuiltInBroker(workbook, fileData.name, detection.parser);
                break;

            case 'custom':
                // Handle custom templates from unified file mapping system
                const template = detection.template;

                if (template.parsingConfig && template.columnMapping) {
                    // Use two-step process: extract data with skip rules, then apply simple mapping
                    console.log('Using GenericBrokerParser for data extraction with skip rules from saved template');
                    const tempTemplate = {
                        parsingConfig: template.parsingConfig,
                        columnMapping: {} // Empty mapping - we'll apply the real mapping after extraction
                    };
                    const genericParser = GenericBrokerParser.fromTemplate(tempTemplate);
                    const extractedData = await genericParser.parse(workbook, fileData.name);

                    // Apply the actual mapping using the existing function
                    console.log('Extracted data length:', extractedData ? extractedData.length : 'null/undefined');
                    console.log('Template columnMapping exists:', !!template.columnMapping);
                    console.log('applyMappingToData function exists:', typeof window.applyMappingToData === 'function');

                    if (template.columnMapping && typeof window.applyMappingToData === 'function') {
                        console.log('Applying column mapping with keys:', Object.keys(template.columnMapping));
                        parsedData = window.applyMappingToData(extractedData, template.columnMapping);
                        console.log('Applied mapping result length:', parsedData ? parsedData.length : 'null/undefined');
                    } else {
                        console.log('No mapping to apply, returning extracted data as-is');
                        parsedData = extractedData; // No mapping to apply
                    }
                } else {
                    // Fallback to direct GenericBrokerParser (legacy behavior)
                    console.log('Using direct GenericBrokerParser for custom template (legacy format)');
                    const genericParser = GenericBrokerParser.fromTemplate(template);
                    parsedData = await genericParser.parse(workbook, fileData.name);
                }
                break;

            case 'manual-mapping':
                // Handle manual mapping from Broker Mapping tab
                if (detection.parsingConfig) {
                    // Use GenericBrokerParser for data extraction with skip rules, then apply simple mapping
                    console.log('Using GenericBrokerParser for data extraction with skip rules');
                    const tempTemplate = {
                        parsingConfig: detection.parsingConfig,
                        columnMapping: {} // Empty mapping - we'll apply the real mapping after extraction
                    };
                    const genericParser = GenericBrokerParser.fromTemplate(tempTemplate);
                    const extractedData = await genericParser.parse(workbook, fileData.name);

                    // Apply the actual mapping using the existing function
                    console.log('Manual-mapping: Extracted data length:', extractedData ? extractedData.length : 'null/undefined');
                    console.log('Manual-mapping: detection.mapping exists:', !!detection.mapping);
                    console.log('Manual-mapping: applyMappingToData function exists:', typeof window.applyMappingToData === 'function');

                    if (detection.mapping && typeof window.applyMappingToData === 'function') {
                        console.log('Manual-mapping: Applying column mapping with keys:', Object.keys(detection.mapping));
                        parsedData = window.applyMappingToData(extractedData, detection.mapping);
                        console.log('Manual-mapping: Applied mapping result length:', parsedData ? parsedData.length : 'null/undefined');
                    } else {
                        console.log('Manual-mapping: No mapping to apply, returning extracted data as-is');
                        parsedData = extractedData; // No mapping to apply
                    }
                } else {
                    // Fallback to simple sheet reading (original behavior)
                    console.log('Using simple sheet reading for manual mapping (no parsingConfig)');
                    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
                        cellFormula: false  // Read calculated values instead of formulas
                    });
                    if (detection.mapping && typeof window.applyMappingToData === 'function') {
                        parsedData = window.applyMappingToData(rawData, detection.mapping);
                    } else {
                        throw new Error('Manual mapping data not provided or applyMappingToData function not available');
                    }
                }
                break;

            case 'unknown':
                return {
                    success: false,
                    error: 'Unknown file format',
                    needsTemplate: true,
                    filename: fileData.name,
                    recordCount: 0,
                    data: []
                };

            default:
                throw new Error(`Unsupported detection type: ${detection.type}`);
        }

        return {
            success: true,
            recordCount: parsedData.length,
            data: parsedData,
            brokerInfo: detection
        };
    } catch (error) {
        console.error('Error processing broker file:', error);
        return {
            success: false,
            error: error.message,
            recordCount: 0,
            data: []
        };
    }
}

/**
 * Process file with built-in broker parser
 * @param {Object} workbook - XLSX workbook
 * @param {String} filename - Original filename
 * @param {String} parserType - Parser type (AON, VGA, BCI, Voogt)
 * @returns {Array} Parsed data
 */
async function processBuiltInBroker(workbook, filename, parserType) {
    switch (parserType) {
        case 'AON':
            return AONParser.parse(workbook, filename);
        case 'VGA':
            return VGAParser.parse(workbook, filename);
        case 'BCI':
            return BCIParser.parse(workbook, filename);
        case 'Voogt':
            return VoogtParser.parse(workbook, filename);
        default:
            throw new Error(`Unknown built-in parser: ${parserType}`);
    }
}

// isRowCompletelyEmpty function removed - compaction logic now in ExcelCacheManager

// removeEmptyRowsFromWorksheet function removed - compaction now handled exclusively by ExcelCacheManager

// readExcelFile function removed - now using centralized ExcelCacheManager

// Export functions globally for cross-module access
window.detectBrokerType = detectBrokerType;
window.processBrokerFile = processBrokerFile;
// window.readExcelFile removed - use ExcelCacheManager.getWorkbook() instead