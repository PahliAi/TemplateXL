/**
 * Borderellen Converter - Universal Calculation Engine
 * Professional-grade formula parser with bulletproof function detection
 */

// ========== FUNCTION REGISTRY ==========

/**
 * Definitive list of all supported functions
 * No sketchy includes('(') detection - only these are valid!
 */
const VALID_FUNCTIONS = new Set([
    // Text Functions
    'LEFT', 'RIGHT', 'MID', 'FIND', 'REGEX', 'SPLIT',
    'TRIM', 'UPPER', 'LOWER', 'REPLACE', 'CONCAT',
    'CONTAINS', 'STARTSWITH', 'ENDSWITH', 'LENGTH', 'ISEMPTY',

    // Math Functions
    'ROUND', 'ABS', 'MIN', 'MAX', 'CEILING', 'FLOOR',

    // Logic Functions
    'IF', 'AND', 'OR', 'NOT'
]);

/**
 * Bulletproof function detection - no more sketchy includes('(') hacks!
 * @param {string} text - Text to check
 * @returns {boolean} True if it's a valid function call
 */
function isValidFunction(text) {
    if (!text || typeof text !== 'string') return false;

    const functionMatch = text.match(/^([A-Z]+)\(/);
    if (!functionMatch) return false;

    const functionName = functionMatch[1];
    return VALID_FUNCTIONS.has(functionName);
}

/**
 * Extract function name from function call
 * @param {string} text - Function call like "LEFT(text, 5)"
 * @returns {string|null} Function name or null
 */
function extractFunctionName(text) {
    const match = text.match(/^([A-Z]+)\(/);
    return match ? match[1] : null;
}

/**
 * Check if text is a parenthetical expression (not a function)
 * @param {string} text - Text to check
 * @returns {boolean} True if it's like "(A + B)" not "FUNC(A, B)"
 */
function isParenthetical(text) {
    if (!text || typeof text !== 'string') return false;
    return text.startsWith('(') && text.endsWith(')') && !isValidFunction(text);
}

/**
 * Smart operand resolution - handles functions, parentheses, columns, and numbers
 * @param {string} operand - The operand to resolve
 * @param {Object} rowData - Data context
 * @returns {any} Resolved value
 */
function resolveOperand(operand, rowData) {
    if (!operand || typeof operand !== 'string') return '';

    const trimmed = operand.trim();

    // 1. Valid function call
    if (isValidFunction(trimmed)) {
        console.log(`Resolving function: ${trimmed}`);
        return executeFormula(trimmed, rowData);
    }

    // 2. Parenthetical expression like (A + B)
    if (isParenthetical(trimmed)) {
        console.log(`Resolving parenthetical: ${trimmed}`);
        return executeFormula(trimmed.slice(1, -1), rowData);
    }

    // 3. Column reference
    if (rowData.hasOwnProperty(trimmed)) {
        console.log(`Resolving column: ${trimmed} = ${rowData[trimmed]}`);
        return rowData[trimmed];
    }

    // 4. Number (including European format)
    const numberValue = parseEuropeanNumber(trimmed);
    if (!isNaN(numberValue)) {
        console.log(`Resolving number: ${trimmed} = ${numberValue}`);
        return numberValue;
    }

    // 5. Fallback - treat as literal string
    console.log(`Resolving literal: ${trimmed}`);
    return trimmed;
}

/**
 * Enhanced European number parser
 * @param {any} value - Value to parse
 * @returns {number} Parsed number or NaN
 */
function parseEuropeanNumber(value) {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string') return parseFloat(value) || NaN;

    // Convert European format (1.234,56) to US format (1234.56)
    const cleaned = value.replace(/\./g, '').replace(',', '.');
    const result = parseFloat(cleaned);
    return result;
}

/**
 * Parse formula into segments using bracket-counting approach
 * Handles nested functions and complex expressions properly
 * @param {string} formula - The formula to parse
 * @returns {Array} Array of segments (functions, operators, operands)
 */
function parseFormulaSegments(formula) {
    const segments = [];
    let current = '';
    let bracketCount = 0;
    let inQuotes = false;
    let quoteChar = '';

    console.log('SEGMENT PARSING:', formula);

    for (let i = 0; i < formula.length; i++) {
        const char = formula[i];

        // Handle quotes
        if (!inQuotes && (char === '"' || char === "'")) {
            inQuotes = true;
            quoteChar = char;
            current += char;
        } else if (inQuotes && char === quoteChar) {
            inQuotes = false;
            current += char;
        } else if (!inQuotes && char === '(') {
            bracketCount++;
            current += char;
        } else if (!inQuotes && char === ')') {
            bracketCount--;
            current += char;

            // When bracket count hits 0, we have a complete segment
            if (bracketCount === 0 && current.trim()) {
                segments.push(current.trim());
                current = '';
            }
        } else if (!inQuotes && bracketCount === 0 && /[+\-×÷\/\*><=]/.test(char)) {
            // Operator at top level - complete current segment and add operator
            if (current.trim()) {
                segments.push(current.trim());
                current = '';
            }

            // Handle multi-character operators
            if ((char === '>' || char === '<' || char === '!' || char === '=') && i < formula.length - 1) {
                const nextChar = formula[i + 1];
                if (nextChar === '=' || (char === '!' && nextChar === '=')) {
                    segments.push(char + nextChar);
                    i++; // Skip next character
                    continue;
                }
            }

            segments.push(char);
        } else {
            current += char;
        }
    }

    // Add any remaining content
    if (current.trim()) {
        segments.push(current.trim());
    }

    console.log('PARSED SEGMENTS:', segments);
    return segments;
}

/**
 * Formula Parser and Calculator Engine
 * Executes CALC: formulas during data processing
 */
function executeFormula(formula, rowData) {
    try {
        console.log('executeFormula called with:', formula);

        // Use segment parsing to determine formula type
        const segments = parseFormulaSegments(formula);

        // Single segment - could be function call, column reference, or literal
        if (segments.length === 1) {
            const segment = segments[0];

            // Check if it's a single function call
            const functionMatch = segment.match(/^([A-Z]+)\((.+)\)$/);
            if (functionMatch) {
                const functionName = functionMatch[1];
                const parameters = parseParameters(functionMatch[2]);
                console.log(`Single function call: ${functionName} with params:`, parameters);
                return executeFunction(functionName, parameters, rowData);
            }

            // Direct column reference or literal
            return rowData[segment] || segment;
        }

        // Multiple segments - mathematical expression
        if (segments.length > 1) {
            console.log('Multi-segment expression detected, using executeBasicFormula');
            return executeBasicFormula(formula, rowData);
        }

        // Fallback
        return rowData[formula] || '';
    } catch (error) {
        console.warn('Formula execution error:', error, 'Formula:', formula);
        return '';
    }
}

/**
 * Parse function parameters handling nested functions and quoted strings
 */
function parseParameters(paramString) {
    const params = [];
    let current = '';
    let depth = 0;
    let inQuotes = false;
    let quoteChar = '';

    for (let i = 0; i < paramString.length; i++) {
        const char = paramString[i];

        if (!inQuotes && (char === '"' || char === "'")) {
            inQuotes = true;
            quoteChar = char;
            current += char;
        } else if (inQuotes && char === quoteChar) {
            inQuotes = false;
            current += char;
        } else if (!inQuotes && char === '(') {
            depth++;
            current += char;
        } else if (!inQuotes && char === ')') {
            depth--;
            current += char;
        } else if (!inQuotes && (char === ',' || char === ';') && depth === 0) {
            params.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }

    if (current.trim()) {
        params.push(current.trim());
    }

    return params;
}

/**
 * Execute specific function with parameters
 */
function executeFunction(functionName, parameters, rowData) {
    const getValue = (param) => {
        console.log(`getValue called with param: "${param}"`);

        // Remove quotes from string literals
        if ((param.startsWith('"') && param.endsWith('"')) ||
            (param.startsWith("'") && param.endsWith("'"))) {
            const literalValue = param.slice(1, -1);
            console.log(`Returning literal value: "${literalValue}"`);
            return literalValue;
        }

        // Use bulletproof operand resolution for everything else
        return resolveOperand(param, rowData);
    };

    const getNumericValue = (param) => {
        const value = getValue(param);
        const num = parseFloat(value);
        return isNaN(num) ? 0 : num;
    };

    switch (functionName.toUpperCase()) {
        // Text Extraction Functions
        case 'LEFT':
            const leftText = getValue(parameters[0]).toString();
            const leftCount = getNumericValue(parameters[1]);
            return leftText.substring(0, leftCount);

        case 'RIGHT':
            const rightText = getValue(parameters[0]).toString();
            const rightCount = getNumericValue(parameters[1]);
            return rightText.substring(Math.max(0, rightText.length - rightCount));

        case 'MID':
            const midText = getValue(parameters[0]).toString();
            const start = (getNumericValue(parameters[1]) || 1) - 1; // Convert to 0-based
            const length = getNumericValue(parameters[2]) || 1;
            return midText.substring(start, start + length);

        case 'FIND':
            const searchText = getValue(parameters[0]).toString();
            const targetText = getValue(parameters[1]).toString();
            const startPos = parameters[2] ? (getNumericValue(parameters[2]) || 1) - 1 : 0; // Convert to 0-based
            const foundIndex = targetText.indexOf(searchText, startPos);
            return foundIndex >= 0 ? foundIndex + 1 : 0; // Convert back to 1-based, 0 if not found

        case 'REGEX':
            const regexText = getValue(parameters[0]).toString();
            let pattern = getValue(parameters[1]);
            const group = getNumericValue(parameters[2]) || 1;
            try {
                // Handle escaped backslashes in pattern (\\d becomes \d)
                pattern = pattern.replace(/\\\\/g, '\\');
                const regex = new RegExp(pattern);
                const match = regexText.match(regex);
                return match && match[group] ? match[group] : '';
            } catch (error) {
                console.warn('Regex error:', error, 'Pattern:', pattern);
                return '';
            }

        case 'SPLIT':
            const splitText = getValue(parameters[0]).toString();
            const delimiter = getValue(parameters[1]);
            const index = (getNumericValue(parameters[2]) || 1) - 1; // Convert to 0-based
            const parts = splitText.split(delimiter);
            return parts[index] || '';

        // Text Manipulation Functions
        case 'TRIM':
            return getValue(parameters[0]).toString().trim();

        case 'UPPER':
            return getValue(parameters[0]).toString().toUpperCase();

        case 'LOWER':
            return getValue(parameters[0]).toString().toLowerCase();

        case 'REPLACE':
            const replaceText = getValue(parameters[0]).toString();
            const findText = getValue(parameters[1]);
            const replaceWith = getValue(parameters[2]);
            return replaceText.replace(new RegExp(findText, 'g'), replaceWith);

        case 'CONCAT':
            return parameters.map(p => getValue(p).toString()).join('');

        // Text Analysis Functions
        case 'CONTAINS':
            const containsText = getValue(parameters[0]).toString();
            const searchFor = getValue(parameters[1]);
            return containsText.includes(searchFor);

        case 'STARTSWITH':
            const startsText = getValue(parameters[0]).toString();
            const prefix = getValue(parameters[1]);
            return startsText.startsWith(prefix);

        case 'ENDSWITH':
            const endsText = getValue(parameters[0]).toString();
            const suffix = getValue(parameters[1]);
            return endsText.endsWith(suffix);

        case 'LENGTH':
            return getValue(parameters[0]).toString().length;

        case 'ISEMPTY':
            const value = getValue(parameters[0]);
            return value === '' || value === null || value === undefined;

        // Math Functions
        case 'ROUND':
            const roundValue = getNumericValue(parameters[0]);
            const decimals = getNumericValue(parameters[1]) || 2;
            return parseFloat(roundValue.toFixed(decimals));

        case 'ABS':
            return Math.abs(getNumericValue(parameters[0]));

        case 'MIN':
            const minValues = parameters.map(p => getNumericValue(p));
            return Math.min(...minValues);

        case 'MAX':
            const maxValues = parameters.map(p => getNumericValue(p));
            return Math.max(...maxValues);

        case 'CEILING':
            return Math.ceil(getNumericValue(parameters[0]));

        case 'FLOOR':
            return Math.floor(getNumericValue(parameters[0]));

        // Logic Functions
        case 'IF':
            const condition = parameters[0];
            const trueValue = parameters[1];
            const falseValue = parameters[2];

            // Use bulletproof formula execution for condition evaluation
            const conditionResult = executeFormula(condition, rowData);
            return conditionResult ? getValue(trueValue) : getValue(falseValue);

        case 'AND':
            return parameters.every(p => evaluateCondition(p, rowData));

        case 'OR':
            return parameters.some(p => evaluateCondition(p, rowData));

        case 'NOT':
            return !evaluateCondition(parameters[0], rowData);

        default:
            console.warn('Unknown function:', functionName);
            return '';
    }
}

/**
 * Evaluate logical conditions for IF, AND, OR functions
 */
function evaluateCondition(condition, rowData) {
    try {
        // Handle comparison operators
        if (condition.includes(' = ')) {
            const [left, right] = condition.split(' = ');
            return getValue(left.trim(), rowData) == getValue(right.trim(), rowData);
        }
        if (condition.includes(' ≠ ')) {
            const [left, right] = condition.split(' ≠ ');
            return getValue(left.trim(), rowData) != getValue(right.trim(), rowData);
        }
        if (condition.includes(' > ')) {
            const [left, right] = condition.split(' > ');
            return parseFloat(getValue(left.trim(), rowData)) > parseFloat(getValue(right.trim(), rowData));
        }
        if (condition.includes(' < ')) {
            const [left, right] = condition.split(' < ');
            return parseFloat(getValue(left.trim(), rowData)) < parseFloat(getValue(right.trim(), rowData));
        }

        // Handle function calls in conditions
        if (condition.includes('(')) {
            return executeFormula(condition, rowData);
        }

        // Simple truthiness test
        const value = getValue(condition, rowData);
        return value && value !== '' && value !== '0' && value !== 0;
    } catch (error) {
        console.warn('Condition evaluation error:', error);
        return false;
    }
}


/**
 * Execute mathematical expressions with proper multi-operand support
 * Handles: A + B + C, (A + B) * C, FUNC(x) + Y, etc.
 */
function executeBasicFormula(formula, rowData) {
    console.log('executeBasicFormula called with:', formula);
    console.log('rowData:', rowData);

    // Use our new bracket-counting segment parser
    const segments = parseFormulaSegments(formula);
    console.log('Formula segments:', segments);

    if (segments.length === 0) return '';
    if (segments.length === 1) {
        // Single operand - resolve and return
        const result = resolveOperand(segments[0], rowData);
        const numResult = parseEuropeanNumber(result);
        return isNaN(numResult) ? result : numResult;
    }

    // Evaluate the segmented expression (left-to-right)
    let result = resolveOperand(segments[0], rowData);
    result = parseEuropeanNumber(result);
    if (isNaN(result)) result = 0;

    for (let i = 1; i < segments.length; i += 2) {
        const operator = segments[i];
        const operand = segments[i + 1];

        if (!operand) break;

        const operandResolved = resolveOperand(operand, rowData);
        let operandValue = parseEuropeanNumber(operandResolved);
        if (isNaN(operandValue)) operandValue = 0;

        switch (operator) {
            case '+': result += operandValue; break;
            case '-': result -= operandValue; break;
            case '×':
            case '*': result *= operandValue; break;
            case '÷':
            case '/': result = operandValue !== 0 ? result / operandValue : 0; break;
            case '>': return result > operandValue;
            case '<': return result < operandValue;
            case '>=': return result >= operandValue;
            case '<=': return result <= operandValue;
            case '=':
            case '==': return result == operandValue;
            case '!=':
            case '≠': return result != operandValue;
        }
    }

    console.log('Expression result:', result);

    // Return appropriate type
    if (typeof result === 'boolean') {
        return result;
    } else if (typeof result === 'number') {
        return parseFloat(result.toFixed(4));
    } else {
        return result;
    }
}



// Export functions globally for cross-module access
window.executeFormula = executeFormula;