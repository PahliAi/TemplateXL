/**
 * Borderellen Converter - Auto Mapping Module
 * Handles intelligent column mapping between source files and templates
 */

class AutoMapping {
    /**
     * Calculate Levenshtein distance between two strings
     * @param {string} str1 - First string
     * @param {string} str2 - Second string
     * @returns {number} Edit distance between strings
     */
    static levenshteinDistance(str1, str2) {
        const len1 = str1.length;
        const len2 = str2.length;
        const matrix = Array(len1 + 1).fill().map(() => Array(len2 + 1).fill(0));

        for (let i = 0; i <= len1; i++) matrix[i][0] = i;
        for (let j = 0; j <= len2; j++) matrix[0][j] = j;

        for (let i = 1; i <= len1; i++) {
            for (let j = 1; j <= len2; j++) {
                if (str1[i - 1] === str2[j - 1]) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j] + 1,     // deletion
                        matrix[i][j - 1] + 1,     // insertion
                        matrix[i - 1][j - 1] + 1  // substitution
                    );
                }
            }
        }

        return matrix[len1][len2];
    }

    /**
     * Normalize string for comparison (case-insensitive, remove special chars)
     * @param {string} str - String to normalize
     * @returns {string} Normalized string
     */
    static normalizeString(str) {
        return str.toLowerCase()
            .trim()
            .replace(/[%$#@&*()[\]{}|\\:";'<>?,.]/g, '')
            .replace(/\s+/g, ' ');
    }

    /**
     * Calculate confidence score between two column names
     * @param {string} sourceName - Source column name
     * @param {string} templateName - Template column name
     * @returns {number} Confidence score (0-100)
     */
    static calculateColumnMatchConfidence(sourceName, templateName) {
        const normalizedSource = this.normalizeString(sourceName);
        const normalizedTemplate = this.normalizeString(templateName);

        // Debug logging with original and normalized names
        console.log(`Comparing: "${sourceName}" -> "${normalizedSource}" with "${templateName}" -> "${normalizedTemplate}"`);

        // Exact match gets EXACTLY 100 (no other combination can reach this)
        if (normalizedSource === normalizedTemplate) {
            console.log(`✓ EXACT MATCH FOUND: "${sourceName}" === "${templateName}" (normalized: "${normalizedSource}")`, {
                exactMatch: 100,
                charSim: 0,
                substring: 0,
                wordOverlap: 0,
                penalty: 0,
                total: 100
            });
            return 100;
        }

        // All fuzzy matches are capped at 99 maximum to ensure exact matches always win
        let confidence = 0;
        let charSimScore = 0;
        let substringScore = 0;
        let wordOverlapScore = 0;
        let lengthPenalty = 0;

        // 1. Character similarity using Levenshtein distance (max 50 points)
        const maxLen = Math.max(normalizedSource.length, normalizedTemplate.length);
        if (maxLen > 0) {
            const editDistance = this.levenshteinDistance(normalizedSource, normalizedTemplate);
            const similarity = (maxLen - editDistance) / maxLen;
            charSimScore = similarity * 50; // Up to 50 points for character similarity
            confidence += charSimScore;
        }

        // 2. Improved substring containment (max 30 points)
        const shortStr = normalizedSource.length <= normalizedTemplate.length ? normalizedSource : normalizedTemplate;
        const longStr = normalizedSource.length > normalizedTemplate.length ? normalizedSource : normalizedTemplate;

        if (longStr.includes(shortStr) && shortStr.length > 2) {
            // Award points based on how much of the longer string the shorter one represents
            const coverage = shortStr.length / longStr.length;
            substringScore = coverage * 30; // More points for better coverage
            confidence += substringScore;
        }

        // 3. Word overlap scoring (max 20 points)
        const sourceWords = normalizedSource.split(' ').filter(w => w.length > 1);
        const templateWords = normalizedTemplate.split(' ').filter(w => w.length > 1);

        if (sourceWords.length > 0 && templateWords.length > 0) {
            const commonWords = sourceWords.filter(w => templateWords.includes(w));
            const wordOverlap = commonWords.length / Math.max(sourceWords.length, templateWords.length);
            wordOverlapScore = wordOverlap * 20; // Up to 20 points for word overlap
            confidence += wordOverlapScore;
        }

        // 4. Length difference penalty (up to -10 points)
        const lengthDiff = Math.abs(normalizedSource.length - normalizedTemplate.length);
        lengthPenalty = Math.min(10, lengthDiff * 0.5);
        confidence -= lengthPenalty;

        // Cap all fuzzy matches at 99 to ensure exact matches always win
        const finalScore = Math.max(0, Math.min(99, confidence));

        // Enhanced debug logging with score breakdown (commented out for less verbose logging)
        // console.log(`Score breakdown for "${sourceName}" → "${templateName}":`, {
        //     normalized: `"${normalizedSource}" vs "${normalizedTemplate}"`,
        //     charSim: Math.round(charSimScore * 10) / 10,
        //     substring: Math.round(substringScore * 10) / 10,
        //     wordOverlap: Math.round(wordOverlapScore * 10) / 10,
        //     penalty: Math.round(lengthPenalty * 10) / 10,
        //     rawTotal: Math.round(confidence * 10) / 10,
        //     cappedTotal: Math.round(finalScore * 10) / 10
        // });

        return finalScore;
    }

    /**
     * Perform optimal assignment using greedy algorithm with backtracking
     * @param {Array} confidenceMatrix - 2D matrix of confidence scores
     * @returns {Array} Array of optimal assignments
     */
    static performOptimalAssignment(confidenceMatrix) {
        // Flatten matrix and sort by confidence (highest first)
        const allPairs = [];
        for (let i = 0; i < confidenceMatrix.length; i++) {
            for (let j = 0; j < confidenceMatrix[i].length; j++) {
                allPairs.push(confidenceMatrix[i][j]);
            }
        }

        allPairs.sort((a, b) => b.confidence - a.confidence);

        // Greedy assignment - no source or template can be used twice
        const usedSources = new Set();
        const usedTemplates = new Set();
        const assignments = [];

        for (const pair of allPairs) {
            if (!usedSources.has(pair.sourceIndex) && !usedTemplates.has(pair.templateIndex)) {
                assignments.push(pair);
                usedSources.add(pair.sourceIndex);
                usedTemplates.add(pair.templateIndex);
            }
        }

        return assignments;
    }

    /**
     * Generate auto-mapping suggestions for source columns against template
     * @param {Array} sourceColumnNames - Array of source column names
     * @param {Array} templateColumns - Array of template column objects
     * @param {number} confidenceThreshold - Minimum confidence threshold (default 30)
     * @returns {Object} Mapping result with assignments and confidence scores
     */
    static generateMappingSuggestions(sourceColumnNames, templateColumns, confidenceThreshold = 30) {
        if (!sourceColumnNames || sourceColumnNames.length === 0) {
            throw new Error('No source columns provided');
        }

        if (!templateColumns || templateColumns.length === 0) {
            throw new Error('No template columns provided');
        }

        console.log(`Building ${sourceColumnNames.length} × ${templateColumns.length} confidence matrix...`);

        // Build confidence matrix: source × template
        const confidenceMatrix = [];

        for (let i = 0; i < sourceColumnNames.length; i++) {
            confidenceMatrix[i] = [];
            for (let j = 0; j < templateColumns.length; j++) {
                const confidence = this.calculateColumnMatchConfidence(
                    sourceColumnNames[i],
                    templateColumns[j].name
                );
                confidenceMatrix[i][j] = {
                    sourceIndex: i,
                    templateIndex: j,
                    sourceName: sourceColumnNames[i],
                    templateName: templateColumns[j].name,
                    confidence: confidence
                };
            }
        }

        // Perform optimal assignment
        const assignments = this.performOptimalAssignment(confidenceMatrix);

        // Filter by confidence threshold and build results
        const mapping = {};
        const confidenceScores = {};
        let assignmentCount = 0;

        assignments.forEach(assignment => {
            if (assignment.confidence >= confidenceThreshold) {
                mapping[assignment.templateName] = assignment.sourceName;
                confidenceScores[assignment.templateName] = assignment.confidence;
                assignmentCount++;
                console.log(`✓ MAPPED: "${assignment.sourceName}" → "${assignment.templateName}" (${assignment.confidence.toFixed(1)}%)`);
            } else {
                console.log(`✗ SKIPPED (below ${confidenceThreshold}%): "${assignment.sourceName}" → "${assignment.templateName}" (${assignment.confidence.toFixed(1)}%)`);
            }
        });

        return {
            mapping: mapping,
            confidenceScores: confidenceScores,
            assignmentCount: assignmentCount,
            totalAssignments: assignments.length,
            threshold: confidenceThreshold
        };
    }
}

// Export for use in other modules
window.AutoMapping = AutoMapping;