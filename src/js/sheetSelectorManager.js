/**
 * Borderellen Converter - Sheet Selector Manager
 * Allows users to select which worksheet to use when Excel file has multiple sheets
 */

class SheetSelectorManager {
    /**
     * Show sheet selector modal and return user's choice
     * @param {Array} sheetNames - Array of sheet names
     * @param {string} filename - Filename for context
     * @returns {Promise<Object>} Selected sheet info: {sheetName: string, pattern: string|null}
     */
    static async selectSheet(sheetNames, filename) {
        return new Promise((resolve, reject) => {
            // Create modal overlay
            const modal = document.createElement('div');
            modal.id = 'sheet-selector-modal';
            modal.style.cssText = `
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0,0,0,0.7);
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: 10000;
            `;

            // Detect pattern in sheet names (e.g., Q2_2025, Q3_2025)
            const pattern = this.detectSheetPattern(sheetNames);
            const suggestedSheet = pattern ? this.findSheetByPattern(sheetNames, pattern.regex) : sheetNames[0];

            modal.innerHTML = `
                <div style="background: white; padding: 30px; border-radius: 8px; max-width: 600px; width: 90%;">
                    <h2 style="margin-top: 0; color: #333;">Select Worksheet</h2>
                    <p style="color: #666;">
                        This Excel file contains <strong>${sheetNames.length} worksheets</strong>.
                        Select which sheet contains the data to process.
                    </p>

                    <div style="margin: 20px 0;">
                        <strong style="display: block; margin-bottom: 10px;">File: ${filename}</strong>

                        <div style="margin-top: 15px;">
                            <label style="display: block; margin-bottom: 8px; font-weight: bold;">Select Sheet:</label>
                            <select id="sheet-name-select" style="width: 100%; padding: 10px; font-size: 14px; border: 2px solid #ddd; border-radius: 4px;">
                                ${sheetNames.map(name => `
                                    <option value="${name}" ${name === suggestedSheet ? 'selected' : ''}>
                                        ${name}${name === suggestedSheet && pattern ? ' (matched pattern)' : ''}
                                    </option>
                                `).join('')}
                            </select>
                        </div>

                        ${pattern ? `
                            <div style="margin-top: 20px; padding: 15px; background: #f0f7ff; border-left: 4px solid #0066cc; border-radius: 4px;">
                                <strong style="color: #0066cc;">Pattern Detected:</strong>
                                <p style="margin: 5px 0 0 0; color: #666;">
                                    Sheet names follow pattern: <code style="background: white; padding: 2px 6px; border-radius: 3px;">${pattern.description}</code>
                                </p>
                                <label style="display: block; margin-top: 10px;">
                                    <input type="checkbox" id="use-pattern-checkbox" ${pattern ? 'checked' : ''}>
                                    <span>Use pattern for future files (e.g., Q1_2025, Q2_2025, Q3_2026)</span>
                                </label>
                            </div>
                        ` : `
                            <div style="margin-top: 20px; padding: 15px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                                <strong style="color: #856404;">No pattern detected</strong>
                                <p style="margin: 5px 0 0 0; color: #856404;">
                                    This selection will only apply to this file. Consider using consistent sheet names for automation.
                                </p>
                            </div>
                        `}
                    </div>

                    <div style="display: flex; gap: 10px; justify-content: flex-end; margin-top: 25px;">
                        <button id="cancel-sheet-btn" style="padding: 10px 20px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">
                            Cancel
                        </button>
                        <button id="confirm-sheet-btn" style="padding: 10px 20px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: bold;">
                            Confirm Selection
                        </button>
                    </div>
                </div>
            `;

            document.body.appendChild(modal);

            // Event listeners
            const confirmBtn = document.getElementById('confirm-sheet-btn');
            const cancelBtn = document.getElementById('cancel-sheet-btn');
            const selectElement = document.getElementById('sheet-name-select');
            const patternCheckbox = document.getElementById('use-pattern-checkbox');

            confirmBtn.addEventListener('click', () => {
                const selectedSheet = selectElement.value;
                const usePattern = patternCheckbox ? patternCheckbox.checked : false;

                modal.remove();
                resolve({
                    sheetName: selectedSheet,
                    pattern: usePattern && pattern ? pattern.regex : null,
                    patternDescription: usePattern && pattern ? pattern.description : null
                });
            });

            cancelBtn.addEventListener('click', () => {
                modal.remove();
                reject(new Error('Sheet selection cancelled by user'));
            });

            // Close on overlay click
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    modal.remove();
                    reject(new Error('Sheet selection cancelled by user'));
                }
            });
        });
    }

    /**
     * Detect common patterns in sheet names
     * @param {Array} sheetNames - Array of sheet names
     * @returns {Object|null} Pattern info with regex and description
     */
    static detectSheetPattern(sheetNames) {
        // Pattern: Q1_2025, Q2_2025, Q3_2025, Q4_2025
        const quarterPattern = /^Q[1-4]_\d{4}$/;
        const quarterMatch = sheetNames.find(name => quarterPattern.test(name));
        if (quarterMatch) {
            return {
                regex: '^Q[1-4]_\\d{4}$',
                description: 'Q1_2025, Q2_2025, Q3_2025, etc.'
            };
        }

        // Pattern: 2025_Q1, 2025_Q2, etc.
        const yearQuarterPattern = /^\d{4}_Q[1-4]$/;
        const yearQuarterMatch = sheetNames.find(name => yearQuarterPattern.test(name));
        if (yearQuarterMatch) {
            return {
                regex: '^\\d{4}_Q[1-4]$',
                description: '2025_Q1, 2025_Q2, etc.'
            };
        }

        // Pattern: Jan 2025, Feb 2025, etc.
        const monthYearPattern = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/i;
        const monthYearMatch = sheetNames.find(name => monthYearPattern.test(name));
        if (monthYearMatch) {
            return {
                regex: '^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+\\d{4}$',
                description: 'Jan 2025, Feb 2025, etc.'
            };
        }

        // Pattern: 2025-01, 2025-02, etc.
        const yearMonthPattern = /^\d{4}-\d{2}$/;
        const yearMonthMatch = sheetNames.find(name => yearMonthPattern.test(name));
        if (yearMonthMatch) {
            return {
                regex: '^\\d{4}-\\d{2}$',
                description: '2025-01, 2025-02, etc.'
            };
        }

        // No pattern detected
        return null;
    }

    /**
     * Find sheet name matching a regex pattern
     * @param {Array} sheetNames - Array of sheet names
     * @param {string} regexPattern - Regex pattern as string
     * @returns {string|null} Matching sheet name or null
     */
    static findSheetByPattern(sheetNames, regexPattern) {
        try {
            const regex = new RegExp(regexPattern);
            return sheetNames.find(name => regex.test(name)) || null;
        } catch (error) {
            console.error('Invalid regex pattern:', regexPattern, error);
            return null;
        }
    }
}

// Export globally
window.SheetSelectorManager = SheetSelectorManager;
