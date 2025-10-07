/**
 * Borderellen Converter - Custom Broker Template Manager
 * Manages CRUD operations for custom broker templates and JSON import/export
 */

class CustomBrokerTemplateManager {
    /**
     * Saves a custom broker template to IndexedDB and exports JSON backup
     * @param {Object} template - Template object to save
     * @returns {Promise<Object>} Saved template
     */
    static async saveTemplate(template) {
        // Validate template structure
        this.validateTemplate(template);

        // Check for duplicate template names if this is a new template (no existing id)
        if (!template.id && template.name) {
            const existingTemplates = await this.getAllTemplates();
            const duplicateName = existingTemplates.find(t => t.name === template.name);

            if (duplicateName) {
                throw new Error(`A template named "${template.name}" already exists. Please choose a different name.`);
            }
        }

        // Check for duplicate keywords if keyword is provided
        if (template.matchingKeyword && template.matchingKeyword.trim()) {
            const existingTemplates = await this.getAllTemplates();
            const duplicateKeyword = existingTemplates.find(t =>
                (!template.id || t.id !== template.id) && // Exclude current template if updating
                t.matchingKeyword &&
                t.matchingKeyword.toLowerCase() === template.matchingKeyword.trim().toLowerCase()
            );

            if (duplicateKeyword) {
                throw new Error(`A template with keyword "${template.matchingKeyword.trim()}" already exists (Template: "${duplicateKeyword.name}"). Keywords must be unique to ensure predictable template selection.`);
            }
        }

        // Ensure template has required metadata
        const completeTemplate = {
            ...template,
            id: template.id || `custom-${Date.now()}`,
            version: template.version || '1.0',
            created: template.created || new Date().toISOString(),
            lastModified: new Date().toISOString(),
            matchingKeyword: template.matchingKeyword || '' // Add keyword field for file matching
        };

        try {
            // Convert to unified file mapping format
            const fileMapping = {
                ...completeTemplate,
                id: completeTemplate.id || `mapping-${Date.now()}`,
                creationMethod: 'template-builder',
                sourceType: completeTemplate.brokerType || 'Unknown',
                sourceName: completeTemplate.brokerName || 'Custom Template',
                parsingConfig: completeTemplate.parsingConfig || {},
                columnMapping: completeTemplate.mapping || {}
            };

            // Store in unified IndexedDB store
            await window.saveFileMapping(fileMapping);

            // Also export as JSON file for backup
            this.exportTemplateAsJSON(completeTemplate);

            console.log('Custom broker template saved to unified store:', fileMapping.id);
            return completeTemplate;
        } catch (error) {
            console.error('Error saving custom broker template:', error);
            throw error;
        }
    }



    /**
     * Gets all saved custom broker templates
     * @returns {Promise<Array>} Array of all templates
     */
    static async getAllTemplates() {
        try {
            return await window.loadAllFileMappings();
        } catch (error) {
            console.error('Error loading all templates:', error);
            return [];
        }
    }

    /**
     * Gets a specific template by ID
     * @param {String} templateId - Template ID to find
     * @returns {Promise<Object|null>} Template or null if not found
     */
    static async getTemplateById(templateId) {
        try {
            const mappings = await window.loadAllFileMappings();
            return mappings.find(m => m.id === templateId) || null;
        } catch (error) {
            console.error('Error loading template by ID:', error);
            return null;
        }
    }


    /**
     * Updates an existing template
     * @param {String} templateId - Template ID to update
     * @param {Object} updates - Partial template object with updates
     * @returns {Promise<Object>} Updated template
     */
    static async updateTemplate(templateId, updates) {
        try {
            const existingTemplate = await this.getTemplateById(templateId);
            if (!existingTemplate) {
                throw new Error(`Template with ID ${templateId} not found`);
            }

            const updatedTemplate = {
                ...existingTemplate,
                ...updates,
                lastModified: new Date().toISOString(),
                version: this.incrementVersion(existingTemplate.version)
            };

            return await this.saveTemplate(updatedTemplate);
        } catch (error) {
            console.error('Error updating template:', error);
            throw error;
        }
    }

    /**
     * Exports a template as downloadable JSON file
     * @param {Object} template - Template to export
     */
    static async exportTemplateAsJSON(template) {
        try {
            const jsonData = JSON.stringify(template, null, 2);
            const filename = `${template.name.replace(/[^a-z0-9]/gi, '_')}.broker-template.json`;

            // Try to use preferred folder if available
            try {
                if ('showSaveFilePicker' in window && window.appSettings && window.appSettings.downloadFolderHandle) {
                    const fileHandle = await window.appSettings.downloadFolderHandle.getFileHandle(filename, { create: true });
                    const writable = await fileHandle.createWritable();
                    await writable.write(jsonData);
                    await writable.close();
                    return;
                }
            } catch (error) {
                console.log('Failed to write to preferred folder, falling back to browser download:', error);
            }

            // Fallback to browser download
            const blob = new Blob([jsonData], { type: 'application/json' });
            const url = URL.createObjectURL(blob);

            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            URL.revokeObjectURL(url);
            console.log('Template exported as JSON:', filename);
        } catch (error) {
            console.error('Error exporting template:', error);
        }
    }



    /**
     * Validates template structure and required fields
     * @param {Object} template - Template to validate
     * @throws {Error} If template is invalid
     */
    static validateTemplate(template) {
        const required = ['name', 'parsingConfig', 'columnMapping'];
        const missing = required.filter(field => !template[field]);

        if (missing.length > 0) {
            throw new Error(`Template missing required fields: ${missing.join(', ')}`);
        }

        // Validate parsing config structure
        if (!template.parsingConfig.dataStartMethod) {
            throw new Error('Template parsingConfig must specify dataStartMethod');
        }

        // Validate file pattern if provided
        if (template.filePattern && template.filePattern.regex) {
            try {
                new RegExp(template.filePattern.regex);
            } catch (error) {
                throw new Error(`Invalid regex pattern in template: ${error.message}`);
            }
        }
    }


    /**
     * Increments version number (simple numeric increment)
     * @param {String} currentVersion - Current version string
     * @returns {String} Incremented version
     */
    static incrementVersion(currentVersion = '1.0') {
        const parts = currentVersion.split('.');
        const major = parseInt(parts[0]) || 1;
        const minor = parseInt(parts[1]) || 0;

        return `${major}.${minor + 1}`;
    }
}