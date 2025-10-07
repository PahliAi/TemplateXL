/**
 * Borderellen Converter - Storage Manager
 * Handles IndexedDB operations and JSON export/import functionality
 */

let db = null;

/**
 * Initialize IndexedDB database
 * @returns {Promise<IDBDatabase>} Database instance
 */
function initIndexedDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open('BorderellenDB', 5);

        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
            db = request.result;
            window.db = db; // Make db globally accessible
            resolve(db);
        };

        request.onupgradeneeded = (event) => {
            const database = event.target.result;

            // Create templates store
            if (!database.objectStoreNames.contains('templates')) {
                const templateStore = database.createObjectStore('templates', { keyPath: 'id' });
                templateStore.createIndex('name', 'name', { unique: false });
            }

            // Create settings store
            if (!database.objectStoreNames.contains('settings')) {
                const settingsStore = database.createObjectStore('settings', { keyPath: 'key' });
            }

            // Create unified file mappings store (replaces brokerTemplates + customBrokerTemplates)
            if (!database.objectStoreNames.contains('fileMappings')) {
                const fileMappingsStore = database.createObjectStore('fileMappings', { keyPath: 'id' });
                fileMappingsStore.createIndex('name', 'name', { unique: false });
                fileMappingsStore.createIndex('creationMethod', 'creationMethod', { unique: false });
                fileMappingsStore.createIndex('matchingKeyword', 'matchingKeyword', { unique: false });
                fileMappingsStore.createIndex('created', 'created', { unique: false });
            }


            // Create broker contacts store
            if (!database.objectStoreNames.contains('brokerContacts')) {
                const contactsStore = database.createObjectStore('brokerContacts', { keyPath: 'id' });
                contactsStore.createIndex('brokerName', 'brokerName', { unique: false });
                contactsStore.createIndex('email', 'email', { unique: false });
                contactsStore.createIndex('created', 'created', { unique: false });
            }

            // Create template associations store for file-template persistence
            if (!database.objectStoreNames.contains('templateAssociations')) {
                const associationsStore = database.createObjectStore('templateAssociations', { keyPath: 'key' });
            }
        };

    });
}

/**
 * Save template to IndexedDB
 * @param {string} templateId - Template ID
 * @param {Object} template - Template object
 * @returns {Promise<boolean>} Success status
 */
async function saveTemplateToStorage(templateId, template) {
    try {
        if (!db) await initIndexedDB();

        // Save to IndexedDB
        const transaction = db.transaction(['templates'], 'readwrite');
        const store = transaction.objectStore('templates');

        await new Promise((resolve, reject) => {
            const request = store.put(template);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        // Save current template ID
        const settingsTransaction = db.transaction(['settings'], 'readwrite');
        const settingsStore = settingsTransaction.objectStore('settings');

        await new Promise((resolve, reject) => {
            const request = settingsStore.put({ key: 'currentTemplateId', value: templateId });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        return true;
    } catch (error) {
        console.error('Error saving template to IndexedDB:', error);
        return false;
    }
}

/**
 * Load all templates from IndexedDB
 * @returns {Promise<Object>} Templates object and current template info
 */
async function loadTemplatesFromStorage() {
    try {
        if (!db) await initIndexedDB();
        console.log('Loading templates from storage...');

        // Load all templates
        const transaction = db.transaction(['templates'], 'readonly');
        const store = transaction.objectStore('templates');

        const templates = await new Promise((resolve, reject) => {
            const request = store.getAll();
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        console.log('Found templates:', templates.length);

        // Rebuild savedTemplates object
        const savedTemplates = {};
        templates.forEach(template => {
            savedTemplates[template.id] = template;
            console.log('Loaded template:', template.name);
        });

        // Load current template ID
        const settingsTransaction = db.transaction(['settings'], 'readonly');
        const settingsStore = settingsTransaction.objectStore('settings');

        const currentIdSetting = await new Promise((resolve, reject) => {
            const request = settingsStore.get('currentTemplateId');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        console.log('Current template ID setting:', currentIdSetting);

        return {
            templates: savedTemplates,
            currentTemplateId: currentIdSetting ? currentIdSetting.value : null,
            hasTemplates: templates.length > 0
        };
    } catch (error) {
        console.error('Error loading templates from IndexedDB:', error);
        return {
            templates: {},
            currentTemplateId: null,
            hasTemplates: false
        };
    }
}

/**
 * Delete template from IndexedDB
 * @param {string} templateId - Template ID to delete
 * @returns {Promise<boolean>} Success status
 */
async function deleteTemplateFromStorage(templateId) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['templates'], 'readwrite');
        const store = transaction.objectStore('templates');

        await new Promise((resolve, reject) => {
            const request = store.delete(templateId);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        return true;
    } catch (error) {
        console.error('Error deleting template from IndexedDB:', error);
        return false;
    }
}

/**
 * Load user settings from IndexedDB
 * @returns {Promise<Object>} Settings object
 */
async function loadSettings() {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['settings'], 'readonly');
        const store = transaction.objectStore('settings');

        const userNameSetting = await new Promise((resolve, reject) => {
            const request = store.get('userName');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        const userEmailSetting = await new Promise((resolve, reject) => {
            const request = store.get('userEmail');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        const userSignatureSetting = await new Promise((resolve, reject) => {
            const request = store.get('userSignature');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        const downloadFolderSetting = await new Promise((resolve, reject) => {
            const request = store.get('downloadFolder');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        const emailSubjectSetting = await new Promise((resolve, reject) => {
            const request = store.get('emailSubject');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        const emailBodyTemplateSetting = await new Promise((resolve, reject) => {
            const request = store.get('emailBodyTemplate');
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        const settings = {
            userName: userNameSetting ? userNameSetting.value : 'User',
            userEmail: userEmailSetting ? userEmailSetting.value : '',
            userSignature: userSignatureSetting ? userSignatureSetting.value : '',
            downloadFolder: downloadFolderSetting ? downloadFolderSetting.value : '',
            emailSubject: emailSubjectSetting ? emailSubjectSetting.value : '',
            emailBodyTemplate: emailBodyTemplateSetting ? emailBodyTemplateSetting.value : ''
        };

        // Load folder handle for startIn suggestion (no permission needed)
        if ('showDirectoryPicker' in window) {
            console.log('Attempting to load folder handle from IndexedDB...');
            const folderHandleSetting = await new Promise((resolve, reject) => {
                const request = store.get('downloadFolderHandle');
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error);
            });

            console.log('Folder handle setting from DB:', folderHandleSetting);

            if (folderHandleSetting && folderHandleSetting.value) {
                try {
                    settings.downloadFolderHandle = folderHandleSetting.value;
                    console.log('Restored folder handle for startIn suggestion:', folderHandleSetting.value.name);
                } catch (error) {
                    console.log('Error restoring folder handle:', error);
                }
            } else {
                console.log('No folder handle found in IndexedDB');
            }
        }

        return settings;
    } catch (error) {
        console.error('Error loading settings:', error);
        return {
            userName: 'User',
            userEmail: '',
            userSignature: '',
            downloadFolder: ''
        };
    }
}

/**
 * Save user settings to IndexedDB
 * @param {Object} settings - Settings object
 * @returns {Promise<boolean>} Success status
 */
async function saveSettings(settings) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['settings'], 'readwrite');
        const store = transaction.objectStore('settings');

        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'userName', value: settings.userName });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'userEmail', value: settings.userEmail });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'userSignature', value: settings.userSignature });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'downloadFolder', value: settings.downloadFolder });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        // Save email template settings if provided
        if (settings.emailSubject !== undefined) {
            await new Promise((resolve, reject) => {
                const request = store.put({ key: 'emailSubject', value: settings.emailSubject });
                request.onsuccess = () => resolve();
                request.onerror = () => reject(request.error);
            });
        }

        if (settings.emailBodyTemplate !== undefined) {
            await new Promise((resolve, reject) => {
                const request = store.put({ key: 'emailBodyTemplate', value: settings.emailBodyTemplate });
                request.onsuccess = () => resolve();
                request.onerror = () => reject(request.error);
            });
        }

        // Save folder handle if available (File System Access API handles are serializable to IndexedDB)
        if (settings.downloadFolderHandle) {
            console.log('Saving folder handle to IndexedDB:', settings.downloadFolderHandle.name);
            await new Promise((resolve, reject) => {
                const request = store.put({ key: 'downloadFolderHandle', value: settings.downloadFolderHandle });
                request.onsuccess = () => {
                    console.log('Folder handle saved successfully');
                    resolve();
                };
                request.onerror = () => {
                    console.error('Failed to save folder handle:', request.error);
                    reject(request.error);
                };
            });
        } else {
            console.log('No folder handle to save');
        }

        return true;
    } catch (error) {
        console.error('Error saving settings:', error);
        return false;
    }
}

/**
 * Export template as JSON file
 * @param {Object} template - Template object
 * @param {string} fileName - Optional filename
 * @param {Object} appSettings - App settings for folder handling
 */
async function exportTemplateAsJSON(template, fileName, appSettings) {
    const dataStr = JSON.stringify(template, null, 2);
    const finalFileName = fileName || `${template.name.replace(/[^a-z0-9]/gi, '_')}.json`;

    try {
        // Try to use File System Access API if available and user has selected a folder
        if ('showSaveFilePicker' in window && appSettings && appSettings.downloadFolderHandle) {
            const fileHandle = await appSettings.downloadFolderHandle.getFileHandle(finalFileName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(dataStr);
            await writable.close();
            return;
        }
    } catch (error) {
        console.log('File System Access failed, falling back to download:', error);
    }

    // Fallback to regular download
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = finalFileName;
    a.click();

    URL.revokeObjectURL(url);
}

/**
 * Load template from JSON file
 * @param {File} file - JSON file to load
 * @returns {Promise<Object>} Template object
 */
async function loadTemplateFromJSON(file) {
    try {
        const text = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsText(file);
        });

        const imported = JSON.parse(text);

        // Validate basic structure
        if (!imported.columns || !Array.isArray(imported.columns)) {
            throw new Error('Invalid template format - missing columns array');
        }

        return {
            success: true,
            template: imported
        };
    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}


// ========== BROKER CONTACTS MANAGEMENT ==========

/**
 * Save broker contact to IndexedDB
 * @param {Object} contact - Contact object
 * @returns {Promise<boolean>} Success status
 */
async function saveBrokerContact(contact) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['brokerContacts'], 'readwrite');
        const store = transaction.objectStore('brokerContacts');

        await new Promise((resolve, reject) => {
            const request = store.put(contact);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        console.log('Broker contact saved:', contact);
        return true;
    } catch (error) {
        console.error('Error saving broker contact:', error);
        return false;
    }
}

/**
 * Load all broker contacts from IndexedDB
 * @returns {Promise<Array>} Array of contact objects
 */
async function loadAllBrokerContacts() {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['brokerContacts'], 'readonly');
        const store = transaction.objectStore('brokerContacts');

        const contacts = await new Promise((resolve, reject) => {
            const request = store.getAll();
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        console.log('Loaded broker contacts:', contacts.length);
        return contacts.sort((a, b) => a.brokerName.localeCompare(b.brokerName));
    } catch (error) {
        console.error('Error loading broker contacts:', error);
        return [];
    }
}

/**
 * Delete broker contact from IndexedDB
 * @param {string} contactId - Contact ID to delete
 * @returns {Promise<boolean>} Success status
 */
async function deleteBrokerContact(contactId) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['brokerContacts'], 'readwrite');
        const store = transaction.objectStore('brokerContacts');

        await new Promise((resolve, reject) => {
            const request = store.delete(contactId);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        console.log('Broker contact deleted:', contactId);
        return true;
    } catch (error) {
        console.error('Error deleting broker contact:', error);
        return false;
    }
}

/**
 * Export broker contacts as JSON file
 * @param {Array} contacts - Array of contact objects
 * @param {string} fileName - Optional filename
 * @param {Object} appSettings - App settings for folder handling
 */
async function exportBrokerContactsAsJSON(contacts, fileName, appSettings) {
    const timestamp = new Date().toISOString().slice(0, 10);
    const defaultFileName = fileName || `broker-contacts-${timestamp}.json`;

    const exportData = {
        exportDate: new Date().toISOString(),
        version: '1.0',
        type: 'brokerContacts',
        contacts: contacts.map(contact => ({
            brokerName: contact.brokerName,
            firstName: contact.firstName,
            lastName: contact.lastName,
            email: contact.email,
            created: contact.created
        }))
    };

    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });

    try {
        if (appSettings.downloadFolderHandle) {
            try {
                const fileHandle = await appSettings.downloadFolderHandle.getFileHandle(defaultFileName, { create: true });
                const writable = await fileHandle.createWritable();
                await writable.write(blob);
                await writable.close();
                console.log('Broker contacts exported to selected folder:', defaultFileName);
                return;
            } catch (folderError) {
                console.log('Folder access failed, falling back to browser download:', folderError.message);
            }
        }

        // Fallback to browser download
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = defaultFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        console.log('Broker contacts exported via browser download:', defaultFileName);
    } catch (error) {
        console.error('Error exporting broker contacts:', error);
        alert('Failed to export broker contacts: ' + error.message);
    }
}

/**
 * Load broker contacts from JSON file
 * @param {File} file - JSON file to import
 * @returns {Promise<Object>} Import result with success status and contacts array
 */
async function loadBrokerContactsFromJSON(file) {
    try {
        const text = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsText(file);
        });

        const imported = JSON.parse(text);

        // Validate contacts structure
        if (!imported.contacts || !Array.isArray(imported.contacts)) {
            throw new Error('Invalid contacts format - missing contacts array');
        }

        // Validate each contact
        const validContacts = imported.contacts.filter(contact => {
            return contact.brokerName && contact.firstName && contact.lastName && contact.email;
        });

        if (validContacts.length === 0) {
            throw new Error('No valid contacts found in import file');
        }

        // Generate new IDs and timestamps for imported contacts
        const processedContacts = validContacts.map(contact => ({
            id: `contact-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
            brokerName: contact.brokerName,
            firstName: contact.firstName,
            lastName: contact.lastName,
            email: contact.email,
            created: new Date().toISOString(),
            imported: true,
            importedAt: new Date().toISOString()
        }));

        return {
            success: true,
            contacts: processedContacts,
            imported: validContacts.length,
            skipped: imported.contacts.length - validContacts.length
        };
    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}


/**
 * Save file mapping to unified store
 * @param {Object} mapping - File mapping to save
 * @returns {Promise<boolean>} Success status
 */
async function saveFileMapping(mapping) {
    try {
        if (!db) await initIndexedDB();

        // Ensure required fields
        const completeMapping = {
            ...mapping,
            id: mapping.id || `mapping-${Date.now()}`,
            created: mapping.created || new Date().toISOString(),
            lastModified: new Date().toISOString(),
            matchingKeyword: mapping.matchingKeyword || '',
            creationMethod: mapping.creationMethod || 'unknown'
        };

        const transaction = db.transaction(['fileMappings'], 'readwrite');
        const store = transaction.objectStore('fileMappings');

        await new Promise((resolve, reject) => {
            const request = store.put(completeMapping);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        return true;
    } catch (error) {
        console.error('Error saving file mapping:', error);
        return false;
    }
}

/**
 * Load all file mappings from unified store
 * @returns {Promise<Array>} Array of file mappings
 */
async function loadAllFileMappings() {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['fileMappings'], 'readonly');
        const store = transaction.objectStore('fileMappings');

        const mappings = await new Promise((resolve, reject) => {
            const request = store.getAll();
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });

        console.log('Loaded file mappings:', mappings.length);
        return mappings;
    } catch (error) {
        console.error('Error loading file mappings:', error);
        return [];
    }
}

/**
 * Load file mapping by keyword matching
 * @param {string} filename - Filename to match against keywords
 * @returns {Promise<Object|null>} Matching mapping or null
 */
async function loadFileMappingByKeyword(filename) {
    try {
        const mappings = await loadAllFileMappings();

        const matchingMapping = mappings.find(mapping => {
            if (!mapping.matchingKeyword || mapping.matchingKeyword.trim() === '') {
                return false;
            }

            const keyword = mapping.matchingKeyword.trim().toLowerCase();
            const filenameLC = filename.toLowerCase();

            return filenameLC.includes(keyword);
        });

        if (matchingMapping) {
            console.log(`Found matching file mapping for ${filename} using keyword "${matchingMapping.matchingKeyword}":`, matchingMapping.name);
        }

        return matchingMapping || null;
    } catch (error) {
        console.error('Error loading file mapping by keyword:', error);
        return null;
    }
}

/**
 * Delete file mapping from unified store
 * @param {string} mappingId - Mapping ID to delete
 * @returns {Promise<boolean>} Success status
 */
async function deleteFileMapping(mappingId) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction(['fileMappings'], 'readwrite');
        const store = transaction.objectStore('fileMappings');

        await new Promise((resolve, reject) => {
            const request = store.delete(mappingId);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        console.log('File mapping deleted:', mappingId);
        return true;
    } catch (error) {
        console.error('Error deleting file mapping:', error);
        return false;
    }
}

/**
 * Export file mapping as JSON file
 * @param {Object} mapping - File mapping object
 * @param {string} fileName - Optional filename
 * @param {Object} appSettings - App settings for folder handling
 */
async function exportFileMappingAsJSON(mapping, fileName, appSettings) {
    // Create a clean filename from mapping name
    const safeName = mapping.name.replace(/[^a-z0-9\s-]/gi, '_').replace(/\s+/g, '_');
    const finalFileName = fileName || `${safeName}_file_mapping.json`;

    // Create export data with additional metadata
    const exportData = {
        ...mapping,
        exportedAt: new Date().toISOString(),
        exportedBy: appSettings?.userName || 'User',
        type: 'file_mapping',
        version: mapping.version || '1.0'
    };

    const dataStr = JSON.stringify(exportData, null, 2);

    try {
        // Try to use File System Access API if available and user has selected a folder
        if ('showSaveFilePicker' in window && appSettings && appSettings.downloadFolderHandle) {
            const fileHandle = await appSettings.downloadFolderHandle.getFileHandle(finalFileName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(dataStr);
            await writable.close();
            console.log('File mapping exported to selected folder:', finalFileName);
            return;
        }
    } catch (error) {
        console.log('File System Access failed, falling back to download:', error);
    }

    // Fallback to regular download
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = finalFileName;
    a.click();

    URL.revokeObjectURL(url);
    console.log('File mapping downloaded as JSON:', finalFileName);
}

/**
 * Load file mapping from JSON file
 * @param {File} file - JSON file to load
 * @returns {Promise<Object>} Result object with success status and mapping data
 */
async function loadFileMappingFromJSON(file) {
    try {
        const text = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsText(file);
        });

        const imported = JSON.parse(text);

        // Validate file mapping structure
        if (!imported.columnMapping || typeof imported.columnMapping !== 'object') {
            throw new Error('Invalid file mapping format - missing columnMapping object');
        }

        // Generate new ID and update timestamps for imported mapping
        const importedMapping = {
            ...imported,
            id: `mapping-${Date.now()}`,
            imported: true,
            importedAt: new Date().toISOString(),
            originalId: imported.id,
            creationMethod: imported.creationMethod || 'imported'
        };

        return {
            success: true,
            mapping: importedMapping
        };
    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Save current template ID to IndexedDB
 * @param {string} templateId - Template ID to set as current
 * @returns {Promise<boolean>} Success status
 */
async function saveCurrentTemplateId(templateId) {
    try {
        if (!db) await initIndexedDB();

        const settingsTransaction = db.transaction(['settings'], 'readwrite');
        const settingsStore = settingsTransaction.objectStore('settings');

        await new Promise((resolve, reject) => {
            const request = settingsStore.put({ key: 'currentTemplateId', value: templateId });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        console.log('Current template ID saved to IndexedDB:', templateId);
        return true;
    } catch (error) {
        console.error('Error saving current template ID:', error);
        return false;
    }
}

// Export essential functions to window for cross-module access
window.initIndexedDB = initIndexedDB;
window.loadSettings = loadSettings;
window.saveSettings = saveSettings;
window.saveCurrentTemplateId = saveCurrentTemplateId;
window.saveBrokerContact = saveBrokerContact;
window.loadAllBrokerContacts = loadAllBrokerContacts;
window.deleteBrokerContact = deleteBrokerContact;
window.exportBrokerContactsAsJSON = exportBrokerContactsAsJSON;
window.loadBrokerContactsFromJSON = loadBrokerContactsFromJSON;

// Unified file mappings functions
window.saveFileMapping = saveFileMapping;
window.loadAllFileMappings = loadAllFileMappings;
window.loadFileMappingByKeyword = loadFileMappingByKeyword;
window.deleteFileMapping = deleteFileMapping;
window.exportFileMappingAsJSON = exportFileMappingAsJSON;
window.loadFileMappingFromJSON = loadFileMappingFromJSON;

/**
 * Generic function to save data to IndexedDB
 * @param {string} storeName - Name of the object store
 * @param {string} key - Key to save data under
 * @param {*} data - Data to save
 * @returns {Promise<boolean>} Success status
 */
async function saveToIndexedDB(storeName, key, data) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction([storeName], 'readwrite');
        const store = transaction.objectStore(storeName);

        await new Promise((resolve, reject) => {
            const request = store.put({ key: key, data: data });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        return true;
    } catch (error) {
        console.error(`Error saving to IndexedDB store ${storeName}:`, error);
        return false;
    }
}

/**
 * Generic function to load data from IndexedDB
 * @param {string} storeName - Name of the object store
 * @param {string} key - Key to load data from
 * @returns {Promise<*>} Loaded data or null if not found
 */
async function loadFromIndexedDB(storeName, key) {
    try {
        if (!db) await initIndexedDB();

        const transaction = db.transaction([storeName], 'readonly');
        const store = transaction.objectStore(storeName);

        return new Promise((resolve, reject) => {
            const request = store.get(key);
            request.onsuccess = () => {
                const result = request.result;
                resolve(result ? result.data : null);
            };
            request.onerror = () => reject(request.error);
        });
    } catch (error) {
        console.error(`Error loading from IndexedDB store ${storeName}:`, error);
        return null;
    }
}

// Generic IndexedDB functions
window.saveToIndexedDB = saveToIndexedDB;
window.loadFromIndexedDB = loadFromIndexedDB;