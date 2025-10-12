/**
 * Borderellen Converter - Email Management Module
 * Handles all email-related functionality for missing data analysis and notifications
 */

// ========== EMAIL STATE VARIABLES ==========

window.emailUploadedFiles = [];
window.missingDataAnalysis = [];
window.pendingContactMatching = [];

// ========== EMAIL INITIALIZATION ==========

/**
 * Initialize email brokers tab functionality
 */
function initializeEmailBrokersTab() {
    // Data source radio buttons
    const useResultsData = document.getElementById('use-results-data');
    const useUploadData = document.getElementById('use-upload-data');
    const emailUploadSection = document.getElementById('email-upload-section');

    useResultsData.addEventListener('change', () => {
        if (useResultsData.checked) {
            emailUploadSection.style.display = 'none';
            analyzeCurrentResultsData();
        }
    });

    useUploadData.addEventListener('change', () => {
        if (useUploadData.checked) {
            emailUploadSection.style.display = 'block';
        }
    });

    // File upload functionality
    const emailUploadZone = document.getElementById('email-upload-zone');
    const emailFileInput = document.getElementById('email-file-input');
    const emailBrowseBtn = document.getElementById('email-browse-btn');
    const emailClearFilesBtn = document.getElementById('email-clear-files-btn');

    emailUploadZone.addEventListener('click', () => emailFileInput.click());
    emailBrowseBtn.addEventListener('click', () => emailFileInput.click());
    emailClearFilesBtn.addEventListener('click', clearEmailFiles);

    emailUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        emailUploadZone.style.background = '#444';
    });

    emailUploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        emailUploadZone.style.background = '#333';
    });

    emailUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        emailUploadZone.style.background = '#333';
        handleEmailFilesDrop(e.dataTransfer.files);
    });

    emailFileInput.addEventListener('change', (e) => {
        handleEmailFilesDrop(e.target.files);
    });

    // Analysis and email generation
    const refreshAnalysisBtn = document.getElementById('refresh-analysis-btn');
    const sendEmailsBtn = document.getElementById('send-emails-btn');
    const saveEmailTemplateBtn = document.getElementById('save-email-template-btn');
    const resetEmailTemplateBtn = document.getElementById('reset-email-template-btn');

    refreshAnalysisBtn.addEventListener('click', refreshMissingDataAnalysis);
    sendEmailsBtn.addEventListener('click', generateBrokerEmails);
    saveEmailTemplateBtn.addEventListener('click', saveEmailTemplate);
    resetEmailTemplateBtn.addEventListener('click', resetEmailTemplate);

    // Contact matching modal
    const closeContactMatching = document.getElementById('close-contact-matching');
    const skipContactBtn = document.getElementById('skip-contact-btn');
    const confirmContactBtn = document.getElementById('confirm-contact-btn');

    closeContactMatching.addEventListener('click', closeContactMatchingModal);
    skipContactBtn.addEventListener('click', skipCurrentBrokerContact);
    confirmContactBtn.addEventListener('click', confirmBrokerContact);

    // Filename assignment modal
    const closeFilenameAssignment = document.getElementById('close-filename-assignment');
    const skipFilenameAssignments = document.getElementById('skip-filename-assignments');
    const confirmFilenameAssignments = document.getElementById('confirm-filename-assignments');

    closeFilenameAssignment.addEventListener('click', closeFilenameAssignmentModal);
    skipFilenameAssignments.addEventListener('click', skipFilenameAssignmentProcess);
    confirmFilenameAssignments.addEventListener('click', confirmFilenameAssignmentProcess);

    // Initialize with current results data if available and "Use current Results data" is selected
    setTimeout(() => {
        if (useResultsData.checked && window.currentCombinedData && window.currentCombinedData.length > 0) {
            analyzeCurrentResultsData();
        }
    }, 100); // Small delay to ensure DOM is ready
}

// ========== FILE HANDLING ==========

/**
 * Handle email files drop/upload
 */
async function handleEmailFilesDrop(files) {
    const analysisStatus = document.getElementById('analysis-status');
    const analysisProgress = document.getElementById('analysis-progress');

    analysisStatus.style.display = 'block';
    analysisProgress.innerHTML = '<p style="color: #ffa726;">Processing uploaded files...</p>';

    window.emailUploadedFiles = [];

    for (let file of files) {
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            alert(`Skipping ${file.name}: Only Excel files (.xlsx, .xls) are supported.`);
            continue;
        }

        try {
            // Read and validate file against active template
            const workbook = await ExcelCacheManager.getWorkbook(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(worksheet);

            if (data.length === 0) {
                alert(`File ${file.name} appears to be empty.`);
                continue;
            }

            // Check if columns match active template
            const templateColumns = window.borderellenTemplate?.columns || [];
            const fileColumns = Object.keys(data[0]);
            const templateColumnNames = templateColumns.map(col => col.name);

            // Check for exact column match
            const missingColumns = templateColumnNames.filter(col => !fileColumns.includes(col));
            const extraColumns = fileColumns.filter(col => !templateColumnNames.includes(col));

            if (missingColumns.length > 0 || extraColumns.length > 0) {
                let message = `File ${file.name} does not match the active template:\n`;
                if (missingColumns.length > 0) {
                    message += `Missing columns: ${missingColumns.join(', ')}\n`;
                }
                if (extraColumns.length > 0) {
                    message += `Extra columns: ${extraColumns.join(', ')}\n`;
                }
                message += '\nPlease ensure the file was generated using this application with the same template.';
                alert(message);
                continue;
            }

            window.emailUploadedFiles.push({
                name: file.name,
                data: data,
                workbook: workbook
            });

        } catch (error) {
            console.error('Error processing file:', file.name, error);
            alert(`Error processing ${file.name}: ${error.message}`);
        }
    }

    if (window.emailUploadedFiles.length > 0) {
        analysisProgress.innerHTML = `<p style="color: #4caf50;">Successfully loaded ${window.emailUploadedFiles.length} file(s). Running missing data analysis...</p>`;
        await analyzeMissingData(window.emailUploadedFiles);
    } else {
        analysisProgress.innerHTML = '<p style="color: #f44336;">No valid files were processed.</p>';
    }
}

/**
 * Clear uploaded files for email analysis
 */
function clearEmailFiles() {
    window.emailUploadedFiles = [];
    document.getElementById('email-file-input').value = '';
    document.getElementById('analysis-status').style.display = 'none';
    document.getElementById('email-template-section').style.display = 'none';
    document.getElementById('missing-data-results').style.display = 'none';
}

// ========== DATA ANALYSIS ==========

/**
 * Analyze current results data for missing values
 */
async function analyzeCurrentResultsData() {
    if (!window.currentCombinedData || window.currentCombinedData.length === 0) {
        alert('No results data available. Please process some broker files first.');
        return;
    }

    const analysisStatus = document.getElementById('analysis-status');
    const analysisProgress = document.getElementById('analysis-progress');

    analysisStatus.style.display = 'block';
    analysisProgress.innerHTML = '<p style="color: #ffa726;">Analyzing current results data...</p>';

    // Convert current combined data to file-like format for analysis
    // Group by original source file to maintain filename associations
    const groupedByFile = {};

    window.currentCombinedData.forEach(record => {
        const sourceFile = record._sourceFile || 'Unknown File';
        if (!groupedByFile[sourceFile]) {
            groupedByFile[sourceFile] = [];
        }
        groupedByFile[sourceFile].push(record);
    });

    // Convert to file data format
    const fileData = Object.keys(groupedByFile).map(filename => ({
        name: filename,
        data: groupedByFile[filename]
    }));

    await analyzeMissingData(fileData);
}

/**
 * Analyze missing data for all files
 */
async function analyzeMissingData(filesData) {
    try {
        const templateColumns = window.borderellenTemplate?.columns || [];
        const requiredColumns = templateColumns.filter(col => col.required);

        if (requiredColumns.length === 0) {
            alert(`No required columns are defined in the active template.\n\nTemplate has ${templateColumns.length} total columns, but none are marked as "required".\n\nTo fix this:\n1. Go to Template Manager (Tab 1)\n2. Check the "Required" boxes for mandatory columns\n3. Save the template`);
            return;
        }

        console.log(`Email analysis: Found ${requiredColumns.length} required columns out of ${templateColumns.length} total columns`);
        console.log('Required columns:', requiredColumns.map(col => col.name));

        window.missingDataAnalysis = [];

        // Group data by broker
        const brokerGroups = {};

        for (const fileData of filesData) {
            for (const row of fileData.data) {
                const brokerName = row.Makelaar || row['Broker Name'] || 'Unknown';

                if (!brokerGroups[brokerName]) {
                    brokerGroups[brokerName] = {
                        brokerName: brokerName,
                        rows: [],
                        filenames: [],
                        primaryFilename: fileData.name
                    };
                }

                // Add filename if not already present
                if (!brokerGroups[brokerName].filenames.includes(fileData.name)) {
                    brokerGroups[brokerName].filenames.push(fileData.name);
                }

                brokerGroups[brokerName].rows.push({ ...row, _sourceFilename: fileData.name });
            }
        }

        // Analyze missing data for each broker
        for (const [brokerName, brokerData] of Object.entries(brokerGroups)) {
            // Create a display filename
            const displayFilename = brokerData.filenames.length === 1
                ? brokerData.filenames[0]
                : `${brokerData.filenames.length} files: ${brokerData.filenames.join(', ')}`;

            // Find the best filename for this broker (one containing broker name)
            const brokerSpecificFilename = findBrokerSpecificFilename(brokerName, brokerData.filenames);

            const analysis = {
                brokerName: brokerName,
                filename: displayFilename,
                emailFilename: brokerSpecificFilename || brokerData.primaryFilename,
                primaryFilename: brokerData.primaryFilename,
                allFilenames: brokerData.filenames,
                totalRows: brokerData.rows.length,
                missingColumns: [],
                overallCompletionRate: 0
            };

            let totalRequiredFields = 0;
            let totalFilledFields = 0;

            for (const column of requiredColumns) {
                const columnName = column.name;
                let filledCount = 0;

                for (const row of brokerData.rows) {
                    const value = row[columnName];
                    if (value !== null && value !== undefined && value !== '') {
                        filledCount++;
                    }
                }

                const completionRate = (filledCount / brokerData.rows.length) * 100;

                analysis.missingColumns.push({
                    columnName: columnName,
                    filledCount: filledCount,
                    totalCount: brokerData.rows.length,
                    completionRate: completionRate
                });

                totalRequiredFields += brokerData.rows.length;
                totalFilledFields += filledCount;
            }

            analysis.overallCompletionRate = (totalFilledFields / totalRequiredFields) * 100;

            // Only include brokers that have missing data
            if (analysis.overallCompletionRate < 100) {
                window.missingDataAnalysis.push(analysis);
            }
        }

        await displayMissingDataResults();

    } catch (error) {
        console.error('Error analyzing missing data:', error);
        document.getElementById('analysis-progress').innerHTML = `<p style="color: #f44336;">Error during analysis: ${error.message}</p>`;
    }
}

/**
 * Display missing data analysis results
 */
async function displayMissingDataResults() {
    const analysisProgress = document.getElementById('analysis-progress');
    const emailTemplateSection = document.getElementById('email-template-section');
    const missingDataResults = document.getElementById('missing-data-results');
    const summaryDiv = document.getElementById('broker-analysis-summary');
    const tableBody = document.getElementById('missing-data-table-body');

    if (window.missingDataAnalysis.length === 0) {
        analysisProgress.innerHTML = '<p style="color: #4caf50;">✓ All brokers have complete data - no emails needed!</p>';
        emailTemplateSection.style.display = 'none';
        missingDataResults.style.display = 'none';
        return;
    }

    analysisProgress.innerHTML = `<p style="color: #4caf50;">✓ Analysis complete - found ${window.missingDataAnalysis.length} broker(s) with missing data</p>`;

    // Show summary
    const totalBrokers = window.missingDataAnalysis.length;
    const avgCompletion = window.missingDataAnalysis.reduce((sum, broker) => sum + broker.overallCompletionRate, 0) / totalBrokers;

    summaryDiv.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; background: #333; padding: 16px; border-radius: 8px;">
            <div>
                <div style="color: #00bcd4; font-size: 24px; font-weight: bold;">${totalBrokers}</div>
                <div style="color: #888; font-size: 14px;">Brokers with missing data</div>
            </div>
            <div>
                <div style="color: #ffa726; font-size: 24px; font-weight: bold;">${avgCompletion.toFixed(1)}%</div>
                <div style="color: #888; font-size: 14px;">Average completion rate</div>
            </div>
        </div>
    `;

    // Display table
    tableBody.innerHTML = await Promise.all(window.missingDataAnalysis.map(async (broker) => {
        const contact = await findBrokerContact(broker.brokerName);
        const missingColumnsText = broker.missingColumns
            .map(col => `${col.columnName} (${col.completionRate.toFixed(1)}%)`)
            .join(', ');

        return `
            <tr>
                <td>${escapeHtml(broker.brokerName)}</td>
                <td>${contact ? escapeHtml(`${contact.firstName} ${contact.lastName}`) : '<span style="color: #f44336;">No contact found</span>'}</td>
                <td>${broker.totalRows}</td>
                <td style="max-width: 300px; word-wrap: break-word;">${missingColumnsText}</td>
                <td>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <div style="background: #555; border-radius: 4px; overflow: hidden; flex: 1; height: 8px;">
                            <div style="background: ${broker.overallCompletionRate >= 80 ? '#4caf50' : broker.overallCompletionRate >= 60 ? '#ffa726' : '#f44336'}; width: ${broker.overallCompletionRate}%; height: 100%;"></div>
                        </div>
                        <span style="font-size: 12px; color: #888; min-width: 40px;">${broker.overallCompletionRate.toFixed(1)}%</span>
                    </div>
                </td>
                <td>
                    <button class="btn btn-secondary" onclick="previewEmail('${broker.brokerName}')" style="font-size: 12px; padding: 4px 8px;">Preview Email</button>
                </td>
            </tr>
        `;
    })).then(rows => rows.join(''));

    emailTemplateSection.style.display = 'block';
    missingDataResults.style.display = 'block';
}

/**
 * Find broker contact by name
 */
async function findBrokerContact(brokerName) {
    try {
        const contacts = await window.loadAllBrokerContacts();
        return contacts.find(contact =>
            contact.brokerName.toLowerCase().trim() === brokerName.toLowerCase().trim()
        );
    } catch (error) {
        console.error('Error loading contacts:', error);
        return null;
    }
}

/**
 * Refresh missing data analysis
 */
async function refreshMissingDataAnalysis() {
    const useResultsData = document.getElementById('use-results-data');

    if (useResultsData.checked) {
        await analyzeCurrentResultsData();
    } else if (emailUploadedFiles.length > 0) {
        await analyzeMissingData(window.emailUploadedFiles);
    } else {
        alert('No data available for analysis. Please select a data source.');
    }
}

// ========== EMAIL GENERATION ==========

/**
 * Validate all placeholders before email generation
 */
async function validateEmailPlaceholders() {
    const subject = document.getElementById('email-subject').value;
    const bodyTemplate = document.getElementById('email-body-template').value;

    const errors = [];
    const brokersNeedingFilenames = [];

    // Check user settings
    if (!window.appSettings.userEmail || window.appSettings.userEmail.trim() === '') {
        errors.push('User email address is not configured. Please set it in Settings.');
    }

    // Validate email template placeholders for each broker
    for (const broker of window.missingDataAnalysis) {
        const contact = broker.manualContact || await findBrokerContact(broker.brokerName);

        if (!contact && (bodyTemplate.includes('{contact_first_name}') || bodyTemplate.includes('{contact_last_name}'))) {
            errors.push(`No contact found for broker "${broker.brokerName}". Contact is needed for name placeholders.`);
        }

        if ((!broker.emailFilename || broker.emailFilename.trim() === '') && bodyTemplate.includes('{filename}')) {
            brokersNeedingFilenames.push(broker);
        }
    }

    // If brokers need filenames and user can provide them, show filename assignment modal
    if (brokersNeedingFilenames.length > 0 && errors.length === 0) {
        const proceed = await showFilenameAssignmentModal(brokersNeedingFilenames);
        if (!proceed) {
            errors.push('Filename assignment was cancelled.');
        }
    }

    return errors;
}

/**
 * Generate emails for all brokers with missing data
 */
async function generateBrokerEmails() {
    if (window.missingDataAnalysis.length === 0) {
        alert('No brokers with missing data found.');
        return;
    }

    // First, validate all placeholders
    const validationErrors = await validateEmailPlaceholders();
    if (validationErrors.length > 0) {
        alert('Please fix the following issues before generating emails:\n\n' + validationErrors.join('\n'));
        return;
    }

    window.pendingContactMatching = [];

    // Check which brokers need contact selection
    for (const broker of window.missingDataAnalysis) {
        const contact = await findBrokerContact(broker.brokerName);
        if (!contact) {
            window.pendingContactMatching.push(broker);
        }
    }

    if (window.pendingContactMatching.length > 0) {
        await handleContactMatching();
    } else {
        await generateAllEmails();
    }
}

/**
 * Generate all emails
 */
async function generateAllEmails() {
    const subject = document.getElementById('email-subject').value;
    const bodyTemplate = document.getElementById('email-body-template').value;

    let emailsGenerated = 0;

    for (const broker of window.missingDataAnalysis) {
        const contact = broker.manualContact || await findBrokerContact(broker.brokerName);

        if (contact) {
            await generateSingleEmail(broker, contact, subject, bodyTemplate);
            emailsGenerated++;
        }
    }

    alert(`Generated ${emailsGenerated} email(s). Check your email client.`);
}

/**
 * Generate single email for a broker
 */
async function generateSingleEmail(broker, contact, subject, bodyTemplate) {
    try {
        // Use template system for all email generation (including table)
        await createAndDownloadEMLFileFromTemplate(
            contact.email,
            subject,
            bodyTemplate,
            contact.firstName,
            contact.lastName,
            broker.emailFilename || broker.filename,
            broker.totalRows,
            broker.missingColumns, // Pass all required columns, not just incomplete ones
            window.appSettings.userSignature || window.appSettings.userName || 'User'
        );

        // Small delay between emails to avoid overwhelming the system
        await new Promise(resolve => setTimeout(resolve, 1000));

    } catch (error) {
        console.error('Error generating email for broker:', broker.brokerName, error);
    }
}

/**
 * Preview email for a specific broker
 */
async function previewEmail(brokerName) {
    const broker = window.missingDataAnalysis.find(b => b.brokerName === brokerName);
    if (!broker) return;

    const contact = broker.manualContact || await findBrokerContact(broker.brokerName);
    if (!contact) {
        alert('No contact found for this broker. Please run the email generation process to select a contact.');
        return;
    }

    const subject = document.getElementById('email-subject').value;
    const bodyTemplate = document.getElementById('email-body-template').value;

    // Use template system for email preview
    await createAndDownloadEMLFileFromTemplate(
        contact.email,
        subject,
        bodyTemplate,
        contact.firstName,
        contact.lastName,
        broker.emailFilename || broker.filename,
        broker.totalRows,
        broker.missingColumns, // Show ALL required columns, not just incomplete ones
        window.appSettings.userSignature || window.appSettings.userName || 'User'
    );
}

// ========== CONTACT MATCHING ==========

/**
 * Handle contact matching for brokers without exact matches
 */
async function handleContactMatching() {
    if (window.pendingContactMatching.length === 0) {
        await generateAllEmails();
        return;
    }

    const broker = window.pendingContactMatching[0];
    await showContactMatchingModal(broker);
}

/**
 * Show contact matching modal
 */
async function showContactMatchingModal(broker) {
    const modal = document.getElementById('contact-matching-modal');
    const messageEl = document.getElementById('contact-matching-message');
    const selector = document.getElementById('broker-contact-selector');

    messageEl.textContent = `No exact contact match found for broker "${broker.brokerName}". Please select a contact or skip this broker.`;

    // Populate contact selector
    const contacts = await window.loadAllBrokerContacts();
    selector.innerHTML = '<option value="">Choose a contact...</option>' +
        contacts.map(contact =>
            `<option value="${contact.id}">${contact.brokerName} - ${contact.firstName} ${contact.lastName}</option>`
        ).join('');

    modal.classList.add('show');
}

/**
 * Close contact matching modal
 */
function closeContactMatchingModal() {
    document.getElementById('contact-matching-modal').classList.remove('show');
}

/**
 * Skip current broker contact
 */
async function skipCurrentBrokerContact() {
    window.pendingContactMatching.shift(); // Remove first broker
    closeContactMatchingModal();
    await handleContactMatching(); // Continue with next broker
}

/**
 * Confirm broker contact selection
 */
async function confirmBrokerContact() {
    const selector = document.getElementById('broker-contact-selector');
    const selectedContactId = selector.value;

    if (!selectedContactId) {
        alert('Please select a contact or skip this broker.');
        return;
    }

    const broker = window.pendingContactMatching[0];
    const contacts = await window.loadAllBrokerContacts();
    const selectedContact = contacts.find(c => c.id === selectedContactId);

    if (selectedContact) {
        // Store the manual mapping for this broker
        broker.manualContact = selectedContact;
    }

    window.pendingContactMatching.shift(); // Remove first broker
    closeContactMatchingModal();
    await handleContactMatching(); // Continue with next broker
}

// ========== FILENAME ASSIGNMENT ==========

/**
 * Show filename assignment modal
 */
async function showFilenameAssignmentModal(brokers) {
    return new Promise((resolve) => {
        const modal = document.getElementById('filename-assignment-modal');
        const container = document.getElementById('filename-assignments-container');

        // Store resolve function for later use
        window.filenameAssignmentResolve = resolve;

        // Create form for each broker
        container.innerHTML = brokers.map(broker => `
            <div class="form-group" style="border: 1px solid #444; padding: 12px; border-radius: 4px; margin-bottom: 12px;">
                <label class="form-label">Filename for "${broker.brokerName}"</label>
                <input type="text" class="form-input filename-input" data-broker="${broker.brokerName}"
                       placeholder="Enter filename (e.g., ${broker.brokerName}_borderel.xlsx)"
                       value="${broker.primaryFilename || ''}" />
                <small style="color: #888; display: block; margin-top: 4px;">
                    Available files: ${broker.allFilenames.join(', ')}
                </small>
            </div>
        `).join('');

        modal.classList.add('show');
    });
}

/**
 * Close filename assignment modal
 */
function closeFilenameAssignmentModal() {
    document.getElementById('filename-assignment-modal').classList.remove('show');
    if (window.filenameAssignmentResolve) {
        window.filenameAssignmentResolve(false);
    }
}

/**
 * Skip filename assignment process
 */
function skipFilenameAssignmentProcess() {
    // Use default names (broker name + "_borderel.xlsx")
    const inputs = document.querySelectorAll('.filename-input');
    inputs.forEach(input => {
        const brokerName = input.dataset.broker;
        const defaultFilename = `${brokerName}_borderel.xlsx`;

        // Find the broker and set default filename
        const broker = window.missingDataAnalysis.find(b => b.brokerName === brokerName);
        if (broker) {
            broker.emailFilename = defaultFilename;
        }
    });

    closeFilenameAssignmentModal();
    if (window.filenameAssignmentResolve) {
        window.filenameAssignmentResolve(true);
    }
}

/**
 * Confirm filename assignment process
 */
function confirmFilenameAssignmentProcess() {
    const inputs = document.querySelectorAll('.filename-input');
    let allValid = true;

    inputs.forEach(input => {
        const brokerName = input.dataset.broker;
        const filename = input.value.trim();

        if (!filename) {
            allValid = false;
            return;
        }

        // Find the broker and set filename
        const broker = window.missingDataAnalysis.find(b => b.brokerName === brokerName);
        if (broker) {
            broker.emailFilename = filename;
        }
    });

    if (!allValid) {
        alert('Please provide filenames for all brokers or use "Use Default Names".');
        return;
    }

    closeFilenameAssignmentModal();
    if (window.filenameAssignmentResolve) {
        window.filenameAssignmentResolve(true);
    }
}

// ========== EMAIL CREATION ==========

/**
 * Create and download EML file using email template with placeholders
 */
async function createAndDownloadEMLFileFromTemplate(toEmail, subject, bodyTemplate, contactFirstName, contactLastName, filename, totalRows, missingColumns, userSignature) {
    // Process template with placeholders
    const htmlEmailBody = processEmailTemplate(bodyTemplate, {
        contact_email: toEmail,
        contact_first_name: contactFirstName,
        contact_last_name: contactLastName,
        filename: filename,
        total_rows: totalRows,
        missing_data_table: createMissingDataTableHTML(missingColumns),
        user_signature: userSignature
    });

    // Create EML file content with proper headers
    const emlContent = `To: ${toEmail}
From: ${window.appSettings.userEmail || 'noreply@borderellenconverter.nl'}
Subject: ${subject}
Date: ${new Date().toUTCString()}
MIME-Version: 1.0
Content-Type: text/html; charset=UTF-8
Content-Transfer-Encoding: 8bit

${htmlEmailBody}`;

    // Generate filename based on broker name with date and time
    const brokerName = filename.split(/[._]/)[0] || 'Email';
    const now = new Date();
    const timestamp = now.toISOString().slice(0, 19).replace(/[:.]/g, '-');
    const emlFilename = `${brokerName}_OntbrekendeData_${timestamp}.eml`;

    // Create blob and download EML file
    const blob = new Blob([emlContent], { type: 'message/rfc822' });

    // Download EML file using the existing download helper
    const success = await window.downloadToPreferredFolder(blob, emlFilename, 'message/rfc822');

    if (success) {
        console.log(`EML file created: ${emlFilename}. Please open manually from your downloads folder.`);
    }
}

/**
 * Process email template by replacing placeholders with actual values
 */
function processEmailTemplate(template, values) {
    // First, separate signature from main content BEFORE any processing
    let mainTemplate = template;
    let signatureTemplate = '';

    // Find signature placeholder and split there
    if (template.includes('{user_signature}')) {
        const signatureIndex = template.indexOf('{user_signature}');
        // Find the start of the line containing the signature
        const beforeSignature = template.substring(0, signatureIndex);
        const lastNewline = beforeSignature.lastIndexOf('\n');

        mainTemplate = template.substring(0, lastNewline >= 0 ? lastNewline : signatureIndex);
        signatureTemplate = template.substring(lastNewline >= 0 ? lastNewline : signatureIndex);
    }

    // Process main template (everything except signature)
    let processedMainContent = mainTemplate;
    Object.keys(values).forEach(placeholder => {
        if (placeholder === 'user_signature') return; // Skip signature in main content

        const value = values[placeholder];
        const regex = new RegExp(`\\{${placeholder}\\}`, 'g');

        if (placeholder === 'contact_email') {
            // Just insert the email address without special styling
            processedMainContent = processedMainContent.replace(regex, value);
        } else if (placeholder === 'missing_data_table') {
            // Insert the missing data table
            processedMainContent = processedMainContent.replace(regex, value);
        } else {
            // Regular placeholder replacement - all body text should be black
            let styledValue = value;
            if (placeholder === 'contact_first_name' || placeholder === 'contact_last_name' || placeholder === 'filename') {
                styledValue = `<strong style="color: black;">${value}</strong>`;
            } else if (placeholder === 'total_rows') {
                styledValue = `<strong style="color: black;">${value}</strong>`;
            }
            processedMainContent = processedMainContent.replace(regex, styledValue);
        }
    });

    // Process signature separately and keep it as-is (Allianz blue)
    let processedSignature = '';
    if (signatureTemplate.includes('{user_signature}')) {
        processedSignature = signatureTemplate.replace(/{user_signature}/g, createHTMLSignature(values.user_signature));
    }

    // Wrap main content in paragraphs, but preserve signature HTML structure
    const wrappedMainContent = processedMainContent.split('\n').map(line =>
        line.trim() ? `<p style="font-family: Arial, sans-serif; font-size: 10pt; color: black; margin: 0 0 4px 0;">${line.trim()}</p>` : '<br>'
    ).join('');

    return `<div style="font-family: Arial, sans-serif; font-size: 10pt; color: black; line-height: 1.6; max-width: 600px;">
${wrappedMainContent}${processedSignature}
</div>`;
}

/**
 * Create HTML for missing data table
 */
function createMissingDataTableHTML(allRequiredColumns) {
    if (!allRequiredColumns || allRequiredColumns.length === 0) {
        return `
<div style="background-color: #f1f8e9; padding: 15px; border-radius: 6px; border: 1px solid #c8e6c9; text-align: center; margin: 4px 0;">
    <h4 style="font-family: Arial, sans-serif; font-size: 10pt; color: rgb(0, 55, 129); margin: 0 0 8px 0; font-weight: bold;">
        ✅ Uitstekend!
    </h4>
    <p style="font-family: Arial, sans-serif; font-size: 10pt; color: #388e3c; margin: 0; font-weight: bold;">
        Alle verplichte velden zijn volledig ingevuld.
    </p>
</div>`;
    }

    // Calculate totals for the summary row
    let totalFilled = 0;
    let totalFields = 0;

    allRequiredColumns.forEach(col => {
        totalFilled += col.filledCount;
        totalFields += col.totalCount;
    });

    const overallPercentage = totalFields > 0 ? (totalFilled / totalFields * 100).toFixed(1) : '0.0';

    let tableHTML = `<div style="margin: 0 0 4px 0;">
    <table cellpadding="3" cellspacing="0" style="border-collapse: collapse; width: auto; font-family: Arial, sans-serif; font-size: 10pt; border: 1px solid #dee2e6;">
        <thead>
            <tr style="background-color: #e8f4f8;">
                <th style="text-align: left; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Gegevensveld</th>
                <th style="text-align: center; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Gevuld</th>
                <th style="text-align: center; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Totaal</th>
                <th style="text-align: center; padding: 3px 3px; font-weight: bold; color: black; border: 1px solid #dee2e6;">Percentage</th>
            </tr>
        </thead>
        <tbody>`;

    // Show ALL required columns (complete and incomplete)
    allRequiredColumns.forEach((col, index) => {
        const percentage = col.completionRate.toFixed(1);
        const rowBg = index % 2 === 0 ? '#ffffff' : '#f8f9fa';

        // Color-code the percentage based on completion
        const percentageColor = col.completionRate >= 100 ? '#4caf50' : col.completionRate >= 80 ? '#ffa726' : '#f44336';

        tableHTML += `
            <tr style="background-color: ${rowBg};">
                <td style="padding: 3px 3px; color: black; border: 1px solid #dee2e6;">${col.columnName}</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6;">${col.filledCount}</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6;">${col.totalCount}</td>
                <td style="text-align: center; padding: 3px 3px; color: ${percentageColor}; border: 1px solid #dee2e6; font-weight: bold;">${percentage}%</td>
            </tr>`;
    });

    // Add total score row
    const totalRowBg = '#e8f4f8';
    const totalPercentageColor = parseFloat(overallPercentage) >= 100 ? '#4caf50' : parseFloat(overallPercentage) >= 80 ? '#ffa726' : '#f44336';

    tableHTML += `
            <tr style="background-color: ${totalRowBg}; border-top: 2px solid #dee2e6;">
                <td style="padding: 3px 3px; color: black; border: 1px solid #dee2e6; font-weight: bold;">TOTAAL SCORE</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6; font-weight: bold;">${totalFilled}</td>
                <td style="text-align: center; padding: 3px 3px; color: black; border: 1px solid #dee2e6; font-weight: bold;">${totalFields}</td>
                <td style="text-align: center; padding: 3px 3px; color: ${totalPercentageColor}; border: 1px solid #dee2e6; font-weight: bold; font-size: 11pt;">${overallPercentage}%</td>
            </tr>`;

    tableHTML += `
        </tbody>
    </table>
</div>`;

    return tableHTML;
}

/**
 * Create HTML signature with Allianz branding
 */
function createHTMLSignature(userSignature) {
    // Parse signature for structured information
    const signatureLines = userSignature.split('\n').filter(line => line.trim());

    if (signatureLines.length === 0) {
        return `<p style="font-family: Arial, sans-serif; font-size: 10pt; color: rgb(0, 55, 129);">${userSignature}</p>`;
    }

    // Try to detect if signature has structured format (Name, Function, Company details)
    const hasStructuredFormat = signatureLines.length >= 2;

    if (hasStructuredFormat && signatureLines.length >= 3) {
        // Structured signature: Name, Function, Company details
        const [name, functionTitle, ...companyDetails] = signatureLines;

        return `
<div style="font-family: Arial, sans-serif; font-size: 10pt; margin-top: 4px;">
    <div style="color: rgb(0, 55, 129); font-weight: bold; font-size: 10pt; margin-bottom: 2px;">
        ${name}
    </div>
    <div style="color: rgb(0, 55, 129); font-size: 10pt; margin-bottom: 2px;">
        ${functionTitle}
    </div>
    <div style="color: rgb(0, 55, 129); font-size: 10pt; margin-bottom: 10px;">
        ${companyDetails.join('<br>')}
    </div>
    ${createSocialMediaLinks()}
</div>`;
    } else {
        // Simple signature - ALWAYS Allianz blue (NEVER black)
        return `
<div style="font-family: Arial, sans-serif; font-size: 10pt; margin-top: 4px;">
    <div style="color: rgb(0, 55, 129); font-weight: bold; font-size: 10pt; margin-bottom: 10px;">
        ${signatureLines[0]}
    </div>
    ${signatureLines.slice(1).map(line =>
        `<div style="color: rgb(0, 55, 129); font-size: 10pt; margin-bottom: 2px;">${line}</div>`
    ).join('')}
    ${createSocialMediaLinks()}
</div>`;
    }
}

/**
 * Create social media links section according to Allianz branding
 */
function createSocialMediaLinks() {
    // Social media links temporarily disabled - Belgian channels need proper icons
    return '';
}

// ========== UTILITY FUNCTIONS ==========

/**
 * Find the most appropriate filename for a broker
 */
function findBrokerSpecificFilename(brokerName, filenames) {
    // First, try to find a filename containing the broker name (case insensitive)
    const brokerNameLower = brokerName.toLowerCase();
    const brokerWords = brokerNameLower.split(/\s+/);

    for (const filename of filenames) {
        const filenameLower = filename.toLowerCase();

        // Check if filename contains the full broker name
        if (filenameLower.includes(brokerNameLower)) {
            return filename;
        }

        // Check if filename contains any significant words from broker name (length > 2)
        for (const word of brokerWords) {
            if (word.length > 2 && filenameLower.includes(word)) {
                return filename;
            }
        }
    }

    // If no match found, return null so primary filename is used
    return null;
}

/**
 * HTML escape utility function
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ========== EMAIL TEMPLATE MANAGEMENT ==========

/**
 * Save email template to settings
 */
async function saveEmailTemplate() {
    try {
        const subject = document.getElementById('email-subject').value;
        const body = document.getElementById('email-body-template').value;

        // Direct IndexedDB save without going through complex folder handling
        if (!window.db) await window.initIndexedDB();

        const transaction = window.db.transaction(['settings'], 'readwrite');
        const store = transaction.objectStore('settings');

        // Save email subject
        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'emailSubject', value: subject });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        // Save email body template
        await new Promise((resolve, reject) => {
            const request = store.put({ key: 'emailBodyTemplate', value: body });
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });

        // Update local appSettings
        window.appSettings.emailSubject = subject;
        window.appSettings.emailBodyTemplate = body;

        alert('Email template saved successfully!');
        console.log('Email template saved directly to IndexedDB');

    } catch (error) {
        console.error('Error saving email template:', error);
        alert('Error saving email template: ' + error.message);
    }
}

/**
 * Reset email template to default
 */
async function resetEmailTemplate() {
    if (confirm('Are you sure you want to reset the email template to default values?')) {
        const defaultSubject = 'Ontbrekende data in uw borderellen';
        const defaultBody = `{contact_email}

Geachte {contact_first_name} {contact_last_name},

Hartelijk dank voor het aanleveren van uw borderel. Deze informatie is erg waardevol voor ons.

Uit onze analyse is gebleken dat u al een groot deel van de benodigde informatie heeft aangeleverd om de boeking op de juiste manier voor u te verwerken. De enige informatie die wij nog nodig hebben om de boeking te kunnen verwerken is:

{missing_data_table}

Om een soepele en snelle afwikkeling te kunnen garanderen, verzoeken wij u vriendelijk de benodigde gegevens te vermelden in uw borderel volgens bijgevoegd format. Dit stelt ons in staat om de zaak efficiënt te verwerken en vertraging te voorkomen.

Bij voorbaat dank voor uw medewerking. Mocht u nog vragen hebben, dan horen wij dat graag.

Met vriendelijke groet,

{user_signature}`;

        // Update UI
        document.getElementById('email-subject').value = defaultSubject;
        document.getElementById('email-body-template').value = defaultBody;

        // Save to IndexedDB immediately
        try {
            window.appSettings.emailSubject = defaultSubject;
            window.appSettings.emailBodyTemplate = defaultBody;

            const success = await window.saveSettings(window.appSettings);
            if (success) {
                alert('Email template reset to default and saved successfully!');
            } else {
                alert('Email template was reset but could not be saved. Please try saving manually.');
            }
        } catch (error) {
            console.error('Error saving reset template:', error);
            alert('Email template was reset but could not be saved: ' + error.message);
        }
    }
}

/**
 * Load email template from settings
 */
function loadEmailTemplate() {
    console.log('Loading email template...');

    if (!window.appSettings) {
        console.warn('appSettings not yet loaded, skipping email template load');
        return;
    }

    console.log('appSettings.emailSubject:', window.appSettings.emailSubject);
    console.log('appSettings.emailBodyTemplate exists:', !!window.appSettings.emailBodyTemplate);
    console.log('Full appSettings keys:', Object.keys(window.appSettings));

    const subjectEl = document.getElementById('email-subject');
    const bodyEl = document.getElementById('email-body-template');

    console.log('Elements found - subject:', !!subjectEl, 'body:', !!bodyEl);

    if (window.appSettings.emailSubject && subjectEl) {
        subjectEl.value = window.appSettings.emailSubject;
        console.log('✅ Set subject to:', window.appSettings.emailSubject);
    } else {
        console.log('❌ Did not set subject - emailSubject:', window.appSettings.emailSubject, 'element:', !!subjectEl);
    }

    if (window.appSettings.emailBodyTemplate && bodyEl) {
        bodyEl.value = window.appSettings.emailBodyTemplate;
        console.log('✅ Set body template (first 50 chars):', window.appSettings.emailBodyTemplate.substring(0, 50));
    } else {
        console.log('❌ Did not set body - emailBodyTemplate exists:', !!window.appSettings.emailBodyTemplate, 'element:', !!bodyEl);
    }
}

// ========== GLOBAL EXPORTS ==========

// Make email functions globally accessible
window.initializeEmailBrokersTab = initializeEmailBrokersTab;
window.previewEmail = previewEmail;
window.closeContactMatchingModal = closeContactMatchingModal;
window.skipCurrentBrokerContact = skipCurrentBrokerContact;
window.confirmBrokerContact = confirmBrokerContact;
window.closeFilenameAssignmentModal = closeFilenameAssignmentModal;
window.skipFilenameAssignmentProcess = skipFilenameAssignmentProcess;
window.confirmFilenameAssignmentProcess = confirmFilenameAssignmentProcess;
window.loadEmailTemplate = loadEmailTemplate;