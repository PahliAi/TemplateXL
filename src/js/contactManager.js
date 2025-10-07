/**
 * Borderellen Converter - Contact Manager Module
 * Handles all contact management functionality including CRUD operations
 */

// Global contact state variables
window.currentEditingContactId = null;
window.allContacts = [];

// ========== CONTACT MANAGEMENT FUNCTIONS ==========

/**
 * Load and display contacts
 */
async function loadAndDisplayContacts() {
    try {
        window.allContacts = await window.loadAllBrokerContacts();
        displayContactsTable(window.allContacts);
    } catch (error) {
        console.error('Error loading contacts:', error);
        alert('Error loading contacts: ' + error.message);
    }
}

/**
 * Display contacts in the table
 */
function displayContactsTable(contacts) {
    const noContactsMessage = document.getElementById('no-contacts-message');
    const tableContainer = document.getElementById('contacts-table-container');
    const tableBody = document.getElementById('contacts-table-body');

    if (!contacts || contacts.length === 0) {
        noContactsMessage.style.display = 'block';
        tableContainer.style.display = 'none';
        return;
    }

    noContactsMessage.style.display = 'none';
    tableContainer.style.display = 'block';

    tableBody.innerHTML = contacts.map(contact => `
        <tr>
            <td>${escapeHtml(contact.brokerName)}</td>
            <td>${escapeHtml(contact.firstName)} ${escapeHtml(contact.lastName)}</td>
            <td>${escapeHtml(contact.email)}</td>
            <td>
                <button class="btn btn-secondary" onclick="editContact('${contact.id}')" style="margin-right: 8px; font-size: 12px; padding: 4px 8px;">Edit</button>
                <button class="btn" onclick="deleteContact('${contact.id}')" style="background: #dc3545; border-color: #dc3545; font-size: 12px; padding: 4px 8px;">Delete</button>
            </td>
        </tr>
    `).join('');
}

/**
 * Save contact (create or update)
 */
async function saveContact(event) {
    event.preventDefault();

    const brokerName = document.getElementById('contact-broker-name').value.trim();
    const firstName = document.getElementById('contact-first-name').value.trim();
    const lastName = document.getElementById('contact-last-name').value.trim();
    const email = document.getElementById('contact-email').value.trim();

    if (!brokerName || !firstName || !lastName || !email) {
        alert('Please fill in all required fields.');
        return;
    }

    // Email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        alert('Please enter a valid email address.');
        return;
    }

    const contact = {
        id: window.currentEditingContactId || `contact-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        brokerName,
        firstName,
        lastName,
        email,
        created: window.currentEditingContactId ?
            window.allContacts.find(c => c.id === window.currentEditingContactId)?.created || new Date().toISOString() :
            new Date().toISOString(),
        lastModified: new Date().toISOString()
    };

    try {
        const success = await window.saveBrokerContact(contact);
        if (success) {
            resetContactForm();
            await loadAndDisplayContacts();
        } else {
            alert('Error saving contact. Please try again.');
        }
    } catch (error) {
        console.error('Error saving contact:', error);
        alert('Error saving contact: ' + error.message);
    }
}

/**
 * Edit contact
 */
function editContact(contactId) {
    const contact = window.allContacts.find(c => c.id === contactId);
    if (!contact) {
        alert('Contact not found.');
        return;
    }

    window.currentEditingContactId = contactId;

    document.getElementById('contact-broker-name').value = contact.brokerName;
    document.getElementById('contact-first-name').value = contact.firstName;
    document.getElementById('contact-last-name').value = contact.lastName;
    document.getElementById('contact-email').value = contact.email;

    document.getElementById('form-section-title').textContent = 'Edit Contact';
    document.getElementById('save-contact-btn').textContent = 'Update Contact';

    // Scroll to form for better visibility
    document.getElementById('contact-form-section').scrollIntoView({ behavior: 'smooth' });
}

/**
 * Delete contact
 */
async function deleteContact(contactId) {
    const contact = window.allContacts.find(c => c.id === contactId);
    if (!contact) {
        alert('Contact not found.');
        return;
    }

    if (!confirm(`Are you sure you want to delete the contact for ${contact.firstName} ${contact.lastName} (${contact.brokerName})?`)) {
        return;
    }

    try {
        const success = await window.deleteBrokerContact(contactId);
        if (success) {
            await loadAndDisplayContacts();
            alert('Contact deleted successfully!');
        } else {
            alert('Error deleting contact. Please try again.');
        }
    } catch (error) {
        console.error('Error deleting contact:', error);
        alert('Error deleting contact: ' + error.message);
    }
}

/**
 * Reset contact form
 */
function resetContactForm() {
    window.currentEditingContactId = null;
    document.getElementById('contact-form').reset();
    document.getElementById('form-section-title').textContent = 'Add New Contact';
    document.getElementById('save-contact-btn').textContent = 'Save Contact';
}

/**
 * Cancel contact form
 */
function cancelContactForm() {
    resetContactForm();
}

/**
 * Export contacts as JSON
 */
async function exportContacts() {
    try {
        if (!window.allContacts || window.allContacts.length === 0) {
            alert('No contacts to export.');
            return;
        }

        await window.exportBrokerContactsAsJSON(window.allContacts, null, window.appSettings);
        alert(`Exported ${window.allContacts.length} contacts successfully!`);
    } catch (error) {
        console.error('Error exporting contacts:', error);
        alert('Error exporting contacts: ' + error.message);
    }
}

/**
 * Import contacts from JSON
 */
async function importContacts(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        const result = await window.loadBrokerContactsFromJSON(file);

        if (result.success) {
            // Save imported contacts to IndexedDB
            let savedCount = 0;
            for (const contact of result.contacts) {
                const success = await window.saveBrokerContact(contact);
                if (success) savedCount++;
            }

            // Refresh contact list
            await loadAndDisplayContacts();

            let message = `Successfully imported ${savedCount} contacts.`;
            if (result.skipped > 0) {
                message += `\n${result.skipped} contacts were skipped due to missing required fields.`;
            }
            alert(message);
        } else {
            alert('Error importing contacts: ' + result.error);
        }
    } catch (error) {
        console.error('Error importing contacts:', error);
        alert('Error importing contacts: ' + error.message);
    }

    // Reset file input
    event.target.value = '';
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ========== GLOBAL EXPORTS ==========

// Contact management functions
window.loadAndDisplayContacts = loadAndDisplayContacts;
window.displayContactsTable = displayContactsTable;
window.saveContact = saveContact;
window.editContact = editContact;
window.deleteContact = deleteContact;
window.resetContactForm = resetContactForm;
window.cancelContactForm = cancelContactForm;
window.exportContacts = exportContacts;
window.importContacts = importContacts;
window.escapeHtml = escapeHtml;