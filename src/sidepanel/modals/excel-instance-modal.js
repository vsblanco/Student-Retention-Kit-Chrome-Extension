// Excel Instance Modal - Allows selecting which Excel tab to use when multiple are open
import { elements } from '../ui-manager.js';

// Excel Instance Modal state
let excelInstanceResolve = null;

/**
 * Excel/SharePoint URL patterns to check for
 */
const EXCEL_URL_PATTERNS = [
    "https://excel.office.com/*",
    "https://*.officeapps.live.com/*",
    "https://*.sharepoint.com/*"
];

/**
 * Gets all open Excel tabs
 * @returns {Promise<Array>} Array of tab objects with id, title, and url
 */
export async function getExcelTabs() {
    try {
        const tabs = await chrome.tabs.query({ url: EXCEL_URL_PATTERNS });
        return tabs.map(tab => ({
            id: tab.id,
            title: tab.title || 'Untitled',
            url: tab.url
        }));
    } catch (error) {
        console.error('Error getting Excel tabs:', error);
        return [];
    }
}

/**
 * Opens the Excel Instance selection modal and returns a promise that resolves with the selected tab ID
 * @param {Array} excelTabs - Array of Excel tab objects
 * @param {string} [customMessage] - Optional custom message to display in the modal
 * @returns {Promise<number|null>} Promise that resolves with selected tab ID or null if cancelled
 */
export function openExcelInstanceModal(excelTabs, customMessage = null) {
    return new Promise((resolve) => {
        if (!elements.excelInstanceModal || !elements.excelInstanceList) {
            resolve(null);
            return;
        }

        // Store resolve function for later
        excelInstanceResolve = resolve;

        // Update message if custom message provided
        if (elements.excelInstanceMessage) {
            elements.excelInstanceMessage.textContent = customMessage ||
                'Multiple Excel instances detected. Select which one to send the master list to:';
        }

        // Clear existing buttons
        elements.excelInstanceList.innerHTML = '';

        // Create a button for each Excel tab
        excelTabs.forEach(tab => {
            const button = document.createElement('button');
            button.className = 'btn-secondary';
            button.style.cssText = 'width: 100%; text-align: left; padding: 12px 15px; display: flex; align-items: center; gap: 10px;';

            // Extract a cleaner name from the title
            let displayName = tab.title;
            // Remove common suffixes like "- Excel" or "- Microsoft Excel"
            displayName = displayName.replace(/\s*-\s*(Microsoft\s*)?Excel.*$/i, '');
            // Truncate if too long
            if (displayName.length > 50) {
                displayName = displayName.substring(0, 47) + '...';
            }

            button.innerHTML = `
                <i class="fas fa-file-excel" style="color: #217346; font-size: 1.2em;"></i>
                <span style="flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${displayName}</span>
            `;
            button.title = tab.title; // Full title on hover

            button.addEventListener('click', () => {
                closeExcelInstanceModal(tab.id);
            });

            elements.excelInstanceList.appendChild(button);
        });

        // Show modal
        elements.excelInstanceModal.style.display = 'flex';
    });
}

/**
 * Closes the Excel Instance selection modal
 * @param {number|null} selectedTabId - The selected tab ID or null if cancelled
 */
export function closeExcelInstanceModal(selectedTabId = null) {
    if (elements.excelInstanceModal) {
        elements.excelInstanceModal.style.display = 'none';
    }

    // Resolve the promise with the selected tab ID
    if (excelInstanceResolve) {
        excelInstanceResolve(selectedTabId);
        excelInstanceResolve = null;
    }
}
