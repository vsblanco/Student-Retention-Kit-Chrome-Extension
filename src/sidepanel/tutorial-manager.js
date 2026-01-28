/**
 * Tutorial Manager
 *
 * Manages the tutorial flow for new users, including:
 * - Checking if user is new
 * - Displaying tutorial pages
 * - Handling page navigation
 * - Greying out tabs during tutorial
 * - Saving tutorial completion state
 */

import { TUTORIAL_PAGES, TUTORIAL_SETTINGS } from '../constants/tutorial.js';
import { STORAGE_KEYS, MESSAGE_TYPES, SHEET_DEFINITIONS } from '../constants/index.js';
import { checkExcelConnectionStatus, sendConnectionPing } from './excel-integration.js';

/**
 * Tutorial Manager Class
 */
class TutorialManager {
    constructor() {
        this.currentPageIndex = 0;
        this.pages = TUTORIAL_PAGES;
        this.isActive = false;

        // DOM elements
        this.elements = {
            tutorialPage: null,
            tutorialHeader: null,
            tutorialBody: null,
            tutorialProgress: null,
            tutorialProgressText: null,
            tutorialSkipBtn: null,
            tutorialPrevBtn: null,
            tutorialNextBtn: null,
            tabButtons: [],
            settingsBtn: null,
            versionText: null
        };
    }

    /**
     * Initialize the tutorial manager
     */
    async init() {
        // Cache DOM elements
        this.elements.tutorialPage = document.getElementById('tutorialPage');
        this.elements.tutorialHeader = document.getElementById('tutorialHeader');
        this.elements.tutorialBody = document.getElementById('tutorialBody');
        this.elements.tutorialProgress = document.getElementById('tutorialProgress');
        this.elements.tutorialProgressText = document.getElementById('tutorialProgressText');
        this.elements.tutorialSkipBtn = document.getElementById('tutorialSkipBtn');
        this.elements.tutorialPrevBtn = document.getElementById('tutorialPrevBtn');
        this.elements.tutorialNextBtn = document.getElementById('tutorialNextBtn');
        this.elements.tabButtons = Array.from(document.querySelectorAll('.tab-button'));
        this.elements.settingsBtn = document.getElementById('headerSettingsBtn');
        this.elements.versionText = document.getElementById('versionText');

        // Set up event listeners
        this.setupEventListeners();

        // Check if user should see tutorial
        const shouldShowTutorial = await this.shouldShowTutorial();

        if (shouldShowTutorial && TUTORIAL_SETTINGS.showForNewUsers) {
            await this.startTutorial();
        }
    }

    /**
     * Set up event listeners for tutorial navigation
     */
    setupEventListeners() {
        this.elements.tutorialSkipBtn?.addEventListener('click', () => this.skipTutorial());
        this.elements.tutorialPrevBtn?.addEventListener('click', () => this.previousPage());
        this.elements.tutorialNextBtn?.addEventListener('click', () => this.nextPage());

        // Listen for Excel connection messages to update status in real-time
        chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
            if (!this.isActive) return;

            // Check for Excel connection related messages
            if (message.type === MESSAGE_TYPES.SRK_OFFICE_ADDIN_CONNECTED ||
                message.type === MESSAGE_TYPES.SRK_TASKPANE_PONG) {

                // If we're on the Initial Setup page, update the connection status
                const currentPage = this.pages[this.currentPageIndex];
                if (currentPage && currentPage.id === 'initial-setup') {
                    this.updateExcelConnectionStatus();
                }
            }

            // Handle sheet list response
            if (message.type === MESSAGE_TYPES.SRK_SHEET_LIST_RESPONSE) {
                const currentPage = this.pages[this.currentPageIndex];
                if (currentPage && currentPage.id === 'initial-setup') {
                    this.updateSheetButtonsFromList(message.sheets);
                }
            }
        });
    }

    /**
     * Check if tutorial should be shown
     * @returns {Promise<boolean>} True if tutorial should be shown
     */
    async shouldShowTutorial() {
        try {
            const data = await chrome.storage.local.get([STORAGE_KEYS.TUTORIAL_COMPLETED]);
            return !data[STORAGE_KEYS.TUTORIAL_COMPLETED];
        } catch (error) {
            console.error('Error checking tutorial status:', error);
            return false;
        }
    }

    /**
     * Start the tutorial
     */
    async startTutorial() {
        this.isActive = true;
        this.currentPageIndex = 0;

        // Switch to tutorial tab
        this.switchToTutorialTab();

        // Grey out other tabs if enabled
        if (TUTORIAL_SETTINGS.greyOutTabs) {
            this.greyOutTabs();
        }

        // Render first page
        this.renderPage();
    }

    /**
     * Restart the tutorial (for users who want to review it)
     */
    async restartTutorial() {
        // Reset to first page and start tutorial
        this.currentPageIndex = 0;
        await this.startTutorial();
    }

    /**
     * Switch to the tutorial tab
     */
    switchToTutorialTab() {
        // Remove active class from all tabs and contents
        document.querySelectorAll('.tab-button').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

        // Show tutorial tab
        const tutorialContent = document.getElementById('tutorial');
        if (tutorialContent) {
            tutorialContent.classList.add('active');
        }
    }

    /**
     * Grey out all tab buttons, settings, and version buttons during tutorial
     */
    greyOutTabs() {
        this.elements.tabButtons.forEach(button => {
            button.classList.add('greyed-out');
        });

        // Disable settings button
        if (this.elements.settingsBtn) {
            this.elements.settingsBtn.classList.add('greyed-out');
            this.elements.settingsBtn.style.pointerEvents = 'none';
            this.elements.settingsBtn.style.opacity = '0.3';
        }

        // Disable version text
        if (this.elements.versionText) {
            this.elements.versionText.classList.add('greyed-out');
            this.elements.versionText.style.pointerEvents = 'none';
            this.elements.versionText.style.opacity = '0.3';
        }
    }

    /**
     * Un-grey all tab buttons and restore settings/version buttons
     */
    unGreyTabs() {
        this.elements.tabButtons.forEach(button => {
            button.classList.remove('greyed-out');
        });

        // Re-enable settings button
        if (this.elements.settingsBtn) {
            this.elements.settingsBtn.classList.remove('greyed-out');
            this.elements.settingsBtn.style.pointerEvents = '';
            this.elements.settingsBtn.style.opacity = '';
        }

        // Re-enable version text
        if (this.elements.versionText) {
            this.elements.versionText.classList.remove('greyed-out');
            this.elements.versionText.style.pointerEvents = '';
            this.elements.versionText.style.opacity = '';
        }
    }

    /**
     * Render the current tutorial page
     */
    renderPage() {
        const page = this.pages[this.currentPageIndex];
        if (!page) return;

        // Add animation class
        this.elements.tutorialPage?.classList.remove('tutorial-page-enter');
        void this.elements.tutorialPage?.offsetWidth; // Trigger reflow
        this.elements.tutorialPage?.classList.add('tutorial-page-enter');

        // Update header
        if (this.elements.tutorialHeader) {
            this.elements.tutorialHeader.textContent = page.header;
        }

        // Update body (support both string and function)
        if (this.elements.tutorialBody) {
            this.elements.tutorialBody.innerHTML = typeof page.body === 'function' ? page.body() : page.body;
        }

        // Update progress bar
        const progressPercent = ((this.currentPageIndex + 1) / this.pages.length) * 100;
        if (this.elements.tutorialProgress) {
            this.elements.tutorialProgress.style.width = `${progressPercent}%`;
        }

        // Update progress text
        if (this.elements.tutorialProgressText) {
            this.elements.tutorialProgressText.textContent =
                `Page ${this.currentPageIndex + 1} of ${this.pages.length}`;
        }

        // Update button visibility
        this.updateButtonVisibility(page);

        // If we're on the Initial Setup page, update Excel connection status
        if (page.id === 'initial-setup') {
            this.updateExcelConnectionStatus();
            this.setupSheetCreationButtons();

            // Send ping to ensure connection is active before requesting sheet list
            sendConnectionPing();

            // Request sheet list after a brief delay to allow ping to establish connection
            setTimeout(() => {
                this.requestSheetList();
            }, 100);
        }

        // If we're on the Personalized Emails page, setup preview buttons
        if (page.id === 'personalized-emails') {
            this.setupEmailPreviewButtons();
        }

        // If we're on the Missing Assignments page, setup preview buttons
        if (page.id === 'missing-assignments') {
            this.setupMissingAssignmentsPreviewButtons();
        }
    }

    /**
     * Update button visibility based on page settings
     * @param {Object} page - The current page object
     */
    updateButtonVisibility(page) {
        // All buttons are always visible

        // Previous button - disable on first page, enable otherwise
        if (this.elements.tutorialPrevBtn) {
            const isFirstPage = this.currentPageIndex === 0;
            this.elements.tutorialPrevBtn.disabled = isFirstPage;
        }

        // Next button - update label based on whether it's the last page
        if (this.elements.tutorialNextBtn) {
            const isLastPage = this.currentPageIndex === this.pages.length - 1;

            if (isLastPage) {
                this.elements.tutorialNextBtn.innerHTML = `Finish <i class="fas fa-check"></i>`;
            } else {
                const label = page.nextLabel || 'Next';
                this.elements.tutorialNextBtn.innerHTML = `${label} <i class="fas fa-arrow-right"></i>`;
            }
        }

        // Skip button is always visible and enabled
    }

    /**
     * Update Excel connection status on the Initial Setup page
     */
    async updateExcelConnectionStatus() {
        const statusElement = document.getElementById('tutorialExcelStatus');
        const sheetItems = document.querySelectorAll('.sheet-item');

        if (!statusElement) return;

        try {
            const status = await checkExcelConnectionStatus();

            // Update status text and styling
            if (status === 'connected') {
                statusElement.textContent = 'Connected';
                statusElement.className = 'connection-status connected';

                // Enable sheet items
                sheetItems.forEach(item => {
                    item.classList.remove('disabled');
                });
            } else {
                statusElement.textContent = 'Waiting...';
                statusElement.className = 'connection-status waiting';

                // Disable sheet items
                sheetItems.forEach(item => {
                    item.classList.add('disabled');
                });
            }
        } catch (error) {
            console.error('Error updating Excel connection status:', error);
            statusElement.textContent = 'Waiting...';
            statusElement.className = 'connection-status waiting';

            // Disable sheet items on error
            sheetItems.forEach(item => {
                item.classList.add('disabled');
            });
        }
    }

    /**
     * Set up event listeners for sheet creation buttons
     */
    setupSheetCreationButtons() {
        // Get all create buttons in the tutorial
        const createButtons = document.querySelectorAll('.tutorial-create-btn');

        createButtons.forEach(button => {
            // Clone and replace to remove any existing event listeners
            const newButton = button.cloneNode(true);
            button.parentNode.replaceChild(newButton, button);

            // Add click event listener
            newButton.addEventListener('click', (e) => {
                e.preventDefault();

                // Determine which sheet to create based on parent item ID
                const parentItem = newButton.closest('.sheet-item');
                if (!parentItem) return;

                let sheetDefinition = null;

                if (parentItem.id === 'masterListItem') {
                    sheetDefinition = SHEET_DEFINITIONS.MASTER_LIST;
                } else if (parentItem.id === 'studentHistoryItem') {
                    sheetDefinition = SHEET_DEFINITIONS.STUDENT_HISTORY;
                } else if (parentItem.id === 'missingAssignmentsItem') {
                    sheetDefinition = SHEET_DEFINITIONS.MISSING_ASSIGNMENTS;
                }

                if (sheetDefinition) {
                    this.sendCreateSheetMessage(sheetDefinition);
                }
            });
        });
    }

    /**
     * Set up event listeners for email preview buttons
     */
    setupEmailPreviewButtons() {
        const previewButtons = document.querySelectorAll('.email-preview-btn');
        const emailTemplateBox = document.getElementById('emailTemplateBox');

        if (!emailTemplateBox) return;

        // Template HTML with parameters
        const templateHTML = `
            <p style="margin: 0 0 10px 0;">Hi <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{First}</span>,</p>
            <p style="margin: 0 0 10px 0;">I hope your week is going well!</p>
            <p style="margin: 0 0 10px 0;">I was looking over your Canvas course today and noticed it has been <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{DaysOut}</span> days since you last engaged. You currently have a <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{Grade}</span> in the class, and I want to make sure we keep that momentum going so you can successfully complete the course on time.</p>
            <p style="margin: 0 0 10px 0;">I saw that the last assignment you submitted was <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{LastAssignment}</span>. Would you be open to a quick call today? I'd love to discuss a game plan to help you tackle the upcoming workload and ensure you feel supported. Reminder that your next assignment is due on <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{NextAssignmentDue}</span>.</p>
            <p style="margin: 0 0 10px 0;">Best regards,</p>
            <p style="margin: 0;"><span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{AssignedName}</span></p>
        `;

        // Preview data
        const previews = {
            preview1: `
                <p style="margin: 0 0 10px 0;">Hi <strong>Sarah</strong>,</p>
                <p style="margin: 0 0 10px 0;">I hope your week is going well!</p>
                <p style="margin: 0 0 10px 0;">I was looking over your Canvas course today and noticed it has been <strong>7</strong> days since you last engaged. You currently have a <strong>78%</strong> in the class, and I want to make sure we keep that momentum going so you can successfully complete the course on time.</p>
                <p style="margin: 0 0 10px 0;">I saw that the last assignment you submitted was <strong>Discussion Post 1.0</strong>. Would you be open to a quick call today? I'd love to discuss a game plan to help you tackle the upcoming workload and ensure you feel supported. Reminder that your next assignment is due on <strong>January 18, 2026</strong>.</p>
                <p style="margin: 0 0 10px 0;">Best regards,</p>
                <p style="margin: 0;"><strong>Advisor 1</strong></p>
            `,
            preview2: `
                <p style="margin: 0 0 10px 0;">Hi <strong>Marcus</strong>,</p>
                <p style="margin: 0 0 10px 0;">I hope your week is going well!</p>
                <p style="margin: 0 0 10px 0;">I was looking over your Canvas course today and noticed it has been <strong>12</strong> days since you last engaged. You currently have a <strong>62%</strong> in the class, and I want to make sure we keep that momentum going so you can successfully complete the course on time.</p>
                <p style="margin: 0 0 10px 0;">I saw that the last assignment you submitted was <strong>Week 3 Quiz</strong>. Would you be open to a quick call today? I'd love to discuss a game plan to help you tackle the upcoming workload and ensure you feel supported. Reminder that your next assignment is due on <strong>January 15, 2026</strong>.</p>
                <p style="margin: 0 0 10px 0;">Best regards,</p>
                <p style="margin: 0;"><strong>Advisor 2</strong></p>
            `
        };

        previewButtons.forEach(button => {
            // Clone and replace to remove any existing event listeners
            const newButton = button.cloneNode(true);
            button.parentNode.replaceChild(newButton, button);

            // Add click event listener
            newButton.addEventListener('click', (e) => {
                e.preventDefault();

                const previewType = newButton.dataset.preview;
                const isActive = newButton.dataset.active === 'true';

                if (isActive) {
                    // Clicking the same button - deactivate and show template
                    emailTemplateBox.innerHTML = templateHTML;
                    emailTemplateBox.style.fontFamily = 'monospace';
                    emailTemplateBox.style.color = '#374151';
                    newButton.style.background = 'rgba(0,0,0,0.06)';
                    newButton.style.color = 'inherit';
                    newButton.dataset.active = 'false';
                } else {
                    // Clicking a different button or activating from template
                    // First, deactivate all buttons
                    const allButtons = document.querySelectorAll('.email-preview-btn');
                    allButtons.forEach(btn => {
                        btn.style.background = 'rgba(0,0,0,0.06)';
                        btn.style.color = 'inherit';
                        btn.dataset.active = 'false';
                    });

                    // Then activate the clicked button and show its preview
                    emailTemplateBox.innerHTML = previews[previewType];
                    emailTemplateBox.style.fontFamily = 'inherit';
                    emailTemplateBox.style.color = 'inherit';
                    newButton.style.background = 'var(--primary-color)';
                    newButton.style.color = 'white';
                    newButton.dataset.active = 'true';
                }
            });
        });
    }

    /**
     * Set up event listeners for missing assignments preview buttons
     */
    setupMissingAssignmentsPreviewButtons() {
        const previewButtons = document.querySelectorAll('.missing-preview-btn');
        const missingTemplateBox = document.getElementById('missingTemplateBox');

        if (!missingTemplateBox) return;

        // Template HTML with parameters
        const templateHTML = `
            <p style="margin: 0 0 10px 0;">Hello <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{First}</span>,</p>
            <p style="margin: 0 0 10px 0;">I hope you are having a great day. I noticed it's been <span style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500;">{DaysOut}</span> days since you last engaged with your course.</p>
            <p style="margin: 0 0 5px 0;">You are currently missing:</p>
            <div style="background-color: #fef3c7; padding: 2px 4px; border-radius: 3px; font-weight: 500; display: inline-block;">{MissingAssignmentsList}</div>
        `;

        // Preview data
        const previews = {
            preview1: `
                <p style="margin: 0 0 10px 0;">Hello <strong>Jane</strong>,</p>
                <p style="margin: 0 0 10px 0;">I hope you are having a great day. I noticed it's been <strong>5</strong> days since you last engaged with your course.</p>
                <p style="margin: 0 0 5px 0;">You are currently missing:</p>
                <ul style="margin: 5px 0 10px 20px; padding: 0;">
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">Discussion Post 2.0</span></li>
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">Mid Term</span></li>
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">Mind Tap 2.2 - Human Anatomy</span></li>
                </ul>
            `,
            preview2: `
                <p style="margin: 0 0 10px 0;">Hello <strong>Michael</strong>,</p>
                <p style="margin: 0 0 10px 0;">I hope you are having a great day. I noticed it's been <strong>8</strong> days since you last engaged with your course.</p>
                <p style="margin: 0 0 5px 0;">You are currently missing:</p>
                <ul style="margin: 5px 0 10px 20px; padding: 0;">
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">1.6 MindTap Assignment</span></li>
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">2.2 Lab</span></li>
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">2.4 Mid Term</span></li>
                    <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">3.3 WebAssign Laboratory</span></li>
                </ul>
            `
        };

        previewButtons.forEach(button => {
            // Clone and replace to remove any existing event listeners
            const newButton = button.cloneNode(true);
            button.parentNode.replaceChild(newButton, button);

            // Add click event listener
            newButton.addEventListener('click', (e) => {
                e.preventDefault();

                const previewType = newButton.dataset.preview;
                const isActive = newButton.dataset.active === 'true';

                if (isActive) {
                    // Clicking the same button - deactivate and show template
                    missingTemplateBox.innerHTML = templateHTML;
                    missingTemplateBox.style.fontFamily = 'monospace';
                    missingTemplateBox.style.color = '#374151';
                    newButton.style.background = 'rgba(0,0,0,0.06)';
                    newButton.style.color = 'inherit';
                    newButton.dataset.active = 'false';
                } else {
                    // Clicking a different button or activating from template
                    // First, deactivate all buttons
                    const allButtons = document.querySelectorAll('.missing-preview-btn');
                    allButtons.forEach(btn => {
                        btn.style.background = 'rgba(0,0,0,0.06)';
                        btn.style.color = 'inherit';
                        btn.dataset.active = 'false';
                    });

                    // Then activate the clicked button and show its preview
                    missingTemplateBox.innerHTML = previews[previewType];
                    missingTemplateBox.style.fontFamily = 'inherit';
                    missingTemplateBox.style.color = 'inherit';
                    newButton.style.background = 'var(--primary-color)';
                    newButton.style.color = 'white';
                    newButton.dataset.active = 'true';
                }
            });
        });
    }

    /**
     * Send SRK_CREATE_SHEET message to Excel add-in
     * @param {Object} sheetDefinition - Sheet definition with name and headers
     */
    async sendCreateSheetMessage(sheetDefinition) {
        try {
            const payload = {
                type: MESSAGE_TYPES.SRK_CREATE_SHEET,
                sheetName: sheetDefinition.name,
                headers: sheetDefinition.headers
            };

            console.log(`📊 Creating sheet: ${sheetDefinition.name}`, sheetDefinition.headers);

            // Send message to background script, which will relay to Excel
            await chrome.runtime.sendMessage({
                type: MESSAGE_TYPES.SRK_CREATE_SHEET,
                payload: payload
            });

            console.log(`✅ Sheet creation request sent: ${sheetDefinition.name}`);

            // Request updated sheet list after a short delay to allow Excel to create the sheet
            setTimeout(() => {
                this.requestSheetList();
            }, 500);
        } catch (error) {
            console.error('Error sending create sheet message:', error);
        }
    }

    /**
     * Request list of sheets from Excel workbook
     */
    async requestSheetList() {
        try {
            const payload = {
                type: MESSAGE_TYPES.SRK_REQUEST_SHEET_LIST
            };

            console.log('📊 Requesting sheet list from Excel workbook');

            // Send message to background script, which will relay to Excel
            await chrome.runtime.sendMessage({
                type: MESSAGE_TYPES.SRK_REQUEST_SHEET_LIST,
                payload: payload
            });
        } catch (error) {
            console.error('Error requesting sheet list:', error);
        }
    }

    /**
     * Update sheet creation buttons based on existing sheets
     * @param {Array<string>} sheets - Array of sheet names from Excel workbook
     */
    updateSheetButtonsFromList(sheets) {
        if (!sheets || !Array.isArray(sheets)) {
            console.warn('Invalid sheet list received:', sheets);
            return;
        }

        console.log('📋 Received sheet list:', sheets);

        // Check each sheet definition and update buttons
        const sheetMappings = [
            { itemId: 'masterListItem', sheetName: SHEET_DEFINITIONS.MASTER_LIST.name },
            { itemId: 'studentHistoryItem', sheetName: SHEET_DEFINITIONS.STUDENT_HISTORY.name },
            { itemId: 'missingAssignmentsItem', sheetName: SHEET_DEFINITIONS.MISSING_ASSIGNMENTS.name }
        ];

        sheetMappings.forEach(mapping => {
            const item = document.getElementById(mapping.itemId);
            if (!item) return;

            const button = item.querySelector('.tutorial-create-btn');
            if (!button) return;

            // Check if sheet exists in the workbook
            const sheetExists = sheets.includes(mapping.sheetName);

            if (sheetExists) {
                // Sheet exists - show "Created" and disable button
                button.textContent = 'Created';
                button.disabled = true;
                button.style.opacity = '0.6';
                button.style.cursor = 'not-allowed';
                button.classList.add('sheet-created');
            } else {
                // Sheet doesn't exist - show "Create" and enable button
                button.textContent = 'Create';
                button.disabled = false;
                button.style.opacity = '';
                button.style.cursor = '';
                button.classList.remove('sheet-created');
            }
        });
    }

    /**
     * Go to the next page
     */
    nextPage() {
        if (this.currentPageIndex < this.pages.length - 1) {
            this.currentPageIndex++;
            this.renderPage();
        } else {
            // Last page - finish tutorial
            this.completeTutorial();
        }
    }

    /**
     * Go to the previous page
     */
    previousPage() {
        if (this.currentPageIndex > 0) {
            this.currentPageIndex--;
            this.renderPage();
        }
    }

    /**
     * Skip the tutorial
     */
    async skipTutorial() {
        await this.completeTutorial();
    }

    /**
     * Complete the tutorial and unlock the extension
     */
    async completeTutorial() {
        this.isActive = false;

        // Save tutorial completion state
        await chrome.storage.local.set({
            [STORAGE_KEYS.TUTORIAL_COMPLETED]: true
        });

        // Un-grey tabs
        this.unGreyTabs();

        // Switch to checker tab
        const checkerTab = document.querySelector('.tab-button[data-tab="checker"]');
        if (checkerTab) {
            checkerTab.click();
        }
    }

    /**
     * Reset tutorial (for testing or re-showing)
     */
    async resetTutorial() {
        await chrome.storage.local.set({
            [STORAGE_KEYS.TUTORIAL_COMPLETED]: false
        });
        this.currentPageIndex = 0;
    }

    /**
     * Check if tutorial is currently active
     * @returns {boolean} True if tutorial is active
     */
    isActiveTutorial() {
        return this.isActive;
    }
}

// Export singleton instance
export const tutorialManager = new TutorialManager();
