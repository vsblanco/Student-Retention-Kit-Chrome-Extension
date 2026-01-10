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
import { STORAGE_KEYS } from '../constants/index.js';

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

        // Update body
        if (this.elements.tutorialBody) {
            this.elements.tutorialBody.innerHTML = page.body;
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
    }

    /**
     * Update button visibility based on page settings
     * @param {Object} page - The current page object
     */
    updateButtonVisibility(page) {
        // Skip button
        if (this.elements.tutorialSkipBtn) {
            this.elements.tutorialSkipBtn.style.display = page.showSkip ? 'block' : 'none';
        }

        // Previous button
        if (this.elements.tutorialPrevBtn) {
            this.elements.tutorialPrevBtn.style.display = page.showPrevious ? 'flex' : 'none';
        }

        // Next button
        if (this.elements.tutorialNextBtn) {
            this.elements.tutorialNextBtn.style.display = page.showNext ? 'flex' : 'none';

            // Update next button label
            const label = page.nextLabel || 'Next';
            const isLastPage = this.currentPageIndex === this.pages.length - 1;

            if (isLastPage) {
                this.elements.tutorialNextBtn.innerHTML = `Finish <i class="fas fa-check"></i>`;
            } else {
                this.elements.tutorialNextBtn.innerHTML = `${label} <i class="fas fa-arrow-right"></i>`;
            }
        }
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
