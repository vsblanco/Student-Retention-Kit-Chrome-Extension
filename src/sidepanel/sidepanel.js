// Sidepanel Main - Orchestrates all modules and manages app lifecycle
import { STORAGE_KEYS, EXTENSION_STATES, MESSAGE_TYPES, GUIDES, UI_FEATURES } from '../constants/index.js';
import { hasDispositionCode } from '../constants/dispositions.js';
import { getCacheStats, clearAllCache } from '../utils/canvasCache.js';
import { loadAndRenderMarkdown } from '../utils/markdownRenderer.js';
import CallManager from './callManager.js';
import { tutorialManager } from './tutorial-manager.js';

// Import all module functions
import {
    elements,
    cacheDomElements,
    switchTab,
    updateTabBadge,
    updateButtonVisuals,
    updateDebugModeUI,
    updateEmbedHelperUI,
    updateHighlightColorUI,
    blockTextSelection
} from './ui-manager.js';

import {
    setActiveStudent,
    renderFoundList,
    filterFoundList,
    renderMasterList,
    filterMasterList,
    sortMasterList
} from './student-renderer.js';

import {
    handleFileImport,
    resetQueueUI,
    restoreDefaultQueueUI,
    exportMasterListCSV,
    sendMasterListToExcel,
    sendMasterListWithMissingAssignmentsToExcel
} from './file-handler.js';

import { processStep2, processStep3, processStep4 } from './canvas-integration.js';

import {
    openScanFilterModal,
    closeScanFilterModal,
    updateScanFilterCount,
    toggleFailingFilter,
    saveScanFilterSettings,
    openQueueModal,
    closeQueueModal,
    renderQueueModal,
    openConnectionsModal,
    closeConnectionsModal,
    saveConnectionsSettings,
    updatePowerAutomateStatus,
    updateCanvasStatus,
    updateFive9Status,
    toggleEmbedHelperModal,
    toggleDebugModeModal,
    toggleSyncActiveStudentModal,
    toggleSendMasterListModal,
    toggleReformatNameModal,
    toggleHighlightStudentRowModal,
    clearCacheFromModal,
    shouldShowDailyUpdateModal,
    openDailyUpdateModal,
    closeDailyUpdateModal
} from './modal-manager.js';

import { QueueManager } from './queue-manager.js';

import {
    updateFive9ConnectionIndicator,
    startFive9ConnectionMonitor,
    stopFive9ConnectionMonitor,
    setupFive9StatusListeners
} from './five9-integration.js';

import {
    startExcelConnectionMonitor,
    stopExcelConnectionMonitor,
    pingExcelAddIn,
    sendConnectionPing
} from './excel-integration.js';

// --- STATE MANAGEMENT ---
let isScanning = false;
let callManager;
let queueManager;
let isDebugMode = false;
let embedHelperEnabled = true;
let highlightColor = '#ffff00';

// --- RESEND HIGHLIGHT PING FUNCTIONS ---
async function resendHighlightPing(entry) {
    await chrome.runtime.sendMessage({
        type: MESSAGE_TYPES.RESEND_HIGHLIGHT_PING,
        entry: entry
    });
}

async function resendAllHighlightPings() {
    await chrome.runtime.sendMessage({
        type: MESSAGE_TYPES.RESEND_ALL_HIGHLIGHT_PINGS
    });
}

// --- INITIALIZATION ---
document.addEventListener('DOMContentLoaded', () => {
    blockTextSelection();
    cacheDomElements();
    initializeApp();
});

async function initializeApp() {
    // Set version from manifest
    const manifest = chrome.runtime.getManifest();
    if (elements.versionText && manifest.version) {
        elements.versionText.textContent = `Version ${manifest.version}`;
    }

    // Initialize call manager with UI callbacks
    const uiCallbacks = {
        updateCurrentStudent: (student) => {
            setActiveStudent(student, callManager);
        },
        finalizeAutomation: (lastStudent) => {
            queueManager.setQueue([lastStudent]);
            setActiveStudent(lastStudent, callManager);
        },
        cancelAutomation: (currentStudent) => {
            queueManager.setQueue([currentStudent]);
            setActiveStudent(currentStudent, callManager);
        }
    };
    callManager = new CallManager(elements, uiCallbacks);

    // Initialize queue manager
    queueManager = new QueueManager(callManager);

    // Ensure checker is stopped when side panel opens
    await chrome.storage.local.set({ [STORAGE_KEYS.EXTENSION_STATE]: EXTENSION_STATES.OFF });

    setupEventListeners();
    initializeCallControlButtons();
    await loadStorageData();
    setActiveStudent(null, callManager);
    populateGuides();

    // Load and display last call timestamp
    await callManager.loadLastCallTimestamp();

    // Start Five9 connection monitoring
    startFive9ConnectionMonitor(() => queueManager.getQueue());

    // Check Five9 connection immediately on sidepanel open
    updateFive9ConnectionIndicator(queueManager.getQueue());

    // Setup Five9 status listeners
    setupFive9StatusListeners(callManager);

    // Start Excel connection monitoring
    startExcelConnectionMonitor();

    // Send ping to Excel add-in to check connectivity on taskpane open
    await pingExcelAddIn();

    // Send simple SRK_PING to instantly test connection when side panel opens
    await sendConnectionPing();

    // Initialize tutorial manager (must be done before daily update modal)
    await tutorialManager.init();

    // Check if daily update modal should be shown (only if tutorial is not active)
    if (!tutorialManager.isActiveTutorial()) {
        const showModal = await shouldShowDailyUpdateModal();
        if (showModal) {
            openDailyUpdateModal();
        }
    }

    // Periodically check Canvas connection status (every 5 seconds)
    setInterval(async () => {
        await updateCanvasStatus();
    }, 5000);
}

// --- ABOUT PAGE ---
let aboutContentLoaded = false;
async function loadAboutContent() {
    // Only load once
    if (aboutContentLoaded) return;

    const aboutContainer = document.getElementById('aboutContent');
    if (!aboutContainer) {
        console.error('About content container not found');
        return;
    }

    try {
        await loadAndRenderMarkdown('../../README.md', aboutContainer);
        aboutContentLoaded = true;
    } catch (error) {
        console.error('Failed to load about content:', error);
    }
}

// --- GUIDES ---
/**
 * Populates the guides section with PDF links from GUIDES constant
 */
function populateGuides() {
    const guidesContainer = document.getElementById('guidesContainer');
    if (!guidesContainer) {
        console.error('Guides container not found');
        return;
    }

    // Clear existing content
    guidesContainer.innerHTML = '';

    // Create guide cards
    GUIDES.forEach(guide => {
        const guideCard = document.createElement('div');
        guideCard.className = 'setting-card';
        guideCard.style.cursor = 'pointer';
        guideCard.innerHTML = `
            <div style="display:flex; align-items:center; gap:10px; flex-grow:1;">
                <i class="fas ${guide.icon}" style="color:var(--primary-color); font-size:1.5em; width:28px; text-align:center;"></i>
                <span style="font-weight:500;">${guide.name}</span>
            </div>
            <i class="fas fa-external-link-alt" style="color:var(--text-secondary); font-size:1em;"></i>
        `;

        // Add click handler to open PDF in new tab
        guideCard.addEventListener('click', () => {
            const guideUrl = chrome.runtime.getURL(guide.path);
            chrome.tabs.create({ url: guideUrl });
        });

        guidesContainer.appendChild(guideCard);
    });
}

// --- CONSOLE MESSAGE HANDLER ---
function addConsoleMessage(type, args) {
    if (!elements.consoleContent) return;

    const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const message = Array.from(args).map(arg => {
        if (typeof arg === 'object') {
            try {
                return JSON.stringify(arg, null, 2);
            } catch (e) {
                return String(arg);
            }
        }
        return String(arg);
    }).join(' ');

    // Detect specific message patterns and apply custom colors
    let customType = type;
    if (message.includes('Sending payload to Office Add-in') || message.includes('SRK_HIGHLIGHT_STUDENT_ROW')) {
        customType = 'ping';
    } else if (message.includes('onSubmissionFound triggered') || message.includes('Submission Found')) {
        customType = 'submission';
    }

    const logEntry = document.createElement('div');
    logEntry.className = `console-log ${customType}`;
    logEntry.innerHTML = `<span class="timestamp">[${timestamp}]</span>${message}`;

    elements.consoleContent.appendChild(logEntry);
    elements.consoleContent.scrollTop = elements.consoleContent.scrollHeight;

    // Limit to 100 entries
    const entries = elements.consoleContent.querySelectorAll('.console-log');
    if (entries.length > 100) {
        entries[0].remove();
    }
}

// --- EVENT LISTENERS ---
function setupEventListeners() {
    // Tab switching
    elements.tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            switchTab(tab.dataset.tab);
            if (tab.dataset.tab === 'settings') {
                updateCacheStats();
            } else if (tab.dataset.tab === 'about') {
                loadAboutContent();
            } else if (tab.dataset.tab === 'contact') {
                // Check Five9 connection when switching to contact tab
                updateFive9ConnectionIndicator(queueManager.getQueue());
            }
        });
    });

    // CTRL key release detection for automation mode
    document.addEventListener('keyup', (e) => {
        if (e.key === 'Control' || e.key === 'Meta') {
            if (queueManager.getLength() > 1) {
                switchTab('contact');
                // Check Five9 connection when switching to contact tab
                updateFive9ConnectionIndicator(queueManager.getQueue());
            }
        }
    });

    // Header and modals
    if (elements.headerSettingsBtn) {
        elements.headerSettingsBtn.addEventListener('click', () => switchTab('settings'));
    }

    if (elements.versionText) {
        elements.versionText.addEventListener('click', () => {
            switchTab('about');
            loadAboutContent();
        });
    }

    if (elements.clearMasterListBtn) {
        elements.clearMasterListBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to clear the master list? This cannot be undone.')) {
                chrome.storage.local.set({
                    [STORAGE_KEYS.MASTER_ENTRIES]: [],
                    [STORAGE_KEYS.LAST_UPDATED]: null
                });
            }
        });
    }

    // Specific Submission Date Toggle
    if (elements.useSpecificDateToggle) {
        elements.useSpecificDateToggle.addEventListener('click', async () => {
            const isOn = elements.useSpecificDateToggle.classList.contains('fa-toggle-on');
            const newState = !isOn;

            // Update toggle visual
            elements.useSpecificDateToggle.classList.toggle('fa-toggle-on', newState);
            elements.useSpecificDateToggle.classList.toggle('fa-toggle-off', !newState);
            elements.useSpecificDateToggle.style.color = newState ? '#22c55e' : 'gray';

            // Show/hide date picker
            if (elements.specificDatePicker) {
                elements.specificDatePicker.style.display = newState ? 'block' : 'none';
            }

            // If turning on, set to today's date if no date is set
            if (newState && elements.specificDateInput) {
                const currentValue = elements.specificDateInput.value;
                if (!currentValue) {
                    const today = new Date().toISOString().split('T')[0];
                    elements.specificDateInput.value = today;
                    await chrome.storage.local.set({
                        [STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE]: today
                    });
                }
            }

            // Save toggle state
            await chrome.storage.local.set({
                [STORAGE_KEYS.USE_SPECIFIC_DATE]: newState
            });
        });
    }

    // Specific Date Input Change
    if (elements.specificDateInput) {
        elements.specificDateInput.addEventListener('change', async (e) => {
            const selectedDate = e.target.value;
            await chrome.storage.local.set({
                [STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE]: selectedDate
            });
        });
    }

    // Clear Specific Date Button
    if (elements.clearSpecificDateBtn) {
        elements.clearSpecificDateBtn.addEventListener('click', async () => {
            // Clear the date input
            if (elements.specificDateInput) {
                elements.specificDateInput.value = '';
            }

            // Turn off the toggle
            if (elements.useSpecificDateToggle) {
                elements.useSpecificDateToggle.classList.remove('fa-toggle-on');
                elements.useSpecificDateToggle.classList.add('fa-toggle-off');
                elements.useSpecificDateToggle.style.color = 'gray';
            }

            // Hide the date picker
            if (elements.specificDatePicker) {
                elements.specificDatePicker.style.display = 'none';
            }

            // Clear from storage
            await chrome.storage.local.set({
                [STORAGE_KEYS.USE_SPECIFIC_DATE]: false,
                [STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE]: null
            });
        });
    }

    // Review Tutorial Button
    const reviewTutorialBtn = document.getElementById('reviewTutorialBtn');
    if (reviewTutorialBtn) {
        reviewTutorialBtn.addEventListener('click', () => {
            tutorialManager.restartTutorial();
        });
    }

    // Connections Modal
    if (elements.configureExcelBtn) {
        elements.configureExcelBtn.addEventListener('click', () => openConnectionsModal('excel'));
    }

    if (elements.configurePowerAutomateBtn) {
        elements.configurePowerAutomateBtn.addEventListener('click', () => openConnectionsModal('powerAutomate'));
    }

    if (elements.configureCanvasBtn) {
        elements.configureCanvasBtn.addEventListener('click', () => openConnectionsModal('canvas'));
    }

    if (elements.configureFive9Btn) {
        elements.configureFive9Btn.addEventListener('click', () => openConnectionsModal('five9'));
    }

    if (elements.closeConnectionsBtn) {
        elements.closeConnectionsBtn.addEventListener('click', closeConnectionsModal);
    }

    if (elements.saveConnectionsBtn) {
        elements.saveConnectionsBtn.addEventListener('click', async () => {
            await saveConnectionsSettings();
            // Update Five9 indicator immediately after settings change
            updateFive9ConnectionIndicator(queueManager.getQueue());
        });
    }

    // Canvas Modal Settings
    if (elements.embedHelperToggleModal) {
        elements.embedHelperToggleModal.addEventListener('click', toggleEmbedHelperModal);
    }

    if (elements.clearCacheBtnModal) {
        elements.clearCacheBtnModal.addEventListener('click', clearCacheFromModal);
    }

    // Five9 Modal Settings
    if (elements.debugModeToggleModal) {
        elements.debugModeToggleModal.addEventListener('click', toggleDebugModeModal);
    }

    // Excel Modal Settings
    if (elements.syncActiveStudentToggleModal) {
        elements.syncActiveStudentToggleModal.addEventListener('click', toggleSyncActiveStudentModal);
    }

    if (elements.sendMasterListToggleModal) {
        elements.sendMasterListToggleModal.addEventListener('click', toggleSendMasterListModal);
    }

    if (elements.reformatNameToggleModal) {
        elements.reformatNameToggleModal.addEventListener('click', toggleReformatNameModal);
    }

    if (elements.highlightStudentRowToggleModal) {
        elements.highlightStudentRowToggleModal.addEventListener('click', toggleHighlightStudentRowModal);
    }

    // Highlight Row Color Sync
    if (elements.highlightRowColorInput && elements.highlightRowColorTextInput) {
        elements.highlightRowColorInput.addEventListener('input', (e) => {
            elements.highlightRowColorTextInput.value = e.target.value;
        });
        elements.highlightRowColorTextInput.addEventListener('input', (e) => {
            const color = e.target.value;
            if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
                elements.highlightRowColorInput.value = color;
            }
        });
    }

    // Scan Filter Modal
    if (elements.scanFilterBtn) {
        elements.scanFilterBtn.addEventListener('click', openScanFilterModal);
    }

    if (elements.closeScanFilterBtn) {
        elements.closeScanFilterBtn.addEventListener('click', closeScanFilterModal);
    }

    if (elements.failingToggle) {
        elements.failingToggle.addEventListener('click', () => {
            toggleFailingFilter();
            updateScanFilterCount();
        });
    }

    if (elements.daysOutOperator) {
        elements.daysOutOperator.addEventListener('change', updateScanFilterCount);
    }

    if (elements.daysOutValue) {
        elements.daysOutValue.addEventListener('input', updateScanFilterCount);
    }

    if (elements.saveScanFilterBtn) {
        elements.saveScanFilterBtn.addEventListener('click', saveScanFilterSettings);
    }

    // Queue Modal
    if (elements.manageQueueBtn) {
        elements.manageQueueBtn.addEventListener('click', () => {
            openQueueModal(
                queueManager.getQueue(),
                (fromIndex, toIndex) => {
                    queueManager.reorderQueue(fromIndex, toIndex);
                    renderQueueModal(
                        queueManager.getQueue(),
                        (fromIdx, toIdx) => queueManager.reorderQueue(fromIdx, toIdx),
                        (index) => handleQueueRemoval(index)
                    );
                },
                (index) => handleQueueRemoval(index)
            );
        });
    }

    if (elements.closeQueueModalBtn) {
        elements.closeQueueModalBtn.addEventListener('click', closeQueueModal);
    }

    // Daily Update Modal
    if (elements.closeDailyUpdateBtn) {
        elements.closeDailyUpdateBtn.addEventListener('click', closeDailyUpdateModal);
    }

    if (elements.dailyUpdateLaterBtn) {
        elements.dailyUpdateLaterBtn.addEventListener('click', closeDailyUpdateModal);
    }

    if (elements.dailyUpdateBtn) {
        elements.dailyUpdateBtn.addEventListener('click', async () => {
            // Close the modal
            await closeDailyUpdateModal();

            // Switch to data tab
            switchTab('data');

            // Trigger the update master list process
            if (elements.updateMasterBtn) {
                elements.updateMasterBtn.click();
            }
        });
    }

    // Modal outside click handlers
    window.addEventListener('click', (e) => {
        if (elements.scanFilterModal && e.target === elements.scanFilterModal) {
            closeScanFilterModal();
        }
        if (elements.queueModal && e.target === elements.queueModal) {
            closeQueueModal();
        }
        if (elements.connectionsModal && e.target === elements.connectionsModal) {
            closeConnectionsModal();
        }
        if (elements.dailyUpdateModal && e.target === elements.dailyUpdateModal) {
            closeDailyUpdateModal();
        }
    });

    // Cache Management
    if (elements.clearCacheBtn) {
        elements.clearCacheBtn.addEventListener('click', async () => {
            if (confirm('Are you sure you want to clear all cached Canvas API data?')) {
                await clearAllCache();
                updateCacheStats();
            }
        });
    }

    // Debug Mode Toggle
    if (elements.debugModeToggle) {
        elements.debugModeToggle.addEventListener('click', toggleDebugMode);
    }

    // Embed Helper Toggle
    if (elements.embedHelperToggle) {
        elements.embedHelperToggle.addEventListener('click', toggleEmbedHelper);
    }

    // Highlight Color Picker
    if (elements.highlightColorPicker) {
        elements.highlightColorPicker.addEventListener('input', updateHighlightColor);
    }

    // Checker Tab
    if (elements.startBtn) {
        elements.startBtn.addEventListener('click', toggleScanState);
    }

    if (elements.clearListBtn) {
        elements.clearListBtn.addEventListener('click', () => {
            chrome.storage.local.set({ [STORAGE_KEYS.FOUND_ENTRIES]: [] });
        });
    }

    if (elements.foundSearch) {
        elements.foundSearch.addEventListener('input', filterFoundList);
    }

    // Call Tab
    if (elements.dialBtn) {
        elements.dialBtn.addEventListener('click', () => callManager.toggleCallState());
    }

    if (elements.skipStudentBtn) {
        elements.skipStudentBtn.addEventListener('click', () => {
            if (callManager) {
                callManager.skipToNext();
            }
        });
    }

    // Disposition buttons
    const dispositionContainer = document.querySelector('.disposition-grid');
    if (dispositionContainer) {
        dispositionContainer.addEventListener('click', (e) => {
            const btn = e.target.closest('.disposition-btn');
            if (!btn) return;

            if (btn.classList.contains('disabled')) {
                console.warn('This disposition does not have a code set yet.');
                return;
            }

            callManager.handleDisposition(btn.innerText.trim());
        });

        initializeDispositionButtons();
    }

    // Data Tab
    if (elements.updateMasterBtn) {
        elements.updateMasterBtn.addEventListener('click', handleUpdateMasterList);

        // Right-click context menu for Update Master List button
        elements.updateMasterBtn.addEventListener('contextmenu', (e) => {
            e.preventDefault();

            if (!elements.updateMasterContextMenu) return;

            // Position the context menu at the mouse position
            elements.updateMasterContextMenu.style.left = `${e.pageX}px`;
            elements.updateMasterContextMenu.style.top = `${e.pageY}px`;
            elements.updateMasterContextMenu.style.display = 'block';
        });
    }

    // Context menu item - Send List to Excel
    if (elements.sendListToExcelMenuItem) {
        elements.sendListToExcelMenuItem.addEventListener('click', async () => {
            // Hide context menu
            if (elements.updateMasterContextMenu) {
                elements.updateMasterContextMenu.style.display = 'none';
            }

            // Get current master list from storage
            const data = await chrome.storage.local.get([STORAGE_KEYS.MASTER_ENTRIES]);
            const students = data[STORAGE_KEYS.MASTER_ENTRIES] || [];

            if (students.length === 0) {
                alert('No master list data to send. Please update the master list first.');
                return;
            }

            // Check if any students have missing assignments data
            const hasMissingAssignments = students.some(s => s.missingAssignments && s.missingAssignments.length > 0);

            // Send to Excel - use the appropriate function based on whether we have missing assignments
            if (hasMissingAssignments) {
                await sendMasterListWithMissingAssignmentsToExcel(students);
                console.log('Manually sent master list with missing assignments to Excel');
            } else {
                await sendMasterListToExcel(students);
                console.log('Manually sent master list to Excel');
            }
        });
    }

    // Hide context menu when clicking elsewhere
    document.addEventListener('click', () => {
        if (elements.updateMasterContextMenu) {
            elements.updateMasterContextMenu.style.display = 'none';
        }
        if (elements.checkerContextMenu) {
            elements.checkerContextMenu.style.display = 'none';
        }
    });

    // Variable to track the selected student entry for context menu
    let selectedStudentEntry = null;

    // Right-click context menu for Checker Tab
    const checkerTab = document.getElementById('checker');
    if (checkerTab) {
        checkerTab.addEventListener('contextmenu', async (e) => {
            e.preventDefault();

            if (!elements.checkerContextMenu || !elements.checkerContextMenuText) return;

            // Check if right-clicked on a student list item
            const listItem = e.target.closest('.glass-list li');

            if (listItem && listItem.dataset.entryData) {
                // Right-clicked on a student - show "Resend Highlight Ping"
                selectedStudentEntry = JSON.parse(listItem.dataset.entryData);
                elements.checkerContextMenuText.textContent = 'Resend Highlight Ping';

                // Position the context menu at the mouse position
                elements.checkerContextMenu.style.left = `${e.pageX}px`;
                elements.checkerContextMenu.style.top = `${e.pageY}px`;
                elements.checkerContextMenu.style.display = 'block';
            } else {
                // Right-clicked elsewhere on checker tab - check if there are any students
                const data = await chrome.storage.local.get(STORAGE_KEYS.FOUND_ENTRIES);
                const foundEntries = data[STORAGE_KEYS.FOUND_ENTRIES] || [];

                if (foundEntries.length > 0) {
                    // Show "Resend All Highlight Pings" only if there are students
                    selectedStudentEntry = null;
                    elements.checkerContextMenuText.textContent = 'Resend All Highlight Pings';

                    // Position the context menu at the mouse position
                    elements.checkerContextMenu.style.left = `${e.pageX}px`;
                    elements.checkerContextMenu.style.top = `${e.pageY}px`;
                    elements.checkerContextMenu.style.display = 'block';
                }
                // If no students, don't show the context menu
            }
        });
    }

    // Context menu item - Resend Highlight Ping(s)
    if (elements.resendHighlightPingMenuItem) {
        elements.resendHighlightPingMenuItem.addEventListener('click', async () => {
            // Hide context menu
            if (elements.checkerContextMenu) {
                elements.checkerContextMenu.style.display = 'none';
            }

            if (selectedStudentEntry) {
                // Resend ping for single student
                await resendHighlightPing(selectedStudentEntry);
                console.log('Resent highlight ping for:', selectedStudentEntry.name);
            } else {
                // Resend pings for all students
                await resendAllHighlightPings();
                console.log('Resent all highlight pings');
            }
        });
    }

    // Mini Console toggle functionality is now handled in updateCanvasStatus
    // to allow status text to be a clickable link when Canvas is disconnected
    if (elements.statusText) {
        elements.statusText.style.cursor = 'pointer';
    }

    if (elements.clearConsoleBtn && elements.consoleContent) {
        elements.clearConsoleBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            elements.consoleContent.innerHTML = '';
        });
    }

    // Intercept console logs from sidepanel and display in mini console
    const originalConsole = {
        log: console.log,
        warn: console.warn,
        error: console.error,
        info: console.info
    };

    console.log = function(...args) {
        originalConsole.log.apply(console, args);
        addConsoleMessage('log', args);
    };

    console.warn = function(...args) {
        originalConsole.warn.apply(console, args);
        addConsoleMessage('warn', args);
    };

    console.error = function(...args) {
        originalConsole.error.apply(console, args);
        addConsoleMessage('error', args);
    };

    console.info = function(...args) {
        originalConsole.info.apply(console, args);
        addConsoleMessage('info', args);
    };

    if (elements.studentPopFile) {
        elements.studentPopFile.addEventListener('change', (e) => {
            handleFileImport(e.target.files[0], (students) => {
                renderMasterList(students, (entry, li, evt) => {
                    queueManager.handleStudentClick(entry, li, evt);
                });
                processStep2(students, (updatedStudents) => {
                    renderMasterList(updatedStudents, (entry, li, evt) => {
                        queueManager.handleStudentClick(entry, li, evt);
                    });
                    processStep3(updatedStudents, async (finalStudents) => {
                        renderMasterList(finalStudents, (entry, li, evt) => {
                            queueManager.handleStudentClick(entry, li, evt);
                        });
                        // Send master list with missing assignments to Excel
                        await processStep4(finalStudents);
                    });
                });
            });
        });
    }

    if (elements.queueCloseBtn) {
        elements.queueCloseBtn.addEventListener('click', () => {
            elements.updateQueueSection.style.display = 'none';
        });
    }

    if (elements.masterSearch) {
        elements.masterSearch.addEventListener('input', filterMasterList);
    }

    if (elements.sortSelect) {
        elements.sortSelect.addEventListener('change', sortMasterList);
    }

    if (elements.downloadMasterBtn) {
        elements.downloadMasterBtn.addEventListener('click', exportMasterListCSV);
    }
}

// --- HELPER FUNCTIONS ---

/**
 * Initializes disposition button states
 * Only shows dispositions that have a valid code defined
 */
function initializeDispositionButtons() {
    const dispositionButtons = document.querySelectorAll('.disposition-btn');

    dispositionButtons.forEach(btn => {
        const buttonText = btn.innerText.trim();

        // Hide dispositions that don't have a code (including "Other")
        if (!hasDispositionCode(buttonText)) {
            btn.style.display = 'none';
        } else {
            btn.style.display = 'flex';
            btn.classList.remove('disabled');
            btn.style.opacity = '1';
            btn.style.cursor = 'pointer';
            btn.title = '';
        }
    });
}

/**
 * Initializes call control button visibility based on UI_FEATURES flags
 */
function initializeCallControlButtons() {
    // Mute button visibility
    const muteButton = document.querySelector('.control-btn.mute');
    if (muteButton) {
        muteButton.style.display = UI_FEATURES.SHOW_MUTE_BUTTON ? 'flex' : 'none';
    }

    // Speaker button visibility
    const speakerButton = document.querySelector('.control-btn.speaker');
    if (speakerButton) {
        speakerButton.style.display = UI_FEATURES.SHOW_SPEAKER_BUTTON ? 'flex' : 'none';
    }
}

/**
 * Handles queue removal operations
 */
function handleQueueRemoval(index) {
    const result = queueManager.removeFromQueue(index);
    if (result === 'close') {
        closeQueueModal();
    } else if (result === 'refresh') {
        renderQueueModal(
            queueManager.getQueue(),
            (fromIdx, toIdx) => queueManager.reorderQueue(fromIdx, toIdx),
            (idx) => handleQueueRemoval(idx)
        );
    }
}

/**
 * Handles Update Master List button click
 */
async function handleUpdateMasterList() {
    if (elements.updateQueueSection) {
        elements.updateQueueSection.style.display = 'block';
        elements.updateQueueSection.scrollIntoView({ behavior: 'smooth' });

        resetQueueUI();

        // Direct CSV upload
        restoreDefaultQueueUI();

        const step1 = document.getElementById('step1');
        if (step1) {
            step1.className = 'queue-item active';
            step1.querySelector('i').className = 'fas fa-spinner';
        }

        if (elements.studentPopFile) {
            elements.studentPopFile.click();
        }
    }
}

/**
 * Loads data from storage and updates UI
 */
async function loadStorageData() {
    const data = await chrome.storage.local.get([
        STORAGE_KEYS.FOUND_ENTRIES,
        STORAGE_KEYS.MASTER_ENTRIES,
        STORAGE_KEYS.LAST_UPDATED,
        STORAGE_KEYS.EXTENSION_STATE,
        STORAGE_KEYS.DEBUG_MODE,
        STORAGE_KEYS.EMBED_IN_CANVAS,
        STORAGE_KEYS.HIGHLIGHT_COLOR,
        STORAGE_KEYS.AUTO_UPDATE_MASTER_LIST,
        STORAGE_KEYS.POWER_AUTOMATE_URL,
        STORAGE_KEYS.USE_SPECIFIC_DATE,
        STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE
    ]);

    const foundEntries = data[STORAGE_KEYS.FOUND_ENTRIES] || [];
    renderFoundList(foundEntries);
    updateTabBadge('checker', foundEntries.length);

    renderMasterList(data[STORAGE_KEYS.MASTER_ENTRIES] || [], (entry, li, evt) => {
        queueManager.handleStudentClick(entry, li, evt);
    });

    if (elements.lastUpdatedText && data[STORAGE_KEYS.LAST_UPDATED]) {
        elements.lastUpdatedText.textContent = data[STORAGE_KEYS.LAST_UPDATED];
    }

    updateButtonVisuals(data[STORAGE_KEYS.EXTENSION_STATE] || EXTENSION_STATES.OFF);

    isDebugMode = data[STORAGE_KEYS.DEBUG_MODE] || false;
    updateDebugModeUI(isDebugMode);
    if (callManager) {
        callManager.setDebugMode(isDebugMode);
    }

    // Load Embed Helper setting (default: true)
    embedHelperEnabled = data[STORAGE_KEYS.EMBED_IN_CANVAS] !== undefined
        ? data[STORAGE_KEYS.EMBED_IN_CANVAS]
        : true;
    updateEmbedHelperUI(embedHelperEnabled);

    // Load Highlight Color setting (default: #ffff00)
    highlightColor = data[STORAGE_KEYS.HIGHLIGHT_COLOR] || '#ffff00';
    updateHighlightColorUI(highlightColor);

    // Load Power Automate URL and update status
    const powerAutomateUrl = data[STORAGE_KEYS.POWER_AUTOMATE_URL] || '';
    updatePowerAutomateStatus(powerAutomateUrl);

    // Update Canvas connection status
    updateCanvasStatus();

    // Update Five9 connection status
    updateFive9Status();

    // Load Specific Submission Date settings
    const useSpecificDate = data[STORAGE_KEYS.USE_SPECIFIC_DATE] || false;
    const specificDate = data[STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE];

    if (elements.useSpecificDateToggle) {
        elements.useSpecificDateToggle.classList.toggle('fa-toggle-on', useSpecificDate);
        elements.useSpecificDateToggle.classList.toggle('fa-toggle-off', !useSpecificDate);
        elements.useSpecificDateToggle.style.color = useSpecificDate ? '#22c55e' : 'gray';
    }

    if (elements.specificDatePicker) {
        elements.specificDatePicker.style.display = useSpecificDate ? 'block' : 'none';
    }

    if (elements.specificDateInput && specificDate) {
        elements.specificDateInput.value = specificDate;
    }
}

// Storage change listener
chrome.storage.onChanged.addListener((changes) => {
    if (changes[STORAGE_KEYS.FOUND_ENTRIES]) {
        renderFoundList(changes[STORAGE_KEYS.FOUND_ENTRIES].newValue);
        updateTabBadge('checker', (changes[STORAGE_KEYS.FOUND_ENTRIES].newValue || []).length);
    }
    if (changes[STORAGE_KEYS.MASTER_ENTRIES]) {
        renderMasterList(changes[STORAGE_KEYS.MASTER_ENTRIES].newValue, (entry, li, evt) => {
            queueManager.handleStudentClick(entry, li, evt);
        });
    }
    if (changes[STORAGE_KEYS.EXTENSION_STATE]) {
        updateButtonVisuals(changes[STORAGE_KEYS.EXTENSION_STATE].newValue);
    }

    // Handle name format toggle changes - re-render all displays
    if (changes.reformatNameEnabled) {
        console.log(`Name format changed to: ${changes.reformatNameEnabled.newValue ? 'First Last' : 'Original'}`);

        // Re-render found list
        chrome.storage.local.get([STORAGE_KEYS.FOUND_ENTRIES], (data) => {
            const foundEntries = data[STORAGE_KEYS.FOUND_ENTRIES] || [];
            renderFoundList(foundEntries);
        });

        // Re-render master list
        chrome.storage.local.get([STORAGE_KEYS.MASTER_ENTRIES], (data) => {
            const masterEntries = data[STORAGE_KEYS.MASTER_ENTRIES] || [];
            renderMasterList(masterEntries, (entry, li, evt) => {
                queueManager.handleStudentClick(entry, li, evt);
            });
        });

        // Re-render active student if one is selected
        if (callManager && callManager.activeStudent) {
            setActiveStudent(callManager.activeStudent, callManager);
        }

        // Re-render queue modal if it's open
        if (elements.queueModal && elements.queueModal.style.display !== 'none') {
            queueManager.renderQueue();
        }
    }
});

// Runtime message listener for Office Add-in student selection sync
chrome.runtime.onMessage.addListener(async (msg, sender, sendResponse) => {
    if (msg.type === MESSAGE_TYPES.SRK_SELECTED_STUDENTS) {
        // IGNORE PINGS IF CALL IS ACTIVE - Don't interrupt current call session
        if (callManager && callManager.getCallActiveState()) {
            console.log('%c [Sidepanel] Ignoring ping - call already in session', 'color: orange; font-weight: bold');
            return; // Exit early, don't process this ping
        }

        const modeText = msg.count === 1 ? 'active student' : 'automation mode';
        console.log(`%c [Sidepanel] Setting ${modeText} from Office Add-in:`, 'color: purple; font-weight: bold', msg.count, 'student(s)');

        if (msg.students && msg.students.length > 0 && callManager && queueManager) {
            // Try to find matching students in master list for complete data
            const data = await chrome.storage.local.get([STORAGE_KEYS.MASTER_ENTRIES]);
            const masterEntries = data[STORAGE_KEYS.MASTER_ENTRIES] || [];

            // Match all students with master list
            const studentsToSet = msg.students.map(student => {
                // Try to match by SyStudentId or name
                if (masterEntries.length > 0) {
                    const matchedStudent = masterEntries.find(entry => {
                        // Match by SyStudentId if available
                        if (student.SyStudentId && entry.SyStudentId) {
                            return entry.SyStudentId === student.SyStudentId;
                        }
                        // Otherwise match by name
                        return entry.name === student.name;
                    });

                    if (matchedStudent) {
                        console.log(`Matched with master list: ${matchedStudent.name}`);
                        return matchedStudent;
                    } else {
                        console.log(`No match for ${student.name}, using Office add-in data`);
                    }
                }
                return student;
            });

            // Set queue using queue manager (handles both single and multiple)
            queueManager.setQueue(studentsToSet);

            // Switch to contact tab to show the selected student(s)
            switchTab('contact');
            // Check Five9 connection when switching to contact tab
            updateFive9ConnectionIndicator(queueManager.getQueue());

            if (msg.count === 1) {
                console.log(`Active student set to: ${studentsToSet[0].name}`);
            } else {
                console.log(`Automation mode enabled with ${msg.count} students`);
            }
        }
    }

    // Handle logs from background script
    if (msg.type === MESSAGE_TYPES.LOG_TO_PANEL) {
        addConsoleMessage(msg.level, msg.args);
    }
});

/**
 * Toggles scanning state
 */
function toggleScanState() {
    // Don't toggle if button is disabled (no Canvas connection)
    if (elements.startBtn && elements.startBtn.disabled) {
        return;
    }

    isScanning = !isScanning;
    const newState = isScanning ? EXTENSION_STATES.ON : EXTENSION_STATES.OFF;
    chrome.storage.local.set({ [STORAGE_KEYS.EXTENSION_STATE]: newState });
}

/**
 * Toggles debug mode
 */
async function toggleDebugMode() {
    isDebugMode = !isDebugMode;
    await chrome.storage.local.set({ [STORAGE_KEYS.DEBUG_MODE]: isDebugMode });
    updateDebugModeUI(isDebugMode);
    if (callManager) {
        callManager.setDebugMode(isDebugMode);
    }
    updateFive9ConnectionIndicator(queueManager.getQueue());
}

/**
 * Updates cache statistics display
 */
async function updateCacheStats() {
    if (!elements.cacheStatsText) return;

    try {
        const stats = await getCacheStats();

        if (stats.totalEntries === 0) {
            elements.cacheStatsText.textContent = 'No cached data';
        } else {
            const validText = stats.validEntries === 1 ? 'entry' : 'entries';
            const expiredText = stats.expiredEntries > 0
                ? ` (${stats.expiredEntries} expired)`
                : '';
            elements.cacheStatsText.textContent = `${stats.validEntries} valid ${validText}${expiredText}`;
        }
    } catch (error) {
        console.error('Error updating cache stats:', error);
        elements.cacheStatsText.textContent = 'Error loading stats';
    }
}

/**
 * Toggles embed helper in Canvas
 */
async function toggleEmbedHelper() {
    embedHelperEnabled = !embedHelperEnabled;
    await chrome.storage.local.set({ [STORAGE_KEYS.EMBED_IN_CANVAS]: embedHelperEnabled });
    updateEmbedHelperUI(embedHelperEnabled);
}

/**
 * Updates highlight color setting
 */
async function updateHighlightColor(event) {
    highlightColor = event.target.value;
    await chrome.storage.local.set({ [STORAGE_KEYS.HIGHLIGHT_COLOR]: highlightColor });
}
