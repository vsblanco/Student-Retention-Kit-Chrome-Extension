// Connections Modal - Handles settings for Excel, Power Automate, Canvas, and Five9
import { STORAGE_KEYS, CANVAS_DOMAIN, FIVE9_CONNECTION_STATES, EXTENSION_STATES } from '../../constants/index.js';
import { storageGet, storageSet, storageGetValue, sessionGetValue } from '../../utils/storage.js';
import { encrypt, decrypt } from '../../utils/encryption.js';
import { elements } from '../ui-manager.js';

/**
 * Opens the connections modal for a specific connection type
 * @param {string} connectionType - 'excel', 'powerAutomate', 'canvas', or 'five9'
 */
export async function openConnectionsModal(connectionType) {
    if (!elements.connectionsModal) return;

    // Show the modal
    elements.connectionsModal.style.display = 'flex';

    // Hide all configuration content
    if (elements.excelConfigContent) {
        elements.excelConfigContent.style.display = 'none';
    }
    if (elements.powerAutomateConfigContent) {
        elements.powerAutomateConfigContent.style.display = 'none';
    }
    if (elements.canvasConfigContent) {
        elements.canvasConfigContent.style.display = 'none';
    }
    if (elements.five9ConfigContent) {
        elements.five9ConfigContent.style.display = 'none';
    }

    // Show the appropriate configuration content
    if (connectionType === 'excel') {
        if (elements.connectionModalTitle) {
            elements.connectionModalTitle.textContent = 'Excel Add-in Settings';
        }
        if (elements.excelConfigContent) {
            elements.excelConfigContent.style.display = 'block';
        }
    } else if (connectionType === 'powerAutomate') {
        if (elements.connectionModalTitle) {
            elements.connectionModalTitle.textContent = 'Power Automate Settings';
        }
        if (elements.powerAutomateConfigContent) {
            elements.powerAutomateConfigContent.style.display = 'block';
        }
    } else if (connectionType === 'canvas') {
        if (elements.connectionModalTitle) {
            elements.connectionModalTitle.textContent = 'Canvas Settings';
        }
        if (elements.canvasConfigContent) {
            elements.canvasConfigContent.style.display = 'block';
        }
        // Load canvas cache stats when Canvas settings are opened
        loadCacheStatsForModal();
    } else if (connectionType === 'five9') {
        if (elements.connectionModalTitle) {
            elements.connectionModalTitle.textContent = 'Five9 Settings';
        }
        if (elements.five9ConfigContent) {
            elements.five9ConfigContent.style.display = 'block';
        }
    }

    // Load current settings into modal
    const result = await storageGet([
        STORAGE_KEYS.AUTO_UPDATE_MASTER_LIST,
        STORAGE_KEYS.SYNC_ACTIVE_STUDENT,
        STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL,
        STORAGE_KEYS.REFORMAT_NAME_ENABLED,
        STORAGE_KEYS.HIGHLIGHT_STUDENT_ROW_ENABLED,
        STORAGE_KEYS.POWER_AUTOMATE_URL,
        STORAGE_KEYS.POWER_AUTOMATE_ENABLED,
        STORAGE_KEYS.POWER_AUTOMATE_DEBUG,
        STORAGE_KEYS.EMBED_IN_CANVAS,
        STORAGE_KEYS.HIGHLIGHT_COLOR,
        STORAGE_KEYS.CANVAS_CACHE_ENABLED,
        STORAGE_KEYS.NON_API_COURSE_FETCH,
        STORAGE_KEYS.NEXT_ASSIGNMENT_ENABLED,
        STORAGE_KEYS.CALL_DEMO,
        STORAGE_KEYS.AUTO_SWITCH_TO_CALL_TAB,
        STORAGE_KEYS.HIGHLIGHT_START_COL,
        STORAGE_KEYS.HIGHLIGHT_END_COL,
        STORAGE_KEYS.HIGHLIGHT_EDIT_COLUMN,
        STORAGE_KEYS.HIGHLIGHT_EDIT_TEXT,
        STORAGE_KEYS.HIGHLIGHT_TARGET_SHEET,
        STORAGE_KEYS.HIGHLIGHT_ROW_COLOR
    ]);

    // Load auto-update setting
    const setting = result[STORAGE_KEYS.AUTO_UPDATE_MASTER_LIST] || 'always';
    if (elements.autoUpdateSelectModal) {
        elements.autoUpdateSelectModal.value = setting;
    }

    // Load sync active student setting
    const syncActiveStudent = result[STORAGE_KEYS.SYNC_ACTIVE_STUDENT] !== undefined ? result[STORAGE_KEYS.SYNC_ACTIVE_STUDENT] : true;
    updateSyncActiveStudentModalUI(syncActiveStudent);

    // Load send master list to Excel setting
    const sendMasterListToExcel = result[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL] !== undefined ? result[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL] : true;
    updateSendMasterListModalUI(sendMasterListToExcel);

    // Load reformat name setting
    const reformatNameEnabled = result[STORAGE_KEYS.REFORMAT_NAME_ENABLED] !== undefined ? result[STORAGE_KEYS.REFORMAT_NAME_ENABLED] : true;
    updateReformatNameModalUI(reformatNameEnabled);

    // Load highlight student row enabled setting
    const highlightStudentRowEnabled = result[STORAGE_KEYS.HIGHLIGHT_STUDENT_ROW_ENABLED] !== undefined ? result[STORAGE_KEYS.HIGHLIGHT_STUDENT_ROW_ENABLED] : true;
    updateHighlightStudentRowModalUI(highlightStudentRowEnabled);

    // Load Power Automate settings (decrypt the URL)
    const encryptedPaUrl = result[STORAGE_KEYS.POWER_AUTOMATE_URL] || '';
    const paUrl = await decrypt(encryptedPaUrl);
    if (elements.powerAutomateUrlInput) {
        elements.powerAutomateUrlInput.value = paUrl;
    }

    const paEnabled = result[STORAGE_KEYS.POWER_AUTOMATE_ENABLED] || false;
    updatePowerAutomateEnabledUI(paEnabled);

    const paDebug = result[STORAGE_KEYS.POWER_AUTOMATE_DEBUG] || false;
    updatePowerAutomateDebugUI(paDebug);

    // Load Canvas settings
    const embedHelper = result[STORAGE_KEYS.EMBED_IN_CANVAS] !== undefined ? result[STORAGE_KEYS.EMBED_IN_CANVAS] : true;
    updateEmbedHelperModalUI(embedHelper);

    const highlightColor = result[STORAGE_KEYS.HIGHLIGHT_COLOR] || '#ffff00';
    if (elements.highlightColorPickerModal) {
        elements.highlightColorPickerModal.value = highlightColor;
    }

    const canvasCacheEnabled = result[STORAGE_KEYS.CANVAS_CACHE_ENABLED] !== undefined ? result[STORAGE_KEYS.CANVAS_CACHE_ENABLED] : true;
    updateCanvasCacheModalUI(canvasCacheEnabled);

    const nonApiCourseFetch = result[STORAGE_KEYS.NON_API_COURSE_FETCH] !== undefined ? result[STORAGE_KEYS.NON_API_COURSE_FETCH] : false;
    updateNonApiCourseFetchUI(nonApiCourseFetch);

    const nextAssignmentEnabled = result[STORAGE_KEYS.NEXT_ASSIGNMENT_ENABLED] !== undefined ? result[STORAGE_KEYS.NEXT_ASSIGNMENT_ENABLED] : false;
    updateNextAssignmentUI(nextAssignmentEnabled);

    // Load Five9 settings (Call Demo mode, formerly debugMode)
    const callDemo = result[STORAGE_KEYS.CALL_DEMO] || false;
    updateDebugModeModalUI(callDemo);

    // Load Auto Switch to Call Tab setting
    const autoSwitchCallTab = result[STORAGE_KEYS.AUTO_SWITCH_TO_CALL_TAB] !== undefined ? result[STORAGE_KEYS.AUTO_SWITCH_TO_CALL_TAB] : true;
    updateAutoSwitchCallTabUI(autoSwitchCallTab);

    // Load Highlight Student Row settings
    if (elements.highlightStartColInput) {
        elements.highlightStartColInput.value = result[STORAGE_KEYS.HIGHLIGHT_START_COL] || 'Student Name';
    }
    if (elements.highlightEndColInput) {
        elements.highlightEndColInput.value = result[STORAGE_KEYS.HIGHLIGHT_END_COL] || 'Outreach';
    }
    if (elements.highlightEditColumnInput) {
        elements.highlightEditColumnInput.value = result[STORAGE_KEYS.HIGHLIGHT_EDIT_COLUMN] || 'Outreach';
    }
    if (elements.highlightEditTextInput) {
        elements.highlightEditTextInput.value = result[STORAGE_KEYS.HIGHLIGHT_EDIT_TEXT] || 'Submitted {assignment}';
    }
    if (elements.highlightTargetSheetInput) {
        elements.highlightTargetSheetInput.value = result[STORAGE_KEYS.HIGHLIGHT_TARGET_SHEET] || 'LDA MM-DD-YYYY';
    }
    if (elements.highlightRowColorInput && elements.highlightRowColorTextInput) {
        const color = result[STORAGE_KEYS.HIGHLIGHT_ROW_COLOR] || '#92d050';
        elements.highlightRowColorInput.value = color;
        elements.highlightRowColorTextInput.value = color;
    }

    // Load cache stats
    loadCacheStatsForModal();
}

/**
 * Closes the connections modal
 */
export function closeConnectionsModal() {
    if (elements.connectionsModal) {
        elements.connectionsModal.style.display = 'none';
    }
}

/**
 * Saves connections settings from the modal
 */
export async function saveConnectionsSettings() {
    const settingsToSave = {};

    // Save auto-update setting
    if (elements.autoUpdateSelectModal) {
        const newSetting = elements.autoUpdateSelectModal.value;
        settingsToSave[STORAGE_KEYS.AUTO_UPDATE_MASTER_LIST] = newSetting;
        console.log(`Auto-update master list setting saved: ${newSetting}`);
    }

    // Save sync active student setting
    if (elements.syncActiveStudentToggleModal) {
        const syncEnabled = elements.syncActiveStudentToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.SYNC_ACTIVE_STUDENT] = syncEnabled;
        console.log(`Sync Active Student setting saved: ${syncEnabled}`);
    }

    // Save send master list to Excel setting
    if (elements.sendMasterListToggleModal) {
        const sendEnabled = elements.sendMasterListToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL] = sendEnabled;
        console.log(`Send Master List to Excel setting saved: ${sendEnabled}`);
    }

    // Save reformat name setting
    if (elements.reformatNameToggleModal) {
        const reformatEnabled = elements.reformatNameToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.REFORMAT_NAME_ENABLED] = reformatEnabled;
        console.log(`Reformat Name setting saved: ${reformatEnabled}`);
    }

    // Save highlight student row enabled setting
    if (elements.highlightStudentRowToggleModal) {
        const highlightEnabled = elements.highlightStudentRowToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_STUDENT_ROW_ENABLED] = highlightEnabled;
        console.log(`Highlight Student Row enabled setting saved: ${highlightEnabled}`);
    }

    // Save Power Automate settings (encrypt the URL)
    if (elements.powerAutomateUrlInput) {
        const paUrl = elements.powerAutomateUrlInput.value.trim();
        const encryptedPaUrl = paUrl ? await encrypt(paUrl) : '';
        settingsToSave[STORAGE_KEYS.POWER_AUTOMATE_URL] = encryptedPaUrl;
        console.log(`Power Automate URL saved: ${paUrl ? 'URL configured (encrypted)' : 'URL cleared'}`);
    }

    let paEnabled = false;
    if (elements.powerAutomateEnabledToggle) {
        paEnabled = elements.powerAutomateEnabledToggle.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.POWER_AUTOMATE_ENABLED] = paEnabled;
        console.log(`Power Automate Enabled saved: ${paEnabled}`);
    }

    if (elements.powerAutomateDebugToggle) {
        const paDebug = elements.powerAutomateDebugToggle.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.POWER_AUTOMATE_DEBUG] = paDebug;
        console.log(`Power Automate Debug saved: ${paDebug}`);
    }

    // Update status text immediately with enabled state
    if (elements.powerAutomateUrlInput) {
        updatePowerAutomateStatus(elements.powerAutomateUrlInput.value.trim(), paEnabled);
    }

    // Save Canvas settings
    if (elements.embedHelperToggleModal) {
        const embedEnabled = elements.embedHelperToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.EMBED_IN_CANVAS] = embedEnabled;
        console.log(`Embed Helper setting saved: ${embedEnabled}`);
    }

    if (elements.highlightColorPickerModal) {
        const highlightColor = elements.highlightColorPickerModal.value;
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_COLOR] = highlightColor;
        console.log(`Highlight Color saved: ${highlightColor}`);
    }

    if (elements.canvasCacheToggleModal) {
        const cacheEnabled = elements.canvasCacheToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.CANVAS_CACHE_ENABLED] = cacheEnabled;
        console.log(`Canvas Cache setting saved: ${cacheEnabled}`);
    }

    if (elements.nonApiCourseFetchToggle) {
        const nonApiCourseFetch = elements.nonApiCourseFetchToggle.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.NON_API_COURSE_FETCH] = nonApiCourseFetch;
        console.log(`Non API Course Fetch setting saved: ${nonApiCourseFetch}`);
    }

    if (elements.nextAssignmentToggle) {
        const nextAssignmentEnabled = elements.nextAssignmentToggle.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.NEXT_ASSIGNMENT_ENABLED] = nextAssignmentEnabled;
        console.log(`Next Assignment setting saved: ${nextAssignmentEnabled}`);
    }

    // Save Five9 settings (Call Demo mode)
    if (elements.debugModeToggleModal) {
        const callDemoEnabled = elements.debugModeToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.CALL_DEMO] = callDemoEnabled;
        console.log(`Call Demo mode setting saved: ${callDemoEnabled}`);
    }

    // Save Auto Switch to Call Tab setting
    if (elements.autoSwitchCallTabToggle) {
        const autoSwitchEnabled = elements.autoSwitchCallTabToggle.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.AUTO_SWITCH_TO_CALL_TAB] = autoSwitchEnabled;
        console.log(`Auto Switch to Call Tab setting saved: ${autoSwitchEnabled}`);
    }

    // Save Highlight Student Row settings
    if (elements.highlightStartColInput) {
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_START_COL] = elements.highlightStartColInput.value || 'Student Name';
    }
    if (elements.highlightEndColInput) {
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_END_COL] = elements.highlightEndColInput.value || 'Outreach';
    }
    if (elements.highlightEditColumnInput) {
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_EDIT_COLUMN] = elements.highlightEditColumnInput.value || 'Outreach';
    }
    if (elements.highlightEditTextInput) {
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_EDIT_TEXT] = elements.highlightEditTextInput.value || 'Submitted {assignment}';
    }
    if (elements.highlightTargetSheetInput) {
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_TARGET_SHEET] = elements.highlightTargetSheetInput.value || 'LDA MM-DD-YYYY';
    }
    if (elements.highlightRowColorTextInput) {
        settingsToSave[STORAGE_KEYS.HIGHLIGHT_ROW_COLOR] = elements.highlightRowColorTextInput.value || '#92d050';
    }

    // Save all settings
    await storageSet(settingsToSave);

    // Update Five9 status to reflect Call Demo mode change
    updateFive9Status();

    // Close modal after saving
    closeConnectionsModal();
}

// ============================================
// Toggle UI Update Functions
// ============================================

function updateEmbedHelperModalUI(isEnabled) {
    if (!elements.embedHelperToggleModal) return;
    if (isEnabled) {
        elements.embedHelperToggleModal.className = 'fas fa-toggle-on';
        elements.embedHelperToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.embedHelperToggleModal.className = 'fas fa-toggle-off';
        elements.embedHelperToggleModal.style.color = 'gray';
    }
}

function updateCanvasCacheModalUI(isEnabled) {
    if (!elements.canvasCacheToggleModal) return;
    if (isEnabled) {
        elements.canvasCacheToggleModal.className = 'fas fa-toggle-on';
        elements.canvasCacheToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.canvasCacheToggleModal.className = 'fas fa-toggle-off';
        elements.canvasCacheToggleModal.style.color = 'gray';
    }
    if (elements.cacheSettingsContainer) {
        if (isEnabled) {
            elements.cacheSettingsContainer.style.opacity = '1';
            elements.cacheSettingsContainer.style.pointerEvents = 'auto';
        } else {
            elements.cacheSettingsContainer.style.opacity = '0.4';
            elements.cacheSettingsContainer.style.pointerEvents = 'none';
        }
    }
}

function updateNonApiCourseFetchUI(isEnabled) {
    if (!elements.nonApiCourseFetchToggle) return;
    if (isEnabled) {
        elements.nonApiCourseFetchToggle.className = 'fas fa-toggle-on';
        elements.nonApiCourseFetchToggle.style.color = 'var(--primary-color)';
    } else {
        elements.nonApiCourseFetchToggle.className = 'fas fa-toggle-off';
        elements.nonApiCourseFetchToggle.style.color = 'gray';
    }
}

function updateNextAssignmentUI(isEnabled) {
    if (!elements.nextAssignmentToggle) return;
    if (isEnabled) {
        elements.nextAssignmentToggle.className = 'fas fa-toggle-on';
        elements.nextAssignmentToggle.style.color = 'var(--primary-color)';
    } else {
        elements.nextAssignmentToggle.className = 'fas fa-toggle-off';
        elements.nextAssignmentToggle.style.color = 'gray';
    }
}

function updateDebugModeModalUI(isEnabled) {
    if (!elements.debugModeToggleModal) return;
    if (isEnabled) {
        elements.debugModeToggleModal.className = 'fas fa-toggle-on';
        elements.debugModeToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.debugModeToggleModal.className = 'fas fa-toggle-off';
        elements.debugModeToggleModal.style.color = 'gray';
    }
}

function updateAutoSwitchCallTabUI(isEnabled) {
    if (!elements.autoSwitchCallTabToggle) return;
    if (isEnabled) {
        elements.autoSwitchCallTabToggle.className = 'fas fa-toggle-on';
        elements.autoSwitchCallTabToggle.style.color = 'var(--primary-color)';
    } else {
        elements.autoSwitchCallTabToggle.className = 'fas fa-toggle-off';
        elements.autoSwitchCallTabToggle.style.color = 'gray';
    }
}

function updateSyncActiveStudentModalUI(isEnabled) {
    if (!elements.syncActiveStudentToggleModal) return;
    if (isEnabled) {
        elements.syncActiveStudentToggleModal.className = 'fas fa-toggle-on';
        elements.syncActiveStudentToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.syncActiveStudentToggleModal.className = 'fas fa-toggle-off';
        elements.syncActiveStudentToggleModal.style.color = 'gray';
    }
}

function updateSendMasterListModalUI(isEnabled) {
    if (!elements.sendMasterListToggleModal) return;
    if (isEnabled) {
        elements.sendMasterListToggleModal.className = 'fas fa-toggle-on';
        elements.sendMasterListToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.sendMasterListToggleModal.className = 'fas fa-toggle-off';
        elements.sendMasterListToggleModal.style.color = 'gray';
    }
}

function updateReformatNameModalUI(isEnabled) {
    if (!elements.reformatNameToggleModal) return;
    if (isEnabled) {
        elements.reformatNameToggleModal.className = 'fas fa-toggle-on';
        elements.reformatNameToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.reformatNameToggleModal.className = 'fas fa-toggle-off';
        elements.reformatNameToggleModal.style.color = 'gray';
    }
}

function updateHighlightStudentRowModalUI(isEnabled) {
    if (!elements.highlightStudentRowToggleModal) return;
    if (isEnabled) {
        elements.highlightStudentRowToggleModal.className = 'fas fa-toggle-on';
        elements.highlightStudentRowToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.highlightStudentRowToggleModal.className = 'fas fa-toggle-off';
        elements.highlightStudentRowToggleModal.style.color = 'gray';
    }
    if (elements.highlightSettingsContainer) {
        if (isEnabled) {
            elements.highlightSettingsContainer.style.opacity = '1';
            elements.highlightSettingsContainer.style.pointerEvents = 'auto';
        } else {
            elements.highlightSettingsContainer.style.opacity = '0.4';
            elements.highlightSettingsContainer.style.pointerEvents = 'none';
        }
    }
}

function updatePowerAutomateEnabledUI(isEnabled) {
    if (!elements.powerAutomateEnabledToggle) return;
    if (isEnabled) {
        elements.powerAutomateEnabledToggle.className = 'fas fa-toggle-on';
        elements.powerAutomateEnabledToggle.style.color = 'var(--primary-color)';
    } else {
        elements.powerAutomateEnabledToggle.className = 'fas fa-toggle-off';
        elements.powerAutomateEnabledToggle.style.color = 'gray';
    }
    if (elements.powerAutomateSettingsContainer) {
        elements.powerAutomateSettingsContainer.style.opacity = isEnabled ? '1' : '0.5';
        elements.powerAutomateSettingsContainer.style.pointerEvents = isEnabled ? 'auto' : 'none';
    }
    if (elements.powerAutomateDebugContainer) {
        elements.powerAutomateDebugContainer.style.opacity = isEnabled ? '1' : '0.5';
        elements.powerAutomateDebugContainer.style.pointerEvents = isEnabled ? 'auto' : 'none';
    }
}

function updatePowerAutomateDebugUI(isEnabled) {
    if (!elements.powerAutomateDebugToggle) return;
    if (isEnabled) {
        elements.powerAutomateDebugToggle.className = 'fas fa-toggle-on';
        elements.powerAutomateDebugToggle.style.color = 'var(--primary-color)';
    } else {
        elements.powerAutomateDebugToggle.className = 'fas fa-toggle-off';
        elements.powerAutomateDebugToggle.style.color = 'gray';
    }
}

// ============================================
// Toggle Export Functions
// ============================================

export function toggleEmbedHelperModal() {
    if (!elements.embedHelperToggleModal) return;
    const isCurrentlyOn = elements.embedHelperToggleModal.classList.contains('fa-toggle-on');
    updateEmbedHelperModalUI(!isCurrentlyOn);
}

export function toggleCanvasCacheModal() {
    if (!elements.canvasCacheToggleModal) return;
    const isCurrentlyOn = elements.canvasCacheToggleModal.classList.contains('fa-toggle-on');
    updateCanvasCacheModalUI(!isCurrentlyOn);
}

export function toggleNonApiCourseFetch() {
    if (!elements.nonApiCourseFetchToggle) return;
    const isCurrentlyOn = elements.nonApiCourseFetchToggle.classList.contains('fa-toggle-on');
    updateNonApiCourseFetchUI(!isCurrentlyOn);
}

export function toggleNextAssignment() {
    if (!elements.nextAssignmentToggle) return;
    const isCurrentlyOn = elements.nextAssignmentToggle.classList.contains('fa-toggle-on');
    updateNextAssignmentUI(!isCurrentlyOn);
}

export function togglePowerAutomateEnabled() {
    if (!elements.powerAutomateEnabledToggle) return;
    const isCurrentlyOn = elements.powerAutomateEnabledToggle.classList.contains('fa-toggle-on');
    updatePowerAutomateEnabledUI(!isCurrentlyOn);
}

export function togglePowerAutomateDebug() {
    if (!elements.powerAutomateDebugToggle) return;
    const isCurrentlyOn = elements.powerAutomateDebugToggle.classList.contains('fa-toggle-on');
    updatePowerAutomateDebugUI(!isCurrentlyOn);
}

export function toggleDebugModeModal() {
    if (!elements.debugModeToggleModal) return;
    const isCurrentlyOn = elements.debugModeToggleModal.classList.contains('fa-toggle-on');
    updateDebugModeModalUI(!isCurrentlyOn);
}

export function toggleAutoSwitchCallTabModal() {
    if (!elements.autoSwitchCallTabToggle) return;
    const isCurrentlyOn = elements.autoSwitchCallTabToggle.classList.contains('fa-toggle-on');
    updateAutoSwitchCallTabUI(!isCurrentlyOn);
}

export function toggleSyncActiveStudentModal() {
    if (!elements.syncActiveStudentToggleModal) return;
    const isCurrentlyOn = elements.syncActiveStudentToggleModal.classList.contains('fa-toggle-on');
    updateSyncActiveStudentModalUI(!isCurrentlyOn);
}

export function toggleSendMasterListModal() {
    if (!elements.sendMasterListToggleModal) return;
    const isCurrentlyOn = elements.sendMasterListToggleModal.classList.contains('fa-toggle-on');
    updateSendMasterListModalUI(!isCurrentlyOn);
}

export function toggleReformatNameModal() {
    if (!elements.reformatNameToggleModal) return;
    const isCurrentlyOn = elements.reformatNameToggleModal.classList.contains('fa-toggle-on');
    updateReformatNameModalUI(!isCurrentlyOn);
}

export function toggleHighlightStudentRowModal() {
    if (!elements.highlightStudentRowToggleModal) return;
    const isCurrentlyOn = elements.highlightStudentRowToggleModal.classList.contains('fa-toggle-on');
    updateHighlightStudentRowModalUI(!isCurrentlyOn);
}

// ============================================
// Status Update Functions
// ============================================

export function updatePowerAutomateStatus(url, enabled = null) {
    if (!elements.powerAutomateStatusText) return;

    if (url && url.trim()) {
        if (enabled === false) {
            elements.powerAutomateStatusText.textContent = 'Disabled';
            elements.powerAutomateStatusText.style.color = 'var(--text-secondary)';
            if (elements.powerAutomateStatusDot) {
                elements.powerAutomateStatusDot.style.backgroundColor = '#9ca3af';
                elements.powerAutomateStatusDot.title = 'Disabled';
            }
        } else {
            elements.powerAutomateStatusText.textContent = 'Connected';
            elements.powerAutomateStatusText.style.color = 'green';
            if (elements.powerAutomateStatusDot) {
                elements.powerAutomateStatusDot.style.backgroundColor = '#10b981';
                elements.powerAutomateStatusDot.title = 'Connected';
            }
        }
    } else {
        elements.powerAutomateStatusText.textContent = 'Not configured';
        elements.powerAutomateStatusText.style.color = 'var(--text-secondary)';
        if (elements.powerAutomateStatusDot) {
            elements.powerAutomateStatusDot.style.backgroundColor = '#9ca3af';
            elements.powerAutomateStatusDot.title = 'Not configured';
        }
    }
}

export async function updateCanvasStatus() {
    if (!elements.canvasStatusText) return;

    try {
        const response = await fetch(`${CANVAS_DOMAIN}/api/v1/users/self`, {
            headers: { 'Accept': 'application/json' },
            credentials: 'include'
        });

        if (response.ok) {
            elements.canvasStatusText.textContent = 'Connected';
            elements.canvasStatusText.style.color = 'green';
            if (elements.canvasStatusDot) {
                elements.canvasStatusDot.style.backgroundColor = '#10b981';
                elements.canvasStatusDot.title = 'Connected';
            }
            await updateStartButtonForMasterList();
        } else {
            elements.canvasStatusText.textContent = 'No user logged in';
            elements.canvasStatusText.style.color = 'var(--text-secondary)';
            if (elements.canvasStatusDot) {
                elements.canvasStatusDot.style.backgroundColor = '#9ca3af';
                elements.canvasStatusDot.title = 'No user logged in';
            }
            if (elements.startBtn) {
                elements.startBtn.disabled = true;
                elements.startBtn.style.opacity = '0.5';
                elements.startBtn.style.cursor = 'not-allowed';
            }
            if (elements.statusText) {
                elements.statusText.textContent = 'Please log into Canvas';
                elements.statusText.style.textDecoration = 'underline';
                elements.statusText.style.color = 'var(--primary-color)';
                elements.statusText.onclick = () => {
                    chrome.tabs.create({ url: CANVAS_DOMAIN });
                };
            }
        }
    } catch (error) {
        console.error('Error checking Canvas status:', error);
        elements.canvasStatusText.textContent = 'No user logged in';
        elements.canvasStatusText.style.color = 'var(--text-secondary)';
        if (elements.canvasStatusDot) {
            elements.canvasStatusDot.style.backgroundColor = '#9ca3af';
            elements.canvasStatusDot.title = 'No user logged in';
        }
        if (elements.startBtn) {
            elements.startBtn.disabled = true;
            elements.startBtn.style.opacity = '0.5';
            elements.startBtn.style.cursor = 'not-allowed';
        }
        if (elements.statusText) {
            elements.statusText.textContent = 'Please log into Canvas';
            elements.statusText.style.textDecoration = 'underline';
            elements.statusText.style.color = 'var(--primary-color)';
            elements.statusText.onclick = () => {
                chrome.tabs.create({ url: CANVAS_DOMAIN });
            };
        }
    }
}

export async function updateStartButtonForMasterList() {
    const data = await storageGet([STORAGE_KEYS.MASTER_ENTRIES]);
    const masterEntries = data[STORAGE_KEYS.MASTER_ENTRIES] || [];

    if (masterEntries.length === 0) {
        if (elements.startBtn) {
            elements.startBtn.disabled = true;
            elements.startBtn.style.opacity = '0.5';
            elements.startBtn.style.cursor = 'not-allowed';
        }
        if (elements.statusText) {
            const currentState = await sessionGetValue(STORAGE_KEYS.EXTENSION_STATE, EXTENSION_STATES.OFF);
            const isCurrentlyScanning = currentState === EXTENSION_STATES.ON;
            if (!isCurrentlyScanning) {
                elements.statusText.textContent = 'No students in Master List';
                elements.statusText.style.textDecoration = 'none';
                elements.statusText.style.color = 'var(--text-secondary)';
                elements.statusText.onclick = null;
            }
        }
        return;
    }

    const studentsWithGradebook = masterEntries.filter(entry =>
        (entry.url && entry.url.trim() !== '') ||
        (entry.Gradebook && entry.Gradebook.trim() !== '')
    );

    if (studentsWithGradebook.length === 0) {
        if (elements.startBtn) {
            elements.startBtn.disabled = true;
            elements.startBtn.style.opacity = '0.5';
            elements.startBtn.style.cursor = 'not-allowed';
        }
        if (elements.statusText) {
            const currentState = await sessionGetValue(STORAGE_KEYS.EXTENSION_STATE, EXTENSION_STATES.OFF);
            const isCurrentlyScanning = currentState === EXTENSION_STATES.ON;
            if (!isCurrentlyScanning) {
                elements.statusText.textContent = 'No gradebook links found';
                elements.statusText.style.textDecoration = 'none';
                elements.statusText.style.color = 'var(--text-secondary)';
                elements.statusText.onclick = null;
            }
        }
        return;
    }

    if (elements.startBtn) {
        elements.startBtn.disabled = false;
        elements.startBtn.style.opacity = '1';
        elements.startBtn.style.cursor = 'pointer';
    }
    if (elements.statusText) {
        const currentState = await sessionGetValue(STORAGE_KEYS.EXTENSION_STATE, EXTENSION_STATES.OFF);
        const isCurrentlyScanning = currentState === EXTENSION_STATES.ON;
        elements.statusText.style.textDecoration = 'none';
        elements.statusText.style.color = '';
        if (!isCurrentlyScanning) {
            elements.statusText.textContent = 'Ready to Scan';
        }
        elements.statusText.onclick = () => {
            if (elements.miniConsole) {
                if (elements.miniConsole.style.display === 'none') {
                    elements.miniConsole.style.display = 'flex';
                } else {
                    elements.miniConsole.style.display = 'none';
                }
            }
        };
    }
}

export async function updateFive9Status() {
    if (!elements.five9StatusText) return;

    try {
        const callDemoEnabled = await storageGetValue(STORAGE_KEYS.CALL_DEMO, false);

        if (callDemoEnabled) {
            elements.five9StatusText.textContent = 'Demo Mode';
            elements.five9StatusText.style.color = '#8b5cf6';
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#8b5cf6';
                elements.five9StatusDot.title = 'Call Demo mode is active - Five9 connection not required';
            }
            return;
        }

        const five9Tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });

        if (five9Tabs.length === 0) {
            elements.five9StatusText.textContent = 'Not connected';
            elements.five9StatusText.style.color = 'var(--text-secondary)';
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#9ca3af';
                elements.five9StatusDot.title = 'No Five9 tab detected';
            }
            return;
        }

        const response = await chrome.runtime.sendMessage({
            type: 'GET_FIVE9_CONNECTION_STATE'
        });

        const connectionState = response ? response.state : FIVE9_CONNECTION_STATES.AWAITING_CONNECTION;

        if (connectionState === FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION) {
            elements.five9StatusText.textContent = 'Active Connection';
            elements.five9StatusText.style.color = 'green';
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#10b981';
                elements.five9StatusDot.title = 'Agent connected and ready';
            }
        } else {
            elements.five9StatusText.textContent = 'Awaiting Agent';
            elements.five9StatusText.style.color = '#f59e0b';
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#f59e0b';
                elements.five9StatusDot.title = 'Tab open, waiting for agent connection';
            }
        }
    } catch (error) {
        console.error('Error checking Five9 status:', error);
        elements.five9StatusText.textContent = 'Not connected';
        elements.five9StatusText.style.color = 'var(--text-secondary)';
        if (elements.five9StatusDot) {
            elements.five9StatusDot.style.backgroundColor = '#9ca3af';
            elements.five9StatusDot.title = 'Error checking status';
        }
    }
}

// ============================================
// Cache Functions
// ============================================

async function loadCacheStatsForModal() {
    if (!elements.cacheStatsTextModal) return;

    try {
        const { getCacheStats } = await import('../../utils/canvasCache.js');
        const stats = await getCacheStats();
        elements.cacheStatsTextModal.textContent =
            `Total: ${stats.totalEntries} | Valid: ${stats.validEntries} | Expired: ${stats.expiredEntries}`;
    } catch (error) {
        console.error('Error loading cache stats:', error);
        elements.cacheStatsTextModal.textContent = 'Error loading stats';
    }
}

export async function clearCacheFromModal() {
    try {
        if (!confirm('Clear all Canvas API cached data? Next update will require fresh API calls.')) {
            return;
        }

        const { clearAllCache } = await import('../../utils/canvasCache.js');
        await clearAllCache();
        await loadCacheStatsForModal();

        alert('✓ Canvas API cache cleared successfully!');
        console.log('Canvas API cache cleared from modal');
    } catch (error) {
        console.error('Error clearing cache from modal:', error);
        alert('Error clearing cache. Check console for details.');
    }
}
