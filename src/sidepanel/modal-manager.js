// Modal Manager - Handles all modal dialogs (scan filter, queue, version history, latest updates)
import { STORAGE_KEYS, CANVAS_DOMAIN, FIVE9_CONNECTION_STATES, EXTENSION_STATES, GENERIC_AVATAR_URL } from '../constants/index.js';
import { storageGet, storageSet, storageGetValue, sessionGetValue } from '../utils/storage.js';
import { encrypt, decrypt } from '../utils/encryption.js';
import { elements } from './ui-manager.js';
import { resolveStudentData } from './student-renderer.js';
import { getReleaseNotes, hasReleaseNotes, getLatestReleaseNotes } from '../constants/release-notes.js';

/**
 * Opens the scan filter modal
 */
export async function openScanFilterModal() {
    if (!elements.scanFilterModal) return;

    // Load current settings
    const settings = await storageGet([
        STORAGE_KEYS.LOOPER_DAYS_OUT_FILTER,
        STORAGE_KEYS.SCAN_FILTER_INCLUDE_FAILING
    ]);

    const daysOutFilter = settings[STORAGE_KEYS.LOOPER_DAYS_OUT_FILTER] || '>=5';
    const includeFailing = settings[STORAGE_KEYS.SCAN_FILTER_INCLUDE_FAILING] || false;

    // Parse days out filter (e.g., ">=5" -> operator: ">=", value: "5")
    const match = daysOutFilter.match(/^\s*([><]=?|=)\s*(\d+)\s*$/);
    if (match && elements.daysOutOperator && elements.daysOutValue) {
        elements.daysOutOperator.value = match[1];
        elements.daysOutValue.value = match[2];
    }

    // Set failing toggle state
    if (elements.failingToggle) {
        if (includeFailing) {
            elements.failingToggle.className = 'fas fa-toggle-on';
            elements.failingToggle.style.color = 'var(--primary-color)';
        } else {
            elements.failingToggle.className = 'fas fa-toggle-off';
            elements.failingToggle.style.color = 'gray';
        }
    }

    // Calculate and display initial count
    await updateScanFilterCount();

    // Show modal
    elements.scanFilterModal.style.display = 'flex';
}

/**
 * Closes the scan filter modal
 */
export function closeScanFilterModal() {
    if (!elements.scanFilterModal) return;
    elements.scanFilterModal.style.display = 'none';
}

/**
 * Updates the student count based on current filter settings
 */
export async function updateScanFilterCount() {
    if (!elements.daysOutOperator || !elements.daysOutValue || !elements.failingToggle || !elements.studentCountValue) return;

    const operator = elements.daysOutOperator.value;
    const value = parseInt(elements.daysOutValue.value, 10);
    const includeFailing = elements.failingToggle.classList.contains('fa-toggle-on');

    const data = await storageGet([STORAGE_KEYS.MASTER_ENTRIES]);
    const masterEntries = data[STORAGE_KEYS.MASTER_ENTRIES] || [];

    let filteredCount = 0;

    masterEntries.forEach(entry => {
        const daysOut = entry.daysOut;

        let meetsDaysOutCriteria = false;
        if (daysOut != null) {
            switch (operator) {
                case '>': meetsDaysOutCriteria = daysOut > value; break;
                case '<': meetsDaysOutCriteria = daysOut < value; break;
                case '>=': meetsDaysOutCriteria = daysOut >= value; break;
                case '<=': meetsDaysOutCriteria = daysOut <= value; break;
                case '=': meetsDaysOutCriteria = daysOut === value; break;
                default: meetsDaysOutCriteria = false;
            }
        }

        let isFailing = false;
        if (includeFailing && entry.grade != null) {
            const grade = parseFloat(entry.grade);
            if (!isNaN(grade) && grade < 60) {
                isFailing = true;
            }
        }

        if (meetsDaysOutCriteria || isFailing) {
            filteredCount++;
        }
    });

    elements.studentCountValue.textContent = filteredCount;
}

/**
 * Toggles the failing filter state
 */
export function toggleFailingFilter() {
    if (!elements.failingToggle) return;

    const isOn = elements.failingToggle.classList.contains('fa-toggle-on');
    if (isOn) {
        elements.failingToggle.className = 'fas fa-toggle-off';
        elements.failingToggle.style.color = 'gray';
    } else {
        elements.failingToggle.className = 'fas fa-toggle-on';
        elements.failingToggle.style.color = 'var(--primary-color)';
    }
}

/**
 * Saves the scan filter settings
 */
export async function saveScanFilterSettings() {
    if (!elements.daysOutOperator || !elements.daysOutValue || !elements.failingToggle) return;

    const operator = elements.daysOutOperator.value;
    const value = elements.daysOutValue.value;
    const daysOutFilter = `${operator}${value}`;
    const includeFailing = elements.failingToggle.classList.contains('fa-toggle-on');

    await storageSet({
        [STORAGE_KEYS.LOOPER_DAYS_OUT_FILTER]: daysOutFilter,
        [STORAGE_KEYS.SCAN_FILTER_INCLUDE_FAILING]: includeFailing
    });

    closeScanFilterModal();
    console.log('Scan filter settings saved:', { daysOutFilter, includeFailing });
}

/**
 * Opens the queue management modal
 */
export function openQueueModal(selectedQueue, onReorder, onRemove) {
    if (!elements.queueModal || !elements.queueList) return;

    renderQueueModal(selectedQueue, onReorder, onRemove);
    elements.queueModal.style.display = 'flex';
}

/**
 * Closes the queue management modal
 */
export function closeQueueModal() {
    if (!elements.queueModal) return;
    elements.queueModal.style.display = 'none';
}

/**
 * Renders the queue modal content
 */
export async function renderQueueModal(selectedQueue, onReorder, onRemove) {
    if (!elements.queueList || !elements.queueCount) return;

    elements.queueList.innerHTML = '';

    if (selectedQueue.length === 0) {
        elements.queueList.innerHTML = '<li style="justify-content:center; color:gray;">No students in queue</li>';
        elements.queueCount.textContent = '0 students';
        return;
    }

    elements.queueCount.textContent = `${selectedQueue.length} student${selectedQueue.length !== 1 ? 's' : ''}`;

    // Get reformat name setting
    const reformatEnabled = await storageGetValue(STORAGE_KEYS.REFORMAT_NAME_ENABLED, true);

    selectedQueue.forEach((student, index) => {
        const li = document.createElement('li');
        li.className = 'queue-item-draggable';
        li.draggable = true;
        li.dataset.index = index;

        const data = resolveStudentData(student, reformatEnabled);

        li.innerHTML = `
            <div style="display: flex; align-items: center; width: 100%; justify-content: space-between;">
                <div style="display: flex; align-items: center; flex-grow: 1;">
                    <i class="fas fa-grip-vertical queue-drag-handle"></i>
                    <div style="margin-right: 10px; font-weight: 600; color: var(--text-secondary); min-width: 20px;">#${index + 1}</div>
                    <div>
                        <div style="font-weight: 500; color: var(--text-main);">${data.name}</div>
                        <div style="font-size: 0.8em; color: var(--text-secondary);">${data.daysOut} Days Out</div>
                    </div>
                </div>
                <button class="queue-remove-btn" data-index="${index}" title="Remove from queue">
                    <i class="fas fa-times"></i>
                </button>
            </div>
        `;

        // Drag events
        li.addEventListener('dragstart', (e) => handleDragStart(e));
        li.addEventListener('dragend', (e) => handleDragEnd(e));
        li.addEventListener('dragover', (e) => handleDragOver(e));
        li.addEventListener('drop', (e) => handleDrop(e, onReorder));
        li.addEventListener('dragleave', (e) => handleDragLeave(e));

        // Remove button
        const removeBtn = li.querySelector('.queue-remove-btn');
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            if (onRemove) {
                onRemove(index);
            }
        });

        elements.queueList.appendChild(li);
    });
}

// Drag and drop state
let draggedElement = null;
let draggedIndex = null;

function handleDragStart(e) {
    draggedElement = e.currentTarget;
    draggedIndex = parseInt(e.currentTarget.dataset.index);
    e.currentTarget.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/html', e.currentTarget.innerHTML);
}

function handleDragEnd(e) {
    e.currentTarget.classList.remove('dragging');
    document.querySelectorAll('.queue-item-draggable').forEach(item => {
        item.classList.remove('drag-over');
    });
}

function handleDragOver(e) {
    if (e.preventDefault) {
        e.preventDefault();
    }
    e.dataTransfer.dropEffect = 'move';

    const afterElement = e.currentTarget;
    if (afterElement !== draggedElement) {
        afterElement.classList.add('drag-over');
    }

    return false;
}

function handleDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
}

function handleDrop(e, onReorder) {
    if (e.stopPropagation) {
        e.stopPropagation();
    }

    const dropIndex = parseInt(e.currentTarget.dataset.index);

    if (draggedIndex !== dropIndex && onReorder) {
        onReorder(draggedIndex, dropIndex);
    }

    return false;
}

/**
 * Opens the connections modal for a specific connection type
 * @param {string} connectionType - 'excel', 'powerAutomate', or 'canvas'
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
    const autoSwitchCallTab = result[STORAGE_KEYS.AUTO_SWITCH_TO_CALL_TAB] || false;
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
 * Updates the embed helper toggle UI in the modal
 * @param {boolean} isEnabled - Whether embed helper is enabled
 */
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

/**
 * Updates the canvas cache toggle UI in the modal
 * @param {boolean} isEnabled - Whether canvas cache is enabled
 */
function updateCanvasCacheModalUI(isEnabled) {
    if (!elements.canvasCacheToggleModal) return;

    if (isEnabled) {
        elements.canvasCacheToggleModal.className = 'fas fa-toggle-on';
        elements.canvasCacheToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.canvasCacheToggleModal.className = 'fas fa-toggle-off';
        elements.canvasCacheToggleModal.style.color = 'gray';
    }

    // Enable/disable the cache settings container
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

/**
 * Updates the non-API course fetch toggle UI in the modal
 * @param {boolean} isEnabled - Whether non-API course fetch is enabled
 */
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

/**
 * Updates the next assignment toggle UI in the modal
 * @param {boolean} isEnabled - Whether next assignment feature is enabled
 */
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

/**
 * Updates the debug mode toggle UI in the modal
 * @param {boolean} isEnabled - Whether debug mode is enabled
 */
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

/**
 * Updates the auto switch to call tab toggle UI in the modal
 * @param {boolean} isEnabled - Whether auto switch to call tab is enabled
 */
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

/**
 * Updates the sync active student toggle UI in the modal
 * @param {boolean} isEnabled - Whether sync active student is enabled
 */
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

/**
 * Updates the send master list to Excel toggle UI in the modal
 * @param {boolean} isEnabled - Whether sending master list to Excel is enabled
 */
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

/**
 * Updates the reformat name toggle UI in the modal
 * @param {boolean} isEnabled - Whether name reformatting is enabled
 */
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

/**
 * Updates the highlight student row toggle UI in the modal
 * @param {boolean} isEnabled - Whether student row highlighting is enabled
 */
function updateHighlightStudentRowModalUI(isEnabled) {
    if (!elements.highlightStudentRowToggleModal) return;

    if (isEnabled) {
        elements.highlightStudentRowToggleModal.className = 'fas fa-toggle-on';
        elements.highlightStudentRowToggleModal.style.color = 'var(--primary-color)';
    } else {
        elements.highlightStudentRowToggleModal.className = 'fas fa-toggle-off';
        elements.highlightStudentRowToggleModal.style.color = 'gray';
    }

    // Enable/disable the settings container
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

/**
 * Loads cache stats for the modal
 */
async function loadCacheStatsForModal() {
    if (!elements.cacheStatsTextModal) return;

    try {
        const { getCacheStats } = await import('../utils/canvasCache.js');
        const stats = await getCacheStats();

        // Format stats as readable text
        elements.cacheStatsTextModal.textContent =
            `Total: ${stats.totalEntries} | Valid: ${stats.validEntries} | Expired: ${stats.expiredEntries}`;
    } catch (error) {
        console.error('Error loading cache stats:', error);
        elements.cacheStatsTextModal.textContent = 'Error loading stats';
    }
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

/**
 * Updates the Power Automate connection status text
 * @param {string} url - The Power Automate URL (empty if not configured)
 * @param {boolean} enabled - Whether Power Automate is enabled (optional)
 */
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

/**
 * Updates the Power Automate enabled toggle UI and greys out settings when disabled
 * @param {boolean} isEnabled - Whether Power Automate is enabled
 */
function updatePowerAutomateEnabledUI(isEnabled) {
    if (!elements.powerAutomateEnabledToggle) return;

    if (isEnabled) {
        elements.powerAutomateEnabledToggle.className = 'fas fa-toggle-on';
        elements.powerAutomateEnabledToggle.style.color = 'var(--primary-color)';
    } else {
        elements.powerAutomateEnabledToggle.className = 'fas fa-toggle-off';
        elements.powerAutomateEnabledToggle.style.color = 'gray';
    }

    // Grey out settings when disabled
    if (elements.powerAutomateSettingsContainer) {
        elements.powerAutomateSettingsContainer.style.opacity = isEnabled ? '1' : '0.5';
        elements.powerAutomateSettingsContainer.style.pointerEvents = isEnabled ? 'auto' : 'none';
    }
    if (elements.powerAutomateDebugContainer) {
        elements.powerAutomateDebugContainer.style.opacity = isEnabled ? '1' : '0.5';
        elements.powerAutomateDebugContainer.style.pointerEvents = isEnabled ? 'auto' : 'none';
    }
}

/**
 * Updates the Power Automate debug toggle UI
 * @param {boolean} isEnabled - Whether debug mode is enabled
 */
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

/**
 * Updates the Canvas connection status text based on login state
 */
export async function updateCanvasStatus() {
    if (!elements.canvasStatusText) return;

    try {
        // Check if user is logged in by attempting to fetch current user profile
        const response = await fetch('https://nuc.instructure.com/api/v1/users/self', {
            headers: { 'Accept': 'application/json' },
            credentials: 'include'
        });

        if (response.ok) {
            // User is logged in
            elements.canvasStatusText.textContent = 'Connected';
            elements.canvasStatusText.style.color = 'green';
            if (elements.canvasStatusDot) {
                elements.canvasStatusDot.style.backgroundColor = '#10b981';
                elements.canvasStatusDot.title = 'Connected';
            }

            // Check master list status to determine if Start button should be enabled
            await updateStartButtonForMasterList();
        } else {
            // Not logged in or authentication failed
            elements.canvasStatusText.textContent = 'No user logged in';
            elements.canvasStatusText.style.color = 'var(--text-secondary)';
            if (elements.canvasStatusDot) {
                elements.canvasStatusDot.style.backgroundColor = '#9ca3af';
                elements.canvasStatusDot.title = 'No user logged in';
            }

            // Disable start button and update status text as clickable link
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

        // Disable start button and update status text as clickable link
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

/**
 * Checks master list for students with gradebook links and updates Start button state
 * Should be called after verifying Canvas connection is OK
 */
export async function updateStartButtonForMasterList() {
    const data = await storageGet([STORAGE_KEYS.MASTER_ENTRIES]);
    const masterEntries = data[STORAGE_KEYS.MASTER_ENTRIES] || [];

    // Check if there are any students
    if (masterEntries.length === 0) {
        if (elements.startBtn) {
            elements.startBtn.disabled = true;
            elements.startBtn.style.opacity = '0.5';
            elements.startBtn.style.cursor = 'not-allowed';
        }
        if (elements.statusText) {
            // Check if scanner is currently running before overwriting status
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

    // Check if any students have gradebook links (url or Gradebook field)
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
            // Check if scanner is currently running before overwriting status
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

    // All checks passed - enable the button
    if (elements.startBtn) {
        elements.startBtn.disabled = false;
        elements.startBtn.style.opacity = '1';
        elements.startBtn.style.cursor = 'pointer';
    }
    if (elements.statusText) {
        // Check if scanner is currently running before overwriting status
        const currentState = await sessionGetValue(STORAGE_KEYS.EXTENSION_STATE, EXTENSION_STATES.OFF);
        const isCurrentlyScanning = currentState === EXTENSION_STATES.ON;

        // Remove link styling
        elements.statusText.style.textDecoration = 'none';
        elements.statusText.style.color = '';

        // Only update text if not currently scanning
        if (!isCurrentlyScanning) {
            elements.statusText.textContent = 'Ready to Scan';
        }

        // Set onclick to toggle mini console when connected
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

/**
 * Toggles the embed helper setting in the modal
 */
export function toggleEmbedHelperModal() {
    if (!elements.embedHelperToggleModal) return;

    const isCurrentlyOn = elements.embedHelperToggleModal.classList.contains('fa-toggle-on');
    updateEmbedHelperModalUI(!isCurrentlyOn);
}

/**
 * Toggles the canvas cache setting in the modal
 */
export function toggleCanvasCacheModal() {
    if (!elements.canvasCacheToggleModal) return;

    const isCurrentlyOn = elements.canvasCacheToggleModal.classList.contains('fa-toggle-on');
    updateCanvasCacheModalUI(!isCurrentlyOn);
}

/**
 * Toggles the non-API course fetch setting in the modal
 */
export function toggleNonApiCourseFetch() {
    if (!elements.nonApiCourseFetchToggle) return;

    const isCurrentlyOn = elements.nonApiCourseFetchToggle.classList.contains('fa-toggle-on');
    updateNonApiCourseFetchUI(!isCurrentlyOn);
}

/**
 * Toggles the next assignment setting in the modal
 */
export function toggleNextAssignment() {
    if (!elements.nextAssignmentToggle) return;

    const isCurrentlyOn = elements.nextAssignmentToggle.classList.contains('fa-toggle-on');
    updateNextAssignmentUI(!isCurrentlyOn);
}

/**
 * Toggles the Power Automate enabled setting in the modal
 */
export function togglePowerAutomateEnabled() {
    if (!elements.powerAutomateEnabledToggle) return;

    const isCurrentlyOn = elements.powerAutomateEnabledToggle.classList.contains('fa-toggle-on');
    updatePowerAutomateEnabledUI(!isCurrentlyOn);
}

/**
 * Toggles the Power Automate debug setting in the modal
 */
export function togglePowerAutomateDebug() {
    if (!elements.powerAutomateDebugToggle) return;

    const isCurrentlyOn = elements.powerAutomateDebugToggle.classList.contains('fa-toggle-on');
    updatePowerAutomateDebugUI(!isCurrentlyOn);
}

/**
 * Toggles the debug mode setting in the modal
 */
export function toggleDebugModeModal() {
    if (!elements.debugModeToggleModal) return;

    const isCurrentlyOn = elements.debugModeToggleModal.classList.contains('fa-toggle-on');
    updateDebugModeModalUI(!isCurrentlyOn);
}

/**
 * Toggles the auto switch to call tab setting in the modal
 */
export function toggleAutoSwitchCallTabModal() {
    if (!elements.autoSwitchCallTabToggle) return;

    const isCurrentlyOn = elements.autoSwitchCallTabToggle.classList.contains('fa-toggle-on');
    updateAutoSwitchCallTabUI(!isCurrentlyOn);
}

/**
 * Toggles the sync active student setting in the modal
 */
export function toggleSyncActiveStudentModal() {
    if (!elements.syncActiveStudentToggleModal) return;

    const isCurrentlyOn = elements.syncActiveStudentToggleModal.classList.contains('fa-toggle-on');
    updateSyncActiveStudentModalUI(!isCurrentlyOn);
}

/**
 * Toggles the send master list to Excel setting in the modal
 */
export function toggleSendMasterListModal() {
    if (!elements.sendMasterListToggleModal) return;

    const isCurrentlyOn = elements.sendMasterListToggleModal.classList.contains('fa-toggle-on');
    updateSendMasterListModalUI(!isCurrentlyOn);
}

/**
 * Toggles the reformat name setting in the modal
 */
export function toggleReformatNameModal() {
    if (!elements.reformatNameToggleModal) return;

    const isCurrentlyOn = elements.reformatNameToggleModal.classList.contains('fa-toggle-on');
    updateReformatNameModalUI(!isCurrentlyOn);
}

/**
 * Toggles the highlight student row setting in the modal
 */
export function toggleHighlightStudentRowModal() {
    if (!elements.highlightStudentRowToggleModal) return;

    const isCurrentlyOn = elements.highlightStudentRowToggleModal.classList.contains('fa-toggle-on');
    updateHighlightStudentRowModalUI(!isCurrentlyOn);
}

/**
 * Updates the Five9 connection status text based on active tabs
 */
export async function updateFive9Status() {
    if (!elements.five9StatusText) return;

    try {
        // Check if Call Demo mode is enabled
        const callDemoEnabled = await storageGetValue(STORAGE_KEYS.CALL_DEMO, false);

        if (callDemoEnabled) {
            // Demo mode - show demo status instead of Five9 connection
            elements.five9StatusText.textContent = 'Demo Mode';
            elements.five9StatusText.style.color = '#8b5cf6'; // Purple for demo
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#8b5cf6';
                elements.five9StatusDot.title = 'Call Demo mode is active - Five9 connection not required';
            }
            return;
        }

        // Check Five9 tab status
        const five9Tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });

        if (five9Tabs.length === 0) {
            // No tab - Not connected
            elements.five9StatusText.textContent = 'Not connected';
            elements.five9StatusText.style.color = 'var(--text-secondary)';
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#9ca3af';
                elements.five9StatusDot.title = 'No Five9 tab detected';
            }
            return;
        }

        // Tab exists - check agent connection state
        const response = await chrome.runtime.sendMessage({
            type: 'GET_FIVE9_CONNECTION_STATE'
        });

        const connectionState = response ? response.state : FIVE9_CONNECTION_STATES.AWAITING_CONNECTION;

        if (connectionState === FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION) {
            // Active connection
            elements.five9StatusText.textContent = 'Active Connection';
            elements.five9StatusText.style.color = 'green';
            if (elements.five9StatusDot) {
                elements.five9StatusDot.style.backgroundColor = '#10b981';
                elements.five9StatusDot.title = 'Agent connected and ready';
            }
        } else {
            // Awaiting connection
            elements.five9StatusText.textContent = 'Awaiting Agent';
            elements.five9StatusText.style.color = '#f59e0b'; // Orange/amber color
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

/**
 * Clears the Canvas API cache from the modal
 */
export async function clearCacheFromModal() {
    try {
        if (!confirm('Clear all Canvas API cached data? Next update will require fresh API calls.')) {
            return;
        }

        const { clearAllCache } = await import('../utils/canvasCache.js');
        await clearAllCache();

        // Reload stats to show empty cache
        await loadCacheStatsForModal();

        alert('✓ Canvas API cache cleared successfully!');
        console.log('Canvas API cache cleared from modal');
    } catch (error) {
        console.error('Error clearing cache from modal:', error);
        alert('Error clearing cache. Check console for details.');
    }
}

/**
 * Checks if the daily update modal should be shown
 * Returns true if the modal should be shown, false otherwise
 */
export async function shouldShowDailyUpdateModal() {
    const lastUpdated = await storageGetValue(STORAGE_KEYS.LAST_UPDATED);

    // If there's no master list yet, don't show the modal
    if (!lastUpdated) {
        return false;
    }

    const now = new Date();
    const todayDateString = now.toLocaleDateString('en-US');

    // Check if master list was updated today
    const lastUpdatedDate = new Date(lastUpdated);
    const lastUpdatedDateString = lastUpdatedDate.toLocaleDateString('en-US');

    if (todayDateString === lastUpdatedDateString) {
        // Master list already updated today
        return false;
    }

    // Show modal if:
    // 1. Master list exists
    // 2. Master list hasn't been updated today
    return true;
}

/**
 * Opens the daily update modal
 */
export function openDailyUpdateModal() {
    if (!elements.dailyUpdateModal) return;
    elements.dailyUpdateModal.style.display = 'flex';
}

/**
 * Closes the daily update modal
 */
export async function closeDailyUpdateModal() {
    if (!elements.dailyUpdateModal) return;
    elements.dailyUpdateModal.style.display = 'none';
}

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

// Campus Selection Modal state
let campusSelectionResolve = null;

/**
 * Gets unique campuses from student array
 * @param {Array} students - Array of student objects
 * @returns {Array} Array of unique campus names, sorted alphabetically
 */
export function getCampusesFromStudents(students) {
    if (!students || students.length === 0) return [];

    const campuses = [...new Set(
        students
            .map(s => s.campus)
            .filter(c => c && c.trim() !== '')
    )].sort();

    return campuses;
}

/**
 * Opens the Campus Selection modal and returns a promise that resolves with the selected campus
 * @param {Array} campuses - Array of campus names
 * @param {string} [customMessage] - Optional custom message to display in the modal
 * @returns {Promise<string|null>} Promise that resolves with selected campus (empty string for all) or null if cancelled
 */
export function openCampusSelectionModal(campuses, customMessage = null) {
    return new Promise((resolve) => {
        if (!elements.campusSelectionModal || !elements.campusSelectionList) {
            resolve(''); // Default to all campuses if modal not available
            return;
        }

        // Store resolve function for later
        campusSelectionResolve = resolve;

        // Update message if custom message provided
        if (elements.campusSelectionMessage) {
            elements.campusSelectionMessage.textContent = customMessage ||
                'Multiple campuses detected. Select which campus data to send:';
        }

        // Clear existing buttons
        elements.campusSelectionList.innerHTML = '';

        // Create "All Campuses" button first
        const allButton = document.createElement('button');
        allButton.className = 'btn-secondary';
        allButton.style.cssText = 'width: 100%; text-align: left; padding: 12px 15px; display: flex; align-items: center; gap: 10px;';
        allButton.innerHTML = `
            <i class="fas fa-globe" style="color: #6366f1; font-size: 1.2em;"></i>
            <span style="flex: 1;">All Campuses</span>
            <span style="font-size: 0.85em; color: var(--text-secondary);">(Send all data)</span>
        `;
        allButton.addEventListener('click', () => {
            closeCampusSelectionModal('');
        });
        elements.campusSelectionList.appendChild(allButton);

        // Create a button for each campus
        campuses.forEach(campus => {
            const button = document.createElement('button');
            button.className = 'btn-secondary';
            button.style.cssText = 'width: 100%; text-align: left; padding: 12px 15px; display: flex; align-items: center; gap: 10px;';

            button.innerHTML = `
                <i class="fas fa-building" style="color: #6366f1; font-size: 1.2em;"></i>
                <span style="flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${campus}</span>
            `;
            button.title = campus;

            button.addEventListener('click', () => {
                closeCampusSelectionModal(campus);
            });

            elements.campusSelectionList.appendChild(button);
        });

        // Show modal
        elements.campusSelectionModal.style.display = 'flex';
    });
}

/**
 * Closes the Campus Selection modal
 * @param {string|null} selectedCampus - The selected campus or null if cancelled
 */
export function closeCampusSelectionModal(selectedCampus = null) {
    if (elements.campusSelectionModal) {
        elements.campusSelectionModal.style.display = 'none';
    }

    // Resolve the promise with the selected campus
    if (campusSelectionResolve) {
        campusSelectionResolve(selectedCampus);
        campusSelectionResolve = null;
    }
}

// ============================================================================
// STUDENT VIEW MODAL
// ============================================================================

// Store the current student for the student view modal
let currentStudentViewStudent = null;

/**
 * Gets the current student displayed in the student view modal
 * @returns {Object|null} The current student data or null
 */
export function getCurrentStudentViewStudent() {
    return currentStudentViewStudent;
}

/**
 * Opens the student view modal with student details
 * @param {Object} student - The student data object
 * @param {boolean} hasMultipleCampuses - Whether there are multiple campuses in the list
 */
export function openStudentViewModal(student, hasMultipleCampuses = false) {
    if (!elements.studentViewModal || !student) return;

    // Store the current student for access by email button
    currentStudentViewStudent = student;

    const data = resolveStudentData(student);

    // Reset to main view
    showStudentViewMain();

    // Generate initials for avatar fallback
    const nameParts = (data.name || '').trim().split(/\s+/);
    let initials = '';
    if (nameParts.length > 0) {
        const firstInitial = nameParts[0][0] || '';
        const lastInitial = nameParts.length > 1 ? nameParts[nameParts.length - 1][0] : '';
        initials = (firstInitial + lastInitial).toUpperCase();
        if (!initials) initials = '?';
    }

    // Avatar
    if (elements.studentViewAvatar) {
        if (data.Photo && data.Photo !== GENERIC_AVATAR_URL) {
            elements.studentViewAvatar.textContent = '';
            elements.studentViewAvatar.style.backgroundImage = `url('${data.Photo}')`;
            elements.studentViewAvatar.style.backgroundSize = 'cover';
            elements.studentViewAvatar.style.backgroundPosition = 'center';
            elements.studentViewAvatar.style.backgroundColor = 'transparent';
            elements.studentViewAvatar.style.color = 'transparent';
        } else {
            elements.studentViewAvatar.style.backgroundImage = 'none';
            elements.studentViewAvatar.textContent = initials;

            // Gender-based avatar colors
            const gender = (student.Gender || student.gender || '').toLowerCase();
            if (gender === 'male' || gender === 'm' || gender === 'boy') {
                elements.studentViewAvatar.style.backgroundColor = 'rgb(18, 120, 255)'; // Blue
                elements.studentViewAvatar.style.color = 'rgb(255, 255, 255)'; // White
            } else if (gender === 'female' || gender === 'f' || gender === 'girl') {
                elements.studentViewAvatar.style.backgroundColor = 'rgb(255, 145, 175)'; // Pastel pink
                elements.studentViewAvatar.style.color = 'rgb(255, 255, 255)'; // White
            } else {
                elements.studentViewAvatar.style.backgroundColor = '#e5e7eb'; // Gray
                elements.studentViewAvatar.style.color = '#6b7280';
            }
        }
    }

    // Name
    if (elements.studentViewName) {
        elements.studentViewName.textContent = data.name || 'Unknown Student';
    }

    // Campus (only show if multiple campuses) - access raw student field
    if (elements.studentViewCampus) {
        const campusSpan = elements.studentViewCampus.querySelector('span');
        const campus = student.campus || student.Campus || '';
        if (hasMultipleCampuses && campus) {
            elements.studentViewCampus.style.display = 'block';
            if (campusSpan) campusSpan.textContent = campus;
        } else {
            elements.studentViewCampus.style.display = 'none';
        }
    }

    // New badge
    if (elements.studentViewNewBadge) {
        elements.studentViewNewBadge.style.display = data.isNew ? 'block' : 'none';
    }

    // Days Out - gray by default, red if >= 14
    if (elements.studentViewDaysOut && elements.studentViewDaysOutCard) {
        const daysOut = data.daysOut !== undefined && data.daysOut !== null ? data.daysOut : '-';
        elements.studentViewDaysOut.textContent = daysOut;

        const daysOutNum = parseInt(daysOut) || 0;
        if (daysOutNum >= 14) {
            elements.studentViewDaysOutCard.style.background = 'rgba(239, 68, 68, 0.15)';
            elements.studentViewDaysOutCard.style.color = '#b91c1c';
        } else {
            elements.studentViewDaysOutCard.style.background = 'rgba(107, 114, 128, 0.15)';
            elements.studentViewDaysOutCard.style.color = '#4b5563';
        }
    }

    // Days Out Detail View population
    if (elements.studentViewDaysOutTitle && elements.studentViewDaysLeftText && elements.studentViewDeadlineText) {
        const daysOut = data.daysOut !== undefined && data.daysOut !== null ? parseInt(data.daysOut) : 0;

        // Title: "X Days Out"
        elements.studentViewDaysOutTitle.textContent = `${daysOut} Days Out`;

        // Calculate days left (assuming 14 days total for a standard period)
        const daysLeft = Math.max(0, 14 - daysOut);
        elements.studentViewDaysLeftText.textContent = `The student has ${daysLeft} day${daysLeft !== 1 ? 's' : ''} left.`;

        // Calculate deadline date
        const today = new Date();
        const deadlineDate = new Date(today);
        deadlineDate.setDate(today.getDate() + daysLeft);

        const options = { year: 'numeric', month: 'long', day: 'numeric' };
        const formattedDate = deadlineDate.toLocaleDateString('en-US', options);
        elements.studentViewDeadlineText.textContent = `They have until ${formattedDate} to submit work.`;
    }

    // Grade - access raw student field, always gray
    if (elements.studentViewGrade && elements.studentViewGradeCard) {
        const rawGrade = student.grade ?? student.Grade ?? student.currentGrade ?? null;
        const grade = rawGrade !== undefined && rawGrade !== null ? rawGrade : '-';
        elements.studentViewGrade.textContent = (typeof grade === 'number' || !isNaN(parseFloat(grade))) && grade !== '-' ? `${parseFloat(grade).toFixed(0)}%` : grade;

        // Always gray
        elements.studentViewGradeCard.style.background = 'rgba(107, 114, 128, 0.15)';
        elements.studentViewGradeCard.style.color = '#4b5563';
    }

    // Missing Assignments Card - show count, always gray
    const missing = student.missingAssignments || [];
    if (elements.studentViewMissingCount && elements.studentViewMissingCard) {
        elements.studentViewMissingCount.textContent = missing.length;

        // Always gray
        elements.studentViewMissingCard.style.background = 'rgba(107, 114, 128, 0.15)';
        elements.studentViewMissingCard.style.color = '#4b5563';
    }

    // Populate missing assignments list for detail view
    if (elements.studentViewMissingList) {
        if (missing.length > 0) {
            elements.studentViewMissingList.innerHTML = missing.map((assignment, index) => {
                const title = assignment.assignmentTitle || assignment.Assignment || assignment.name || `Assignment ${index + 1}`;
                const dueDate = assignment.dueDate || assignment.DueDate || assignment.due_at || '';
                return `
                    <div style="padding: 10px 12px; background: rgba(245, 158, 11, 0.08); border-radius: 8px; margin-bottom: 8px;">
                        <div style="color: var(--text-main); font-weight: 500; font-size: 0.9em;">${title}</div>
                        ${dueDate ? `<div style="color: var(--text-secondary); font-size: 0.8em; margin-top: 4px;"><i class="fas fa-clock" style="margin-right: 4px;"></i>Due: ${dueDate}</div>` : ''}
                    </div>
                `;
            }).join('');
        } else {
            elements.studentViewMissingList.innerHTML = '<div style="color: var(--text-secondary); font-size: 0.9em; text-align: center; padding: 40px 20px;"><i class="fas fa-check-circle" style="font-size: 2em; margin-bottom: 10px; opacity: 0.5; display: block;"></i>No missing assignments</div>';
        }
    }

    // Next Assignment Card - show due date (always gray)
    const nextAssignment = student.nextAssignment || null;
    if (elements.studentViewNextDate && elements.studentViewNextCard) {
        if (nextAssignment && nextAssignment.DueDate) {
            elements.studentViewNextDate.textContent = nextAssignment.DueDate;
        } else {
            elements.studentViewNextDate.textContent = '-';
        }
        // Always use gray color
        elements.studentViewNextCard.style.background = 'rgba(107, 114, 128, 0.15)';
        elements.studentViewNextCard.style.color = '#4b5563';
    }

    // Populate next assignment detail view
    if (nextAssignment && nextAssignment.Assignment) {
        if (elements.studentViewNextAssignment) {
            elements.studentViewNextAssignment.textContent = nextAssignment.Assignment;
        }
        if (elements.studentViewNextAssignmentDate) {
            elements.studentViewNextAssignmentDate.textContent = nextAssignment.DueDate ? `Due: ${nextAssignment.DueDate}` : 'No due date';
        }
        if (elements.studentViewNextDetailContent) {
            elements.studentViewNextDetailContent.style.display = 'block';
        }
        if (elements.studentViewNoNextAssignment) {
            elements.studentViewNoNextAssignment.style.display = 'none';
        }
    } else {
        if (elements.studentViewNextDetailContent) {
            elements.studentViewNextDetailContent.style.display = 'none';
        }
        if (elements.studentViewNoNextAssignment) {
            elements.studentViewNoNextAssignment.style.display = 'block';
        }
    }

    // Store current student reference for button actions
    elements.studentViewModal.dataset.studentName = data.name || '';

    // Show modal
    elements.studentViewModal.style.display = 'flex';
}

/**
 * Shows the main view of the student modal
 */
export function showStudentViewMain() {
    if (elements.studentViewMain) elements.studentViewMain.style.display = 'block';
    if (elements.studentViewMissingDetail) elements.studentViewMissingDetail.style.display = 'none';
    if (elements.studentViewNextDetail) elements.studentViewNextDetail.style.display = 'none';
    if (elements.studentViewDaysOutDetail) elements.studentViewDaysOutDetail.style.display = 'none';
}

/**
 * Shows the missing assignments detail view
 */
export function showStudentViewMissing() {
    if (elements.studentViewMain) elements.studentViewMain.style.display = 'none';
    if (elements.studentViewMissingDetail) elements.studentViewMissingDetail.style.display = 'block';
    if (elements.studentViewNextDetail) elements.studentViewNextDetail.style.display = 'none';
    if (elements.studentViewDaysOutDetail) elements.studentViewDaysOutDetail.style.display = 'none';
}

/**
 * Shows the next assignment detail view
 */
export function showStudentViewNext() {
    if (elements.studentViewMain) elements.studentViewMain.style.display = 'none';
    if (elements.studentViewMissingDetail) elements.studentViewMissingDetail.style.display = 'none';
    if (elements.studentViewNextDetail) elements.studentViewNextDetail.style.display = 'block';
    if (elements.studentViewDaysOutDetail) elements.studentViewDaysOutDetail.style.display = 'none';
}

/**
 * Shows the days out detail view
 */
export function showStudentViewDaysOut() {
    if (elements.studentViewMain) elements.studentViewMain.style.display = 'none';
    if (elements.studentViewMissingDetail) elements.studentViewMissingDetail.style.display = 'none';
    if (elements.studentViewNextDetail) elements.studentViewNextDetail.style.display = 'none';
    if (elements.studentViewDaysOutDetail) elements.studentViewDaysOutDetail.style.display = 'block';
}

/**
 * Closes the student view modal
 */
export function closeStudentViewModal() {
    if (elements.studentViewModal) {
        elements.studentViewModal.style.display = 'none';
    }
    // Clear the stored student reference
    currentStudentViewStudent = null;
    // Reset to main view for next time
    showStudentViewMain();
}

/**
 * Gets the appropriate greeting based on time of day
 * @returns {string} "Good Morning", "Good Afternoon", or "Good Evening"
 */
function getTimeOfDayGreeting() {
    const hour = new Date().getHours();
    if (hour < 12) return 'Good Morning';
    if (hour < 17) return 'Good Afternoon';
    return 'Good Evening';
}

/**
 * Gets the first name from a full name
 * @param {string} fullName - The full name
 * @returns {string} The first name
 */
function getFirstName(fullName) {
    if (!fullName) return '';
    const parts = fullName.trim().split(/\s+/);
    return parts[0] || '';
}

/**
 * Generates a mailto: URL with pre-filled email template for student outreach
 * @param {Object} student - The student data object
 * @returns {string|null} The mailto: URL or null if no student email available
 */
export function generateStudentEmailTemplate(student) {
    if (!student) return null;

    // Get student email - try various field names
    const studentEmail = student.studentEmail || student.StudentEmail || student.email || student.Email || '';

    if (!studentEmail) {
        return null;
    }

    // Get personal email for CC - try various field names
    const personalEmail = student.personalEmail || student.PersonalEmail || student.otherEmail || student.OtherEmail || '';

    const data = resolveStudentData(student);
    const firstName = getFirstName(data.name);
    const greeting = getTimeOfDayGreeting();
    const daysOut = data.daysOut || 0;
    const missingAssignments = student.missingAssignments || [];
    const missingCount = missingAssignments.length;

    // Get grade from raw student object
    const rawGrade = student.grade ?? student.Grade ?? student.currentGrade ?? null;
    const grade = rawGrade !== undefined && rawGrade !== null ? parseFloat(rawGrade).toFixed(0) : null;

    // Build subject line
    const subject = `Checking In - ${data.name}`;

    // Build email body
    let body = `${greeting} ${firstName},\n\n`;

    // Main message
    body += `It has been ${daysOut} day${daysOut !== 1 ? 's' : ''} since you last submitted`;

    if (missingCount > 0) {
        body += ` and you currently have ${missingCount} missing assignment${missingCount !== 1 ? 's' : ''}`;
    }

    if (grade !== null) {
        body += `. Your class grade is ${grade}%`;
    }
    body += '.\n';

    // Add missing assignments bullet list with links if any
    if (missingCount > 0) {
        body += '\nMissing Assignments:\n';
        missingAssignments.forEach(assignment => {
            const title = assignment.assignmentTitle || assignment.Assignment || assignment.title || assignment.name || 'Untitled Assignment';
            const link = assignment.assignmentLink || assignment.AssignmentLink || assignment.link || '';
            if (link) {
                body += `• ${title}\n  ${link}\n`;
            } else {
                body += `• ${title}\n`;
            }
        });
    }

    // Closing
    body += '\nWould you be able to submit an assignment today?\n';

    // Build mailto URL with CC if personal email exists
    let mailtoUrl = `mailto:${encodeURIComponent(studentEmail)}?`;
    if (personalEmail) {
        mailtoUrl += `cc=${encodeURIComponent(personalEmail)}&`;
    }
    mailtoUrl += `subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    return mailtoUrl;
}

// Canvas Auth Error Modal state
let canvasAuthErrorResolve = null;
let canvasAuthErrorShown = false;

/**
 * Updates the non-API toggle UI in the Canvas Auth Error modal
 * @param {boolean} isEnabled - Whether non-API course fetch is enabled
 */
function updateCanvasAuthNonApiToggleUI(isEnabled) {
    if (!elements.canvasAuthNonApiToggle) return;

    if (isEnabled) {
        elements.canvasAuthNonApiToggle.className = 'fas fa-toggle-on';
        elements.canvasAuthNonApiToggle.style.color = 'var(--primary-color)';
    } else {
        elements.canvasAuthNonApiToggle.className = 'fas fa-toggle-off';
        elements.canvasAuthNonApiToggle.style.color = 'gray';
    }
}

/**
 * Toggles the non-API course fetch setting in the Canvas Auth Error modal
 */
export function toggleCanvasAuthNonApi() {
    if (!elements.canvasAuthNonApiToggle) return;

    const isCurrentlyOn = elements.canvasAuthNonApiToggle.classList.contains('fa-toggle-on');
    updateCanvasAuthNonApiToggleUI(!isCurrentlyOn);
}

/**
 * Opens the Canvas Auth Error modal and returns a promise that resolves with the user's choice
 * @returns {Promise<'continue'|'shutdown'>} Promise that resolves with 'continue' or 'shutdown'
 */
export async function openCanvasAuthErrorModal() {
    return new Promise(async (resolve) => {
        // Prevent multiple modals from stacking
        if (canvasAuthErrorShown) {
            resolve('continue'); // Default to continue if already shown
            return;
        }

        if (!elements.canvasAuthErrorModal) {
            resolve('continue');
            return;
        }

        canvasAuthErrorShown = true;
        canvasAuthErrorResolve = resolve;

        // Load current non-API setting and update toggle UI
        const settings = await storageGet([STORAGE_KEYS.NON_API_COURSE_FETCH]);
        const nonApiFetch = settings[STORAGE_KEYS.NON_API_COURSE_FETCH] || false;
        updateCanvasAuthNonApiToggleUI(nonApiFetch);

        // Show modal
        elements.canvasAuthErrorModal.style.display = 'flex';
    });
}

/**
 * Closes the Canvas Auth Error modal
 * @param {'continue'|'shutdown'} choice - The user's choice
 */
export async function closeCanvasAuthErrorModal(choice = 'continue') {
    // Save the non-API toggle setting if user chose to continue
    if (choice === 'continue' && elements.canvasAuthNonApiToggle) {
        const isNonApiEnabled = elements.canvasAuthNonApiToggle.classList.contains('fa-toggle-on');
        await storageSet({ [STORAGE_KEYS.NON_API_COURSE_FETCH]: isNonApiEnabled });
        if (isNonApiEnabled) {
            console.log('[Canvas Auth Error] Non-API course fetch enabled by user');
        }
    }

    if (elements.canvasAuthErrorModal) {
        elements.canvasAuthErrorModal.style.display = 'none';
    }

    canvasAuthErrorShown = false;

    // Resolve the promise with the user's choice
    if (canvasAuthErrorResolve) {
        canvasAuthErrorResolve(choice);
        canvasAuthErrorResolve = null;
    }
}

/**
 * Checks if an error response indicates a Canvas authorization error
 * @param {Response} response - The fetch response object
 * @returns {boolean} True if it's an authorization error
 */
export function isCanvasAuthError(response) {
    return response && (response.status === 401 || response.status === 403);
}

/**
 * Checks if a JSON error body indicates a Canvas authorization error
 * @param {Object} errorBody - The parsed JSON error body
 * @returns {boolean} True if it's an authorization error
 */
export function isCanvasAuthErrorBody(errorBody) {
    if (!errorBody) return false;

    // Check for "unauthorized" status
    if (errorBody.status === 'unauthorized') return true;

    // Check for authorization error messages
    if (errorBody.errors && Array.isArray(errorBody.errors)) {
        return errorBody.errors.some(err =>
            err.message && (
                err.message.toLowerCase().includes('unauthorized') ||
                err.message.toLowerCase().includes('not authorized')
            )
        );
    }

    return false;
}

// ========================================
// Latest Updates Modal
// ========================================

/**
 * Checks if the latest updates modal should be shown
 * Returns true if:
 * 1. The current version has release notes defined
 * 2. The user hasn't seen the updates for this version yet
 * @returns {Promise<boolean>} Whether the modal should be shown
 */
export async function shouldShowLatestUpdatesModal() {
    try {
        // Get the current version from the manifest
        const manifest = chrome.runtime.getManifest();
        const currentVersion = manifest.version;

        // Check if there are release notes for this version
        if (!hasReleaseNotes(currentVersion)) {
            return false;
        }

        // Get the last seen version from storage
        const lastSeenVersion = await storageGetValue(STORAGE_KEYS.LAST_SEEN_VERSION, null);

        // Show modal if user hasn't seen this version's updates
        return lastSeenVersion !== currentVersion;
    } catch (error) {
        console.error('Error checking if latest updates modal should show:', error);
        return false;
    }
}

/**
 * Opens the latest updates modal and populates it with release notes
 */
export function openLatestUpdatesModal() {
    if (!elements.latestUpdatesModal) return;

    // Get the latest release notes (version comes from release-notes.js, not manifest)
    const latest = getLatestReleaseNotes();

    if (!latest) {
        console.warn('No release notes found');
        return;
    }

    const { version, notes: releaseNotes } = latest;

    // Update the modal title
    if (elements.latestUpdatesTitle) {
        elements.latestUpdatesTitle.textContent = releaseNotes.title || "What's New";
    }

    // Update the version display (uses version from release notes)
    if (elements.latestUpdatesVersion) {
        elements.latestUpdatesVersion.textContent = `Version ${version}`;
    }

    // Update the date display
    if (elements.latestUpdatesDate) {
        if (releaseNotes.date) {
            elements.latestUpdatesDate.textContent = `Last Updated: ${releaseNotes.date}`;
            elements.latestUpdatesDate.style.display = '';
        } else {
            elements.latestUpdatesDate.style.display = 'none';
        }
    }

    // Populate the updates list
    if (elements.latestUpdatesList) {
        elements.latestUpdatesList.innerHTML = '';

        if (releaseNotes.updates && Array.isArray(releaseNotes.updates)) {
            releaseNotes.updates.forEach(update => {
                const li = document.createElement('li');
                li.textContent = update;
                elements.latestUpdatesList.appendChild(li);
            });
        }
    }

    // Show the modal
    elements.latestUpdatesModal.style.display = 'flex';
}

/**
 * Closes the latest updates modal and marks the version as seen
 */
export async function closeLatestUpdatesModal() {
    if (!elements.latestUpdatesModal) return;

    // Hide the modal
    elements.latestUpdatesModal.style.display = 'none';

    // Mark this version as seen
    const manifest = chrome.runtime.getManifest();
    const currentVersion = manifest.version;

    await storageSet({ [STORAGE_KEYS.LAST_SEEN_VERSION]: currentVersion });
    console.log(`Marked version ${currentVersion} as seen`);
}
