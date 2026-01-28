// Modal Manager - Handles all modal dialogs (scan filter, queue, version history)
import { STORAGE_KEYS, CANVAS_DOMAIN, FIVE9_CONNECTION_STATES } from '../constants/index.js';
import { storageGet, storageSet, storageGetValue } from '../utils/storage.js';
import { elements } from './ui-manager.js';
import { resolveStudentData } from './student-renderer.js';

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
        STORAGE_KEYS.CALL_DEMO,
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

    // Load Power Automate settings
    const paUrl = result[STORAGE_KEYS.POWER_AUTOMATE_URL] || '';
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

    // Load Five9 settings (Call Demo mode, formerly debugMode)
    const callDemo = result[STORAGE_KEYS.CALL_DEMO] || false;
    updateDebugModeModalUI(callDemo);

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

    // Save Power Automate settings
    if (elements.powerAutomateUrlInput) {
        const paUrl = elements.powerAutomateUrlInput.value.trim();
        settingsToSave[STORAGE_KEYS.POWER_AUTOMATE_URL] = paUrl;
        console.log(`Power Automate URL saved: ${paUrl ? 'URL configured' : 'URL cleared'}`);
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

    // Save Five9 settings (Call Demo mode)
    if (elements.debugModeToggleModal) {
        const callDemoEnabled = elements.debugModeToggleModal.classList.contains('fa-toggle-on');
        settingsToSave[STORAGE_KEYS.CALL_DEMO] = callDemoEnabled;
        console.log(`Call Demo mode setting saved: ${callDemoEnabled}`);
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
 * Updates the Power Automate enabled toggle UI
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

            // Enable start button and restore normal status text
            if (elements.startBtn) {
                elements.startBtn.disabled = false;
                elements.startBtn.style.opacity = '1';
                elements.startBtn.style.cursor = 'pointer';
            }
            if (elements.statusText) {
                // Remove link styling and restore normal text
                elements.statusText.textContent = 'Ready to Scan';
                elements.statusText.style.textDecoration = 'none';
                elements.statusText.style.color = '';
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
 * @returns {Promise<number|null>} Promise that resolves with selected tab ID or null if cancelled
 */
export function openExcelInstanceModal(excelTabs) {
    return new Promise((resolve) => {
        if (!elements.excelInstanceModal || !elements.excelInstanceList) {
            resolve(null);
            return;
        }

        // Store resolve function for later
        excelInstanceResolve = resolve;

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
