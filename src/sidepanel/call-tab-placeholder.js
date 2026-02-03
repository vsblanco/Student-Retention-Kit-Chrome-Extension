// Call Tab Placeholder - Unified placeholder display with priority-based messages
import { FIVE9_CONNECTION_STATES } from '../constants/index.js';
import { elements } from './ui-manager.js';

/**
 * Message types for the Call tab placeholder
 * Higher priority numbers take precedence
 */
export const PLACEHOLDER_MESSAGES = {
    NO_STUDENT_SELECTED: {
        id: 'no_student',
        priority: 1,
        icon: 'fa-user-graduate',
        iconStyle: 'font-size:3em; margin-bottom:15px; opacity:0.5;',
        header: 'No Student Selected',
        message: 'Select a student from the Master List<br>to view details and make calls.'
    },
    FIVE9_NO_TAB: {
        id: 'five9_no_tab',
        priority: 2,
        icon: 'fa-phone-slash',
        iconStyle: 'font-size:3em; margin-bottom:15px; opacity:0.5;',
        header: 'Awaiting Five9 Tab',
        message: 'Please open Five9 in a new tab<br>to enable calling features.'
    },
    FIVE9_AWAITING_AGENT: {
        id: 'five9_awaiting',
        priority: 2,
        icon: 'fa-spinner fa-spin',
        iconStyle: 'font-size:3em; margin-bottom:15px; opacity:0.4;',
        header: 'Awaiting Agent Connection',
        message: 'Five9 tab detected. Waiting for<br>agent to connect...'
    },
    FIVE9_CONNECTION_ERROR: {
        id: 'five9_error',
        priority: 3,
        icon: 'fa-exclamation-triangle',
        iconStyle: 'font-size:3em; margin-bottom:15px; color:#ef4444;',
        header: 'Five9 Connection Error',
        message: 'Could not establish connection.<br>The Five9 tab may need to be refreshed.'
    }
};

// Track current placeholder state to avoid unnecessary re-renders
let currentPlaceholderMessage = null;

// Track if user has bypassed the Five9 awaiting check
let five9BypassActive = false;

// Track if there's a Five9 connection error
let five9ConnectionError = false;

/**
 * Checks Five9 connection status (duplicated here to avoid circular dependency)
 * @returns {Promise<string>} One of FIVE9_CONNECTION_STATES
 */
async function checkFive9ConnectionState() {
    try {
        const tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });

        if (tabs.length === 0) {
            return FIVE9_CONNECTION_STATES.NO_TAB;
        }

        const response = await chrome.runtime.sendMessage({
            type: 'GET_FIVE9_CONNECTION_STATE'
        });

        if (response && response.state === FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION) {
            return FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION;
        }

        return FIVE9_CONNECTION_STATES.AWAITING_CONNECTION;
    } catch (error) {
        console.error("Error checking Five9 connection:", error);
        return FIVE9_CONNECTION_STATES.NO_TAB;
    }
}

/**
 * Refreshes the Five9 tab if one exists
 * @returns {Promise<boolean>} True if tab was refreshed, false if no tab found
 */
async function refreshFive9Tab() {
    try {
        const tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });
        if (tabs.length > 0) {
            await chrome.tabs.reload(tabs[0].id);
            // Clear error state after refresh
            clearConnectionError();
            return true;
        }
        return false;
    } catch (error) {
        console.error("Error refreshing Five9 tab:", error);
        return false;
    }
}

/**
 * Updates the unified placeholder content
 * @param {Object} messageConfig - The message configuration from PLACEHOLDER_MESSAGES
 */
function renderPlaceholder(messageConfig) {
    if (!elements.callTabPlaceholder) return;

    // Skip if already showing this message
    if (currentPlaceholderMessage === messageConfig.id) return;
    currentPlaceholderMessage = messageConfig.id;

    // Add "Continue Anyways" button for awaiting agent message
    let actionButton = '';
    if (messageConfig.id === 'five9_awaiting') {
        actionButton = `<button id="continueAnywaysBtn" style="margin-top:15px; background:none; border:none; color:#3b82f6; cursor:pointer; font-size:0.85em; text-decoration:underline;">Continue Anyways</button>`;
    } else if (messageConfig.id === 'five9_error') {
        actionButton = `<button id="refreshFive9Btn" style="margin-top:15px; padding:8px 16px; background:var(--primary-color); color:white; border:none; border-radius:6px; cursor:pointer; font-size:0.9em; display:flex; align-items:center; gap:6px;"><i class="fas fa-sync-alt"></i> Refresh Five9 Tab</button>`;
    }

    elements.callTabPlaceholder.innerHTML = `
        <i class="fas ${messageConfig.icon}" style="${messageConfig.iconStyle}"></i>
        <span style="font-size:1.1em; font-weight:500;">${messageConfig.header}</span>
        <span style="font-size:0.9em; margin-top:5px; color:#6b7280;">${messageConfig.message}</span>
        ${actionButton}
    `;

    // Add click handler for continue button
    if (messageConfig.id === 'five9_awaiting') {
        const btn = document.getElementById('continueAnywaysBtn');
        if (btn) {
            btn.addEventListener('click', () => {
                bypassFive9Check();
            });
        }
    }

    // Add click handler for refresh button
    if (messageConfig.id === 'five9_error') {
        const btn = document.getElementById('refreshFive9Btn');
        if (btn) {
            btn.addEventListener('click', async () => {
                btn.disabled = true;
                btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Refreshing...';
                const refreshed = await refreshFive9Tab();
                if (!refreshed) {
                    // No Five9 tab found, show the no tab message
                    showConnectionError(false);
                }
            });
        }
    }
}

/**
 * Shows the placeholder with the given message
 */
function showPlaceholder() {
    if (!elements.callTabPlaceholder) return;
    elements.callTabPlaceholder.style.display = 'flex';
}

/**
 * Hides the placeholder
 */
function hidePlaceholder() {
    if (!elements.callTabPlaceholder) return;
    elements.callTabPlaceholder.style.display = 'none';
    currentPlaceholderMessage = null;
}

/**
 * Bypasses the Five9 awaiting connection check
 * Allows user to continue to call section even without agent connection
 */
function bypassFive9Check() {
    five9BypassActive = true;
    hidePlaceholder();
    showCallSection();
}

/**
 * Resets the Five9 bypass state
 * Called when student selection changes or Five9 connection state changes
 */
export function resetFive9Bypass() {
    five9BypassActive = false;
}

/**
 * Shows the Five9 connection error placeholder
 * @param {boolean} hasFive9Tab - Whether a Five9 tab exists
 */
export async function showConnectionError(hasFive9Tab = true) {
    five9ConnectionError = true;
    currentPlaceholderMessage = null; // Force re-render

    if (hasFive9Tab) {
        // Show error with refresh button
        renderPlaceholder(PLACEHOLDER_MESSAGES.FIVE9_CONNECTION_ERROR);
    } else {
        // No Five9 tab - show the no tab message
        renderPlaceholder(PLACEHOLDER_MESSAGES.FIVE9_NO_TAB);
    }
    showPlaceholder();
    hideCallSection();
}

/**
 * Clears the Five9 connection error state
 */
export function clearConnectionError() {
    five9ConnectionError = false;
    currentPlaceholderMessage = null; // Allow next render
}

/**
 * Shows the call section (student card, call interface, etc.)
 */
function showCallSection() {
    const contactTab = document.getElementById('contact');
    if (!contactTab) return;

    Array.from(contactTab.children).forEach(child => {
        if (child.id === 'callTabPlaceholder') {
            child.style.display = 'none';
        } else if (child.classList.contains('section')) {
            child.style.display = '';
        }
    });
}

/**
 * Hides the call section
 */
function hideCallSection() {
    const contactTab = document.getElementById('contact');
    if (!contactTab) return;

    Array.from(contactTab.children).forEach(child => {
        if (child.id === 'callTabPlaceholder') {
            // Placeholder visibility managed separately
        } else if (child.classList.contains('section')) {
            child.style.display = 'none';
        }
    });
}

/**
 * Determines which placeholder message should be shown based on current state
 * Priority order (highest to lowest):
 * 1. Five9 Connection Error - highest priority when error occurred
 * 2. Five9 NO_TAB - always shows first regardless of student selection
 * 3. No student selected
 * 4. Five9 AWAITING_CONNECTION - only when student is selected
 *
 * @param {Object} state - Current state object
 * @param {Array} state.selectedQueue - Currently selected students
 * @param {boolean} state.debugMode - Whether debug/demo mode is enabled
 * @returns {Object} Result with showPlaceholder boolean and message config
 */
export async function determineCallTabState(state = {}) {
    const { selectedQueue = [], debugMode = false } = state;
    const hasStudentSelected = selectedQueue && selectedQueue.length > 0;

    // Check for connection error first - highest priority
    if (five9ConnectionError && !debugMode) {
        // Re-check if Five9 tab still exists
        const tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });
        const hasFive9Tab = tabs.length > 0;

        return {
            showPlaceholder: true,
            message: hasFive9Tab ? PLACEHOLDER_MESSAGES.FIVE9_CONNECTION_ERROR : PLACEHOLDER_MESSAGES.FIVE9_NO_TAB
        };
    }

    // Check Five9 NO_TAB first - highest priority, shows regardless of student selection
    if (!debugMode) {
        const connectionState = await checkFive9ConnectionState();

        if (connectionState === FIVE9_CONNECTION_STATES.NO_TAB) {
            // Reset bypass when Five9 tab is closed
            five9BypassActive = false;
            return {
                showPlaceholder: true,
                message: PLACEHOLDER_MESSAGES.FIVE9_NO_TAB
            };
        }

        // If no student is selected, show that message (second priority)
        if (!hasStudentSelected) {
            return {
                showPlaceholder: true,
                message: PLACEHOLDER_MESSAGES.NO_STUDENT_SELECTED
            };
        }

        // Student is selected, check if awaiting connection (unless bypassed)
        if (connectionState === FIVE9_CONNECTION_STATES.AWAITING_CONNECTION && !five9BypassActive) {
            return {
                showPlaceholder: true,
                message: PLACEHOLDER_MESSAGES.FIVE9_AWAITING_AGENT
            };
        }
    } else {
        // Debug mode - only check student selection
        if (!hasStudentSelected) {
            return {
                showPlaceholder: true,
                message: PLACEHOLDER_MESSAGES.NO_STUDENT_SELECTED
            };
        }
    }

    // All checks passed - show the call section
    return {
        showPlaceholder: false,
        message: null
    };
}

/**
 * Updates the Call tab display based on current state
 * This is the main function to call when state changes
 *
 * @param {Object} state - Current state object
 * @param {Array} state.selectedQueue - Currently selected students
 * @param {boolean} state.debugMode - Whether debug/demo mode is enabled
 */
export async function updateCallTabDisplay(state = {}) {
    const result = await determineCallTabState(state);

    if (result.showPlaceholder) {
        renderPlaceholder(result.message);
        showPlaceholder();
        hideCallSection();
    } else {
        hidePlaceholder();
        showCallSection();
    }
}

/**
 * Forces a refresh of the Call tab display
 * Useful when Five9 connection state changes
 *
 * @param {Function} getSelectedQueue - Function that returns current selected queue
 * @param {Function} getDebugMode - Function that returns current debug mode state
 */
export async function refreshCallTabDisplay(getSelectedQueue, getDebugMode) {
    const selectedQueue = getSelectedQueue ? getSelectedQueue() : [];
    const debugMode = getDebugMode ? await getDebugMode() : false;

    await updateCallTabDisplay({ selectedQueue, debugMode });
}
