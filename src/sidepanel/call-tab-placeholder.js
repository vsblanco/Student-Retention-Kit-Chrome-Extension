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
    }
};

// Track current placeholder state to avoid unnecessary re-renders
let currentPlaceholderMessage = null;

// Track if user has bypassed the Five9 awaiting check
let five9BypassActive = false;

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
 * Updates the unified placeholder content
 * @param {Object} messageConfig - The message configuration from PLACEHOLDER_MESSAGES
 */
function renderPlaceholder(messageConfig) {
    if (!elements.callTabPlaceholder) return;

    // Skip if already showing this message
    if (currentPlaceholderMessage === messageConfig.id) return;
    currentPlaceholderMessage = messageConfig.id;

    // Add "Continue Anyways" button for awaiting agent message
    const continueButton = messageConfig.id === 'five9_awaiting'
        ? `<button id="continueAnywaysBtn" style="margin-top:15px; background:none; border:none; color:#3b82f6; cursor:pointer; font-size:0.85em; text-decoration:underline;">Continue Anyways</button>`
        : '';

    elements.callTabPlaceholder.innerHTML = `
        <i class="fas ${messageConfig.icon}" style="${messageConfig.iconStyle}"></i>
        <span style="font-size:1.1em; font-weight:500;">${messageConfig.header}</span>
        <span style="font-size:0.9em; margin-top:5px; color:#6b7280;">${messageConfig.message}</span>
        ${continueButton}
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
 * 1. Five9 NO_TAB - always shows first regardless of student selection
 * 2. No student selected
 * 3. Five9 AWAITING_CONNECTION - only when student is selected
 *
 * @param {Object} state - Current state object
 * @param {Array} state.selectedQueue - Currently selected students
 * @param {boolean} state.debugMode - Whether debug/demo mode is enabled
 * @returns {Object} Result with showPlaceholder boolean and message config
 */
export async function determineCallTabState(state = {}) {
    const { selectedQueue = [], debugMode = false } = state;
    const hasStudentSelected = selectedQueue && selectedQueue.length > 0;

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
