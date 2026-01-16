// Five9 Integration - Monitors Five9 connection status
import { STORAGE_KEYS, FIVE9_CONNECTION_STATES } from '../constants/index.js';
import { elements } from './ui-manager.js';

let five9ConnectionCheckInterval = null;
let lastFive9ConnectionState = FIVE9_CONNECTION_STATES.NO_TAB;

/**
 * Checks Five9 connection status with three states:
 * - NO_TAB: No Five9 tab detected
 * - AWAITING_CONNECTION: Tab exists but agent not connected
 * - ACTIVE_CONNECTION: Agent connected (network activity detected)
 * @returns {Promise<string>} One of FIVE9_CONNECTION_STATES
 */
export async function checkFive9Connection() {
    try {
        // First check if Five9 tab is open
        const tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });

        if (tabs.length === 0) {
            return FIVE9_CONNECTION_STATES.NO_TAB;
        }

        // Tab exists - now check if agent is connected via network monitoring
        const response = await chrome.runtime.sendMessage({
            type: 'GET_FIVE9_CONNECTION_STATE'
        });

        // If we have active connection from network monitoring, return that
        if (response && response.state === FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION) {
            return FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION;
        }

        // Tab exists but no agent connection detected yet
        return FIVE9_CONNECTION_STATES.AWAITING_CONNECTION;
    } catch (error) {
        console.error("Error checking Five9 connection:", error);
        return FIVE9_CONNECTION_STATES.NO_TAB;
    }
}

/**
 * Updates the Five9 connection indicator based on connection state
 * Shows different messages for NO_TAB, AWAITING_CONNECTION, and ACTIVE_CONNECTION states
 * Only shows when debug mode is OFF and student is selected
 */
export async function updateFive9ConnectionIndicator(selectedQueue) {
    if (!elements.five9ConnectionIndicator) return;

    const isDebugMode = await chrome.storage.local.get(STORAGE_KEYS.DEBUG_MODE)
        .then(data => data[STORAGE_KEYS.DEBUG_MODE] || false);

    const connectionState = await checkFive9Connection();
    const hasStudentSelected = selectedQueue && selectedQueue.length > 0;

    // Show indicator when debug mode is OFF, student is selected, and connection is not active
    const shouldShowIndicator = !isDebugMode &&
                                 hasStudentSelected &&
                                 connectionState !== FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION;

    // Update indicator content based on state
    if (elements.five9ConnectionIndicator) {
        const currentDisplay = elements.five9ConnectionIndicator.style.display;
        const targetDisplay = shouldShowIndicator ? 'flex' : 'none';

        // Update message based on connection state
        if (shouldShowIndicator) {
            const indicatorHTML = connectionState === FIVE9_CONNECTION_STATES.NO_TAB
                ? '<i class="fas fa-spinner fa-spin"></i> Awaiting Five9 tab'
                : '<i class="fas fa-spinner fa-spin"></i> Awaiting agent connection';

            elements.five9ConnectionIndicator.innerHTML = indicatorHTML;
        }

        // Only change display if state is actually changing to prevent re-triggering fade-in animation
        if (currentDisplay !== targetDisplay) {
            elements.five9ConnectionIndicator.style.display = targetDisplay;
        }
    }

    // Log connection state changes
    if (connectionState !== lastFive9ConnectionState) {
        lastFive9ConnectionState = connectionState;

        switch (connectionState) {
            case FIVE9_CONNECTION_STATES.NO_TAB:
                console.log("❌ Five9 disconnected - No tab");
                break;
            case FIVE9_CONNECTION_STATES.AWAITING_CONNECTION:
                console.log("⏳ Five9 tab detected - Awaiting agent connection");
                break;
            case FIVE9_CONNECTION_STATES.ACTIVE_CONNECTION:
                console.log("✅ Five9 Active Connection - Agent connected");
                break;
        }
    }
}

/**
 * Starts monitoring Five9 connection status
 */
export function startFive9ConnectionMonitor(getSelectedQueue) {
    // Initial check
    if (getSelectedQueue) {
        updateFive9ConnectionIndicator(getSelectedQueue());
    }

    // Check every 3 seconds
    five9ConnectionCheckInterval = setInterval(() => {
        if (getSelectedQueue) {
            updateFive9ConnectionIndicator(getSelectedQueue());
        }
    }, 3000);
}

/**
 * Stops monitoring Five9 connection status
 */
export function stopFive9ConnectionMonitor() {
    if (five9ConnectionCheckInterval) {
        clearInterval(five9ConnectionCheckInterval);
        five9ConnectionCheckInterval = null;
    }
}

/**
 * Setup Five9 status listeners from background.js
 */
export function setupFive9StatusListeners(callManager, getSelectedQueue) {
    chrome.runtime.onMessage.addListener((message, sender) => {
        // Handle Five9 call initiation status
        if (message.type === 'callStatus') {
            if (message.success) {
                console.log("✓ Five9 call initiated successfully");
            } else {
                console.error("✗ Five9 call failed:", message.error);
                // Revert call UI state if call failed
                if (callManager && callManager.getCallActiveState()) {
                    callManager.toggleCallState(true);
                }
            }
        }

        // Handle Five9 hangup status
        if (message.type === 'hangupStatus') {
            if (message.success) {
                console.log("✓ Five9 call ended successfully");
            } else {
                console.error("✗ Five9 hangup failed:", message.error);
            }
        }

        // Handle Five9 connection state changes from network monitoring
        if (message.type === 'FIVE9_CONNECTION_STATE_CHANGED') {
            console.log('Five9 connection state changed:', message.state);
            // Immediately update the indicator when state changes
            if (getSelectedQueue) {
                updateFive9ConnectionIndicator(getSelectedQueue());
            }
        }
    });
}
