// Five9 Integration - Monitors Five9 connection status
import { STORAGE_KEYS, FIVE9_CONNECTION_STATES } from '../constants/index.js';
import { storageGet } from '../utils/storage.js';
import { elements } from './ui-manager.js';
import { updateCallTabDisplay, showConnectionError, clearConnectionError } from './call-tab-placeholder.js';

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
 * Updates the Call tab display based on Five9 connection state
 * Uses the unified call-tab-placeholder system for displaying messages
 * @param {Array} selectedQueue - Currently selected students
 * @param {boolean} [debugModeOverride] - Optional debug mode override (if not provided, fetches from storage)
 */
export async function updateFive9ConnectionIndicator(selectedQueue, debugModeOverride = null) {
    console.log('[Five9Debug] updateFive9ConnectionIndicator called, debugModeOverride:', debugModeOverride);

    // Get debug mode from storage if not provided
    let isDebugMode = debugModeOverride;
    if (isDebugMode === null) {
        const data = await storageGet(STORAGE_KEYS.CALL_DEMO);
        isDebugMode = data[STORAGE_KEYS.CALL_DEMO] || false;
        console.log('[Five9Debug] Read CALL_DEMO from storage:', data[STORAGE_KEYS.CALL_DEMO], '-> isDebugMode:', isDebugMode);
    }

    // Skip Five9 monitoring entirely in demo mode
    if (isDebugMode) {
        console.log('[Five9Debug] Demo mode active - skipping Five9 monitoring, clearing error, calling updateCallTabDisplay with debugMode:true');
        // Clear any lingering Five9 error state
        clearConnectionError();
        await updateCallTabDisplay({
            selectedQueue: selectedQueue || [],
            debugMode: true
        });
        return;
    }

    console.log('[Five9Debug] Demo mode NOT active - proceeding with Five9 monitoring');

    // Log connection state changes
    const connectionState = await checkFive9Connection();
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

    // Use the unified placeholder system to update the display
    await updateCallTabDisplay({
        selectedQueue: selectedQueue || [],
        debugMode: false
    });
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
    chrome.runtime.onMessage.addListener(async (message, sender) => {
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

                // Check if this is a connection error (receiving end doesn't exist)
                // Skip in demo mode since it doesn't use Five9
                const demoData = await storageGet(STORAGE_KEYS.CALL_DEMO);
                const isDemo = demoData[STORAGE_KEYS.CALL_DEMO] || false;

                if (!isDemo && message.error && (
                    message.error.includes('disconnected') ||
                    message.error.includes('Receiving end does not exist') ||
                    message.error.includes('Could not establish connection')
                )) {
                    // Check if Five9 tab still exists
                    const tabs = await chrome.tabs.query({ url: "https://app-atl.five9.com/*" });
                    const hasFive9Tab = tabs.length > 0;
                    showConnectionError(hasFive9Tab);
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

        // Handle Five9 call state changes (from call state monitor)
        if (message.type === 'FIVE9_CALL_STATE_CHANGED') {
            console.log(`📞 Five9 call state: ${message.previousState} -> ${message.newState}`);
        }

        // Handle disposition set externally (through Five9 UI)
        if (message.type === 'FIVE9_DISPOSITION_SET') {
            console.log('📋 Five9 disposition was set externally');
            if (callManager && callManager.getCallActiveState()) {
                // Call was disposed through Five9 UI - reset our state
                callManager.handleExternalDisposition();
            }
        }

        // Handle call disconnected externally (through Five9 UI)
        if (message.type === 'FIVE9_CALL_DISCONNECTED') {
            console.log('📞 Five9 call was disconnected externally');
            if (callManager && callManager.getCallActiveState() && !callManager.waitingForDisposition) {
                // Call was disconnected through Five9 UI - update to awaiting disposition
                callManager.handleExternalDisconnect();
            }
        }
    });
}
