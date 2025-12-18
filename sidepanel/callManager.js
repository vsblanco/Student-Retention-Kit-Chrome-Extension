// [2025-12-18] Version 1.0 - Call Manager Module
// Handles all call state management, automation, and Five9 integration

/**
 * CallManager class - Manages call state, timers, and automation sequences
 */
export default class CallManager {
    constructor(elements) {
        this.elements = elements;
        this.isCallActive = false;
        this.callTimerInterval = null;
        this.selectedQueue = [];
        this.debugMode = false;
    }

    /**
     * Updates the reference to the selected queue
     * @param {Array} queue - Array of selected student entries
     */
    updateQueue(queue) {
        this.selectedQueue = queue;
    }

    /**
     * Sets debug mode state
     * @param {boolean} enabled - Whether debug mode is enabled
     */
    setDebugMode(enabled) {
        this.debugMode = enabled;
        this.updateCallInterfaceState();
    }

    /**
     * Gets current call active state
     * @returns {boolean} Whether a call is currently active
     */
    getCallActiveState() {
        return this.isCallActive;
    }

    /**
     * Toggles call state between active and inactive
     * Handles both single calls and automation mode
     * @param {boolean} forceEnd - Force end the call regardless of current state
     */
    toggleCallState(forceEnd = false) {
        // --- CHECK DEBUG MODE ---
        if (!this.debugMode && !this.isCallActive) {
            alert('Call functionality is disabled.\n\nTo enable the call demo, turn on Debug Mode in Settings.');
            return;
        }
        // ----------------------

        // --- CHECK FOR AUTOMATION MODE ---
        if (this.selectedQueue.length > 1 && !this.isCallActive) {
            this.startAutomationSequence();
            return;
        }
        // --------------------------------------

        if (forceEnd && !this.isCallActive) return;
        this.isCallActive = !this.isCallActive;
        if (forceEnd) this.isCallActive = false;

        if (this.isCallActive) {
            this.elements.dialBtn.style.background = '#ef4444';
            this.elements.dialBtn.style.transform = 'rotate(135deg)';
            this.elements.callStatusText.innerHTML = '<span class="status-indicator" style="background:#ef4444; animation: blink 1s infinite;"></span> Connected';

            // Show Disposition Grid
            if (this.elements.callDispositionSection) {
                this.elements.callDispositionSection.style.display = 'flex';
            }

            this.startCallTimer();
        } else {
            this.elements.dialBtn.style.background = '#10b981';
            this.elements.dialBtn.style.transform = 'rotate(0deg)';
            this.elements.callStatusText.innerHTML = '<span class="status-indicator ready"></span> Ready to Connect';

            // Hide Disposition Grid
            if (this.elements.callDispositionSection) {
                this.elements.callDispositionSection.style.display = 'none';
            }

            // Hide custom input area if it was open
            if (this.elements.otherInputArea) {
                this.elements.otherInputArea.style.display = 'none';
            }

            this.stopCallTimer();
        }
    }

    /**
     * Starts the automation sequence for multiple students
     * Currently a placeholder for future implementation
     */
    startAutomationSequence() {
        alert(`Starting automation for ${this.selectedQueue.length} students...\n(Logic to be implemented)`);
        // TODO: Implement automation logic
        // Future implementation will:
        // 1. Loop through selectedQueue
        // 2. For each student, initiate call via Five9
        // 3. Wait for disposition
        // 4. Auto-advance to next student
        // 5. Update UI with progress
    }

    /**
     * Starts the call timer and updates display every second
     */
    startCallTimer() {
        let seconds = 0;
        this.elements.callTimer.textContent = "00:00";
        clearInterval(this.callTimerInterval);

        this.callTimerInterval = setInterval(() => {
            seconds++;
            const m = Math.floor(seconds / 60).toString().padStart(2, '0');
            const s = (seconds % 60).toString().padStart(2, '0');
            this.elements.callTimer.textContent = `${m}:${s}`;
        }, 1000);
    }

    /**
     * Stops the call timer and resets display
     */
    stopCallTimer() {
        clearInterval(this.callTimerInterval);
        this.callTimerInterval = null;
        this.elements.callTimer.textContent = "00:00";
    }

    /**
     * Handles call disposition selection and ends the call
     * @param {string} type - The disposition type selected
     */
    handleDisposition(type) {
        console.log("Logged Disposition:", type);

        // TODO: Store disposition data
        // Future implementation:
        // - Save disposition to chrome.storage
        // - Associate with current student
        // - Track disposition history

        this.toggleCallState(true); // Force end call
    }

    /**
     * Initiates a call through Five9
     * @param {string} phoneNumber - Phone number to dial
     * @returns {Promise<{success: boolean, error?: string}>}
     */
    async initiateCall(phoneNumber) {
        if (!phoneNumber || phoneNumber === "No Phone Listed") {
            return { success: false, error: "No valid phone number" };
        }

        try {
            // Send message to Five9 connector content script
            const response = await chrome.runtime.sendMessage({
                type: 'executeFive9Call',
                phoneNumber: phoneNumber
            });

            return response;
        } catch (error) {
            console.error("Error initiating call:", error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Hangs up the current Five9 call
     * @returns {Promise<{success: boolean, error?: string}>}
     */
    async hangupCall() {
        try {
            const response = await chrome.runtime.sendMessage({
                type: 'executeFive9Hangup'
            });

            return response;
        } catch (error) {
            console.error("Error hanging up call:", error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Updates the call interface visual state based on debug mode
     */
    updateCallInterfaceState() {
        if (!this.elements.dialBtn || !this.elements.callStatusText) return;

        if (this.debugMode) {
            // Enable call interface
            this.elements.dialBtn.style.opacity = '1';
            this.elements.dialBtn.style.cursor = 'pointer';
            this.elements.dialBtn.title = '';
            if (!this.isCallActive) {
                this.elements.callStatusText.innerHTML = '<span class="status-indicator ready"></span> Ready to Connect';
            }
        } else {
            // Disable call interface
            this.elements.dialBtn.style.opacity = '0.5';
            this.elements.dialBtn.style.cursor = 'not-allowed';
            this.elements.dialBtn.title = 'Enable Debug Mode in Settings to use call functionality';
            this.elements.callStatusText.innerHTML = '<span class="status-indicator" style="background:#f59e0b;"></span> Calls Disabled (Enable Debug Mode)';
        }
    }

    /**
     * Cleanup method - call when disposing the manager
     */
    cleanup() {
        this.stopCallTimer();
        this.selectedQueue = [];
        this.isCallActive = false;
    }
}
