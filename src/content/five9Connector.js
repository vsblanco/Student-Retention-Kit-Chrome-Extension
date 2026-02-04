// [2025-12-17] Version 1.0 - Five9 Connector
// This script runs on https://app-atl.five9.com/* to handle call automation.

// Disposition code constants
const DISPOSITION_CODES = {
    "Left Voicemail": "300000000000046",
    "Service Completed": "300000000000043",
    "Outbound Error": "300000000000271",
    "Follow Up": "300000000000048",
    "No Answer": "",        // TODO: Add Five9 disposition code
    "Disconnected": ""      // TODO: Add Five9 disposition code
};

/**
 * Gets the disposition code for a given disposition type
 * @param {string} dispositionType - The disposition type (e.g., "Left Voicemail")
 * @returns {string|null} The Five9 disposition code, or null if not found/empty
 */
function getDispositionCode(dispositionType) {
    const code = DISPOSITION_CODES[dispositionType];
    return (code && code !== "") ? code : null;
}

console.log("SRK: Five9 Connector Loaded");

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.type === 'executeFive9Call') {
        handleFive9Call(request.phoneNumber, sendResponse);
        return true; // Keep channel open for async response
    }
    if (request.type === 'executeFive9Hangup') {
        handleFive9Hangup(request.dispositionType, sendResponse);
        return true; // Keep channel open
    }
    if (request.type === 'executeFive9DisposeOnly') {
        handleFive9DisposeOnly(request.dispositionType, sendResponse);
        return true; // Keep channel open
    }
    if (request.type === 'executeFive9RestartStation') {
        handleFive9RestartStation(sendResponse);
        return true; // Keep channel open
    }
});

async function handleFive9Call(phoneNumber, sendResponse) {
    try {
        console.log(`SRK: Dialing ${phoneNumber}...`);
        const metadataResp = await fetch("https://app-atl.five9.com/appsvcs/rs/svc/auth/metadata");
        if (!metadataResp.ok) throw new Error("Could not fetch User Metadata");
        const metadata = await metadataResp.json();
        
        const url = `https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions/make_external_call`;
        const payload = {
            "number": phoneNumber,
            "skipDNCCheck": false,
            "checkMultipleContacts": true,
            "campaignId": "300000000000483" // Ensure this Campaign ID is correct for your org
        };

        const callResp = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (callResp.ok) sendResponse({ success: true });
        else sendResponse({ success: false, error: `${callResp.status} - ${await callResp.text()}` });

    } catch (error) {
        console.error("SRK Call Error:", error);
        sendResponse({ success: false, error: error.message });
    }
}

async function handleFive9Hangup(dispositionType, sendResponse) {
    try {
        console.log("SRK: Attempting TWO-STEP hangup...");
        console.log("SRK: Disposition type:", dispositionType);

        // Get the disposition code from constants
        const dispositionCode = getDispositionCode(dispositionType);
        console.log("SRK: Disposition code:", dispositionCode);

        const metadataResp = await fetch("https://app-atl.five9.com/appsvcs/rs/svc/auth/metadata");
        if (!metadataResp.ok) throw new Error("Could not fetch User Metadata");
        const metadata = await metadataResp.json();

        // Fetch active interactions
        const interactionsResp = await fetch(`https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions`);
        if (!interactionsResp.ok) throw new Error("Could not fetch active interactions");
        const interactions = await interactionsResp.json();

        const activeCall = interactions.find(i => i.channelType === 'CALL');

        // *** ROBUSTNESS FIX: Handle Manual Hangup ***
        if (!activeCall) {
            console.warn("SRK: No active CALL found (assuming already ended).");
            // We return SUCCESS so the automation continues to the next number
            sendResponse({ success: true, warning: "Call was already ended manually." });
            return;
        }

        console.log(`SRK: STEP 1 - Disconnecting interaction ${activeCall.interactionId}...`);

        // STEP 1: DISCONNECT
        const disconnectUrl = `https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions/calls/${activeCall.interactionId}/disconnect`;
        const disconnectResp = await fetch(disconnectUrl, {
            method: "PUT",
            headers: { "Content-Type": "application/json" }
        });

        if (!disconnectResp.ok) {
             console.warn("Disconnect step warning:", disconnectResp.status);
        }

        await new Promise(r => setTimeout(r, 500));

        // STEP 2: DISPOSE (only if we have a valid disposition code)
        if (dispositionCode) {
            console.log(`SRK: STEP 2 - Disposing interaction with code ${dispositionCode}...`);

            const disposeUrl = `https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions/calls/${activeCall.interactionId}/dispose`;
            const payload = { "dispositionId": dispositionCode };

            const disposeResp = await fetch(disposeUrl, {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            if (disposeResp.ok) {
                console.log("SRK: Hangup Complete.");

                // Get the interaction state after dispose
                const interactionData = await disposeResp.json();
                const state = interactionData?.state || 'UNKNOWN';
                console.log("SRK: Interaction state after dispose:", state);

                sendResponse({ success: true, state: state });
            } else {
                const errorText = await disposeResp.text();
                console.error("Dispose Error:", disposeResp.status, errorText);

                if (disposeResp.status === 404 || disposeResp.status === 435) {
                    sendResponse({ success: true, state: 'UNKNOWN' });
                } else {
                    sendResponse({ success: false, error: `${disposeResp.status} - ${errorText}` });
                }
            }
        } else {
            console.warn("SRK: No disposition code available - skipping dispose step.");
            console.warn("SRK: Call disconnected but not disposed. Add disposition code to constants/dispositions.js");

            // Fetch interaction to get current state (should be WRAP_UP)
            try {
                const interactionsResp = await fetch(`https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions`);
                if (interactionsResp.ok) {
                    const interactions = await interactionsResp.json();
                    const recentCall = interactions.find(i => i.channelType === 'CALL');
                    const state = recentCall?.state || 'WRAP_UP';
                    console.log("SRK: Interaction state after disconnect (no dispose):", state);
                    sendResponse({ success: true, warning: "No disposition code - call disconnected only", state: state });
                } else {
                    sendResponse({ success: true, warning: "No disposition code - call disconnected only", state: 'WRAP_UP' });
                }
            } catch (e) {
                sendResponse({ success: true, warning: "No disposition code - call disconnected only", state: 'WRAP_UP' });
            }
        }

    } catch (error) {
        console.error("SRK Hangup Error:", error);
        sendResponse({ success: false, error: error.message });
    }
}

/**
 * Handles dispose-only operation (when call already disconnected)
 * This is used when user ends call first, then selects disposition
 */
async function handleFive9DisposeOnly(dispositionType, sendResponse) {
    try {
        console.log("SRK: Attempting DISPOSE-ONLY operation...");
        console.log("SRK: Disposition type:", dispositionType);

        // Get the disposition code from constants
        const dispositionCode = getDispositionCode(dispositionType);
        console.log("SRK: Disposition code:", dispositionCode);

        if (!dispositionCode) {
            console.warn("SRK: No disposition code available - cannot dispose.");
            sendResponse({ success: false, error: "No disposition code for: " + dispositionType });
            return;
        }

        const metadataResp = await fetch("https://app-atl.five9.com/appsvcs/rs/svc/auth/metadata");
        if (!metadataResp.ok) throw new Error("Could not fetch User Metadata");
        const metadata = await metadataResp.json();

        // Fetch recently ended interactions (they may still be disposable)
        const interactionsResp = await fetch(`https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions`);
        if (!interactionsResp.ok) throw new Error("Could not fetch interactions");
        const interactions = await interactionsResp.json();

        // Look for the most recent call interaction (may be in WRAP_UP state)
        const recentCall = interactions.find(i => i.channelType === 'CALL');

        if (!recentCall) {
            console.warn("SRK: No recent CALL found to dispose.");
            sendResponse({ success: false, error: "No recent call found to dispose" });
            return;
        }

        console.log(`SRK: Disposing interaction ${recentCall.interactionId} with code ${dispositionCode}...`);

        const disposeUrl = `https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions/calls/${recentCall.interactionId}/dispose`;
        const payload = { "dispositionId": dispositionCode };

        const disposeResp = await fetch(disposeUrl, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (disposeResp.ok) {
            console.log("SRK: Dispose-only Complete.");

            // Get the interaction state after dispose
            const interactionData = await disposeResp.json();
            const state = interactionData?.state || 'UNKNOWN';
            console.log("SRK: Interaction state after dispose-only:", state);

            sendResponse({ success: true, state: state });
        } else {
            const errorText = await disposeResp.text();
            console.error("Dispose Error:", disposeResp.status, errorText);

            if (disposeResp.status === 404 || disposeResp.status === 435) {
                console.warn("SRK: Interaction may have already been disposed or timed out.");
                sendResponse({ success: true, warning: "Interaction already disposed or not found", state: 'UNKNOWN' });
            } else {
                sendResponse({ success: false, error: `${disposeResp.status} - ${errorText}` });
            }
        }

    } catch (error) {
        console.error("SRK Dispose-Only Error:", error);
        sendResponse({ success: false, error: error.message });
    }
}

/**
 * Handles restarting the Five9 station to re-establish connection
 */
async function handleFive9RestartStation(sendResponse) {
    try {
        console.log("SRK: Restarting Five9 station...");

        const metadataResp = await fetch("https://app-atl.five9.com/appsvcs/rs/svc/auth/metadata");
        if (!metadataResp.ok) throw new Error("Could not fetch User Metadata");
        const metadata = await metadataResp.json();

        const restartUrl = `https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/station/restart`;

        const restartResp = await fetch(restartUrl, {
            method: "PUT",
            headers: { "Content-Type": "application/json" }
        });

        if (restartResp.ok || restartResp.status === 204) {
            console.log("SRK: Station restart successful");
            sendResponse({ success: true });
        } else {
            const errorText = await restartResp.text();
            console.error("SRK Station Restart Error:", restartResp.status, errorText);
            sendResponse({ success: false, error: `${restartResp.status} - ${errorText}` });
        }

    } catch (error) {
        console.error("SRK Station Restart Error:", error);
        sendResponse({ success: false, error: error.message });
    }
}