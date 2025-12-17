// [2025-12-17] Version 1.0 - Five9 Connector
// This script runs on https://app-atl.five9.com/* to handle call automation.

console.log("SRK: Five9 Connector Loaded");

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.type === 'executeFive9Call') {
        handleFive9Call(request.phoneNumber, sendResponse);
        return true; // Keep channel open for async response
    }
    if (request.type === 'executeFive9Hangup') {
        handleFive9Hangup(sendResponse);
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

async function handleFive9Hangup(sendResponse) {
    try {
        console.log("SRK: Attempting TWO-STEP hangup...");

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

        console.log(`SRK: STEP 2 - Disposing interaction...`);

        // STEP 2: DISPOSE
        const disposeUrl = `https://app-atl.five9.com/appsvcs/rs/svc/agents/${metadata.userId}/interactions/calls/${activeCall.interactionId}/dispose`;
        const payload = { "dispositionId": "300000000000046" }; // Ensure this Disposition ID is correct

        const disposeResp = await fetch(disposeUrl, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (disposeResp.ok) {
            console.log("SRK: Hangup Complete.");
            sendResponse({ success: true });
        } else {
            const errorText = await disposeResp.text();
            console.error("Dispose Error:", disposeResp.status, errorText);
            
            if (disposeResp.status === 404 || disposeResp.status === 435) {
                sendResponse({ success: true });
            } else {
                sendResponse({ success: false, error: `${disposeResp.status} - ${errorText}` });
            }
        }

    } catch (error) {
        console.error("SRK Hangup Error:", error);
        sendResponse({ success: false, error: error.message });
    }
}