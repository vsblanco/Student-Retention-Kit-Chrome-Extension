// manifestSideloader.js
// Utility to auto-sideload the Office Add-in manifest into Excel Online

/**
 * Reads the manifest XML from the extension's assets folder
 * @returns {Promise<string>} The manifest XML content
 */
export async function getManifestXML() {
    try {
        const manifestUrl = chrome.runtime.getURL('assets/Excel Add-In Manifest.xml');
        const response = await fetch(manifestUrl);
        const manifestXML = await response.text();
        return manifestXML;
    } catch (error) {
        console.error('[SRK Sideloader] Failed to read manifest:', error);
        throw error;
    }
}

/**
 * Extracts the Add-in ID from the manifest XML
 * @param {string} manifestXML - The manifest XML content
 * @returns {string|null} The Add-in ID or null if not found
 */
export function extractAddinId(manifestXML) {
    const idMatch = manifestXML.match(/<Id>([^<]+)<\/Id>/);
    return idMatch ? idMatch[1] : null;
}

/**
 * Finds existing Office session IDs in localStorage or generates a new one
 * @returns {string} The session ID to use
 */
function findOrGenerateSessionId() {
    // Look for existing __OSF_UPLOADFILE keys to find the session ID
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith('__OSF_UPLOADFILE.MyAddins.16.')) {
            // Extract session ID from key like "__OSF_UPLOADFILE.MyAddins.16.3735224676"
            const parts = key.split('.');
            if (parts.length >= 4) {
                return parts[3];
            }
        }
    }

    // If no existing session found, generate a new one
    // Excel uses what appears to be a random number
    return Math.floor(Math.random() * 9999999999).toString();
}

/**
 * Checks if the add-in is already sideloaded
 * @param {string} addinId - The Add-in ID to check
 * @returns {boolean} True if already sideloaded
 */
function isAddinAlreadySideloaded(addinId) {
    // Check if manifest key exists
    const manifestKey = `__OSF_UPLOADFILE.Manifest.16.${addinId}`;
    return localStorage.getItem(manifestKey) !== null;
}

/**
 * Injects the manifest into Excel's localStorage
 * This mimics what happens when a user manually sideloads an add-in
 * @param {string} manifestXML - The manifest XML content
 * @param {string} addinId - The Add-in ID
 * @returns {Object} Result object with success status and message
 */
export function injectManifestToLocalStorage(manifestXML, addinId) {
    try {
        console.log(`[SRK Sideloader] Starting manifest injection for add-in: ${addinId}`);

        // Check if already sideloaded
        if (isAddinAlreadySideloaded(addinId)) {
            console.log('[SRK Sideloader] Add-in is already sideloaded, refreshing manifest...');
        }

        // Find or generate session ID
        const sessionId = findOrGenerateSessionId();
        console.log(`[SRK Sideloader] Using session ID: ${sessionId}`);

        // Create the manifest storage key and value (matching Excel's format)
        const manifestKey = `__OSF_UPLOADFILE.Manifest.16.${addinId}`;
        const manifestValue = JSON.stringify({
            data: manifestXML,
            createdOn: Date.now(),
            refreshRate: 3
        });

        // Create the MyAddins key (stores array of add-in IDs)
        const myAddinsKey = `__OSF_UPLOADFILE.MyAddins.16.${sessionId}`;
        let myAddinsValue;

        // Get existing add-ins or create new array
        const existingMyAddins = localStorage.getItem(myAddinsKey);
        if (existingMyAddins) {
            try {
                const parsed = JSON.parse(existingMyAddins);
                const addinIds = parsed.data || [];
                // Add our add-in ID if not already present
                if (!addinIds.includes(addinId)) {
                    addinIds.push(addinId);
                }
                myAddinsValue = JSON.stringify({
                    data: addinIds,
                    createdOn: parsed.createdOn || Date.now(),
                    refreshRate: 3
                });
            } catch (e) {
                // If parsing fails, create fresh entry
                myAddinsValue = JSON.stringify({
                    data: [addinId],
                    createdOn: Date.now(),
                    refreshRate: 3
                });
            }
        } else {
            myAddinsValue = JSON.stringify({
                data: [addinId],
                createdOn: Date.now(),
                refreshRate: 3
            });
        }

        // Create the AddinCommandsMyAddins key (for command add-ins)
        const commandsKey = `__OSF_UPLOADFILE.AddinCommandsMyAddins.16.${sessionId}`;
        let commandsValue;

        const existingCommands = localStorage.getItem(commandsKey);
        if (existingCommands) {
            try {
                const parsed = JSON.parse(existingCommands);
                const addinIds = parsed.data || [];
                if (!addinIds.includes(addinId)) {
                    addinIds.push(addinId);
                }
                commandsValue = JSON.stringify({
                    data: addinIds,
                    createdOn: parsed.createdOn || Date.now(),
                    refreshRate: 3
                });
            } catch (e) {
                commandsValue = JSON.stringify({
                    data: [addinId],
                    createdOn: Date.now(),
                    refreshRate: 3
                });
            }
        } else {
            commandsValue = JSON.stringify({
                data: [addinId],
                createdOn: Date.now(),
                refreshRate: 3
            });
        }

        // Write to localStorage
        localStorage.setItem(manifestKey, manifestValue);
        localStorage.setItem(myAddinsKey, myAddinsValue);
        localStorage.setItem(commandsKey, commandsValue);

        console.log('[SRK Sideloader] ✓ Manifest injected successfully!');
        console.log('[SRK Sideloader] localStorage keys written:');
        console.log(`  - ${manifestKey}`);
        console.log(`  - ${myAddinsKey}`);
        console.log(`  - ${commandsKey}`);

        return {
            success: true,
            message: 'Manifest injected successfully. Please refresh Excel to load the add-in.',
            sessionId: sessionId,
            addinId: addinId
        };

    } catch (error) {
        console.error('[SRK Sideloader] Failed to inject manifest:', error);
        return {
            success: false,
            message: `Failed to inject manifest: ${error.message}`,
            error: error
        };
    }
}

/**
 * Main function to auto-sideload the manifest
 * This is called from the content script running in Excel tabs
 */
export async function autoSideloadManifest() {
    try {
        console.log('%c [SRK Auto-Sideloader] Starting...', 'background: #4CAF50; color: white; font-weight: bold; padding: 2px 4px;');

        // Get the manifest XML
        const manifestXML = await getManifestXML();

        // Extract the add-in ID
        const addinId = extractAddinId(manifestXML);

        if (!addinId) {
            throw new Error('Could not extract Add-in ID from manifest');
        }

        console.log(`[SRK Auto-Sideloader] Add-in ID: ${addinId}`);

        // Inject into localStorage
        const result = injectManifestToLocalStorage(manifestXML, addinId);

        if (result.success) {
            console.log('%c [SRK Auto-Sideloader] ✓ SUCCESS!', 'background: #4CAF50; color: white; font-weight: bold; padding: 2px 4px;');
            console.log('[SRK Auto-Sideloader] The add-in should load automatically.');
            console.log('[SRK Auto-Sideloader] If not visible, try refreshing Excel (F5).');
        }

        return result;

    } catch (error) {
        console.error('%c [SRK Auto-Sideloader] FAILED:', 'background: #f44336; color: white; font-weight: bold; padding: 2px 4px;', error);
        return {
            success: false,
            message: error.message,
            error: error
        };
    }
}
