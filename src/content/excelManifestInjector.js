// excelManifestInjector.js
// Content script that injects the manifest sideloader into Excel's page context

console.log('%c [SRK] Excel Manifest Injector Script LOADED', 'background: #FF9800; color: white; font-size: 14px; font-weight: bold; padding: 2px 4px;');

// Check if auto-sideload is enabled
chrome.storage.local.get(['autoSideloadManifest'], async (result) => {
    const autoSideloadEnabled = result.autoSideloadManifest !== undefined
        ? result.autoSideloadManifest
        : true; // Default to enabled

    if (!autoSideloadEnabled) {
        console.log('[SRK] Auto-sideload is disabled in settings');
        return;
    }

    console.log('[SRK] Auto-sideload is enabled, proceeding...');

    try {
        // Get the manifest XML from the extension
        const manifestUrl = chrome.runtime.getURL('assets/Excell Add-In Manifest.xml');
        const response = await fetch(manifestUrl);
        const manifestXML = await response.text();

        // Extract Add-in ID from manifest
        const idMatch = manifestXML.match(/<Id>([^<]+)<\/Id>/);
        const addinId = idMatch ? idMatch[1] : null;

        if (!addinId) {
            console.error('[SRK] Could not extract Add-in ID from manifest');
            return;
        }

        console.log(`[SRK] Add-in ID extracted: ${addinId}`);

        // Create and inject the script that will run in page context
        const script = document.createElement('script');
        script.textContent = `
(function() {
    console.log('%c [SRK Auto-Sideloader] Injecting manifest into localStorage...', 'background: #4CAF50; color: white; font-weight: bold; padding: 2px 4px;');

    const addinId = ${JSON.stringify(addinId)};
    const manifestXML = ${JSON.stringify(manifestXML)};

    try {
        // Check if already sideloaded
        const manifestKey = '__OSF_UPLOADFILE.Manifest.16.' + addinId;
        const alreadyExists = localStorage.getItem(manifestKey) !== null;

        if (alreadyExists) {
            console.log('[SRK Auto-Sideloader] Manifest already exists in localStorage, updating...');
        } else {
            console.log('[SRK Auto-Sideloader] First-time sideload detected');
        }

        // Find or generate session ID
        function findOrGenerateSessionId() {
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                if (key && key.startsWith('__OSF_UPLOADFILE.MyAddins.16.')) {
                    const parts = key.split('.');
                    if (parts.length >= 4) {
                        return parts[3];
                    }
                }
            }
            return Math.floor(Math.random() * 9999999999).toString();
        }

        const sessionId = findOrGenerateSessionId();
        console.log('[SRK Auto-Sideloader] Using session ID:', sessionId);

        // Write manifest
        const manifestValue = JSON.stringify({
            data: manifestXML,
            createdOn: Date.now(),
            refreshRate: 3
        });
        localStorage.setItem(manifestKey, manifestValue);

        // Update MyAddins
        const myAddinsKey = '__OSF_UPLOADFILE.MyAddins.16.' + sessionId;
        const existingMyAddins = localStorage.getItem(myAddinsKey);
        let myAddinsValue;

        if (existingMyAddins) {
            try {
                const parsed = JSON.parse(existingMyAddins);
                const addinIds = parsed.data || [];
                if (!addinIds.includes(addinId)) {
                    addinIds.push(addinId);
                }
                myAddinsValue = JSON.stringify({
                    data: addinIds,
                    createdOn: parsed.createdOn || Date.now(),
                    refreshRate: 3
                });
            } catch (e) {
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
        localStorage.setItem(myAddinsKey, myAddinsValue);

        // Update AddinCommandsMyAddins
        const commandsKey = '__OSF_UPLOADFILE.AddinCommandsMyAddins.16.' + sessionId;
        const existingCommands = localStorage.getItem(commandsKey);
        let commandsValue;

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
        localStorage.setItem(commandsKey, commandsValue);

        console.log('%c [SRK Auto-Sideloader] ✓ SUCCESS!', 'background: #4CAF50; color: white; font-weight: bold; padding: 2px 4px;');
        console.log('[SRK Auto-Sideloader] Manifest injected. Keys written:');
        console.log('  •', manifestKey);
        console.log('  •', myAddinsKey);
        console.log('  •', commandsKey);
        console.log('[SRK Auto-Sideloader] The Student Retention Add-in should load automatically!');

    } catch (error) {
        console.error('%c [SRK Auto-Sideloader] FAILED:', 'background: #f44336; color: white; font-weight: bold; padding: 2px 4px;', error);
    }
})();
        `;

        // Inject the script into the page
        (document.head || document.documentElement).appendChild(script);
        script.remove(); // Clean up the script tag

        console.log('[SRK] Manifest injection script executed');

        // Notify the background script
        chrome.runtime.sendMessage({
            type: 'SRK_MANIFEST_INJECTED',
            addinId: addinId,
            timestamp: Date.now()
        }).catch(() => {
            // Extension might not be ready, that's ok
        });

    } catch (error) {
        console.error('[SRK] Failed to inject manifest:', error);
    }
});
