// background.js
import { startLoop, stopLoop, processNextInQueue, addToFoundUrlCache } from './looper.js';

// --- CORE LISTENERS ---

chrome.action.onClicked.addListener((tab) => chrome.sidePanel.open({ windowId: tab.windowId }));
chrome.commands.onCommand.addListener((command, tab) => {
  if (command === '_execute_action') chrome.sidePanel.open({ windowId: tab.windowId });
});
chrome.runtime.onStartup.addListener(() => {
  updateBadge();
  chrome.storage.local.get('extensionState', data => handleStateChange(data.extensionState));
});
chrome.storage.onChanged.addListener((changes) => {
  if (changes.extensionState) handleStateChange(changes.extensionState.newValue);
  if (changes.extensionState || changes.foundEntries) updateBadge();
});
chrome.runtime.onMessage.addListener(async (msg, sender) => {
  if (msg.action === 'inspectionResult') {
    if (msg.found && msg.entry) {
      await addStudentToFoundList(msg.entry);
      sendConnectionPings(msg.entry);
    }
    if (sender.tab?.id) {
      chrome.tabs.remove(sender.tab.id).catch(e => {});
      processNextInQueue(sender.tab.id);
    }
  } else if (msg.type === 'test-connection-pa') {
    handlePaConnectionTest(msg.connection);
  } else if (msg.type === 'send-debug-payload') {
    if (msg.payload) {
      sendConnectionPings(msg.payload);
    }
  }
});

// --- CONNECTION HANDLING ---

async function sendConnectionPings(payload) {
    const { connections = [], debugMode = false } = await chrome.storage.local.get(['connections', 'debugMode']);
    const bodyPayload = { ...payload };
    
    // If the payload itself isn't marked as a debug send, check the global setting.
    if (!bodyPayload.debug && debugMode) {
      bodyPayload.debug = true;
    }


    for (const conn of connections) {
        if (conn.type === 'power-automate') {
            triggerPowerAutomate(conn, bodyPayload);
        } else if (conn.type === 'pusher') {
            chrome.runtime.sendMessage({
                type: 'trigger-pusher',
                connection: conn,
                payload: bodyPayload
            });
        }
    }
}

async function handlePaConnectionTest(connection) {
    const testPayload = {
        name: 'Test Submission',
        url: '#',
        grade: '100',
        timestamp: new Date().toISOString(),
        test: true
    };

    const result = await triggerPowerAutomate(connection, testPayload);
    
    chrome.runtime.sendMessage({
        type: 'connection-test-result',
        connectionType: 'power-automate',
        success: result.success,
        error: result.error || 'Check console for details.'
    });
}

async function triggerPowerAutomate(connection, payload) {
  try {
    const resp = await fetch(connection.url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    if (!resp.ok && resp.status !== 202) { // 202 is "Accepted" for flows
      throw new Error(`HTTP ${resp.status}`);
    }
    console.log("Power Automate flow triggered successfully. Status:", resp.status);
    return { success: true };
  } catch (e) {
    console.error("Power Automate flow error:", e);
    return { success: false, error: e.message };
  }
}

// --- STATE & DATA MANAGEMENT ---

function updateBadge() {
  chrome.storage.local.get(['extensionState', 'foundEntries'], (data) => {
    const isExtensionOn = data.extensionState === 'on';
    const foundCount = data.foundEntries?.length || 0;
    if (isExtensionOn) {
      chrome.action.setBadgeBackgroundColor({ color: '#0052cc' });
      chrome.action.setBadgeText({ text: foundCount > 0 ? foundCount.toString() : 'ON' });
    } else {
      chrome.action.setBadgeText({ text: '' });
    }
  });
}

function handleStateChange(state) {
    if (state === 'on') startLoop();
    else stopLoop();
}

async function addStudentToFoundList(entry) {
    const { foundEntries = [] } = await chrome.storage.local.get('foundEntries');
    const map = new Map(foundEntries.map(e => [e.url, e]));
    map.set(entry.url, entry);
    addToFoundUrlCache(entry.url);
    await chrome.storage.local.set({ foundEntries: Array.from(map.values()) });
}

// Initial setup on load
updateBadge();
chrome.storage.local.get('extensionState', data => handleStateChange(data.extensionState));
