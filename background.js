// background.js

import { startLoop, stopLoop, openNextTabInLoop } from './looper.js';
import { SUBMISSION_FOUND_URL } from './constants.js';

// --- THIS IS THE NEW PART ---
// Listen for the keyboard shortcut command
chrome.commands.onCommand.addListener((command) => {
  // The _execute_action command is triggered by the shortcut
  if (command === '_execute_action') {
    // Set a flag in storage that the popup will check for on launch.
    // This tells the popup to perform a special action.
    chrome.storage.local.set({ openAction: 'focusMasterSearch' });
  }
});


// All other functions and listeners remain the same.
async function triggerPowerAutomate(payload) {
  try {
    const resp = await fetch(SUBMISSION_FOUND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    console.log("Flow triggered successfully. Status:", resp.status);
  } catch (e) {
    console.error("Flow error", e);
  }
}

function updateBadge() {
  chrome.storage.local.get(['extensionState', 'foundEntries'], (data) => {
    const isExtensionOn = data.extensionState === 'on';
    const foundCount = data.foundEntries?.length || 0;

    if (isExtensionOn) {
      chrome.action.setBadgeBackgroundColor({ color: '#0052cc' });
      if (foundCount > 0) {
        chrome.action.setBadgeText({ text: foundCount.toString() });
      } else {
        chrome.action.setBadgeText({ text: 'ON' });
      }
    } else {
      chrome.action.setBadgeText({ text: '' });
    }
  });
}

function handleStateChange(state) {
    if (state === 'on') {
        startLoop();
    } else {
        stopLoop();
    }
}

chrome.runtime.onStartup.addListener(() => {
  updateBadge();
  chrome.storage.local.get('extensionState', data => handleStateChange(data.extensionState));
});

chrome.storage.onChanged.addListener((changes) => {
  if (changes.extensionState) {
    handleStateChange(changes.extensionState.newValue);
  }
  if (changes.extensionState || changes.foundEntries) {
    updateBadge();
  }
});

updateBadge();
chrome.storage.local.get('extensionState', data => handleStateChange(data.extensionState));

chrome.runtime.onMessage.addListener((msg, sender) => {
  switch (msg.action) {
    case 'inspectionResult':
      if (!sender.tab || !sender.tab.id) return;
      if (msg.found) {
        stopLoop({ keepTabOpen: true });
        chrome.scripting.executeScript({
            target: { tabId: sender.tab.id },
            files: ['highlighter.js']
        });
      } else {
        chrome.tabs.remove(sender.tab.id).catch(e => {});
        openNextTabInLoop();
      }
      break;
    case 'highlightingComplete':
        if (sender.tab?.id) {
          chrome.tabs.remove(sender.tab.id).catch(e => {});
        }
        openNextTabInLoop();
        break;
    case 'focusTab':
      if (sender.tab?.id) {
        chrome.tabs.update(sender.tab.id, { active: true });
      }
      break;
    case 'addNames':
      chrome.storage.local.get({ foundEntries: [] }, data => {
        const map = new Map(data.foundEntries.map(e => [e.name, e]));
        msg.entries.forEach(e => map.set(e.name, e));
        chrome.storage.local.set({ foundEntries: Array.from(map.values()) });
      });
      break;
    case 'runFlow':
      chrome.storage.local.get('extensionState', data => {
        if (data.extensionState === 'on') {
          triggerPowerAutomate(msg.payload);
        }
      });
      break;
  }
});