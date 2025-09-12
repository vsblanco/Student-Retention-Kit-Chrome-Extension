// background.js

const FLOW_URL =
  "https://prod-12.westus.logic.azure.com:443/workflows/8ac3b733279b4838833c5454b31d005d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=RccL3_U3q2nYCWoNfmGisH7rBUIyHrx4SD1vc6bzo7w";

/**
 * Sends a POST to your Power Automate flow.
 * @param {Object} payload — JSON data to send (e.g. info about the keyword event)
 */
async function triggerPowerAutomate(payload) {
  try {
    const resp = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    if (!resp.ok) {
      throw new Error(`HTTP ${resp.status} – ${await resp.text()}`);
    }
    const json = await resp.json();
    console.log("✅ Power Automate flow triggered:", json);
  } catch (err) {
    console.error("❌ Flow trigger error:", err);
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'focusTab' && sender.tab && sender.tab.id) {
    chrome.tabs.update(sender.tab.id, { active: true });
  } else if (msg.action === 'addNames' && Array.isArray(msg.entries)) {
    chrome.storage.local.get({ foundEntries: [] }, data => {
      const map = new Map(
        data.foundEntries.map(e => [e.name, { time: e.time, url: e.url }])
      );
      msg.entries.forEach(({ name, time, url }) => map.set(name, { time, url }));
      const merged = Array.from(map, ([name, { time, url }]) => ({ name, time, url }));
      chrome.storage.local.set({ foundEntries: merged });
    });
  } else if (msg.action === 'runFlow') {
    triggerPowerAutomate(msg.payload || {});
    sendResponse({ status: 'flow_sent' });
  }
});