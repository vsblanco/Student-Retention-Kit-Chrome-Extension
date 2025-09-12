// looper.js

let currentLoopIndex = 0;
let isLooping = false;
let masterListCache = [];
let foundUrlCache = new Set(); // Using URLs for checking duplicates.
let activeTabs = new Map();
let maxConcurrentTabs = 3;

// Allows the background script to update our URL cache in real-time.
export function addToFoundUrlCache(url) {
  if (!foundUrlCache.has(url)) {
    console.log(`Adding ${url} to the found cache to prevent re-checks.`);
    foundUrlCache.add(url);
  }
}

async function loadSettings() {
    const { concurrentTabs = 3 } = await chrome.storage.local.get('concurrentTabs');
    maxConcurrentTabs = concurrentTabs;
}

// --- FIX 1 of 3: Add a 'force' parameter to the function signature ---
export async function startLoop(force = false) {
  // --- FIX 2 of 3: Modify the check to respect the 'force' parameter ---
  if (isLooping && !force) return;

  console.log('START command received.');
  isLooping = true;
  currentLoopIndex = 0;
  activeTabs.clear();

  await loadSettings();
  
  const { masterEntries, foundEntries = [] } = await chrome.storage.local.get(['masterEntries', 'foundEntries']);
  
  // Create a Set of URLs from the found list for fast lookups.
  // We filter out any entries that might not have a valid URL.
  foundUrlCache = new Set(foundEntries.map(e => e.url).filter(Boolean));
  if (foundUrlCache.size > 0) {
    console.log(`${foundUrlCache.size} URLs already in 'Found' list will be skipped.`);
  }

  if (!masterEntries || masterEntries.length === 0) {
    console.log('Master list is empty.');
    stopLoop();
    return;
  }
  masterListCache = masterEntries;
  
  // Store the initial state of the loop for the UI.
  chrome.storage.local.set({ loopStatus: { current: 0, total: masterListCache.length } });
  
  for (let i = 0; i < maxConcurrentTabs; i++) {
    processNextInQueue();
  }
}

export function stopLoop() {
  if (!isLooping) return;
  console.log('STOP command received.');
  isLooping = false;
  
  // Clear the loop status from storage when stopped.
  chrome.storage.local.remove('loopStatus');
  
  for (const tabId of activeTabs.keys()) {
    chrome.tabs.remove(tabId).catch(e => {});
  }
  activeTabs.clear();
  chrome.storage.local.set({ extensionState: 'off' });
}

export function processNextInQueue(finishedTabId = null) {
  if (!isLooping) return;

  if (finishedTabId) {
    activeTabs.delete(finishedTabId);
  }

  if (currentLoopIndex >= masterListCache.length && activeTabs.size === 0) {
    console.log('Looped through entire list. Starting over.');
    // Resetting the counter for the new loop.
    chrome.storage.local.set({ loopStatus: { current: 0, total: masterListCache.length } }); 
    
    // --- FIX 3 of 3: Force the loop to restart ---
    startLoop(true);
    return;
  }

  if (currentLoopIndex < masterListCache.length) {
    const entry = masterListCache[currentLoopIndex];
    currentLoopIndex++;
    
    // Update the current index in storage for the UI to read.
    chrome.storage.local.set({ loopStatus: { current: currentLoopIndex, total: masterListCache.length } });
    
    // --- THE NEW CHECK, NOW USING URL ---
    if (foundUrlCache.has(entry.url)) {
      console.log(`Skipping already found URL: ${entry.url}`);
      setTimeout(() => processNextInQueue(), 0);
      return;
    }

    if (!entry.url || !entry.url.startsWith('http')) {
        console.warn(`Skipping invalid URL: ${entry.url}`);
        setTimeout(() => processNextInQueue(), 50);
        return;
    }
    
    openTab(entry);
  }
}

async function openTab(entry) {
    if (!isLooping) return;

    const urlToOpen = new URL(entry.url);
    urlToOpen.searchParams.set('looper', 'true');
    
    console.log(`Opening tab for index #${currentLoopIndex - 1}: ${urlToOpen.href}`);
    try {
        const tab = await chrome.tabs.create({ url: urlToOpen.href, active: false });
        activeTabs.set(tab.id, entry);
    } catch (error) {
        console.error(`Failed to create tab: ${error.message}. Continuing.`);
        setTimeout(() => processNextInQueue(), 100);
    }
}