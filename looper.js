// looper.js

let currentLoopIndex = 0;
let isLooping = false;
let masterListCache = [];
let foundUrlCache = new Set(); // Using URLs for checking duplicates.
let activeTabs = new Map();
let maxConcurrentTabs = 3;

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

// --- THIS IS THE UPDATED FUNCTION ---
export async function startLoop(force = false) {
  if (isLooping && !force) return;

  console.log('START command received.');
  isLooping = true;
  currentLoopIndex = 0;
  activeTabs.clear();

  await loadSettings();
  
  // Fetch master list, found entries, AND the new filter setting
  const { masterEntries, foundEntries = [], looperDaysOutFilter = 'all' } = await chrome.storage.local.get(['masterEntries', 'foundEntries', 'looperDaysOutFilter']);
  
  foundUrlCache = new Set(foundEntries.map(e => e.url).filter(Boolean));
  if (foundUrlCache.size > 0) {
    console.log(`${foundUrlCache.size} URLs already in 'Found' list will be skipped.`);
  }

  if (!masterEntries || masterEntries.length === 0) {
    console.log('Master list is empty.');
    stopLoop();
    return;
  }
  
  // --- NEW FILTERING LOGIC ---
  let filteredMasterList = masterEntries;
  const filterText = looperDaysOutFilter.trim().toLowerCase();

  if (filterText !== 'all' && filterText !== '') {
    const advancedFilterRegex = /^\s*([><]=?|=)\s*(\d+)\s*$/;
    const match = filterText.match(advancedFilterRegex);

    if (match) {
        filteredMasterList = masterEntries.filter(entry => {
            const operator = match[1];
            const value = parseInt(match[2], 10);
            const daysout = entry.daysout;

            if (daysout == null) return false;

            switch (operator) {
                case '>':  return daysout > value;
                case '<':  return daysout < value;
                case '>=': return daysout >= value;
                case '<=': return daysout <= value;
                case '=':  return daysout === value;
                default:   return false;
            }
        });
        console.log(`Looper will run on a filtered list of ${filteredMasterList.length} students (condition: ${filterText}).`);
    } else {
        console.warn(`Invalid looper filter format: "${filterText}". Running on all ${masterEntries.length} students.`);
    }
  }
  // --- END OF NEW LOGIC ---

  masterListCache = filteredMasterList; // Use the (potentially) filtered list
  
  chrome.storage.local.set({ loopStatus: { current: 0, total: masterListCache.length } });
  
  for (let i = 0; i < maxConcurrentTabs; i++) {
    processNextInQueue();
  }
}

export function stopLoop() {
  if (!isLooping) return;
  console.log('STOP command received.');
  isLooping = false;
  
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
    chrome.storage.local.set({ loopStatus: { current: 0, total: masterListCache.length } }); 
    
    startLoop(true);
    return;
  }

  if (currentLoopIndex < masterListCache.length) {
    const entry = masterListCache[currentLoopIndex];
    currentLoopIndex++;
    
    chrome.storage.local.set({ loopStatus: { current: currentLoopIndex, total: masterListCache.length } });
    
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