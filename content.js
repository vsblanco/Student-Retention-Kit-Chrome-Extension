// content.js

(async function() {
  // Fetch all settings from storage at once.
  const { 
    extensionState = 'off', 
    highlightColor = '#ffff00',
    customKeyword = '' 
  } = await chrome.storage.local.get(['extensionState', 'highlightColor', 'customKeyword']);
  
  const isLooperRun = new URLSearchParams(window.location.search).has('looper');
  let keywordFound = false;
  let foundEntry = null;

  // Decide which keyword to use: the custom one, or the default date.
  let todayStr;
  if (customKeyword) {
    todayStr = customKeyword;
  } else {
    const now = new Date();
    const opts = { month: 'short', day: 'numeric' };
    todayStr = now.toLocaleDateString('en-US', opts).replace(',', '') + ' at';
  }

  function getFirstStudentName() {
    const re = /Grades for\s*([\w ,'-]+)/g;
    let match;
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null);
    while (walker.nextNode()) {
      const txt = walker.currentNode.nodeValue;
      while ((match = re.exec(txt))) return match[1].trim();
    }
    return 'Unknown student';
  }

  function highlightAndNotify(node) {
    if (keywordFound) return;

    const cell = node.parentElement?.closest('td.submitted');
    if (!cell) return;

    const idx = node.nodeValue.indexOf(todayStr);
    if (idx < 0) return;

    keywordFound = true;

    const parent = node.parentNode;
    const span = document.createElement('span');
    span.textContent = node.nodeValue.slice(idx, idx + todayStr.length);
    span.style.backgroundColor = highlightColor;
    span.style.fontWeight = 'bold';
    span.style.fontSize = '1.1em';
    
    parent.insertBefore(document.createTextNode(node.nodeValue.slice(0, idx)), node);
    parent.insertBefore(span, node);
    parent.insertBefore(document.createTextNode(node.nodeValue.slice(idx + todayStr.length)), node);
    parent.removeChild(node);

    if (extensionState === 'off') {
        span.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    if (isLooperRun) {
      console.log('Keyword FOUND on a looper-opened tab.');
      
      const studentName = getFirstStudentName().trim();
      const timeStr = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
      
      const urlObject = new URL(window.location.href);
      urlObject.searchParams.delete('looper');
      const cleanUrl = urlObject.href;

      foundEntry = {
          name: studentName,
          time: timeStr,
          url: cleanUrl,
          timestamp: new Date().toISOString()
      };
    }
  }

  // --- REVISED FUNCTION TO AVOID BACKGROUND THROTTLING ---
  // This version uses a single, synchronous loop instead of setTimeout,
  // which ensures it runs at full speed even in background tabs.
  function walkTheDOM(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while (!keywordFound && (node = walker.nextNode())) {
        highlightAndNotify(node);
    }
    // After the loop is done (either by finding the keyword or finishing the page),
    // call finishCheck.
    finishCheck();
  }

  function finishCheck() {
    if (isLooperRun) {
        chrome.runtime.sendMessage({ action: 'inspectionResult', found: keywordFound, entry: foundEntry });
    }
  }

  console.log(`Content script loaded. Using keyword: "${todayStr}"`);
  walkTheDOM(document.body);

})();