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
      const studentName = getFirstStudentName();
      const timeStr = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
      const url = window.location.href;

      chrome.runtime.sendMessage({ action: 'addNames', entries: [{ name: studentName, time: timeStr, url }] });
      chrome.runtime.sendMessage({ action: 'runFlow', payload: { name: studentName, url, timestamp: new Date().toISOString() } });
      chrome.runtime.sendMessage({ action: 'focusTab' });
    }
  }

  function walkTheDOM(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    function processChunk() {
        let count = 0;
        while ((node = walker.nextNode()) && count < 200) {
            highlightAndNotify(node);
            count++;
        }
        if (node) {
            setTimeout(processChunk, 50);
        } else {
            finishCheck();
        }
    }
    processChunk();
  }

  function finishCheck() {
    if (isLooperRun) {
        chrome.runtime.sendMessage({ action: 'inspectionResult', found: keywordFound });
    }
  }

  console.log(`Content script loaded. Using keyword: "${todayStr}"`);
  walkTheDOM(document.body);

})();