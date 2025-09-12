// content.js

// 1) Compute today's date string, e.g. "Jun 10 at"
const now = new Date();
const opts = { month: 'short', day: 'numeric' };
const todayStr = now.toLocaleDateString('en-US', opts).replace(',', '') + ' at';

let keywordFound = false;
let playedSound = false;
let focusedTab = false;

// Play sound once
function playSoundOnce() {
  if (playedSound) return;
  playedSound = true;
  const audio = new Audio(chrome.runtime.getURL('assets/audio/song.mp3'));
  audio.preload = 'auto';
  audio.play().catch(() => {
    document.addEventListener('click', () => audio.play().catch(() => {}), { once: true, capture: true });
  });
}

// Focus once
function focusTabOnce() {
  if (focusedTab) return;
  focusedTab = true;
  chrome.runtime.sendMessage({ action: 'focusTab' });
}

// Get first student name
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

// Highlight and conditionally trigger
function highlightInTextNode(node, keyword) {
  const cell = node.parentElement?.closest('td.submitted');
  if (!cell) return;

  const orig = node.nodeValue;
  const norm = orig.replace(/\u00A0/g, ' ');
  const idx = norm.indexOf(keyword);
  if (idx < 0) return;

  // Always highlight
  const before = orig.slice(0, idx);
  const matchStr = orig.slice(idx, idx + keyword.length);
  const after = orig.slice(idx + keyword.length);
  const parent = node.parentNode;
  parent.insertBefore(document.createTextNode(before), node);
  const span = document.createElement('span');
  span.textContent = matchStr;
  span.style.backgroundColor = 'yellow';
  span.style.fontWeight = 'bold';
  span.style.fontSize = '1.1em';
  parent.insertBefore(span, node);
  parent.insertBefore(document.createTextNode(after), node);
  parent.removeChild(node);

  // Only once: play, focus, then check storage & send
  if (!keywordFound) {
    keywordFound = true;
    playSoundOnce();
    focusTabOnce();

    const studentName = getFirstStudentName();
    const timeStr = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const url = window.location.href;

    chrome.storage.local.get({ foundEntries: [] }, data => {
      const exists = data.foundEntries.some(e => e.name === studentName);
      if (!exists) {
        // trigger and store
        chrome.runtime.sendMessage({
          action: 'runFlow',
          payload: { name: studentName, url, timestamp: now.toISOString() }
        });
        chrome.runtime.sendMessage({
          action: 'addNames',
          entries: [{ name: studentName, time: timeStr, url }]
        });
      }
    });
  }

  // Continue in same cell
  const nextNode = parent.childNodes[Array.from(parent.childNodes).indexOf(span) + 1];
  if (nextNode) highlightInTextNode(nextNode, keyword);
}

// Walk and apply
function walkAndHighlight(root, keyword) {
  const SKIP = ['SCRIPT','STYLE','NOSCRIPT','IFRAME'];
  for (const node of Array.from(root.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      highlightInTextNode(node, keyword);
    } else if (node.nodeType === Node.ELEMENT_NODE && !SKIP.includes(node.tagName)) {
      walkAndHighlight(node, keyword);
    }
  }
}

// Initial run
walkAndHighlight(document.body, todayStr);