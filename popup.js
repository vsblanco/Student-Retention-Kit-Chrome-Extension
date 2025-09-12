// popup.js

import { MASTER_LIST_URL } from './constants.js';

function renderFoundList(entries) {
  const list = document.getElementById('foundList');
  list.innerHTML = '';
  entries.forEach(({ name, time, url }) => {
    const li = document.createElement('li');
    const a  = document.createElement('a');
    a.textContent = name;
    a.href = '#';
    a.style.color = 'var(--accent-color)';
    a.style.textDecoration = 'none';
    a.addEventListener('click', e => {
      e.preventDefault();
      chrome.tabs.create({ url });
      window.close();
    });
    li.appendChild(a);
    if (time) {
        const timeBadge = document.createElement('span');
        timeBadge.className = 'pill-badge align-right'; 
        timeBadge.textContent = time;
        li.appendChild(timeBadge);
    }
    list.appendChild(li);
  });
}

function renderMasterList(entries) {
  const list = document.getElementById('masterList');
  list.innerHTML = '';
  entries.forEach(({ name, time, url, phone }) => { 
    const li = document.createElement('li');
    if (url && url !== '#N/A' && url.startsWith('http')) {
      const a  = document.createElement('a');
      a.textContent = name;
      a.href = url;
      a.style.color = 'var(--accent-color)';
      a.style.textDecoration = 'none';
      a.addEventListener('click', e => {
        e.preventDefault();
        chrome.tabs.create({ url });
        window.close();
      });
      li.appendChild(a);
    } else {
      const nameSpan = document.createElement('span');
      nameSpan.textContent = name;
      nameSpan.style.color = '#888';
      nameSpan.title = 'Invalid URL. Please update on the master list.';
      li.appendChild(nameSpan);
    }
    if (phone) {
        const phoneSpan = document.createElement('span');
        phoneSpan.className = 'pill-badge';
        phoneSpan.textContent = phone;
        li.appendChild(phoneSpan);
    }
    list.appendChild(li);
  });
}

async function updateMaster() {
  const list = document.getElementById('masterList');
  list.innerHTML = '<li>Loading…</li>';
  try {
    const resp = await fetch(MASTER_LIST_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: '{}'  
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    const students = data.students || [];
    const entries = students.map(s => ({
      name: s.name, time: s.time || '', url:  s.url, phone: s.phone || ''
    }));
    const now = new Date();
    const timestampStr = now.toLocaleDateString('en-US', {
        month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true
    }).replace(',', '');
    await new Promise(res => chrome.storage.local.set({ masterEntries: entries, lastUpdated: timestampStr }, res));
    const badge = document.querySelector('.tab-button[data-tab="master"] .count');
    if(badge) badge.textContent = entries.length;
    renderMasterList(entries);
    const lastUpdatedSpan = document.getElementById('lastUpdatedTime');
    if(lastUpdatedSpan) lastUpdatedSpan.textContent = `Last updated: ${timestampStr}`;
  } catch (e) {
    console.error('Failed to update master list', e);
    list.innerHTML = '<li>Error loading list</li>';
  }
}

function createRipple(event) {
    const button = event.currentTarget;
    const rect = button.getBoundingClientRect();
    const ripple = document.createElement("span");
    ripple.className = 'ripple';
    ripple.style.height = ripple.style.width = Math.max(rect.width, rect.height) + "px";
    const x = event.clientX - rect.left - ripple.offsetWidth / 2;
    const y = event.clientY - rect.top - ripple.offsetHeight / 2;
    ripple.style.left = `${x}px`;
    ripple.style.top = `${y}px`;
    button.appendChild(ripple);
    setTimeout(() => {
      ripple.remove();
    }, 600);
}

document.addEventListener('DOMContentLoaded', () => {
  let isStarted; 
  const manifest = chrome.runtime.getManifest();
  document.getElementById('version-display').textContent = `Version ${manifest.version}`;
  const keywordDisplay = document.getElementById('keyword');

  function updateKeywordDisplay() {
    chrome.storage.local.get({ customKeyword: '' }, (data) => {
        if (data.customKeyword) {
            keywordDisplay.textContent = data.customKeyword;
        } else {
            const now = new Date();
            const opts = { month: 'short', day: 'numeric' };
            keywordDisplay.textContent = now.toLocaleDateString('en-US', opts).replace(',', '') + ' at';
        }
    });
  }
  updateKeywordDisplay();

  chrome.storage.local.get({ foundEntries: [] }, data => {
    const badge = document.querySelector('.tab-button[data-tab="found"] .count');
    if(badge) badge.textContent = data.foundEntries.length;
    if (data.foundEntries.length) renderFoundList(data.foundEntries);
    else { document.getElementById('foundList').innerHTML = '<li>None yet</li>'; }
  });

  chrome.storage.local.get(['masterEntries', 'lastUpdated'], data => {
    const badge = document.querySelector('.tab-button[data-tab="master"] .count');
    if(badge) badge.textContent = data.masterEntries?.length || 0;
    if (data.masterEntries?.length) { renderMasterList(data.masterEntries); }
    else { document.getElementById('masterList').innerHTML = '<li>None yet</li>'; }
    const lastUpdatedSpan = document.getElementById('lastUpdatedTime');
    if (lastUpdatedSpan && data.lastUpdated) { lastUpdatedSpan.textContent = `Last updated: ${data.lastUpdated}`; }
  });

  document.getElementById('clearBtn').addEventListener('click', () => {
    chrome.storage.local.set({ foundEntries: [] }, () => location.reload());
  });

  document.getElementById('updateMasterBtn').addEventListener('click', (event) => {
    createRipple(event);
    updateMaster();
  });

  const tabs = document.querySelectorAll('.tab-button');
  const panes = document.querySelectorAll('.tab-content');
  tabs.forEach(btn => {
    btn.addEventListener('click', () => {
      tabs.forEach(b => b.classList.remove('active'));
      panes.forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    });
  });

  const embedToggle = document.getElementById('embedToggle');
  const colorPicker = document.getElementById('colorPicker');
  const customKeywordInput = document.getElementById('customKeywordInput');
  const sharepointBtn = document.getElementById('sharepointBtn');

  if (embedToggle) {
    chrome.storage.local.get({ embedInCanvas: true }, (data) => {
      embedToggle.checked = data.embedInCanvas;
    });
    embedToggle.addEventListener('change', (event) => {
      chrome.storage.local.set({ embedInCanvas: event.target.checked });
    });
  }
  if (colorPicker) {
    chrome.storage.local.get({ highlightColor: '#ffff00' }, (data) => { colorPicker.value = data.highlightColor; });
    colorPicker.addEventListener('input', (event) => { chrome.storage.local.set({ highlightColor: event.target.value }); });
  }
  if (customKeywordInput) {
    chrome.storage.local.get({ customKeyword: '' }, (data) => { customKeywordInput.value = data.customKeyword; });
    customKeywordInput.addEventListener('input', (event) => {
        const newKeyword = event.target.value.trim();
        chrome.storage.local.set({ customKeyword: newKeyword }, () => { updateKeywordDisplay(); });
    });
  }
  if (sharepointBtn) {
    sharepointBtn.addEventListener('click', (event) => {
        createRipple(event);
        chrome.tabs.create({ url: "https://edukgroup365.sharepoint.com/sites/SM-StudentServices/SitePages/CollabHome.aspx" });
        window.close();
    });
  }

  const startBtn = document.getElementById('startBtn');
  const startBtnText = document.getElementById('startBtnText');
  function updateButtonState(state) {
    isStarted = (state === 'on');
    if (isStarted) {
      startBtn.style.backgroundColor = 'var(--light-accent-color)';
      startBtnText.textContent = 'Stop';
    } else {
      startBtn.style.backgroundColor = 'var(--accent-color)';
      startBtnText.textContent = 'Start';
    }
  }
  if (startBtn && startBtnText) {
    chrome.storage.local.get({ extensionState: 'off' }, data => { updateButtonState(data.extensionState); });
    startBtn.addEventListener('click', (event) => {
      createRipple(event);
      const newState = !isStarted ? 'on' : 'off';
      chrome.storage.local.set({ extensionState: newState });
      updateButtonState(newState);
    });
  }
});