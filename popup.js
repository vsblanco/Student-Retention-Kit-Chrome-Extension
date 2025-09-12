// popup.js
document.addEventListener('DOMContentLoaded', () => {
  // show today’s keyword
  const opts     = { month: 'short', day: 'numeric' };
  const todayStr = new Date()
    .toLocaleDateString('en-US', opts)
    .replace(',', '') + ' at';
  document.getElementById('keyword').textContent = todayStr;

  const list     = document.getElementById('foundList');
  const clearBtn = document.getElementById('clearBtn');

  function render(entries) {
    list.innerHTML = '';
    if (!entries.length) {
      const li = document.createElement('li');
      li.textContent = 'None';
      list.appendChild(li);
    } else {
      entries.forEach(({ name, time, url }) => {
        const li = document.createElement('li');
        const a  = document.createElement('a');
        a.textContent = `${name} – ${time}`;
        a.href        = url;
        a.target      = '_blank';
        a.style.textDecoration = 'none';
        a.style.color = 'blue';
        li.appendChild(a);
        list.appendChild(li);
      });
    }
  }

  // 1) ask the page to rescan
  chrome.tabs.query({ active: true, currentWindow: true }, tabs => {
    chrome.tabs.sendMessage(tabs[0].id, { action: 'rescanNames' }, () => {
      // 2) then load & render the global list
      chrome.storage.local.get({ foundEntries: [] }, data => {
        render(data.foundEntries);
      });
    });
  });

  // clear handler
  clearBtn.addEventListener('click', () => {
    chrome.storage.local.set({ foundEntries: [] }, () => {
      render([]);
    });
  });
});
