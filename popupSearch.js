// popupSearch.js

// --- THIS IS THE CHANGE ---
// Import both render functions from popup.js
import { renderMasterList, renderFoundList } from './popup.js';

document.addEventListener('DOMContentLoaded', () => {
  // --- Master List Search Logic ---
  const searchMasterInput = document.getElementById('newItemInput');
  if (searchMasterInput) {
    searchMasterInput.addEventListener('input', () => {
      const term = searchMasterInput.value.trim().toLowerCase();
      const numericTerm = term.replace(/[^0-9]/g, '');

      chrome.storage.local.get({ masterEntries: [] }, data => {
        const entries = data.masterEntries;
        const filtered = entries.filter(entry => {
          const nameMatch = entry.name.toLowerCase().includes(term);
          let phoneMatch = false;
          if (entry.phone && numericTerm.length > 0) {
              const numericPhone = entry.phone.replace(/[^0-9]/g, '');
              phoneMatch = numericPhone.includes(numericTerm);
          }
          return nameMatch || phoneMatch;
        });

        renderMasterList(filtered);
        
        const badge = document.querySelector('.tab-button[data-tab="master"] .count');
        if (badge) badge.textContent = filtered.length;
      });
    });
  }

  // --- NEW: Found List Search Logic ---
  const searchFoundInput = document.getElementById('searchFoundListInput');
  if(searchFoundInput) {
    searchFoundInput.addEventListener('input', () => {
      const term = searchFoundInput.value.trim().toLowerCase();

      chrome.storage.local.get({ foundEntries: [] }, data => {
        const entries = data.foundEntries;

        const filtered = entries.filter(entry => {
          // Only search by name for the found list
          return entry.name.toLowerCase().includes(term);
        });

        // Re-render the list with only matching entries
        renderFoundList(filtered);

        // Update the badge to reflect number of matches
        const badge = document.querySelector('.tab-button[data-tab="found"] .count');
        if (badge) {
          badge.textContent = filtered.length;
        }
      });
    });
  }
});