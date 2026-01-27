// [2025-10-24 03:00 PM]
// Version: 18.6 - Updated for nested storage structure
// This script is now purely for VISUAL enhancement (Passive Mode).
// It runs automatically when a user views a Gradebook page.

(async function() {
  // Only run on grades page
  if (!window.location.href.includes('/grades/')) return;

  // Legacy flat keys (for backwards compatibility)
  const FLAT_KEYS = {
      CUSTOM_KEYWORD: 'customKeyword',
      HIGHLIGHT_COLOR: 'highlightColor',
      EMBED_IN_CANVAS: 'embedInCanvas',
      EXTENSION_STATE: 'extensionState',
      USE_SPECIFIC_DATE: 'useSpecificDate',
      SPECIFIC_SUBMISSION_DATE: 'specificSubmissionDate'
  };

  // Helper to get value from nested storage structure
  // Checks new nested path first, then falls back to legacy flat key
  function getSettingValue(data, key, defaultValue) {
      // Check nested paths first
      if (key === 'embedInCanvas') {
          if (data.settings?.canvas?.embedInCanvas !== undefined) {
              return data.settings.canvas.embedInCanvas;
          }
      } else if (key === 'highlightColor') {
          if (data.settings?.canvas?.highlightColor !== undefined) {
              return data.settings.canvas.highlightColor;
          }
      } else if (key === 'extensionState') {
          if (data.state?.extensionState !== undefined) {
              return data.state.extensionState;
          }
      }
      // Fall back to legacy flat key
      return data[key] !== undefined ? data[key] : defaultValue;
  }

  // FETCH SETTINGS - get both new nested structure and legacy flat keys
  const data = await chrome.storage.local.get(['settings', 'state', ...Object.values(FLAT_KEYS)]);

  // [CRITICAL FIX] Check if Embed In Canvas is enabled.
  // If false or undefined, default to TRUE (or FALSE depending on preference).
  // Assuming default is true based on previous context, but strictly respecting the toggle.
  const isEnabled = getSettingValue(data, 'embedInCanvas', true);

  // If explicitly set to false, stop execution.
  if (isEnabled === false) {
      console.log("[Visual Inspector] Disabled via settings.");
      return;
  }

  const highlightColor = getSettingValue(data, 'highlightColor', '#ffff00');
  const customKeyword = getSettingValue(data, 'customKeyword', '');
  const useSpecificDate = getSettingValue(data, 'useSpecificDate', false);
  const specificDate = getSettingValue(data, 'specificSubmissionDate', null);

  // Determine Keyword
  let searchKeyword;
  if (customKeyword) {
      searchKeyword = customKeyword;
  } else {
      let targetDate = new Date();

      // If using specific date, parse and use that instead of today's date
      if (useSpecificDate && specificDate) {
          const [year, month, day] = specificDate.split('-').map(Number);
          targetDate = new Date(year, month - 1, day); // month is 0-indexed
          console.log(`[Visual Inspector] Using specific date: ${specificDate} -> ${targetDate.toDateString()}`);
      }

      const opts = { month: 'short', day: 'numeric' };
      // Format: "Oct 23 at"
      searchKeyword = targetDate.toLocaleDateString('en-US', opts).replace(',', '') + ' at';
  }

  console.log(`[Visual Inspector] Running. Keyword: "${searchKeyword}"`);

  // --- 1. SUBMISSION HIGHLIGHTER ---
  function highlightSubmissions() {
      if (!searchKeyword) return;

      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      const nodesToHighlight = [];
      
      let node;
      while (node = walker.nextNode()) {
          // Check if text contains keyword and is inside a submission row
          if (node.nodeValue.includes(searchKeyword)) {
              const parent = node.parentElement;
              // Ensure we are looking at a submitted date cell or similar context
              if (parent && (parent.closest('td.submitted') || parent.closest('.submission_date'))) {
                  nodesToHighlight.push(node);
              }
          }
      }

      nodesToHighlight.forEach(node => {
          const idx = node.nodeValue.indexOf(searchKeyword);
          if (idx >= 0) {
              const range = document.createRange();
              range.setStart(node, idx);
              range.setEnd(node, idx + searchKeyword.length);
              
              const span = document.createElement('span');
              span.style.backgroundColor = highlightColor;
              span.style.fontWeight = 'bold';
              span.style.borderRadius = '3px';
              span.textContent = searchKeyword; 
              
              range.deleteContents();
              range.insertNode(span);
          }
      });
  }

  // --- 2. MISSING ASSIGNMENT PILLS ---
  function markMissingAssignments() {
      const rows = document.querySelectorAll('tr.student_assignment');
      const now = new Date();
      const monthMap = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };

      rows.forEach(row => {
          // Skip if already marked
          if (row.querySelector('.submission-missing-pill')) return;

          // Exclusion Logic: Skip Category/Group rows
          if (row.classList.contains('hard_coded_group') || 
              row.classList.contains('group_total') || 
              row.classList.contains('final_grade')) {
              return;
          }

          // 1. Get Due Date
          const dueCell = row.querySelector('td.due');
          let dueDate = null;
          if (dueCell) {
              const dueText = dueCell.textContent.trim();
              const match = dueText.match(/(\w{3})\s(\d{1,2})/); // "Oct 23"
              if (match) {
                  const year = now.getFullYear();
                  const month = monthMap[match[1]];
                  const day = parseInt(match[2], 10);
                  dueDate = new Date(year, month, day);
                  dueDate.setHours(23, 59, 59, 999); // End of day
              }
          }

          // 2. Get Score
          const scoreCell = row.querySelector('td.assignment_score .grade');
          let score = null;
          if (scoreCell) {
              const scoreText = scoreCell.textContent.trim().replace('%', '');
              // Check for numeric 0 or "-"
              if (scoreText === '0' || scoreText === '0.0') score = 0;
              else if (scoreText === '-') score = null; // Ungraded
              else if (!isNaN(parseFloat(scoreText))) score = parseFloat(scoreText);
          }

          // 3. Get Submission Status
          const submittedCell = row.querySelector('td.submitted');
          const hasSubmission = submittedCell && submittedCell.textContent.trim().length > 5; // Rough check for timestamp text

          // 4. Calculate Logic Conditions
          const isManualZero = (score === 0) && !hasSubmission;

          // Global Future Filter with Exception
          if (dueDate && dueDate > now && !isManualZero) return; 

          const isOfficialMissing = row.classList.contains('missing') || (dueCell && dueCell.classList.contains('missing'));
          const isUnsubmittedPastDue = !hasSubmission && (dueDate && dueDate < now);

          if (isOfficialMissing || isUnsubmittedPastDue || isManualZero) {
              injectMissingPill(row);
          }
      });
  }

  function injectMissingPill(row) {
      const statusCell = row.querySelector('td.status');
      if (statusCell) {
          const pill = document.createElement('div');
          
          pill.className = 'submission-missing-pill';
          pill.style.display = 'inline-block';
          pill.style.backgroundColor = 'transparent'; // No fill
          pill.style.border = '1px solid #EE352E';    // Red outline
          pill.style.color = '#EE352E';               // Red text
          pill.style.borderRadius = '12px';
          pill.style.padding = '2px 8px';
          pill.style.fontSize = '14px';
          pill.style.fontWeight = 400;
          pill.style.textTransform = 'lowercase';     // Lowercase
          pill.style.marginTop = '4px';
          pill.textContent = 'missing';               // Lowercase text
          
          statusCell.appendChild(pill);
      }
  }

  // Run on load
  if (document.readyState === 'complete') {
      highlightSubmissions();
      markMissingAssignments();
  } else {
      window.addEventListener('load', () => {
          highlightSubmissions();
          markMissingAssignments();
      });
  }

  setTimeout(() => {
      highlightSubmissions();
      markMissingAssignments();
  }, 2000);

})();