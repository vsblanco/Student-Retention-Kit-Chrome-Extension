// Student Renderer - Handles rendering of student lists and active student display
import { elements } from './ui-manager.js';
import { GENERIC_AVATAR_URL } from '../constants/index.js';

/**
 * Converts student name from "Last, First" format to "First Last" format if a comma is present.
 * @param {string} name - The student name to convert
 * @returns {string} The converted name in "First Last" format
 */
function convertNameFormat(name) {
    if (!name || typeof name !== 'string') return name;

    // Check if the name contains a comma
    if (!name.includes(',')) {
        return name.trim();
    }

    // Split by comma and trim whitespace
    const parts = name.split(',').map(part => part.trim());

    // If we don't have exactly 2 parts, return the original name
    if (parts.length !== 2) {
        return name.trim();
    }

    // Convert from "Last, First" to "First Last"
    const [lastName, firstName] = parts;
    return `${firstName} ${lastName}`;
}

/**
 * Gets the display name for a student based on the reformat name setting
 * @param {Object} entry - The student entry
 * @param {boolean} reformatEnabled - Whether name reformatting is enabled
 * @returns {string} The formatted or original name
 */
export function getDisplayName(entry, reformatEnabled = true) {
    const originalName = entry.nameOriginal || entry.name || 'Unknown Student';

    if (reformatEnabled) {
        return convertNameFormat(originalName);
    }

    return originalName;
}

/**
 * Normalizes student data for consistent rendering
 * @param {Object} entry - The student entry
 * @param {boolean} reformatEnabled - Whether name reformatting is enabled
 */
export function resolveStudentData(entry, reformatEnabled = true) {
    return {
        name: getDisplayName(entry, reformatEnabled),
        nameOriginal: entry.nameOriginal || entry.name || 'Unknown Student',
        sortable_name: entry.sortable_name || null,
        phone: entry.phone || null,
        daysOut: parseInt(entry.daysOut || 0),
        missing: parseInt(entry.missingCount || 0),
        StudentNumber: entry.StudentNumber || null,
        SyStudentId: entry.SyStudentId || null,
        url: entry.url || entry.Gradebook || null, // Fallback to Gradebook for legacy data
        Photo: entry.Photo || null,
        isNew: entry.isNew || false,
        created_at: entry.created_at || null,
        timestamp: entry.timestamp || null,
        assignment: entry.assignment || null
    };
}

/**
 * Sets the active student in the contact tab
 * @param {Object|null} rawEntry - The student data or null to clear
 * @param {Object} callManager - Reference to call manager for state updates
 */
export async function setActiveStudent(rawEntry, callManager) {
    const contactTab = document.getElementById('contact');
    if (!contactTab) return;

    // Reset automation styles when switching (but not during active automation)
    if (!callManager?.automationMode) {
        if (elements.dialBtn) {
            elements.dialBtn.classList.remove('automation');
            elements.dialBtn.innerHTML = '<i class="fas fa-phone"></i>';
        }
        if (callManager) {
            callManager.updateCallInterfaceState();

            // Hide disposition section when new student is selected
            callManager.waitingForDisposition = false;
            if (elements.callDispositionSection) {
                elements.callDispositionSection.style.display = 'none';
            }
        }
        if (elements.upNextCard) {
            elements.upNextCard.style.display = 'none';
        }
        if (elements.manageQueueBtn) {
            elements.manageQueueBtn.style.display = 'none';
        }
    }

    // 1. Handle "No Student Selected" State
    if (!rawEntry) {
        Array.from(contactTab.children).forEach(child => {
            if (child.id === 'contactPlaceholder') {
                child.style.display = 'flex';
            } else if (child.id === 'five9ConnectionIndicator') {
                // Five9 indicator managed separately
            } else {
                child.style.display = 'none';
            }
        });
        return;
    }

    // 2. Handle "Student Selected" State
    Array.from(contactTab.children).forEach(child => {
        if (child.id === 'contactPlaceholder') {
            child.style.display = 'none';
        } else if (child.id === 'five9ConnectionIndicator') {
            // Five9 indicator managed separately
        } else {
            child.style.display = '';
        }
    });

    // Get reformat name setting
    const settings = await chrome.storage.local.get(['reformatNameEnabled']);
    const reformatEnabled = settings.reformatNameEnabled !== undefined ? settings.reformatNameEnabled : true;

    const data = resolveStudentData(rawEntry, reformatEnabled);

    // Generate initials from name
    const nameParts = data.name.trim().split(/\s+/);
    let initials = '';
    if (nameParts.length > 0) {
        const firstInitial = nameParts[0][0] || '';
        const lastInitial = nameParts.length > 1 ? nameParts[nameParts.length - 1][0] : '';
        initials = (firstInitial + lastInitial).toUpperCase();
        if (!initials) initials = '?';
    }

    const displayPhone = data.phone ? data.phone : "No Phone Listed";

    // AVATAR LOGIC
    if (elements.contactAvatar) {
        elements.contactAvatar.style.color = '';
        if (data.Photo && data.Photo !== GENERIC_AVATAR_URL) {
            elements.contactAvatar.textContent = '';
            elements.contactAvatar.style.backgroundImage = `url('${data.Photo}')`;
            elements.contactAvatar.style.backgroundSize = 'cover';
            elements.contactAvatar.style.backgroundPosition = 'center';
            elements.contactAvatar.style.backgroundColor = 'transparent';
        } else {
            elements.contactAvatar.style.backgroundImage = 'none';
            elements.contactAvatar.textContent = initials;
            elements.contactAvatar.style.backgroundColor = '#e0e7ff';
        }
    }

    if (elements.contactName) elements.contactName.textContent = data.name;
    if (elements.contactPhone) elements.contactPhone.textContent = displayPhone;

    if (elements.contactDetail) {
        elements.contactDetail.textContent = `${data.daysOut} Days Out`;
        elements.contactDetail.style.display = 'block';
    }

    let colorCode = '#10b981';
    if (data.daysOut > 10) colorCode = '#ef4444';
    else if (data.daysOut > 5) colorCode = '#f97316';
    else if (data.daysOut > 2) colorCode = '#f59e0b';

    if (elements.contactCard) {
        elements.contactCard.style.borderLeftColor = colorCode;
    }
}

/**
 * Sets the automation mode UI with gray styling
 * @param {number} queueLength - Number of students in queue
 */
export function setAutomationModeUI(queueLength) {
    const contactTab = document.getElementById('contact');
    if (!contactTab) return;

    // Ensure content is visible (hide placeholder)
    Array.from(contactTab.children).forEach(child => {
        if (child.id === 'contactPlaceholder') {
            child.style.display = 'none';
        } else {
            child.style.display = '';
        }
    });

    // Update Contact Card
    if (elements.contactName) elements.contactName.textContent = "Automation Mode";
    if (elements.contactDetail) elements.contactDetail.textContent = `${queueLength} Students Selected`;
    if (elements.contactPhone) elements.contactPhone.textContent = "Multi-Dial Queue";

    // Create visual badge for count
    if (elements.contactAvatar) {
        elements.contactAvatar.textContent = queueLength;
        elements.contactAvatar.style.backgroundImage = 'none';
        elements.contactAvatar.style.backgroundColor = '#6b7280';
        elements.contactAvatar.style.color = '#ffffff';
    }

    // Transform the Dial Button to Gray
    if (elements.dialBtn) {
        elements.dialBtn.classList.add('automation');
        elements.dialBtn.innerHTML = '<i class="fas fa-robot"></i>';
    }

    // Update Status Text
    if (elements.callStatusText) {
        elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:#6b7280;"></span> Ready to Auto-Dial`;
    }

    if (elements.contactCard) {
        elements.contactCard.style.borderLeftColor = '#6b7280';
    }

    // Show Manage Queue Button
    if (elements.manageQueueBtn) {
        elements.manageQueueBtn.style.display = 'block';
    }
}

/**
 * Renders the found submissions list
 * @param {Array} rawEntries - Array of found submissions
 */
export async function renderFoundList(rawEntries) {
    if (!elements.foundList) return;
    elements.foundList.innerHTML = '';

    if (!rawEntries || rawEntries.length === 0) {
        elements.foundList.innerHTML = '<li style="justify-content:center; color:gray;">No submissions found yet.</li>';
        if (elements.clearListBtn) {
            elements.clearListBtn.style.display = 'none';
        }
        return;
    }

    // Show clear button when there are entries
    if (elements.clearListBtn) {
        elements.clearListBtn.style.display = 'block';
    }

    // Get reformat name setting
    const settings = await chrome.storage.local.get(['reformatNameEnabled']);
    const reformatEnabled = settings.reformatNameEnabled !== undefined ? settings.reformatNameEnabled : true;

    // Create pairs of raw entries and resolved data, then sort by timestamp
    const entriesWithRaw = rawEntries.map(rawEntry => ({
        raw: rawEntry,
        resolved: resolveStudentData(rawEntry, reformatEnabled)
    }));
    entriesWithRaw.sort((a, b) => new Date(b.resolved.timestamp) - new Date(a.resolved.timestamp));

    entriesWithRaw.forEach(({ raw, resolved }) => {
        const li = document.createElement('li');
        let timeDisplay = 'Just now';
        if (resolved.timestamp) {
            timeDisplay = new Date(resolved.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        }

        const assignmentTitle = resolved.assignment || 'Untitled Assignment';

        li.innerHTML = `
            <div style="display: flex; align-items: center; width:100%;">
                <div class="heatmap-indicator heatmap-green"></div>
                <div style="flex-grow:1; display:flex; justify-content:space-between; align-items:center;">
                    <div style="display:flex; flex-direction:column;">
                        <span class="student-name" style="font-weight:500; color:var(--primary-color); cursor:pointer;">${resolved.name}</span>
                        <span style="font-size:0.8em; color:var(--text-secondary);">${assignmentTitle}</span>
                    </div>
                    <span class="timestamp-pill">${timeDisplay}</span>
                </div>
            </div>
        `;

        const nameLink = li.querySelector('.student-name');
        nameLink.addEventListener('click', (e) => {
            e.stopPropagation();
            if (resolved.url) chrome.tabs.create({ url: resolved.url });
        });
        nameLink.addEventListener('mouseenter', () => nameLink.style.textDecoration = 'underline');
        nameLink.addEventListener('mouseleave', () => nameLink.style.textDecoration = 'none');

        // Store the ORIGINAL raw entry data on the li element for context menu access
        // This ensures all fields like SyStudentId are preserved
        li.dataset.entryData = JSON.stringify(raw);

        elements.foundList.appendChild(li);
    });
}

/**
 * Filters the found list based on search term
 * @param {Event} e - Input event
 */
export function filterFoundList(e) {
    const term = e.target.value.toLowerCase();
    const items = elements.foundList.querySelectorAll('li');
    items.forEach(li => {
        const text = li.textContent.toLowerCase();
        li.style.display = text.includes(term) ? 'flex' : 'none';
    });
}

/**
 * Renders the master student list
 * @param {Array} rawEntries - Array of student data
 * @param {Function} onStudentClick - Callback when student is clicked
 */
export async function renderMasterList(rawEntries, onStudentClick) {
    if (!elements.masterList) return;
    elements.masterList.innerHTML = '';

    // Update total count indicator
    if (elements.totalCountText) {
        const count = rawEntries ? rawEntries.length : 0;
        elements.totalCountText.textContent = `Total Students: ${count}`;
    }

    if (!rawEntries || rawEntries.length === 0) {
        elements.masterList.innerHTML = '<li style="justify-content:center;">Master list is empty.</li>';
        return;
    }

    // Get reformat name setting
    const settings = await chrome.storage.local.get(['reformatNameEnabled']);
    const reformatEnabled = settings.reformatNameEnabled !== undefined ? settings.reformatNameEnabled : true;

    rawEntries.forEach(rawEntry => {
        const data = resolveStudentData(rawEntry, reformatEnabled);

        const li = document.createElement('li');
        li.className = 'expandable';
        li.style.cursor = 'pointer';

        li.setAttribute('data-name', data.name);
        li.setAttribute('data-missing', data.missing);
        li.setAttribute('data-days', data.daysOut);
        li.setAttribute('data-created', data.created_at || '');

        let heatmapClass = data.daysOut > 10 ? 'heatmap-red' : (data.daysOut > 5 ? 'heatmap-orange' : (data.daysOut > 2 ? 'heatmap-yellow' : 'heatmap-green'));

        let missingPillHtml = '';
        if (data.missing > 0) {
            missingPillHtml = `<span class="missing-pill">${data.missing} Missing <i class="fas fa-chevron-down" style="font-size:0.8em; margin-left:4px;"></i></span>`;
        }

        let newTagHtml = '';
        if (data.isNew) {
            newTagHtml = `<span style="background:#e0f2fe; color:#0369a1; font-size:0.7em; padding:2px 6px; border-radius:8px; margin-left:6px; font-weight:bold; border:1px solid #bae6fd;">New</span>`;
        }

        // Build missing assignments details HTML
        let missingDetailsHtml = '<li><em>No missing assignments found.</em></li>';
        if (rawEntry.missingAssignments && rawEntry.missingAssignments.length > 0) {
            missingDetailsHtml = rawEntry.missingAssignments.map(assignment => {
                // Handle both old (title) and new (assignmentTitle) field names
                const title = assignment.assignmentTitle || assignment.title || 'Unknown Assignment';
                const linkHtml = assignment.submissionLink
                    ? `<a href="${assignment.submissionLink}" target="_blank" style="color:#2563eb; text-decoration:none;">${title}</a>`
                    : title;
                return `<li style="margin-bottom:6px;">
                    ${linkHtml}
                    <div style="font-size:0.9em; color:#6b7280; margin-top:2px;">
                        Due: ${assignment.dueDate} | Score: ${assignment.score}
                    </div>
                </li>`;
            }).join('');
        }

        li.innerHTML = `
            <div style="display: flex; align-items: center; width:100%;">
                <div class="heatmap-indicator ${heatmapClass}"></div>
                <div style="flex-grow:1;">
                    <div style="display:flex; justify-content:space-between; align-items:center; width:100%;">
                        <div style="display:flex; align-items:center;">
                            <span class="student-name" style="font-weight: 500; color:var(--text-main); position:relative; z-index:2;">${data.name}</span>
                            ${newTagHtml}
                        </div>
                        ${missingPillHtml}
                    </div>
                    <span style="font-size:0.8em; color:gray;">${data.daysOut} Days Out</span>
                </div>
            </div>
            <div class="missing-details" style="display: none; margin-top: 10px; border-top: 1px solid rgba(0,0,0,0.05); padding-top: 8px; cursor: default;">
                <ul style="padding: 0; margin: 0; font-size: 0.85em; color: #4b5563; list-style-type: none;">
                    ${missingDetailsHtml}
                </ul>
            </div>
        `;

        // Click listener for student selection
        li.addEventListener('click', (e) => {
            if (onStudentClick) {
                onStudentClick(rawEntry, li, e);
            }
        });

        // Student name click - open gradebook
        const nameLink = li.querySelector('.student-name');
        if (nameLink) {
            nameLink.addEventListener('click', (e) => {
                e.stopPropagation();
                if (data.url) chrome.tabs.create({ url: data.url });
            });
            nameLink.addEventListener('mouseenter', () => {
                nameLink.style.textDecoration = 'underline';
                nameLink.style.color = 'var(--primary-color)';
            });
            nameLink.addEventListener('mouseleave', () => {
                nameLink.style.textDecoration = 'none';
                nameLink.style.color = 'var(--text-main)';
            });
        }

        // Missing pill click - toggle details
        const pill = li.querySelector('.missing-pill');
        if (pill) {
            pill.addEventListener('click', (e) => {
                e.stopPropagation();
                const details = li.querySelector('.missing-details');
                const icon = pill.querySelector('i');
                if (details) {
                    const isHidden = details.style.display === 'none' || !details.style.display;
                    details.style.display = isHidden ? 'block' : 'none';
                    if (icon) icon.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
                }
            });
        }

        // Prevent click propagation on details div
        const detailsDiv = li.querySelector('.missing-details');
        if (detailsDiv) {
            detailsDiv.addEventListener('click', (e) => {
                e.stopPropagation();
            });
        }

        elements.masterList.appendChild(li);
    });
}

/**
 * Filters the master list based on search term
 * @param {Event} e - Input event
 */
export function filterMasterList(e) {
    const term = e.target.value.toLowerCase();
    const listItems = elements.masterList.querySelectorAll('li.expandable');
    listItems.forEach(li => {
        const name = li.getAttribute('data-name').toLowerCase();
        li.style.display = name.includes(term) ? 'flex' : 'none';
    });
}

/**
 * Sorts the master list based on selected criteria
 */
export function sortMasterList() {
    const criteria = elements.sortSelect.value;
    const listItems = Array.from(elements.masterList.querySelectorAll('li.expandable'));

    listItems.sort((a, b) => {
        if (criteria === 'name') {
            return a.getAttribute('data-name').localeCompare(b.getAttribute('data-name'));
        } else if (criteria === 'missing') {
            return parseInt(b.getAttribute('data-missing')) - parseInt(a.getAttribute('data-missing'));
        } else if (criteria === 'days') {
            return parseInt(b.getAttribute('data-days')) - parseInt(a.getAttribute('data-days'));
        } else if (criteria === 'newest') {
            const dateA = new Date(a.getAttribute('data-created') || 0);
            const dateB = new Date(b.getAttribute('data-created') || 0);
            return dateB - dateA;
        }
    });
    listItems.forEach(item => elements.masterList.appendChild(item));
}
