import { STORAGE_KEYS, EXTENSION_STATES } from '../constants.js';

// --- CONFIGURATION ---
const CANVAS_DOMAIN = "https://nuc.instructure.com";
const GENERIC_AVATAR_URL = "https://nuc.instructure.com/images/messages/avatar-50.png";

// --- STATE MANAGEMENT ---
let isScanning = false;
let isCallActive = false;
let callTimerInterval = null;
let selectedQueue = []; // Tracks multiple selected students

// --- DOM ELEMENTS CACHE ---
const elements = {};

document.addEventListener('DOMContentLoaded', () => {
    // --- Block Text Highlighting Globally ---
    const style = document.createElement('style');
    style.textContent = `
        * {
            -webkit-user-select: none; /* Safari/Chrome */
            user-select: none;
        }
        input, textarea {
            -webkit-user-select: text;
            user-select: text;
        }
    `;
    document.head.appendChild(style);

    cacheDomElements();
    initializeApp();
});

function cacheDomElements() {
    // Navigation
    elements.tabs = document.querySelectorAll('.tab-button');
    elements.contents = document.querySelectorAll('.tab-content');
    
    // Header
    elements.headerSettingsBtn = document.getElementById('headerSettingsBtn');
    elements.versionText = document.getElementById('versionText');
    
    // Checker Tab
    elements.startBtn = document.getElementById('startBtn');
    elements.startBtnText = document.getElementById('startBtnText');
    elements.startBtnIcon = elements.startBtn ? elements.startBtn.querySelector('i') : null;
    elements.statusDot = document.getElementById('statusDot');
    elements.statusText = document.getElementById('statusText');
    elements.foundList = document.querySelector('#checker .glass-list');
    elements.clearListBtn = document.querySelector('#checker .btn-secondary');
    elements.foundSearch = document.querySelector('#checker input[type="text"]');

    // Call Tab
    elements.dialBtn = document.getElementById('dialBtn');
    elements.callStatusText = document.querySelector('.call-status-bar');
    elements.callTimer = document.querySelector('.call-timer');
    elements.otherInputArea = document.getElementById('otherInputArea');
    elements.customNote = document.getElementById('customNote');
    elements.confirmNoteBtn = elements.otherInputArea ? elements.otherInputArea.querySelector('.btn-primary') : null;
    elements.dispositionGrid = document.querySelector('.disposition-grid'); 
    
    // Call Tab - Student Card & Placeholder Logic
    const contactTab = document.getElementById('contact');
    if (contactTab) {
        elements.contactCard = contactTab.querySelector('.setting-card');
        
        // --- INJECT PLACEHOLDER ---
        let placeholder = document.getElementById('contactPlaceholder');
        if (!placeholder) {
            placeholder = document.createElement('div');
            placeholder.id = 'contactPlaceholder';
            placeholder.style.cssText = 'display:none; flex-direction:column; align-items:center; justify-content:flex-start; padding-top:80px; height:100%; min-height:400px; color:#9ca3af; text-align:center; padding-left:20px; padding-right:20px;';
            placeholder.innerHTML = `
                <i class="fas fa-user-graduate" style="font-size:3em; margin-bottom:15px; opacity:0.5;"></i>
                <span style="font-size:1.1em; font-weight:500;">No Student Selected</span>
                <span style="font-size:0.9em; margin-top:5px; color:#6b7280;">Select a student from the Master List<br>to view details and make calls.</span>
            `;
            contactTab.insertBefore(placeholder, contactTab.firstChild);
        }
        elements.contactPlaceholder = placeholder;

        // Cache Card Details
        if (elements.contactCard) {
            elements.contactAvatar = contactTab.querySelector('.setting-card div[style*="border-radius:50%"]');
            const infoContainer = contactTab.querySelector('.setting-card div > div:not([style])'); 
            if (infoContainer) {
                elements.contactName = infoContainer.children[0];
                elements.contactDetail = infoContainer.children[1]; 
            } else {
                elements.contactName = contactTab.querySelector('.setting-card div > div:first-child');
                elements.contactDetail = contactTab.querySelector('.setting-card div > div:last-child');
            }
            elements.contactPhone = contactTab.querySelector('.phone-number-display');
        }
    }

    // Data Tab
    elements.masterList = document.getElementById('masterList');
    elements.masterSearch = document.getElementById('masterSearch');
    elements.sortSelect = document.getElementById('sortSelect');
    elements.updateMasterBtn = document.getElementById('updateMasterBtn');
    elements.downloadMasterBtn = document.getElementById('downloadMasterBtn');
    elements.updateQueueSection = document.getElementById('updateQueueSection');
    elements.queueCloseBtn = elements.updateQueueSection ? elements.updateQueueSection.querySelector('.icon-btn') : null;
    elements.lastUpdatedText = document.getElementById('lastUpdatedText');
    
    // Step 1 File Input
    elements.studentPopFile = document.getElementById('studentPopFile');

    // Modals & Settings
    elements.versionModal = document.getElementById('versionModal');
    elements.closeVersionBtn = document.getElementById('closeVersionBtn');
}

function initializeApp() {
    setupEventListeners();
    loadStorageData();
    setActiveStudent(null); 
}

// --- EVENT LISTENERS ---
function setupEventListeners() {
    elements.tabs.forEach(tab => {
        tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });

    // --- NEW: DETECT CTRL KEY RELEASE TO SWITCH TAB ---
    document.addEventListener('keyup', (e) => {
        // Check for Control (Windows) or Meta (Mac Command)
        if (e.key === 'Control' || e.key === 'Meta') {
            // Only switch if we have multiple students selected (Automation Mode)
            if (selectedQueue.length > 1) {
                switchTab('contact');
            }
        }
    });
    // --------------------------------------------------

    if (elements.headerSettingsBtn) elements.headerSettingsBtn.addEventListener('click', () => switchTab('settings'));
    
    if (elements.versionText) elements.versionText.addEventListener('click', () => elements.versionModal.style.display = 'flex');
    if (elements.closeVersionBtn) elements.closeVersionBtn.addEventListener('click', () => elements.versionModal.style.display = 'none');
    window.addEventListener('click', (e) => {
        if (elements.versionModal && e.target === elements.versionModal) elements.versionModal.style.display = 'none';
    });

    if (elements.startBtn) elements.startBtn.addEventListener('click', toggleScanState);
    if (elements.clearListBtn) elements.clearListBtn.addEventListener('click', () => chrome.storage.local.set({ [STORAGE_KEYS.FOUND_ENTRIES]: [] }));

    if (elements.foundSearch) {
        elements.foundSearch.addEventListener('input', filterFoundList);
    }

    if (elements.dialBtn) elements.dialBtn.addEventListener('click', () => toggleCallState());
    
    const dispositionContainer = document.querySelector('.disposition-grid');
    if (dispositionContainer) {
        dispositionContainer.addEventListener('click', (e) => {
            const btn = e.target.closest('.disposition-btn');
            if (!btn) return;
            if (btn.innerText.includes('Other')) elements.otherInputArea.style.display = 'block';
            else handleDisposition(btn.innerText.trim());
        });
    }

    if (elements.confirmNoteBtn) {
        elements.confirmNoteBtn.addEventListener('click', () => {
            const note = elements.customNote.value;
            handleDisposition(`Custom Note: ${note}`);
            elements.otherInputArea.style.display = 'none';
            elements.customNote.value = '';
        });
    }

    if (elements.updateMasterBtn) {
        elements.updateMasterBtn.addEventListener('click', () => {
            if (elements.updateQueueSection) {
                elements.updateQueueSection.style.display = 'block';
                elements.updateQueueSection.scrollIntoView({ behavior: 'smooth' });
                
                resetQueueUI();
                
                const step1 = document.getElementById('step1');
                if(step1) {
                    step1.className = 'queue-item active';
                    step1.querySelector('i').className = 'fas fa-spinner';
                }

                if(elements.studentPopFile) {
                    elements.studentPopFile.click();
                }
            }
        });
    }
    
    if (elements.studentPopFile) {
        elements.studentPopFile.addEventListener('change', (e) => {
            handleFileImport(e.target.files[0]);
        });
    }
    
    if (elements.queueCloseBtn) {
        elements.queueCloseBtn.addEventListener('click', () => {
            elements.updateQueueSection.style.display = 'none';
        });
    }

    if (elements.masterSearch) elements.masterSearch.addEventListener('input', filterMasterList);
    if (elements.sortSelect) elements.sortSelect.addEventListener('change', sortMasterList);
    if (elements.downloadMasterBtn) elements.downloadMasterBtn.addEventListener('click', exportMasterListCSV);
}

// --- FILE IMPORT LOGIC (STEP 1) ---

function handleFileImport(file) {
    if (!file) {
        resetQueueUI();
        return;
    }

    const step1 = document.getElementById('step1');
    const timeSpan = step1.querySelector('.step-time');
    const startTime = Date.now();

    const reader = new FileReader();

    reader.onload = function(e) {
        const content = e.target.result;
        let students = [];

        try {
            if (file.name.endsWith('.csv')) {
                students = parseCSV(content);
            } else {
                alert("Unsupported file type. Please use .csv.");
                resetQueueUI();
                return;
            }

            if(students.length === 0) {
                throw new Error("No valid student data found (Check header row).");
            }

            const lastUpdated = new Date().toLocaleString();
            
            chrome.storage.local.set({ 
                [STORAGE_KEYS.MASTER_ENTRIES]: students,
                [STORAGE_KEYS.LAST_UPDATED]: lastUpdated
            }, () => {
                const duration = ((Date.now() - startTime) / 1000).toFixed(1);
                step1.className = 'queue-item completed';
                step1.querySelector('i').className = 'fas fa-check';
                timeSpan.textContent = `${duration}s`;
                
                if(elements.lastUpdatedText) {
                    elements.lastUpdatedText.textContent = lastUpdated;
                }

                renderMasterList(students);

                // --- TRIGGER STEP 2 AUTOMATICALLY ---
                processStep2(students);
            });

        } catch (error) {
            console.error("Error parsing file:", error);
            step1.querySelector('i').className = 'fas fa-times';
            step1.style.color = '#ef4444';
            timeSpan.textContent = 'Error: ' + error.message;
        }
        
        elements.studentPopFile.value = '';
    };

    reader.readAsText(file);
}

// --- HELPER: PRELOAD IMAGE ---
function preloadImage(url) {
    if (!url) return;
    const img = new Image();
    img.src = url;
}

// --- STEP 2: FETCH CANVAS IDs, COURSES & PHOTOS ---

async function processStep2(students) {
    const step2 = document.getElementById('step2');
    const timeSpan = step2.querySelector('.step-time');
    
    step2.className = 'queue-item active';
    step2.querySelector('i').className = 'fas fa-spinner';
    
    const startTime = Date.now();

    try {
        console.log(`[Step 2] Pinging Canvas API: ${CANVAS_DOMAIN}`);

        const BATCH_SIZE = 5;
        let processedCount = 0;
        let updatedStudents = [...students];

        for (let i = 0; i < updatedStudents.length; i += BATCH_SIZE) {
            const batch = updatedStudents.slice(i, i + BATCH_SIZE);
            const promises = batch.map(student => fetchCanvasDetails(student));
            
            const results = await Promise.all(promises);
            
            results.forEach((updatedStudent, index) => {
                updatedStudents[i + index] = updatedStudent;
            });

            processedCount += batch.length;
            timeSpan.textContent = `${Math.round((processedCount / updatedStudents.length) * 100)}%`;
        }

        await chrome.storage.local.set({ [STORAGE_KEYS.MASTER_ENTRIES]: updatedStudents });
        
        const duration = ((Date.now() - startTime) / 1000).toFixed(1);
        step2.className = 'queue-item completed';
        step2.querySelector('i').className = 'fas fa-check';
        timeSpan.textContent = `${duration}s`;
        
        renderMasterList(updatedStudents);

    } catch (error) {
        console.error("[Step 2 Error]", error);
        step2.querySelector('i').className = 'fas fa-times';
        step2.style.color = '#ef4444';
        timeSpan.textContent = 'Error';
    }
}

async function fetchCanvasDetails(student) {
    if (!student.SyStudentId) return student;

    try {
        const userUrl = `${CANVAS_DOMAIN}/api/v1/users/sis_user_id:${student.SyStudentId}`;
        const userResp = await fetch(userUrl, { headers: { 'Accept': 'application/json' } });
        
        if (!userResp.ok) return student;
        const userData = await userResp.json();
        
        if (userData.name) student.name = userData.name; 
        if (userData.sortable_name) student.sortable_name = userData.sortable_name;

        if (userData.avatar_url && userData.avatar_url !== GENERIC_AVATAR_URL) {
            student.Photo = userData.avatar_url;
            preloadImage(userData.avatar_url);
        }

        if (userData.created_at) {
            student.created_at = userData.created_at;
            const createdDate = new Date(userData.created_at);
            const today = new Date();
            const timeDiff = today - createdDate;
            const daysDiff = timeDiff / (1000 * 3600 * 24);
            
            if (daysDiff < 60) {
                student.isNew = true;
            }
        }

        const canvasUserId = userData.id;

        if (canvasUserId) {
            const coursesUrl = `${CANVAS_DOMAIN}/api/v1/users/${canvasUserId}/courses?include[]=enrollments&enrollment_state=active&per_page=100`;
            const coursesResp = await fetch(coursesUrl, { headers: { 'Accept': 'application/json' } });

            if (coursesResp.ok) {
                const courses = await coursesResp.json();
                
                const now = new Date();
                const validCourses = courses.filter(c => c.name && !c.name.toUpperCase().includes('CAPV'));

                let activeCourse = null;

                activeCourse = validCourses.find(c => {
                    if (!c.start_at || !c.end_at) return false;
                    const start = new Date(c.start_at);
                    const end = new Date(c.end_at);
                    return now >= start && now <= end;
                });

                if (!activeCourse && validCourses.length > 0) {
                    validCourses.sort((a, b) => {
                        const dateA = a.start_at ? new Date(a.start_at) : new Date(0);
                        const dateB = b.start_at ? new Date(b.start_at) : new Date(0);
                        return dateB - dateA;
                    });
                    activeCourse = validCourses[0];
                }

                if (activeCourse) {
                    student.url = `${CANVAS_DOMAIN}/courses/${activeCourse.id}/grades/${canvasUserId}`;
                    
                    if (activeCourse.enrollments && activeCourse.enrollments.length > 0) {
                        const enrollment = activeCourse.enrollments.find(e => e.type === 'StudentEnrollment') || activeCourse.enrollments[0];
                        if (enrollment && enrollment.grades && enrollment.grades.current_score) {
                            student.grade = enrollment.grades.current_score + '%';
                        }
                    }
                } else {
                    student.url = `${CANVAS_DOMAIN}/users/${canvasUserId}/grades`;
                }
            }
        }
        return student;

    } catch (e) {
        return student;
    }
}

function parseCSV(csvText) {
    const lines = csvText.split(/\r\n|\n/).filter(line => line.trim() !== '');
    if (lines.length < 2) return [];

    const aliases = {
        name: ['student name', 'name', 'studentname', 'student'],
        phone: ['primaryphone', 'phone', 'phone number', 'mobile', 'cell', 'cell phone', 'contact', 'telephone'],
        grade: ['grade', 'grade level', 'level'],
        StudentNumber: ['studentnumber', 'student id', 'sis id'],
        SyStudentId: ['systudentid', 'student sis'],
        daysOut: ['days out', 'dayssincepriorlda', 'days inactive', 'days']
    };

    let headerRowIndex = -1;
    let bestHeaders = [];

    for (let i = 0; i < lines.length; i++) {
        const currentHeaders = lines[i].split(',').map(h => h.trim().replace(/^"|"$/g, '').toLowerCase());
        const hasName = currentHeaders.some(h => aliases.name.includes(h));
        if (hasName) {
            headerRowIndex = i;
            bestHeaders = currentHeaders;
            break; 
        }
    }

    if (headerRowIndex === -1) return [];

    const idx = {
        name: bestHeaders.findIndex(h => aliases.name.includes(h)),
        phone: bestHeaders.findIndex(h => aliases.phone.includes(h)),
        grade: bestHeaders.findIndex(h => aliases.grade.includes(h)),
        StudentNumber: bestHeaders.findIndex(h => aliases.StudentNumber.includes(h)),
        SyStudentId: bestHeaders.findIndex(h => aliases.SyStudentId.includes(h)),
        daysOut: bestHeaders.findIndex(h => aliases.daysOut.includes(h))
    };

    const students = [];

    for (let i = headerRowIndex + 1; i < lines.length; i++) {
        const row = lines[i].match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g);
        if (!row) continue;
        
        const cleanRow = row.map(val => val.replace(/^"|"$/g, '').trim());
        const rawName = cleanRow[idx.name];
        
        if (!rawName || /^\d+$/.test(rawName) || rawName.includes('/')) continue;

        const entry = {
            name: rawName,
            phone: idx.phone !== -1 ? cleanRow[idx.phone] : null,
            grade: idx.grade !== -1 ? cleanRow[idx.grade] : null,
            StudentNumber: idx.StudentNumber !== -1 ? cleanRow[idx.StudentNumber] : null,
            SyStudentId: idx.SyStudentId !== -1 ? cleanRow[idx.SyStudentId] : null,
            daysout: idx.daysOut !== -1 ? parseInt(cleanRow[idx.daysOut]) || 0 : 0,
            
            missingCount: 0,
            url: null,
            assignments: []
        };
        students.push(entry);
    }
    
    return students;
}

function resetQueueUI() {
    const steps = ['step1', 'step2', 'step3', 'step4'];
    steps.forEach(id => {
        const el = document.getElementById(id);
        if(!el) return;
        el.className = 'queue-item';
        el.querySelector('i').className = 'far fa-circle';
        el.querySelector('.step-time').textContent = '';
        el.style.color = '';
    });
    const totalTimeDisplay = document.getElementById('queueTotalTime');
    if(totalTimeDisplay) totalTimeDisplay.style.display = 'none';
}

// --- LOGIC: STORAGE & RENDERING ---
async function loadStorageData() {
    const data = await chrome.storage.local.get([
        STORAGE_KEYS.FOUND_ENTRIES,
        STORAGE_KEYS.MASTER_ENTRIES,
        STORAGE_KEYS.LAST_UPDATED,
        STORAGE_KEYS.EXTENSION_STATE
    ]);

    renderFoundList(data[STORAGE_KEYS.FOUND_ENTRIES] || []);
    renderMasterList(data[STORAGE_KEYS.MASTER_ENTRIES] || []);

    if (elements.lastUpdatedText && data[STORAGE_KEYS.LAST_UPDATED]) {
        elements.lastUpdatedText.textContent = data[STORAGE_KEYS.LAST_UPDATED];
    }
    
    updateButtonVisuals(data[STORAGE_KEYS.EXTENSION_STATE] || EXTENSION_STATES.OFF);
}

chrome.storage.onChanged.addListener((changes) => {
    if (changes[STORAGE_KEYS.FOUND_ENTRIES]) {
        renderFoundList(changes[STORAGE_KEYS.FOUND_ENTRIES].newValue);
        updateTabBadge('checker', (changes[STORAGE_KEYS.FOUND_ENTRIES].newValue || []).length);
    }
    if (changes[STORAGE_KEYS.MASTER_ENTRIES]) {
        renderMasterList(changes[STORAGE_KEYS.MASTER_ENTRIES].newValue);
    }
    if (changes[STORAGE_KEYS.EXTENSION_STATE]) {
        updateButtonVisuals(changes[STORAGE_KEYS.EXTENSION_STATE].newValue);
    }
});

function updateTabBadge(tabId, count) {
    const badge = document.querySelector(`.tab-button[data-tab="${tabId}"] .badge`);
    if (badge) {
        badge.textContent = count;
        badge.style.display = count > 0 ? 'flex' : 'none';
    }
}

// --- HELPER: DATA NORMALIZATION ---
function resolveStudentData(entry) {
    return {
        name: entry.name || 'Unknown Student',
        sortable_name: entry.sortable_name || null,
        phone: entry.phone || null,
        daysOut: parseInt(entry.daysout || 0),
        missing: parseInt(entry.missingCount || 0),
        StudentNumber: entry.StudentNumber || null,
        SyStudentId: entry.SyStudentId || null,
        url: entry.url || null,
        Photo: entry.Photo || null,
        isNew: entry.isNew || false, 
        created_at: entry.created_at || null, 
        timestamp: entry.timestamp || null,
        assignment: entry.assignment || null
    };
}

// --- STUDENT & CALL MANAGEMENT ---

function setActiveStudent(rawEntry) {
    const contactTab = document.getElementById('contact');
    if (!contactTab) return;

    // --- NEW: RESET AUTOMATION STYLES WHEN SWITCHING ---
    if (elements.dialBtn) {
        elements.dialBtn.classList.remove('automation');
        elements.dialBtn.innerHTML = '<i class="fas fa-phone"></i>'; 
    }
    if (elements.callStatusText) {
        elements.callStatusText.innerHTML = '<span class="status-indicator ready"></span> Ready to Connect';
    }
    // ---------------------------------------------------

    // 1. Handle "No Student Selected" State
    if (!rawEntry) {
        Array.from(contactTab.children).forEach(child => {
            if (child.id === 'contactPlaceholder') {
                child.style.display = 'flex';
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
        } else {
            child.style.display = ''; 
        }
    });

    const data = resolveStudentData(rawEntry);

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
        elements.contactAvatar.style.color = ''; // Reset potential automation color
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
        elements.contactDetail.textContent = `${data.daysOut} Days Inactive`;
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

// --- NEW: MULTI-SELECT HELPERS ---
function toggleMultiSelection(entry, liElement) {
    const index = selectedQueue.findIndex(s => s.name === entry.name);
    
    if (index > -1) {
        // Deselect
        selectedQueue.splice(index, 1);
        liElement.classList.remove('multi-selected');
    } else {
        // Select
        selectedQueue.push(entry);
        liElement.classList.add('multi-selected');
    }

    // Update UI based on queue size
    if (selectedQueue.length === 1) {
        setActiveStudent(selectedQueue[0]); // Revert to single view
    } else if (selectedQueue.length > 1) {
        setAutomationModeUI(); // Switch to Automation View
    } else {
        setActiveStudent(null); // Clear view
    }
}

function setAutomationModeUI() {
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

    // 1. Update Contact Card (Placeholder for Batch)
    if (elements.contactName) elements.contactName.textContent = "Automation Mode";
    if (elements.contactDetail) elements.contactDetail.textContent = `${selectedQueue.length} Students Selected`;
    if (elements.contactPhone) elements.contactPhone.textContent = "Multi-Dial Queue";
    
    // Create/Update visual badge for count
    if (elements.contactAvatar) {
        elements.contactAvatar.textContent = selectedQueue.length;
        elements.contactAvatar.style.backgroundImage = 'none';
        elements.contactAvatar.style.backgroundColor = '#8b5cf6'; // Purple
        elements.contactAvatar.style.color = '#ffffff';
    }

    // 2. Transform the Dial Button to Purple
    if (elements.dialBtn) {
        elements.dialBtn.classList.add('automation');
        elements.dialBtn.innerHTML = '<i class="fas fa-robot"></i>'; // Change icon to Robot
    }

    // 3. Update Status Text
    if (elements.callStatusText) {
        elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:#8b5cf6;"></span> Ready to Auto-Dial`;
    }
    
    if (elements.contactCard) {
        elements.contactCard.style.borderLeftColor = '#8b5cf6';
    }
    
    // REMOVED: switchTab('contact'); - Now handled by keyup event
}

// --- RENDERING LISTS ---

function renderFoundList(rawEntries) {
    if (!elements.foundList) return;
    elements.foundList.innerHTML = '';

    if (!rawEntries || rawEntries.length === 0) {
        elements.foundList.innerHTML = '<li style="justify-content:center; color:gray;">No submissions found yet.</li>';
        return;
    }

    const entries = rawEntries.map(resolveStudentData);
    entries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    entries.forEach(data => {
        const li = document.createElement('li');
        let timeDisplay = 'Just now';
        if (data.timestamp) {
            timeDisplay = new Date(data.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        }

        const assignmentTitle = data.assignment || 'Untitled Assignment';

        li.innerHTML = `
            <div style="display: flex; align-items: center; width:100%;">
                <div class="heatmap-indicator heatmap-green"></div>
                <div style="flex-grow:1; display:flex; justify-content:space-between; align-items:center;">
                    <div style="display:flex; flex-direction:column;">
                        <span class="student-name" style="font-weight:500; color:var(--primary-color); cursor:pointer;">${data.name}</span>
                        <span style="font-size:0.8em; color:var(--text-secondary);">${assignmentTitle}</span>
                    </div>
                    <span class="timestamp-pill">${timeDisplay}</span>
                </div>
            </div>
        `;
        
        const nameLink = li.querySelector('.student-name');
        nameLink.addEventListener('click', (e) => {
             e.stopPropagation();
             if(data.url) chrome.tabs.create({ url: data.url });
        });
        nameLink.addEventListener('mouseenter', () => nameLink.style.textDecoration = 'underline');
        nameLink.addEventListener('mouseleave', () => nameLink.style.textDecoration = 'none');

        elements.foundList.appendChild(li);
    });
}

function filterFoundList(e) {
    const term = e.target.value.toLowerCase();
    const items = elements.foundList.querySelectorAll('li');
    items.forEach(li => {
        const text = li.textContent.toLowerCase();
        li.style.display = text.includes(term) ? 'flex' : 'none';
    });
}

function renderMasterList(rawEntries) {
    if (!elements.masterList) return;
    elements.masterList.innerHTML = '';

    if (!rawEntries || rawEntries.length === 0) {
        elements.masterList.innerHTML = '<li style="justify-content:center;">Master list is empty.</li>';
        return;
    }

    rawEntries.forEach(rawEntry => {
        const data = resolveStudentData(rawEntry);
        
        const li = document.createElement('li');
        li.className = 'expandable';
        li.style.cursor = 'pointer'; 
        
        li.setAttribute('data-name', data.name);
        li.setAttribute('data-missing', data.missing);
        li.setAttribute('data-days', data.daysOut);
        li.setAttribute('data-created', data.created_at || '');

        let heatmapClass = data.daysOut > 10 ? 'heatmap-red' : (data.daysOut > 5 ? 'heatmap-orange' : (data.daysOut > 2 ? 'heatmap-yellow' : 'heatmap-green'));

        let missingPillHtml = '';
        if(data.missing > 0) {
            missingPillHtml = `<span class="missing-pill">${data.missing} Missing <i class="fas fa-chevron-down" style="font-size:0.8em; margin-left:4px;"></i></span>`;
        }

        let newTagHtml = '';
        if(data.isNew) {
            newTagHtml = `<span style="background:#e0f2fe; color:#0369a1; font-size:0.7em; padding:2px 6px; border-radius:8px; margin-left:6px; font-weight:bold; border:1px solid #bae6fd;">New</span>`;
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
                    <li><em>Details not loaded.</em></li>
                </ul>
            </div>
        `;

        // --- UPDATED CLICK LISTENER FOR MULTI-SELECT ---
        li.addEventListener('click', (e) => {
            // Check for CTRL (Windows) or COMMAND (Mac)
            if (e.ctrlKey || e.metaKey) {
                toggleMultiSelection(rawEntry, li);
            } else {
                // Standard Single Select (Clears previous multi-selection)
                selectedQueue = [rawEntry]; 
                
                // Visually clear other rows
                document.querySelectorAll('.glass-list li').forEach(el => el.classList.remove('multi-selected'));
                li.classList.add('multi-selected');
                
                setActiveStudent(rawEntry);
                switchTab('contact');
            }
        });

        const nameLink = li.querySelector('.student-name');
        if(nameLink) {
            nameLink.addEventListener('click', (e) => {
                e.stopPropagation(); 
                if(data.url) chrome.tabs.create({ url: data.url });
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

        const pill = li.querySelector('.missing-pill');
        if(pill) {
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

        const detailsDiv = li.querySelector('.missing-details');
        if(detailsDiv) {
            detailsDiv.addEventListener('click', (e) => {
                e.stopPropagation();
            });
        }

        elements.masterList.appendChild(li);
    });
}

// --- LOGIC: TABS & SCANNER ---
function switchTab(targetId) {
    elements.tabs.forEach(t => t.classList.remove('active'));
    elements.contents.forEach(c => c.classList.remove('active'));
    
    const targetContent = document.getElementById(targetId);
    if (targetContent) targetContent.classList.add('active');
    
    const targetTab = document.querySelector(`.tab-button[data-tab="${targetId}"]`);
    if (targetTab) targetTab.classList.add('active');
}

function toggleScanState() {
    isScanning = !isScanning;
    const newState = isScanning ? EXTENSION_STATES.ON : EXTENSION_STATES.OFF;
    chrome.storage.local.set({ [STORAGE_KEYS.EXTENSION_STATE]: newState });
}

function updateButtonVisuals(state) {
    if (!elements.startBtn) return;
    isScanning = (state === EXTENSION_STATES.ON);

    if (isScanning) {
        elements.startBtn.style.background = '#ef4444';
        elements.startBtnText.textContent = 'Stop';
        elements.startBtnIcon.className = 'fas fa-stop';
        elements.statusDot.style.background = '#10b981';
        elements.statusDot.style.animation = 'pulse 2s infinite';
        elements.statusText.textContent = 'Monitoring...';
    } else {
        elements.startBtn.style.background = 'rgba(0, 90, 156, 0.7)';
        elements.startBtnText.textContent = 'Start';
        elements.startBtnIcon.className = 'fas fa-play';
        elements.statusDot.style.background = '#cbd5e1';
        elements.statusDot.style.animation = 'none';
        elements.statusText.textContent = 'Ready to Scan';
    }
}

// --- LOGIC: CALL INTERFACE ---
function toggleCallState(forceEnd = false) {
    // --- NEW: CHECK FOR AUTOMATION MODE ---
    if (selectedQueue.length > 1 && !isCallActive) {
        startAutomationSequence();
        return;
    }
    // --------------------------------------

    if (forceEnd && !isCallActive) return;
    isCallActive = !isCallActive;
    if (forceEnd) isCallActive = false;

    if (isCallActive) {
        elements.dialBtn.style.background = '#ef4444'; 
        elements.dialBtn.style.transform = 'rotate(135deg)';
        elements.callStatusText.innerHTML = '<span class="status-indicator" style="background:#ef4444; animation: blink 1s infinite;"></span> Connected';
        startCallTimer();
    } else {
        elements.dialBtn.style.background = '#10b981';
        elements.dialBtn.style.transform = 'rotate(0deg)';
        elements.callStatusText.innerHTML = '<span class="status-indicator ready"></span> Ready to Connect';
        stopCallTimer();
    }
}

function startAutomationSequence() {
    alert(`Starting automation for ${selectedQueue.length} students...\n(Logic to be implemented)`);
    // Placeholder for future logic
}

function startCallTimer() {
    let seconds = 0;
    elements.callTimer.textContent = "00:00";
    clearInterval(callTimerInterval);
    callTimerInterval = setInterval(() => {
        seconds++;
        const m = Math.floor(seconds / 60).toString().padStart(2, '0');
        const s = (seconds % 60).toString().padStart(2, '0');
        elements.callTimer.textContent = `${m}:${s}`;
    }, 1000);
}

function stopCallTimer() {
    clearInterval(callTimerInterval);
    elements.callTimer.textContent = "00:00";
}

function handleDisposition(type) {
    console.log("Logged Disposition:", type);
    toggleCallState(true);
}

// --- LOGIC: QUEUE SIMULATION ---
function runQueueSimulation() {
    // Only runs for later steps now
}

// --- LOGIC: FILTER & SORT ---
function filterMasterList(e) {
    const term = e.target.value.toLowerCase();
    const listItems = elements.masterList.querySelectorAll('li.expandable');
    listItems.forEach(li => {
        const name = li.getAttribute('data-name').toLowerCase();
        li.style.display = name.includes(term) ? 'flex' : 'none';
    });
}

function sortMasterList() {
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

function exportMasterListCSV() {
    const listItems = elements.masterList.querySelectorAll('li.expandable');
    const rows = [["Student Name", "Missing Assignments", "Days Out"]];
    listItems.forEach(li => {
        rows.push([
            `"${li.getAttribute('data-name')}"`,
            li.getAttribute('data-missing'),
            li.getAttribute('data-days')
        ]);
    });
    const csvContent = "data:text/csv;charset=utf-8," + rows.map(e => e.join(",")).join("\n");
    const link = document.createElement("a");
    link.setAttribute("href", encodeURI(csvContent));
    link.setAttribute("download", "master_student_list.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}