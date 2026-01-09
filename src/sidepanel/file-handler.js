// File Handler - CSV/Excel import and export functionality
import {
    STORAGE_KEYS,
    FIELD_ALIASES,
    MASTER_LIST_COLUMNS,
    EXPORT_MISSING_ASSIGNMENTS_COLUMNS,
    LDA_VISIBLE_COLUMNS,
    normalizeFieldName,
    convertExcelDate,
    isExcelDateNumber,
    calculateDaysSinceLastAttendance,
    parseDate,
    formatDateToMMDDYY
} from '../constants/index.js';
import { elements } from './ui-manager.js';

/**
 * Sends master list data to Excel via SRK_IMPORT_MASTER_LIST payload
 * @param {Array} students - Array of student objects
 */
export async function sendMasterListToExcel(students) {
    if (!students || students.length === 0) {
        console.log('No students to send to Excel');
        return;
    }

    // Check if sending master list to Excel is enabled
    const settings = await chrome.storage.local.get([STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL]);
    const isEnabled = settings[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL] !== undefined
        ? settings[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL]
        : true; // Default to enabled

    if (!isEnabled) {
        console.log('Send master list to Excel is disabled - skipping');
        return;
    }

    try {
        // Extract headers from MASTER_LIST_COLUMNS
        const headers = MASTER_LIST_COLUMNS.map(col => col.header);

        // Transform students into data rows using MASTER_LIST_COLUMNS definitions
        const data = students.map(student => {
            return MASTER_LIST_COLUMNS.map(col => {
                // Use getFieldValue which now uses alias-based matching
                let value = getFieldValue(student, col.field, col.fallback);

                // Format LDA dates to MM-DD-YY format
                if (col.field === 'lda') {
                    console.log(`[LDA To Excel Debug] Student: ${student.name}, Retrieved LDA:`, value, 'student.lda:', student.lda);
                    if (value) {
                        const dateObj = parseDate(value);
                        if (dateObj) {
                            value = formatDateToMMDDYY(dateObj);
                        }
                    }
                }

                // Return value or empty string
                return value !== null && value !== undefined ? value : '';
            });
        });

        // Create the payload
        const payload = {
            type: "SRK_IMPORT_MASTER_LIST",
            data: {
                headers: headers,
                data: data
            }
        };

        // Send message to background script to forward to Excel
        chrome.runtime.sendMessage({
            type: "SRK_SEND_IMPORT_MASTER_LIST",
            payload: payload
        }).catch(() => {
            console.log('Background script might not be ready');
        });

        console.log(`Sent ${students.length} students to Excel for import with ${headers.length} columns`);
    } catch (error) {
        console.error('Error sending master list to Excel:', error);
    }
}

/**
 * Sends master list data with missing assignments to Excel via SRK_IMPORT_MASTER_LIST payload
 * This function includes both the traditional headers/data arrays AND the students array with missingAssignments
 * @param {Array} students - Array of student objects with missingAssignments data
 */
export async function sendMasterListWithMissingAssignmentsToExcel(students) {
    if (!students || students.length === 0) {
        console.log('No students to send to Excel');
        return;
    }

    // Check if sending master list to Excel is enabled
    const settings = await chrome.storage.local.get([STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL]);
    const isEnabled = settings[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL] !== undefined
        ? settings[STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL]
        : true; // Default to enabled

    if (!isEnabled) {
        console.log('Send master list to Excel is disabled - skipping');
        return;
    }

    try {
        // Extract headers from MASTER_LIST_COLUMNS
        const headers = MASTER_LIST_COLUMNS.map(col => col.header);

        // Transform students into data rows using MASTER_LIST_COLUMNS definitions
        const data = students.map(student => {
            return MASTER_LIST_COLUMNS.map(col => {
                // Use getFieldValue which now uses alias-based matching
                let value = getFieldValue(student, col.field, col.fallback);

                // Format LDA dates to MM-DD-YY format
                if (col.field === 'lda') {
                    console.log(`[LDA To Excel Debug] Student: ${student.name}, Retrieved LDA:`, value, 'student.lda:', student.lda);
                    if (value) {
                        const dateObj = parseDate(value);
                        if (dateObj) {
                            value = formatDateToMMDDYY(dateObj);
                        }
                    }
                }

                // Return value or empty string
                return value !== null && value !== undefined ? value : '';
            });
        });

        // Create students array with missing assignments
        const studentsWithMissingAssignments = students
            .filter(student => student.missingAssignments && student.missingAssignments.length > 0)
            .map(student => {
                // Format student data for the students array
                const studentData = {
                    "Student Name": getFieldValue(student, 'name'),
                    "Grade": getFieldValue(student, 'currentGrade', 'grade'),
                    "Grade Book": getFieldValue(student, 'url')
                };

                // Format missing assignments according to MissingAssignmentImport interface
                // Handle both 'title' and 'assignmentTitle' for robustness
                studentData.missingAssignments = student.missingAssignments.map(assignment => ({
                    assignmentTitle: assignment.assignmentTitle || assignment.title || '',
                    assignmentLink: assignment.assignmentLink || assignment.assignmentUrl || '',
                    submissionLink: assignment.submissionLink || assignment.submissionUrl || '',
                    dueDate: assignment.dueDate || '',
                    score: assignment.score || ''
                }));

                return studentData;
            });

        // Create the payload with both traditional format and students array
        const payload = {
            type: "SRK_IMPORT_MASTER_LIST",
            data: {
                headers: headers,
                data: data,
                students: studentsWithMissingAssignments
            }
        };

        // Send message to background script to forward to Excel
        chrome.runtime.sendMessage({
            type: "SRK_SEND_IMPORT_MASTER_LIST",
            payload: payload
        }).catch(() => {
            console.log('Background script might not be ready');
        });

        const totalMissingAssignments = studentsWithMissingAssignments.reduce(
            (sum, s) => sum + (s.missingAssignments?.length || 0),
            0
        );

        console.log(`Sent ${students.length} students to Excel for import with ${headers.length} columns`);
        console.log(`Including ${studentsWithMissingAssignments.length} students with ${totalMissingAssignments} total missing assignments`);
    } catch (error) {
        console.error('Error sending master list with missing assignments to Excel:', error);
    }
}

/**
 * Validates if a string is a valid student name
 */
function isValidStudentName(name) {
    if (!name) return false;
    if (/^\d+$/.test(name)) return false;  // All digits
    if (name.includes('/')) return false;   // Contains date-like patterns
    return true;
}

/**
 * Checks if a field name indicates it should contain a date.
 * Used to determine which fields to apply Excel date conversion to.
 *
 * @param {string} fieldName - The field name to check
 * @returns {boolean} True if field appears to be a date field
 */
function isDateField(fieldName) {
    if (!fieldName) return false;
    const normalized = normalizeFieldName(fieldName);
    // Check if field name contains date-related keywords
    return normalized.includes('date') ||
           normalized.includes('lda') ||
           normalized === 'expstartdate' ||
           normalized === 'lastlda';
}

/**
 * Cleans up program version strings by removing year prefix and everything before it.
 * Examples:
 * - "D2-Q-2024 Medical Billing and Coding Specialist" → "Medical Billing and Coding Specialist"
 * - "A1-2025 Nursing Program" → "Nursing Program"
 * - "Medical Billing" → "Medical Billing" (unchanged if no year found)
 *
 * @param {string} value - The program version string to clean
 * @returns {string} The cleaned program version string
 */
function cleanProgramVersion(value) {
    if (!value || typeof value !== 'string') return value;

    // Match a 4-digit year (19xx or 20xx) and remove everything up to and including it
    // Also removes any trailing separator (space, dash, etc.) after the year
    const yearPattern = /.*\b(19\d{2}|20\d{2})\b[\s\-]*/;
    const cleaned = value.replace(yearPattern, '');

    // Return trimmed result, or original if no year was found
    return cleaned.trim() || value.trim();
}

/**
 * Cleans up phone number strings by removing trailing spaces and anything after the phone number.
 * Examples:
 * - "555-1234 " → "555-1234"
 * - "555-1234 ext 123" → "555-1234"
 * - "(555) 123-4567 work" → "(555) 123-4567"
 * - "+1-555-123-4567 mobile" → "+1-555-123-4567"
 *
 * @param {string|number} value - The phone number string to clean
 * @returns {string} The cleaned phone number string
 */
function cleanPhoneNumber(value) {
    if (!value) return value;

    // Convert to string and trim
    let cleaned = String(value).trim();

    // Extract only the phone number part (digits, hyphens, parentheses, spaces, dots, plus signs)
    // This regex captures the phone number portion and stops at any trailing text/spaces
    const phonePattern = /^([+\d\s\-().]+?)(?:\s+[a-zA-Z].*)?$/;
    const match = cleaned.match(phonePattern);

    if (match) {
        // Return the captured phone number, trimmed
        return match[1].trim();
    }

    // If no match, just return trimmed value
    return cleaned;
}

/**
 * Finds the column index for a field using normalized matching and aliases.
 * Matches fields case-insensitively and ignoring spaces/special characters.
 *
 * @param {Array} headers - Array of header names from the file
 * @param {String} fieldName - The internal field name to match
 * @returns {Number} The column index, or -1 if not found
 */
function findColumnIndex(headers, fieldName) {
    // Normalize the target field name
    const normalizedFieldName = normalizeFieldName(fieldName);

    // Get aliases for this field (if any)
    const aliases = FIELD_ALIASES[fieldName] || [];

    // Normalize all aliases
    const normalizedAliases = aliases.map(alias => normalizeFieldName(alias));

    // Find the header that matches either the field name or one of its aliases
    return headers.findIndex(header => {
        const normalizedHeader = normalizeFieldName(header);

        // Check direct match
        if (normalizedHeader === normalizedFieldName) {
            return true;
        }

        // Check alias matches
        return normalizedAliases.includes(normalizedHeader);
    });
}

/**
 * Unified parser for both CSV and Excel files using SheetJS
 * @param {String|ArrayBuffer} data - File content
 * @param {Boolean} isCSV - True if parsing CSV, false for Excel
 * @param {Number} fileModifiedTime - File last modified timestamp (milliseconds since epoch)
 * @returns {Object} Object containing { students: Array, referenceDate: Date }
 */
export function parseFileWithSheetJS(data, isCSV, fileModifiedTime = null) {
    try {
        if (typeof XLSX === 'undefined') {
            throw new Error('XLSX library not loaded. Please refresh the page.');
        }

        let workbook;
        if (isCSV) {
            workbook = XLSX.read(data, { type: 'string' });
        } else {
            workbook = XLSX.read(data, { type: 'array', cellDates: true });
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        // Determine reference date for daysOut calculation
        // For Excel files, try to get creation date from workbook properties
        // Otherwise use file modified time or current date
        let referenceDate = null;
        if (!isCSV && workbook.Props && workbook.Props.CreatedDate) {
            referenceDate = new Date(workbook.Props.CreatedDate);
        } else if (fileModifiedTime) {
            referenceDate = new Date(fileModifiedTime);
        } else {
            referenceDate = new Date(); // Fallback to current date
        }

        if (rows.length < 2) {
            return { students: [], referenceDate: null };
        }

        // Find header row
        let headerRowIndex = -1;
        let headers = [];

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            // Check if this row contains a name field by trying to find a name column
            const hasNameField = row.some(cell => {
                if (!cell) return false;
                const normalized = normalizeFieldName(cell);
                const nameNormalized = normalizeFieldName('name');
                const nameAliases = (FIELD_ALIASES.name || []).map(a => normalizeFieldName(a));
                return normalized === nameNormalized || nameAliases.includes(normalized);
            });

            if (hasNameField) {
                headerRowIndex = i;
                headers = row;
                break;
            }
        }

        if (headerRowIndex === -1) {
            return { students: [], referenceDate: null };
        }

        // Map column indices for all MASTER_LIST_COLUMNS using normalized field matching
        const columnMapping = {};
        MASTER_LIST_COLUMNS.forEach(col => {
            const index = findColumnIndex(headers, col.field);
            if (index !== -1) {
                columnMapping[col.field] = index;
            }
            // Also check fallback field if it exists
            if (col.fallback && index === -1) {
                const fallbackIndex = findColumnIndex(headers, col.fallback);
                if (fallbackIndex !== -1) {
                    columnMapping[col.field] = fallbackIndex;
                }
            }
        });

        // Ensure we have at least a name column
        if (!columnMapping.name) {
            return { students: [], referenceDate: null };
        }

        // Parse data rows
        const students = [];
        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const studentName = row[columnMapping.name];
            if (!isValidStudentName(studentName)) continue;

            // Create entry with ONLY whitelisted columns from MASTER_LIST_COLUMNS
            const entry = {};

            // Map each MASTER_LIST_COLUMNS field to the imported data
            MASTER_LIST_COLUMNS.forEach(col => {
                const colIndex = columnMapping[col.field];
                if (colIndex !== undefined && colIndex < row.length) {
                    let value = row[colIndex];

                    // Special validation for grade field - only accept integers
                    if (col.field === 'grade' && value !== null && value !== undefined && value !== '') {
                        const gradeNum = Number(value);
                        if (isNaN(gradeNum) || !Number.isInteger(gradeNum)) {
                            console.log(`Student ${studentName}: Skipping invalid grade '${value}' (not an integer)`);
                            value = null; // Don't import invalid grade, but still import student
                        }
                    }

                    // Apply Excel date conversion to date fields
                    if (isDateField(col.field)) {
                        value = convertExcelDate(value);
                    }

                    // Clean program version by removing year prefix
                    if (col.field === 'programVersion' && value !== null && value !== undefined && value !== '') {
                        value = cleanProgramVersion(value);
                    }

                    // Clean phone numbers by removing trailing spaces and text
                    if ((col.field === 'phone' || col.field === 'otherPhone') && value !== null && value !== undefined && value !== '') {
                        value = cleanPhoneNumber(value);
                    }

                    // Debug logging for LDA
                    if (col.field === 'lda') {
                        console.log(`[LDA Import Debug] Student: ${studentName}, Raw LDA value:`, row[colIndex], 'Converted value:', value);
                    }

                    // Use the field name from MASTER_LIST_COLUMNS definition
                    entry[col.field] = value !== null && value !== undefined ? value : null;
                }
            });

            // Ensure critical fields are present with proper types
            entry.name = String(studentName);
            if (entry.phone) entry.phone = cleanPhoneNumber(String(entry.phone));
            if (entry.otherPhone) entry.otherPhone = cleanPhoneNumber(String(entry.otherPhone));
            if (entry.grade) entry.grade = String(entry.grade);
            if (entry.StudentNumber) entry.StudentNumber = String(entry.StudentNumber);
            if (entry.SyStudentId) entry.SyStudentId = String(entry.SyStudentId);

            // Calculate daysOut based on LDA if available, otherwise use imported value
            const ldaValue = entry.lda;
            console.log(`[LDA Storage Debug] Student: ${studentName}, entry.lda:`, entry.lda, 'entry.daysOut:', entry.daysOut);
            if (ldaValue && referenceDate) {
                entry.daysOut = calculateDaysSinceLastAttendance(ldaValue, referenceDate);
            } else {
                // Fall back to imported daysOut value if LDA is not available
                entry.daysOut = parseInt(entry.daysOut) || 0;
            }

            // Initialize fields required by the extension
            entry.missingCount = 0;
            if (!entry.url) entry.url = null; // Only set to null if not imported from file
            entry.assignments = [];

            students.push(entry);
        }

        return { students, referenceDate };

    } catch (error) {
        console.error('Error parsing file with SheetJS:', error);
        throw new Error(`File parsing failed: ${error.message}`);
    }
}

/**
 * Handles CSV/Excel file import
 * @param {File} file - The uploaded file
 * @param {Function} onSuccess - Callback after successful import
 */
export function handleFileImport(file, onSuccess) {
    if (!file) {
        resetQueueUI();
        return;
    }

    const step1 = document.getElementById('step1');
    const timeSpan = step1.querySelector('.step-time');
    const startTime = Date.now();

    // Store the overall process start time for total time calculation
    const queueTotalTimeDiv = document.getElementById('queueTotalTime');
    if (queueTotalTimeDiv) {
        queueTotalTimeDiv.dataset.processStartTime = startTime;
        queueTotalTimeDiv.style.display = 'none'; // Hide until complete
    }

    const isCSV = file.name.toLowerCase().endsWith('.csv');
    const isXLSX = file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls');

    if (!isCSV && !isXLSX) {
        alert("Unsupported file type. Please use .csv or .xlsx files.");
        resetQueueUI();
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const content = e.target.result;
        let students = [];
        let referenceDate = null;

        try {
            // Pass file last modified time for daysOut calculation
            const result = parseFileWithSheetJS(content, isCSV, file.lastModified);
            students = result.students;
            referenceDate = result.referenceDate;

            if (students.length === 0) {
                throw new Error("No valid student data found (Check header row).");
            }

            const lastUpdated = new Date().toLocaleString('en-US', {
                year: 'numeric',
                month: 'numeric',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });

            chrome.storage.local.set({
                [STORAGE_KEYS.MASTER_ENTRIES]: students,
                [STORAGE_KEYS.LAST_UPDATED]: lastUpdated,
                [STORAGE_KEYS.REFERENCE_DATE]: referenceDate ? referenceDate.toISOString() : null
            }, async () => {
                const duration = ((Date.now() - startTime) / 1000).toFixed(1);
                step1.className = 'queue-item completed';
                step1.querySelector('i').className = 'fas fa-check';
                timeSpan.textContent = `${duration}s`;

                if (elements.lastUpdatedText) {
                    elements.lastUpdatedText.textContent = lastUpdated;
                }

                // Send master list to Excel
                await sendMasterListToExcel(students);

                if (onSuccess) {
                    onSuccess(students);
                }
            });

        } catch (error) {
            console.error("Error parsing file:", error);
            step1.querySelector('i').className = 'fas fa-times';
            step1.style.color = '#ef4444';
            timeSpan.textContent = 'Error: ' + error.message;
        }

        elements.studentPopFile.value = '';
    };

    if (isCSV) {
        reader.readAsText(file);
    } else if (isXLSX) {
        reader.readAsArrayBuffer(file);
    }
}

/**
 * Resets the queue UI to default state
 */
export function resetQueueUI() {
    const steps = ['step1', 'step2', 'step3', 'step4'];
    steps.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.className = 'queue-item';
        el.querySelector('i').className = 'far fa-circle';
        el.querySelector('.step-time').textContent = '';
        el.style.color = '';
    });
    const totalTimeDisplay = document.getElementById('queueTotalTime');
    if (totalTimeDisplay) {
        totalTimeDisplay.style.display = 'none';
        totalTimeDisplay.textContent = 'Total Time: 0.0s';
        delete totalTimeDisplay.dataset.processStartTime;
    }
}

/**
 * Restores default queue UI for CSV import
 */
export function restoreDefaultQueueUI() {
    const s1 = document.getElementById('step1');
    const s2 = document.getElementById('step2');
    const s3 = document.getElementById('step3');
    const s4 = document.getElementById('step4');

    if (s1) {
        s1.style.display = '';
        s1.querySelector('.queue-content').innerHTML = '<i class="far fa-circle"></i> Student Population Report';
    }
    if (s2) { s2.style.display = ''; }
    if (s3) { s3.style.display = ''; }
    if (s4) {
        s4.style.display = '';
        s4.querySelector('.queue-content').innerHTML = '<i class="far fa-circle"></i> Sending List to Excel';
    }
}

/**
 * Helper function to get nested property value from an object
 */
function getNestedValue(obj, path) {
    return path.split('.').reduce((current, prop) => current?.[prop], obj);
}

/**
 * Finds a field value in an object using normalized matching and aliases.
 * Matches fields case-insensitively and ignoring spaces/special characters.
 *
 * @param {Object} obj - The object to search in
 * @param {String} fieldName - The internal field name
 * @param {*} defaultValue - Default value if field not found
 * @returns {*} The field value or defaultValue
 */
function getFieldWithAlias(obj, fieldName, defaultValue = null) {
    if (!obj || !fieldName) return defaultValue;

    // First try direct access (for exact matches)
    if (fieldName in obj) {
        const value = obj[fieldName];
        return value !== null && value !== undefined ? value : defaultValue;
    }

    // Normalize the target field name
    const normalizedFieldName = normalizeFieldName(fieldName);

    // Get aliases for this field
    const aliases = FIELD_ALIASES[fieldName] || [];
    const normalizedAliases = aliases.map(alias => normalizeFieldName(alias));

    // Search through object keys
    for (const key in obj) {
        const normalizedKey = normalizeFieldName(key);

        // Check direct match
        if (normalizedKey === normalizedFieldName) {
            const value = obj[key];
            return value !== null && value !== undefined ? value : defaultValue;
        }

        // Check alias matches
        if (normalizedAliases.includes(normalizedKey)) {
            const value = obj[key];
            return value !== null && value !== undefined ? value : defaultValue;
        }
    }

    return defaultValue;
}

/**
 * Helper function to get field value with fallback support
 * Now uses alias-based matching for better field resolution
 */
function getFieldValue(obj, field, fallback) {
    let value = getFieldWithAlias(obj, field);
    if ((value === null || value === undefined || value === '') && fallback) {
        value = getFieldWithAlias(obj, fallback);
    }
    return value !== null && value !== undefined ? value : '';
}

/**
 * Applies conditional formatting to grade cells based on value
 * GREEN >= 70, YELLOW 60-69, RED < 60
 * @param {Object} worksheet - The worksheet object
 * @param {number} colIndex - Column index for grades
 * @param {number} startRow - Starting row (usually 2, after header)
 * @param {number} endRow - Ending row
 */
function applyGradeConditionalFormatting(worksheet, colIndex, startRow, endRow) {
    for (let row = startRow; row <= endRow; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: colIndex }); // -1 because XLSX is 0-indexed
        const cell = worksheet[cellAddress];

        if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
            const grade = parseFloat(cell.v);
            if (!isNaN(grade)) {
                // Initialize cell style if not exists
                if (!cell.s) cell.s = {};

                if (grade >= 70) {
                    // GREEN
                    cell.s = {
                        fill: { fgColor: { rgb: '92D050' } },
                        font: { color: { rgb: '000000' } }
                    };
                } else if (grade >= 60) {
                    // YELLOW
                    cell.s = {
                        fill: { fgColor: { rgb: 'FFFF00' } },
                        font: { color: { rgb: '000000' } }
                    };
                } else {
                    // RED
                    cell.s = {
                        fill: { fgColor: { rgb: 'FF0000' } },
                        font: { color: { rgb: 'FFFFFF' } }
                    };
                }
            }
        }
    }
}

/**
 * Calculates column widths based on custom width or auto-fit
 * @param {Array} columns - Column configuration array (e.g., MASTER_LIST_COLUMNS)
 * @param {Array} data - Data rows (including header row)
 * @returns {Array} Array of column width objects { wch: number }
 */
function calculateColumnWidths(columns, data) {
    const colWidths = [];

    for (let i = 0; i < columns.length; i++) {
        const col = columns[i];

        // Use custom width if specified
        if (col.width !== undefined && col.width !== null) {
            colWidths.push({ wch: col.width });
        } else {
            // Auto-fit based on content
            let maxWidth = col.header.length;
            for (let j = 1; j < data.length; j++) {
                const cellValue = data[j][i];
                if (cellValue) {
                    const cellLength = String(cellValue).length;
                    maxWidth = Math.max(maxWidth, cellLength);
                }
            }
            // Add some padding and cap at reasonable max
            colWidths.push({ wch: Math.min(maxWidth + 2, 50) });
        }
    }

    return colWidths;
}

/**
 * Exports master list to Excel file with three sheets
 */
export async function exportMasterListCSV() {
    try {
        const result = await chrome.storage.local.get([
            STORAGE_KEYS.MASTER_ENTRIES,
            STORAGE_KEYS.REFERENCE_DATE
        ]);
        const students = result[STORAGE_KEYS.MASTER_ENTRIES] || [];
        const referenceDateStr = result[STORAGE_KEYS.REFERENCE_DATE];
        const referenceDate = referenceDateStr ? new Date(referenceDateStr) : new Date();

        if (students.length === 0) {
            alert('No data to export. Please update the master list first.');
            return;
        }

        // --- SHEET 1: MASTER LIST ---
        const masterListHeaders = MASTER_LIST_COLUMNS.map(col => col.header);
        const masterListData = [masterListHeaders];
        const masterListHyperlinkMetadata = []; // Store hyperlink info for each row

        students.forEach(student => {
            const rowMetadata = {}; // Store hyperlink info for this row
            const row = MASTER_LIST_COLUMNS.map((col, colIndex) => {
                let value = getFieldValue(student, col.field, col.fallback);

                if (col.field === 'missingCount') {
                    value = value || 0;
                } else if (col.field === 'daysOut') {
                    value = parseInt(value || 0);
                } else if (col.field === 'lda') {
                    console.log(`[LDA Export Debug] Student: ${student.name}, Retrieved LDA:`, value, 'student.lda:', student.lda);
                    if (value) {
                        // Format LDA dates to MM-DD-YY format
                        const dateObj = parseDate(value);
                        if (dateObj) {
                            value = formatDateToMMDDYY(dateObj);
                        }
                    }
                }

                // Store hyperlink metadata for this column
                if (col.hyperlink && col.hyperlinkText && value) {
                    rowMetadata[colIndex] = { url: value, text: col.hyperlinkText };
                }

                return value;
            });
            masterListData.push(row);
            masterListHyperlinkMetadata.push(rowMetadata);
        });

        // --- SHEET 2: MISSING ASSIGNMENTS ---
        const missingAssignmentsHeaders = EXPORT_MISSING_ASSIGNMENTS_COLUMNS.map(col => col.header);
        const missingAssignmentsData = [missingAssignmentsHeaders];
        const hyperlinkMetadata = []; // Store hyperlink info for each row

        students.forEach(student => {
            if (student.missingAssignments && student.missingAssignments.length > 0) {
                student.missingAssignments.forEach(assignment => {
                    // Normalize assignment fields for export
                    // Handle both old (title, assignmentUrl, submissionUrl) and new (assignmentTitle, assignmentLink, submissionLink) field names
                    const normalizedAssignment = {
                        ...assignment,
                        assignmentTitle: assignment.assignmentTitle || assignment.title || '',
                        assignmentLink: assignment.assignmentLink || assignment.assignmentUrl || '',
                        submissionLink: assignment.submissionLink || assignment.submissionUrl || ''
                    };

                    // If assignmentLink is not set but submissionLink is, derive it
                    if (!normalizedAssignment.assignmentLink && normalizedAssignment.submissionLink) {
                        normalizedAssignment.assignmentLink = normalizedAssignment.submissionLink.replace(/\/submissions\/.*$/, '');
                    }

                    const rowMetadata = {}; // Store hyperlink info for this row
                    const row = EXPORT_MISSING_ASSIGNMENTS_COLUMNS.map((col, colIndex) => {
                        let value = '';
                        if (col.field.startsWith('student.')) {
                            const field = col.field.replace('student.', '');
                            value = getFieldValue(student, field, col.fallback?.replace('student.', ''));
                        } else if (col.field.startsWith('assignment.')) {
                            const field = col.field.replace('assignment.', '');
                            value = getFieldValue(normalizedAssignment, field, col.fallback?.replace('assignment.', ''));
                        }

                        // Store hyperlink metadata for this column
                        if (col.hyperlink) {
                            if (col.hyperlinkField) {
                                // Assignment column: use assignmentLink as URL and assignmentTitle as text
                                const linkField = col.hyperlinkField.replace('assignment.', '');
                                const url = getFieldValue(normalizedAssignment, linkField);
                                rowMetadata[colIndex] = { url: url, text: value };
                            } else if (col.hyperlinkText) {
                                // Fixed text columns: Grade Book and Submission
                                rowMetadata[colIndex] = { url: value, text: col.hyperlinkText };
                            }
                        }

                        return value;
                    });
                    missingAssignmentsData.push(row);
                    hyperlinkMetadata.push(rowMetadata);
                });
            }
        });

        // Create workbook with both sheets
        const wb = XLSX.utils.book_new();
        const ws1 = XLSX.utils.aoa_to_sheet(masterListData);
        const ws2 = XLSX.utils.aoa_to_sheet(missingAssignmentsData);

        // Add HYPERLINK formulas to Master List sheet
        masterListHyperlinkMetadata.forEach((rowMeta, rowIndex) => {
            Object.entries(rowMeta).forEach(([colIndex, linkData]) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: parseInt(colIndex) }); // +1 to skip header row
                if (ws1[cellAddress] && linkData.url) {
                    // Create HYPERLINK formula: =HYPERLINK("url", "text")
                    const escapedUrl = linkData.url.replace(/"/g, '""'); // Escape quotes in URL
                    const escapedText = String(linkData.text || 'Link').replace(/"/g, '""'); // Escape quotes in text
                    ws1[cellAddress] = {
                        t: 'f',
                        f: `HYPERLINK("${escapedUrl}","${escapedText}")`,
                        v: linkData.text || 'Link'
                    };
                }
            });
        });

        // Add HYPERLINK formulas to Missing Assignments sheet
        hyperlinkMetadata.forEach((rowMeta, rowIndex) => {
            Object.entries(rowMeta).forEach(([colIndex, linkData]) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: parseInt(colIndex) }); // +1 to skip header row
                if (ws2[cellAddress] && linkData.url) {
                    // Create HYPERLINK formula: =HYPERLINK("url", "text")
                    const escapedUrl = linkData.url.replace(/"/g, '""'); // Escape quotes in URL
                    const escapedText = String(linkData.text || 'Link').replace(/"/g, '""'); // Escape quotes in text
                    ws2[cellAddress] = {
                        t: 'f',
                        f: `HYPERLINK("${escapedUrl}","${escapedText}")`,
                        v: linkData.text || 'Link'
                    };
                }
            });
        });

        // Find Grade column index for conditional formatting in Master List
        const gradeColIndex = MASTER_LIST_COLUMNS.findIndex(col => col.conditionalFormatting === 'grade');

        // Apply conditional formatting to Grade column in Master List (GREEN >= 70, YELLOW 60-69, RED < 60)
        if (gradeColIndex !== -1 && students.length > 0) {
            applyGradeConditionalFormatting(ws1, gradeColIndex, 2, students.length + 1);
        }

        // Find Overall Grade column index for conditional formatting in Missing Assignments
        const missingGradeColIndex = EXPORT_MISSING_ASSIGNMENTS_COLUMNS.findIndex(col => col.conditionalFormatting === 'grade');

        // Apply conditional formatting to Overall Grade column in Missing Assignments (GREEN >= 70, YELLOW 60-69, RED < 60)
        if (missingGradeColIndex !== -1 && missingAssignmentsData.length > 1) {
            applyGradeConditionalFormatting(ws2, missingGradeColIndex, 2, missingAssignmentsData.length);
        }

        // Set column widths for Master List (custom or auto-fit)
        ws1['!cols'] = calculateColumnWidths(MASTER_LIST_COLUMNS, masterListData);

        // Set column widths for Missing Assignments (custom or auto-fit)
        ws2['!cols'] = calculateColumnWidths(EXPORT_MISSING_ASSIGNMENTS_COLUMNS, missingAssignmentsData);

        // --- SHEET 3: LDA (Filtered and Sorted Master List) ---
        // Filter students by daysOut >= 5
        const filteredStudents = students.filter(student => {
            const daysOut = parseInt(student.daysOut || 0);
            return daysOut >= 5;
        });

        // Sort by daysOut from highest to lowest
        filteredStudents.sort((a, b) => {
            const daysOutA = parseInt(a.daysOut || 0);
            const daysOutB = parseInt(b.daysOut || 0);
            return daysOutB - daysOutA; // Descending order
        });

        // Create LDA sheet data using the same structure as Master List
        const ldaHeaders = MASTER_LIST_COLUMNS.map(col => col.header);
        const ldaData = [ldaHeaders];
        const ldaHyperlinkMetadata = []; // Store hyperlink info for each row

        filteredStudents.forEach(student => {
            const rowMetadata = {}; // Store hyperlink info for this row
            const row = MASTER_LIST_COLUMNS.map((col, colIndex) => {
                let value = getFieldValue(student, col.field, col.fallback);

                if (col.field === 'missingCount') {
                    value = value || 0;
                } else if (col.field === 'daysOut') {
                    value = parseInt(value || 0);
                } else if (col.field === 'lda') {
                    if (value) {
                        // Format LDA dates to MM-DD-YY format
                        const dateObj = parseDate(value);
                        if (dateObj) {
                            value = formatDateToMMDDYY(dateObj);
                        }
                    }
                }

                // Store hyperlink metadata for this column
                if (col.hyperlink && col.hyperlinkText && value) {
                    rowMetadata[colIndex] = { url: value, text: col.hyperlinkText };
                }

                return value;
            });
            ldaData.push(row);
            ldaHyperlinkMetadata.push(rowMetadata);
        });

        // Create worksheet for LDA sheet
        const ws3 = XLSX.utils.aoa_to_sheet(ldaData);

        // Add HYPERLINK formulas to LDA sheet
        ldaHyperlinkMetadata.forEach((rowMeta, rowIndex) => {
            Object.entries(rowMeta).forEach(([colIndex, linkData]) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: parseInt(colIndex) }); // +1 to skip header row
                if (ws3[cellAddress] && linkData.url) {
                    // Create HYPERLINK formula: =HYPERLINK("url", "text")
                    const escapedUrl = linkData.url.replace(/"/g, '""'); // Escape quotes in URL
                    const escapedText = String(linkData.text || 'Link').replace(/"/g, '""'); // Escape quotes in text
                    ws3[cellAddress] = {
                        t: 'f',
                        f: `HYPERLINK("${escapedUrl}","${escapedText}")`,
                        v: linkData.text || 'Link'
                    };
                }
            });
        });

        // Apply conditional formatting to Grade column in LDA sheet (GREEN >= 70, YELLOW 60-69, RED < 60)
        if (gradeColIndex !== -1 && filteredStudents.length > 0) {
            applyGradeConditionalFormatting(ws3, gradeColIndex, 2, filteredStudents.length + 1);
        }

        // Set column widths for LDA sheet (custom or auto-fit)
        ws3['!cols'] = calculateColumnWidths(MASTER_LIST_COLUMNS, ldaData);

        // Hide columns not in LDA_VISIBLE_COLUMNS whitelist (applied after width calculation)
        MASTER_LIST_COLUMNS.forEach((col, i) => {
            if (!LDA_VISIBLE_COLUMNS.includes(col.field)) {
                if (!ws3['!cols'][i]) ws3['!cols'][i] = {};
                ws3['!cols'][i].hidden = true;
            }
        });

        // Create sheet name with reference date formatted as MM-DD-YYYY
        const ldaMonth = String(referenceDate.getMonth() + 1).padStart(2, '0');
        const ldaDay = String(referenceDate.getDate()).padStart(2, '0');
        const ldaYear = String(referenceDate.getFullYear());
        const ldaSheetName = `LDA ${ldaMonth}-${ldaDay}-${ldaYear}`;

        XLSX.utils.book_append_sheet(wb, ws1, 'Master List');
        XLSX.utils.book_append_sheet(wb, ws2, 'Missing Assignments');
        XLSX.utils.book_append_sheet(wb, ws3, ldaSheetName);

        const timestamp = new Date().toISOString().split('T')[0];
        const filename = `student_report_${timestamp}.xlsx`;

        // Write file with cell styles enabled
        XLSX.writeFile(wb, filename, { cellStyles: true });

        console.log(`✓ Exported ${students.length} students to Excel file: ${filename}`);
        console.log(`  - Master List: ${students.length} students`);
        console.log(`  - Missing Assignments: ${missingAssignmentsData.length - 1} assignments`);
        console.log(`  - ${ldaSheetName}: ${filteredStudents.length} students (Days Out >= 5)`);

    } catch (error) {
        console.error('Error exporting to Excel:', error);
        alert('Error creating Excel file. Check console for details.');
    }
}
