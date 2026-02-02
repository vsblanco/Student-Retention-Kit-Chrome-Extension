// Canvas Integration - Handles all Canvas API calls for student data and assignments
import { STORAGE_KEYS, CANVAS_DOMAIN, GENERIC_AVATAR_URL } from '../constants/index.js';
import { getCachedData, setCachedData, hasCachedData, getCache } from '../utils/canvasCache.js';
import { openCanvasAuthErrorModal, isCanvasAuthError, isCanvasAuthErrorBody } from './modal-manager.js';
import { storageGet } from '../utils/storage.js';

/**
 * Fetches courses by scraping the Canvas user profile HTML page
 * This is a workaround for when the API endpoint is blocked
 * @param {number|string} canvasUserId - The Canvas user ID
 * @returns {Promise<Array>} Array of course objects with id, name, start_at, end_at, enrollments
 */
async function fetchCoursesFromHtml(canvasUserId) {
    const profileUrl = `${CANVAS_DOMAIN}/users/${canvasUserId}`;
    console.log(`[Non-API] Fetching courses from HTML: ${profileUrl}`);

    try {
        const response = await fetch(profileUrl, {
            headers: { 'Accept': 'text/html' },
            credentials: 'include'
        });

        if (!response.ok) {
            console.warn(`[Non-API] Failed to fetch user profile page: ${response.status}`);
            return [];
        }

        const html = await response.text();

        // Parse the HTML to extract course information
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Find the courses list
        const coursesList = doc.querySelector('#courses_list ul.context_list');
        if (!coursesList) {
            console.warn('[Non-API] Could not find courses list in HTML');
            return [];
        }

        const courseItems = coursesList.querySelectorAll('li');
        const courses = [];

        courseItems.forEach(li => {
            const link = li.querySelector('a');
            if (!link) return;

            const href = link.getAttribute('href');
            // Extract course ID and user ID from href like "/courses/108968/users/178163"
            const match = href.match(/\/courses\/(\d+)\/users\/(\d+)/);
            if (!match) return;

            const courseId = parseInt(match[1], 10);

            // Get course name from the title attribute of the name span
            const nameSpan = link.querySelector('span.name');
            const courseName = nameSpan ? (nameSpan.getAttribute('title') || nameSpan.textContent.trim()) : '';

            // Get subtitles - first one has the term/date, second has enrollment status
            const subtitles = link.querySelectorAll('span.subtitle');
            let termInfo = '';
            let enrollmentStatus = '';

            if (subtitles.length >= 1) {
                termInfo = subtitles[0].textContent.trim();
            }
            if (subtitles.length >= 2) {
                enrollmentStatus = subtitles[1].textContent.trim();
            }

            // Determine if course is active based on li class or enrollment status
            const isActive = li.classList.contains('active') || enrollmentStatus.toLowerCase().includes('active');
            const isCompleted = li.classList.contains('inactive') || enrollmentStatus.toLowerCase().includes('completed');
            const isPending = li.classList.contains('accepted') || enrollmentStatus.toLowerCase().includes('pending');

            // Extract start date from term info like "FTC 2026 01C January (1/12/2026)"
            let startDate = null;
            let endDate = null;
            const dateMatch = termInfo.match(/\((\d{1,2})\/(\d{1,2})\/(\d{4})\)/);
            if (dateMatch) {
                const month = parseInt(dateMatch[1], 10);
                const day = parseInt(dateMatch[2], 10);
                const year = parseInt(dateMatch[3], 10);
                startDate = new Date(year, month - 1, day);
                // Estimate end date as ~5 weeks after start (typical course length)
                endDate = new Date(startDate);
                endDate.setDate(endDate.getDate() + 35);
            }

            // Build a course object similar to what the API returns
            const course = {
                id: courseId,
                name: courseName,
                start_at: startDate ? startDate.toISOString() : null,
                end_at: endDate ? endDate.toISOString() : null,
                workflow_state: isActive ? 'available' : (isCompleted ? 'completed' : 'unpublished'),
                enrollments: [{
                    type: 'StudentEnrollment',
                    enrollment_state: isActive ? 'active' : (isCompleted ? 'completed' : 'invited'),
                    grades: {} // Grades not available from HTML scraping
                }]
            };

            courses.push(course);
            console.log(`[Non-API] Found course: ${courseName} (ID: ${courseId}, Active: ${isActive})`);
        });

        console.log(`[Non-API] Extracted ${courses.length} course(s) from HTML`);
        return courses;

    } catch (error) {
        console.error('[Non-API] Error fetching courses from HTML:', error);
        return [];
    }
}

// Track if we've already shown the auth error modal in this session
let authErrorShownInSession = false;
// Track if shutdown was requested - persists until process fully stops
let shutdownRequested = false;

/**
 * Resets the auth error state - call this when starting a new process
 */
export function resetAuthErrorState() {
    authErrorShownInSession = false;
    shutdownRequested = false;
}

/**
 * Checks if shutdown was requested
 * @returns {boolean} True if shutdown was requested
 */
export function isShutdownRequested() {
    return shutdownRequested;
}

/**
 * Custom error class for Canvas auth shutdown
 */
export class CanvasAuthShutdownError extends Error {
    constructor() {
        super('Canvas authentication shutdown requested by user');
        this.name = 'CanvasAuthShutdownError';
    }
}

/**
 * Throws CanvasAuthShutdownError if shutdown was requested
 * Call this at the start of operations to abort early
 */
function checkShutdown() {
    if (shutdownRequested) {
        throw new CanvasAuthShutdownError();
    }
}

/**
 * Handles Canvas API authorization errors by showing a modal and returning user choice
 * @param {string} context - Context description for logging
 * @returns {Promise<'continue'|'shutdown'>} The user's choice
 */
async function handleCanvasAuthError(context) {
    console.warn(`[Canvas Integration] Authorization error during ${context}`);

    // If shutdown was already requested, don't show modal again
    if (shutdownRequested) {
        return 'shutdown';
    }

    // Prevent multiple modals from stacking
    if (authErrorShownInSession) {
        return 'continue';
    }

    authErrorShownInSession = true;
    const choice = await openCanvasAuthErrorModal();
    authErrorShownInSession = false;

    // If shutdown selected, set the persistent flag
    if (choice === 'shutdown') {
        shutdownRequested = true;
    }

    return choice;
}

/**
 * Formats duration - shows minutes when >= 60 seconds
 * @param {number} seconds - Duration in seconds
 * @returns {string} Formatted duration string
 */
export function formatDuration(seconds) {
    if (seconds >= 60) {
        const mins = (seconds / 60).toFixed(1);
        return `${mins}m`;
    }
    return `${seconds.toFixed(1)}s`;
}

/**
 * Updates the total time display with current elapsed time
 */
export function updateTotalTime() {
    const queueTotalTimeDiv = document.getElementById('queueTotalTime');
    if (queueTotalTimeDiv && queueTotalTimeDiv.dataset.processStartTime) {
        const totalSeconds = (Date.now() - parseInt(queueTotalTimeDiv.dataset.processStartTime)) / 1000;
        queueTotalTimeDiv.textContent = `Total Time: ${formatDuration(totalSeconds)}`;
        queueTotalTimeDiv.style.display = 'block';
    }
}

/**
 * Preload image for faster rendering
 */
function preloadImage(url) {
    if (!url) return;
    const img = new Image();
    img.src = url;
}

/**
 * Fetches Canvas details for a student (user data and courses)
 * @param {Object} student - The student object
 * @param {boolean} cacheEnabled - Whether to use caching (default: true)
 */
export async function fetchCanvasDetails(student, cacheEnabled = true) {
    // Check if shutdown was requested before processing
    checkShutdown();

    if (!student.SyStudentId) return student;

    try {
        // Ensure SyStudentId is a string for consistent cache key lookup
        const syStudentId = String(student.SyStudentId);
        console.log(`[fetchCanvasDetails] Processing student: ${student.name}, SyStudentId: ${syStudentId}`);

        // Check if non-API course fetch is enabled FIRST
        // Use storageGet to properly handle nested storage paths
        const nonApiSettings = await storageGet([STORAGE_KEYS.NON_API_COURSE_FETCH]);
        const useNonApiFetch = nonApiSettings[STORAGE_KEYS.NON_API_COURSE_FETCH] || false;

        // Only check cache if caching is enabled AND non-API fetch is disabled
        // When non-API fetch is enabled, we always fetch fresh course data from HTML
        const cachedData = (cacheEnabled && !useNonApiFetch) ? await getCachedData(syStudentId) : null;

        let userData;
        let courses;

        if (cachedData) {
            console.log(`✓ Cache hit for ${student.name || student.SyStudentId}`);
            userData = cachedData.userData;
            courses = cachedData.courses;
        } else {
            console.log(`→ Fetching fresh data for ${student.name || student.SyStudentId}`);

            const userUrl = `${CANVAS_DOMAIN}/api/v1/users/sis_user_id:${student.SyStudentId}`;
            const userResp = await fetch(userUrl, { headers: { 'Accept': 'application/json' } });

            // Check for Canvas authorization errors
            if (isCanvasAuthError(userResp)) {
                const choice = await handleCanvasAuthError('fetching user data');
                if (choice === 'shutdown') {
                    throw new CanvasAuthShutdownError();
                }
                // Continue with next student
                return student;
            }

            if (!userResp.ok) {
                console.warn(`✗ Failed to fetch user data for ${student.SyStudentId}: ${userResp.status} ${userResp.statusText}`);
                return student;
            }
            userData = await userResp.json();

            const canvasUserId = userData.id;

            if (canvasUserId) {
                if (useNonApiFetch) {
                    // Use HTML scraping method instead of API
                    console.log(`[fetchCanvasDetails] Using non-API course fetch for ${student.name || student.SyStudentId}`);
                    courses = await fetchCoursesFromHtml(canvasUserId);
                } else {
                    // Fetch ALL courses (not just active) to support Time Machine mode
                    // enrollment_state can be: active, invited_or_pending, completed
                    const coursesUrl = `${CANVAS_DOMAIN}/api/v1/users/${canvasUserId}/courses?include[]=enrollments&per_page=100`;
                    const coursesResp = await fetch(coursesUrl, { headers: { 'Accept': 'application/json' } });

                    // Check for Canvas authorization errors
                    if (isCanvasAuthError(coursesResp)) {
                        const choice = await handleCanvasAuthError('fetching courses');
                        if (choice === 'shutdown') {
                            throw new CanvasAuthShutdownError();
                        }
                        // Continue with what we have
                        courses = [];
                    } else if (coursesResp.ok) {
                        courses = await coursesResp.json();
                        console.log(`✓ Fetched ${courses.length} course(s) for ${student.name || student.SyStudentId}`);
                    } else {
                        console.warn(`✗ Failed to fetch courses for ${student.SyStudentId}: ${coursesResp.status} ${coursesResp.statusText}`);
                        courses = [];
                    }
                }

                // Only save to cache if caching is enabled
                if (cacheEnabled) {
                    await setCachedData(syStudentId, userData, courses);
                }
            }
        }

        // Process userData
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

        // Process courses
        if (canvasUserId && courses && courses.length > 0) {
            console.log(`Processing ${courses.length} course(s) for ${student.name || student.SyStudentId}`);

            // Check if using specific date (Time Machine mode)
            const settings = await chrome.storage.local.get([STORAGE_KEYS.USE_SPECIFIC_DATE, STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE]);
            const useSpecificDate = settings[STORAGE_KEYS.USE_SPECIFIC_DATE] || false;
            const specificDateStr = settings[STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE];

            let now;
            if (useSpecificDate && specificDateStr) {
                // Parse the specific date (format: YYYY-MM-DD)
                const [year, month, day] = specificDateStr.split('-').map(Number);
                now = new Date(year, month - 1, day); // month is 0-indexed
                console.log(`🕐 Time Machine mode: Using date ${specificDateStr} for course selection`);
            } else {
                now = new Date();
            }

            // Debug: Show all course names before filtering
            console.log(`Courses before filtering:`, courses.map(c => c.name || '(no name)'));

            const validCourses = courses.filter(c => c.name && !c.name.toUpperCase().includes('CAPV'));

            console.log(`${validCourses.length} course(s) after CAPV filter`);
            if (validCourses.length > 0) {
                console.log(`Valid courses:`, validCourses.map(c => c.name));
            }

            let activeCourse = null;

            activeCourse = validCourses.find(c => {
                if (!c.start_at || !c.end_at) return false;
                const start = new Date(c.start_at);
                const end = new Date(c.end_at);
                return now >= start && now <= end;
            });

            // If no course is active on the reference date, pick the most recent course
            // This ensures we always have a course-based URL for Step 3
            if (!activeCourse && validCourses.length > 0) {
                validCourses.sort((a, b) => {
                    const dateA = a.start_at ? new Date(a.start_at) : new Date(0);
                    const dateB = b.start_at ? new Date(b.start_at) : new Date(0);
                    return dateB - dateA;
                });
                activeCourse = validCourses[0];
                console.log(`No active course found for ${now.toLocaleDateString()}, using most recent: ${activeCourse.name}`);
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
                // This should rarely happen - only if student has no courses at all
                console.warn(`${student.name}: No courses found, Step 3 will be skipped`);
                student.url = null;
            }
        }

        return student;

    } catch (e) {
        console.error(`✗ Error fetching Canvas details for ${student.SyStudentId}:`, e);
        return student;
    }
}

/**
 * Process Step 2: Fetch Canvas IDs, courses, and photos for all students
 * Optimized to process cached students first for faster initial progress
 * Uses reverse cache lookup: iterates through cache entries and matches to master list
 */
export async function processStep2(students, renderCallback) {
    // Reset auth error state at the start of a new process
    resetAuthErrorState();

    const step2 = document.getElementById('step2');
    const timeSpan = step2.querySelector('.step-time');

    step2.className = 'queue-item active';
    step2.querySelector('i').className = 'fas fa-spinner';

    const startTime = Date.now();

    try {
        console.log(`[Step 2] Pinging Canvas API: ${CANVAS_DOMAIN}`);

        // Check if cache is enabled
        const settings = await chrome.storage.local.get([STORAGE_KEYS.CANVAS_CACHE_ENABLED]);
        const cacheEnabled = settings[STORAGE_KEYS.CANVAS_CACHE_ENABLED] !== undefined
            ? settings[STORAGE_KEYS.CANVAS_CACHE_ENABLED]
            : true;

        console.log(`[Step 2] Cache enabled: ${cacheEnabled}`);

        // Separate students into cached and uncached groups
        const cachedStudents = [];
        const uncachedStudents = [];

        if (cacheEnabled) {
            // REVERSE LOOKUP: Get all cached entries first, then match to master list
            // This is more efficient when cache is empty (no need to check each student)
            console.log(`[Step 2] Using reverse cache lookup...`);

            const cache = await getCache();
            const cachedIds = Object.keys(cache);
            const now = new Date();

            // Filter out expired entries
            const validCachedIds = cachedIds.filter(id => {
                const entry = cache[id];
                if (!entry || !entry.expiresAt) return false;
                return new Date(entry.expiresAt) > now;
            });

            console.log(`[Step 2] Found ${validCachedIds.length} valid cached entries`);

            // Create a map of SyStudentId -> student for quick lookup
            const studentBySyId = new Map();
            students.forEach(student => {
                if (student.SyStudentId) {
                    studentBySyId.set(String(student.SyStudentId), student);
                }
            });

            // Match cached entries to students in master list
            for (const cachedId of validCachedIds) {
                if (studentBySyId.has(cachedId)) {
                    cachedStudents.push(studentBySyId.get(cachedId));
                    studentBySyId.delete(cachedId); // Remove from map so remaining are uncached
                }
            }

            // Remaining students in map are uncached
            uncachedStudents.push(...studentBySyId.values());

            // Update progress for cache check phase (quick since we just read cache once)
            timeSpan.textContent = `15%`;
        } else {
            // Cache disabled - all students are uncached
            console.log(`[Step 2] Cache disabled - processing all ${students.length} students fresh`);
            uncachedStudents.push(...students);
            timeSpan.textContent = `15%`;
        }

        console.log(`[Step 2] Found ${cachedStudents.length} cached, ${uncachedStudents.length} uncached`);
        console.log(`[Step 2] Processing cached students first for faster progress...`);

        const BATCH_SIZE = 20;
        const BATCH_DELAY_MS = 100;

        let processedCount = 0;
        let updatedStudents = [...students];

        // Create a map to track student indices
        const studentIndexMap = new Map();
        students.forEach((student, index) => {
            studentIndexMap.set(student, index);
        });

        // Process cached students first
        const totalCachedBatches = Math.ceil(cachedStudents.length / BATCH_SIZE);
        for (let i = 0; i < cachedStudents.length; i += BATCH_SIZE) {
            // Check if shutdown was requested before processing next batch
            checkShutdown();

            const batch = cachedStudents.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;

            console.log(`[Step 2] Cached batch ${batchNumber}/${totalCachedBatches} (students ${i + 1}-${Math.min(i + BATCH_SIZE, cachedStudents.length)})`);

            // Use Promise.allSettled to handle auth errors gracefully
            const promises = batch.map(student => fetchCanvasDetails(student, cacheEnabled));
            const settledResults = await Promise.allSettled(promises);

            // Check for shutdown errors
            for (const result of settledResults) {
                if (result.status === 'rejected' && result.reason instanceof CanvasAuthShutdownError) {
                    throw result.reason;
                }
            }

            settledResults.forEach((result, batchIndex) => {
                const originalStudent = batch[batchIndex];
                const originalIndex = studentIndexMap.get(originalStudent);
                // If fulfilled, use the result; if rejected (but not shutdown), keep original student
                updatedStudents[originalIndex] = result.status === 'fulfilled' ? result.value : originalStudent;
            });

            processedCount += batch.length;
            // Progress starts at 15% (after cache check) and goes to 100%
            const processingProgress = 15 + Math.round((processedCount / students.length) * 85);
            timeSpan.textContent = `${processingProgress}%`;

            // No delay needed for cached students (they're fast)
        }

        // Process uncached students (requires API calls)
        if (uncachedStudents.length > 0) {
            console.log(`[Step 2] Now processing uncached students (API calls required)...`);
        }

        const totalUncachedBatches = Math.ceil(uncachedStudents.length / BATCH_SIZE);
        for (let i = 0; i < uncachedStudents.length; i += BATCH_SIZE) {
            // Check if shutdown was requested before processing next batch
            checkShutdown();

            const batch = uncachedStudents.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;

            console.log(`[Step 2] Uncached batch ${batchNumber}/${totalUncachedBatches} (students ${i + 1}-${Math.min(i + BATCH_SIZE, uncachedStudents.length)})`);

            // Use Promise.allSettled to handle auth errors gracefully
            const promises = batch.map(student => fetchCanvasDetails(student, cacheEnabled));
            const settledResults = await Promise.allSettled(promises);

            // Check for shutdown errors
            for (const result of settledResults) {
                if (result.status === 'rejected' && result.reason instanceof CanvasAuthShutdownError) {
                    throw result.reason;
                }
            }

            settledResults.forEach((result, batchIndex) => {
                const originalStudent = batch[batchIndex];
                const originalIndex = studentIndexMap.get(originalStudent);
                // If fulfilled, use the result; if rejected (but not shutdown), keep original student
                updatedStudents[originalIndex] = result.status === 'fulfilled' ? result.value : originalStudent;
            });

            processedCount += batch.length;
            // Progress starts at 15% (after cache check) and goes to 100%
            const processingProgress = 15 + Math.round((processedCount / students.length) * 85);
            timeSpan.textContent = `${processingProgress}%`;

            if (i + BATCH_SIZE < uncachedStudents.length) {
                await new Promise(resolve => setTimeout(resolve, BATCH_DELAY_MS));
            }
        }

        await chrome.storage.local.set({ [STORAGE_KEYS.MASTER_ENTRIES]: updatedStudents });

        const durationSeconds = (Date.now() - startTime) / 1000;
        step2.className = 'queue-item completed';
        step2.querySelector('i').className = 'fas fa-check';
        timeSpan.textContent = formatDuration(durationSeconds);

        // Update total time counter
        updateTotalTime();

        console.log(`[Step 2] ✓ Complete in ${formatDuration(durationSeconds)} - ${students.length} students processed`);

        if (renderCallback) {
            renderCallback(updatedStudents);
        }

        return updatedStudents;

    } catch (error) {
        // Handle Canvas auth shutdown gracefully
        if (error instanceof CanvasAuthShutdownError) {
            console.log('[Step 2] Stopped by user due to Canvas auth error');
            step2.querySelector('i').className = 'fas fa-times';
            step2.style.color = '#ef4444';
            timeSpan.textContent = 'Stopped by user';
            throw error; // Re-throw to stop the pipeline
        }

        console.error("[Step 2 Error]", error);
        step2.querySelector('i').className = 'fas fa-times';
        step2.style.color = '#ef4444';
        timeSpan.textContent = 'Error';
        throw error;
    }
}

/**
 * Parses a Canvas gradebook URL to extract course and student IDs
 */
function parseGradebookUrl(url) {
    try {
        const urlObj = new URL(url);
        const regex = /courses\/(\d+)\/grades\/(\d+)/;
        const match = urlObj.pathname.match(regex);
        if (match) {
            return {
                origin: urlObj.origin,
                courseId: match[1],
                studentId: match[2]
            };
        }
    } catch (e) {
        console.warn('Invalid gradebook URL:', url);
    }
    return null;
}

/**
 * Fetches paginated data from Canvas API
 */
async function fetchPaged(url, items = []) {
    // Check if shutdown was requested before making request
    checkShutdown();

    const headers = {
        'Accept': 'application/json',
        'X-Requested-With': 'XMLHttpRequest'
    };

    try {
        const response = await fetch(url, { method: 'GET', credentials: 'include', headers });

        // Check for Canvas authorization errors
        if (isCanvasAuthError(response)) {
            const choice = await handleCanvasAuthError('fetching paged data');
            if (choice === 'shutdown') {
                throw new CanvasAuthShutdownError();
            }
            // Continue with partial data
            return items;
        }

        if (!response.ok) {
            if (items.length > 0) return items;
            throw new Error(`HTTP ${response.status}`);
        }

        const newItems = await response.json();
        const allItems = items.concat(newItems);

        const linkHeader = response.headers.get('Link');
        const nextUrl = getNextPageUrl(linkHeader);

        if (nextUrl) {
            return fetchPaged(nextUrl, allItems);
        }

        return allItems;
    } catch (e) {
        // Re-throw shutdown errors
        if (e instanceof CanvasAuthShutdownError) {
            throw e;
        }
        console.warn('Fetch error:', e);
        return items;
    }
}

/**
 * Extracts next page URL from Link header
 */
function getNextPageUrl(linkHeader) {
    if (!linkHeader) return null;
    const links = linkHeader.split(',');
    const nextLink = links.find(link => link.includes('rel="next"'));
    if (!nextLink) return null;
    const match = nextLink.match(/<([^>]+)>/);
    return match ? match[1] : null;
}

/**
 * Analyzes submissions to find missing assignments
 * @param {Date} referenceDate - The date to use for checking if assignments are past due (defaults to now)
 */
function analyzeMissingAssignments(submissions, userObject, studentName, courseId, origin, referenceDate = new Date()) {
    const now = referenceDate;
    const collectedAssignments = [];

    let currentGrade = "";
    if (userObject && userObject.enrollments) {
        const enrollment = userObject.enrollments.find(e => e.type === 'StudentEnrollment') || userObject.enrollments[0];

        if (enrollment && enrollment.grades) {
            if (enrollment.grades.current_score != null) {
                currentGrade = enrollment.grades.current_score;
            } else if (enrollment.grades.final_score != null) {
                currentGrade = enrollment.grades.final_score;
            } else if (enrollment.grades.current_grade != null) {
                currentGrade = String(enrollment.grades.current_grade).replace(/%/g, '');
            }
        }
    }

    submissions.forEach(sub => {
        const dueDate = sub.cached_due_date ? new Date(sub.cached_due_date) : null;

        if (dueDate && dueDate > now) return;

        const scoreStr = String(sub.score || sub.grade || '').toLowerCase();
        const isComplete = scoreStr === 'complete';

        if (isComplete) return;

        const isMissing = (sub.missing === true) ||
            ((sub.workflow_state === 'unsubmitted' || sub.workflow_state === 'unsubmitted (ungraded)') && (dueDate && dueDate < now)) ||
            (sub.score === 0);

        if (isMissing) {
            // Generate assignment URL from assignment ID
            const assignmentId = sub.assignment ? sub.assignment.id : null;
            const assignmentUrl = assignmentId ? `${origin}/courses/${courseId}/assignments/${assignmentId}` : '';

            // Format score as "points earned/points possible"
            let formattedScore = '-';
            if (sub.score !== null && sub.assignment && sub.assignment.points_possible !== null) {
                formattedScore = `${sub.score}/${sub.assignment.points_possible}`;
            } else if (sub.grade) {
                formattedScore = sub.grade;
            }

            collectedAssignments.push({
                assignmentTitle: sub.assignment ? sub.assignment.name : 'Unknown Assignment',
                submissionLink: sub.preview_url || '',
                assignmentLink: assignmentUrl,
                dueDate: sub.cached_due_date ? new Date(sub.cached_due_date).toLocaleDateString() : 'No Date',
                score: formattedScore,
                workflow_state: sub.workflow_state
            });
        }
    });

    return {
        currentGrade: currentGrade,
        count: collectedAssignments.length,
        assignments: collectedAssignments
    };
}

/**
 * Finds the next upcoming assignment that hasn't been submitted yet
 * @param {Array} submissions - Array of submission objects from Canvas API
 * @param {string} courseId - The course ID
 * @param {string} origin - The Canvas domain origin
 * @param {Date} referenceDate - The date to use as "today" (defaults to now)
 * @returns {Object|null} Next assignment object with Assignment, DueDate, AssignmentLink or null if none found
 */
function findNextAssignment(submissions, courseId, origin, referenceDate = new Date()) {
    const now = referenceDate;
    // Set time to start of day for comparison
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    console.log(`[findNextAssignment] Starting - Total submissions: ${submissions.length}, Reference date: ${now.toISOString()}, Today start: ${todayStart.toISOString()}`);

    // Filter and collect upcoming assignments that haven't been submitted
    const upcomingAssignments = [];

    // Debug counters
    let noAssignmentData = 0;
    let noDueDate = 0;
    let dueDatePast = 0;
    let alreadySubmitted = 0;
    let isCompleteCount = 0;

    submissions.forEach((sub, index) => {
        // Skip if no assignment data
        if (!sub.assignment) {
            noAssignmentData++;
            return;
        }

        const rawDueAt = sub.assignment.due_at;
        const dueDate = rawDueAt ? new Date(rawDueAt) : null;

        // Log first few submissions for debugging
        if (index < 5) {
            console.log(`[findNextAssignment] Sub[${index}]: "${sub.assignment.name}" | due_at raw: ${rawDueAt} | parsed: ${dueDate ? dueDate.toISOString() : 'null'} | workflow: ${sub.workflow_state} | score: ${sub.score} | submitted_at: ${sub.submitted_at}`);
        }

        // Skip assignments without a due date or with due dates before today
        if (!dueDate) {
            noDueDate++;
            return;
        }
        if (dueDate < todayStart) {
            dueDatePast++;
            return;
        }

        // Check if the assignment has already been submitted
        // A submission is considered "submitted" if:
        // - workflow_state is 'submitted', 'graded', or 'pending_review'
        // - submitted_at has an actual value (date string)
        // - score is not null/undefined and score > 0 (has been graded with a passing score)
        const submittedStates = ['submitted', 'graded', 'pending_review'];
        const hasSubmittedAt = sub.submitted_at != null && sub.submitted_at !== '';
        const hasPassingScore = sub.score != null && sub.score > 0;
        const isSubmitted = submittedStates.includes(sub.workflow_state) ||
                           hasSubmittedAt ||
                           hasPassingScore;

        // Also check for "complete" grade which indicates completion
        const scoreStr = String(sub.score || sub.grade || '').toLowerCase();
        const isComplete = scoreStr === 'complete';

        // Skip if already submitted or complete
        if (isSubmitted) {
            alreadySubmitted++;
            return;
        }
        if (isComplete) {
            isCompleteCount++;
            return;
        }

        // Generate assignment URL
        const assignmentId = sub.assignment.id;
        const assignmentUrl = assignmentId ? `${origin}/courses/${courseId}/assignments/${assignmentId}` : '';

        // Format due date - show "Today", "Tomorrow", or "Feb 3" style
        const tomorrow = new Date(todayStart);
        tomorrow.setDate(tomorrow.getDate() + 1);
        const dueDateStart = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
        const isToday = dueDateStart.getTime() === todayStart.getTime();
        const isTomorrow = dueDateStart.getTime() === tomorrow.getTime();
        const formattedDueDate = isToday ? 'Today' : isTomorrow ? 'Tomorrow' : dueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

        console.log(`[findNextAssignment] ✓ Found upcoming: "${sub.assignment.name}" due ${formattedDueDate}`);

        upcomingAssignments.push({
            Assignment: sub.assignment.name || 'Unknown Assignment',
            DueDate: formattedDueDate,
            AssignmentLink: assignmentUrl,
            _dueDateObj: dueDate // Keep for sorting, will be removed before returning
        });
    });

    console.log(`[findNextAssignment] Filter results: noAssignmentData=${noAssignmentData}, noDueDate=${noDueDate}, dueDatePast=${dueDatePast}, alreadySubmitted=${alreadySubmitted}, isComplete=${isCompleteCount}, upcoming=${upcomingAssignments.length}`);

    // If no upcoming assignments found, return null
    if (upcomingAssignments.length === 0) {
        console.log(`[findNextAssignment] No upcoming assignments found - returning null`);
        return null;
    }

    // Sort by due date (nearest first)
    upcomingAssignments.sort((a, b) => a._dueDateObj - b._dueDateObj);

    // Get the nearest assignment and remove the internal sorting field
    const nextAssignment = upcomingAssignments[0];
    delete nextAssignment._dueDateObj;

    return nextAssignment;
}

/**
 * Fetches missing assignments for a single student
 * @param {Date} referenceDate - The date to use for checking if assignments are past due
 * @param {boolean} includeNextAssignment - Whether to also find the next upcoming assignment
 */
async function fetchMissingAssignments(student, referenceDate = new Date(), includeNextAssignment = false) {
    // Check if shutdown was requested before processing
    checkShutdown();

    // Support both 'url' field and legacy 'Gradebook' field
    const gradebookUrl = student.url || student.Gradebook;

    if (!gradebookUrl) {
        console.log(`[Step 3] ${student.name}: No gradebook URL, skipping`);
        return { ...student, missingCount: 0, missingAssignments: [], nextAssignment: null };
    }

    const parsed = parseGradebookUrl(gradebookUrl);
    if (!parsed) {
        console.warn(`[Step 3] ${student.name}: Failed to parse gradebook URL: ${gradebookUrl}`);
        return { ...student, missingCount: 0, missingAssignments: [], nextAssignment: null };
    }

    const { origin, courseId, studentId } = parsed;

    try {
        const submissionsUrl = `${origin}/api/v1/courses/${courseId}/students/submissions?student_ids[]=${studentId}&include[]=assignment&per_page=100`;
        const submissions = await fetchPaged(submissionsUrl);

        const usersUrl = `${origin}/api/v1/courses/${courseId}/users?user_ids[]=${studentId}&include[]=enrollments&per_page=100`;
        const users = await fetchPaged(usersUrl);
        const userObject = users && users.length > 0 ? users[0] : null;

        const result = analyzeMissingAssignments(submissions, userObject, student.name, courseId, origin, referenceDate);

        if (result.count > 0) {
            console.log(`[Step 3] ${student.name}: Found ${result.count} missing assignment(s), Grade: ${result.currentGrade || 'N/A'}`);
        }

        // Find next assignment if enabled
        let nextAssignment = null;
        if (includeNextAssignment) {
            nextAssignment = findNextAssignment(submissions, courseId, origin, referenceDate);
            if (nextAssignment) {
                console.log(`[Step 3] ${student.name}: Next assignment due: ${nextAssignment.Assignment} (${nextAssignment.DueDate})`);
            }
        }

        return {
            ...student,
            missingCount: result.count,
            missingAssignments: result.assignments,
            currentGrade: result.currentGrade,
            nextAssignment: nextAssignment
        };

    } catch (e) {
        // Re-throw shutdown errors
        if (e instanceof CanvasAuthShutdownError) {
            throw e;
        }
        console.error(`[Step 3] ${student.name}: Error fetching data:`, e);
        return { ...student, missingCount: 0, missingAssignments: [], nextAssignment: null };
    }
}

/**
 * Process Step 3: Check missing assignments and grades for all students
 */
export async function processStep3(students, renderCallback) {
    const step3 = document.getElementById('step3');
    const timeSpan = step3.querySelector('.step-time');

    step3.className = 'queue-item active';
    step3.querySelector('i').className = 'fas fa-spinner';

    const startTime = Date.now();

    try {
        console.log(`[Step 3] Checking student gradebooks for missing assignments`);
        console.log(`[Step 3] Processing ${students.length} students in batches of 20`);

        // Check if using specific date (Time Machine mode) for missing assignments check
        // Also check if Next Assignment feature is enabled
        // Use storageGet to properly handle nested storage paths (e.g., settings.canvas.nextAssignmentEnabled)
        const settings = await storageGet([
            STORAGE_KEYS.USE_SPECIFIC_DATE,
            STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE,
            STORAGE_KEYS.NEXT_ASSIGNMENT_ENABLED
        ]);
        const useSpecificDate = settings[STORAGE_KEYS.USE_SPECIFIC_DATE] || false;
        const specificDateStr = settings[STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE];
        const nextAssignmentEnabled = settings[STORAGE_KEYS.NEXT_ASSIGNMENT_ENABLED] || false;

        if (nextAssignmentEnabled) {
            console.log(`[Step 3] Next Assignment feature enabled - will find next due assignment for each student`);
        } else {
            console.log(`[Step 3] Next Assignment feature is DISABLED - enable in settings to find upcoming assignments`);
        }

        let referenceDate;
        if (useSpecificDate && specificDateStr) {
            // Parse the specific date (format: YYYY-MM-DD)
            const [year, month, day] = specificDateStr.split('-').map(Number);
            referenceDate = new Date(year, month - 1, day); // month is 0-indexed
            console.log(`🕐 Time Machine mode: Checking missing assignments as of ${specificDateStr}`);
        } else {
            referenceDate = new Date();
        }

        const BATCH_SIZE = 20;
        const BATCH_DELAY_MS = 100;

        let processedCount = 0;
        let updatedStudents = [...students];

        const totalBatches = Math.ceil(updatedStudents.length / BATCH_SIZE);

        for (let i = 0; i < updatedStudents.length; i += BATCH_SIZE) {
            // Check if shutdown was requested before processing next batch
            checkShutdown();

            const batch = updatedStudents.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;

            console.log(`[Step 3] Processing batch ${batchNumber}/${totalBatches} (students ${i + 1}-${Math.min(i + BATCH_SIZE, updatedStudents.length)})`);

            // Use Promise.allSettled to handle auth errors gracefully
            const promises = batch.map(student => fetchMissingAssignments(student, referenceDate, nextAssignmentEnabled));
            const settledResults = await Promise.allSettled(promises);

            // Check for shutdown errors
            for (const result of settledResults) {
                if (result.status === 'rejected' && result.reason instanceof CanvasAuthShutdownError) {
                    throw result.reason;
                }
            }

            settledResults.forEach((result, index) => {
                // If fulfilled, use the result; if rejected (but not shutdown), keep original student with zeros
                updatedStudents[i + index] = result.status === 'fulfilled'
                    ? result.value
                    : { ...batch[index], missingCount: 0, missingAssignments: [], nextAssignment: null };
            });

            processedCount += batch.length;
            timeSpan.textContent = `${Math.round((processedCount / updatedStudents.length) * 100)}%`;

            if (i + BATCH_SIZE < updatedStudents.length) {
                await new Promise(resolve => setTimeout(resolve, BATCH_DELAY_MS));
            }
        }

        await chrome.storage.local.set({ [STORAGE_KEYS.MASTER_ENTRIES]: updatedStudents });

        const durationSeconds = (Date.now() - startTime) / 1000;
        step3.className = 'queue-item completed';
        step3.querySelector('i').className = 'fas fa-check';
        timeSpan.textContent = formatDuration(durationSeconds);

        // Update total time counter
        updateTotalTime();

        const totalMissing = updatedStudents.reduce((sum, s) => sum + (s.missingCount || 0), 0);
        console.log(`[Step 3] ✓ Complete in ${formatDuration(durationSeconds)} - Found ${totalMissing} total missing assignments`);

        if (renderCallback) {
            renderCallback(updatedStudents);
        }

        return updatedStudents;

    } catch (error) {
        // Handle Canvas auth shutdown gracefully
        if (error instanceof CanvasAuthShutdownError) {
            console.log('[Step 3] Stopped by user due to Canvas auth error');
            step3.querySelector('i').className = 'fas fa-times';
            step3.style.color = '#ef4444';
            timeSpan.textContent = 'Stopped by user';
            throw error; // Re-throw to stop the pipeline
        }

        console.error("[Step 3 Error]", error);
        step3.querySelector('i').className = 'fas fa-times';
        step3.style.color = '#ef4444';
        timeSpan.textContent = 'Error';
        throw error;
    }
}

/**
 * Process Step 4: Send master list with missing assignments to Excel
 */
export async function processStep4(students) {
    const step4 = document.getElementById('step4');
    if (!step4) return students;

    const timeSpan = step4.querySelector('.step-time');

    step4.className = 'queue-item active';
    step4.querySelector('i').className = 'fas fa-spinner';
    step4.querySelector('.queue-content').innerHTML = '<i class="fas fa-spinner"></i> Sending List to Excel';

    const startTime = Date.now();

    try {
        console.log(`[Step 4] Sending master list with missing assignments to Excel`);

        // Dynamically import to avoid circular dependency
        const { sendMasterListWithMissingAssignmentsToExcel } = await import('./file-handler.js');
        const { getExcelTabs, openExcelInstanceModal, getCampusesFromStudents, openCampusSelectionModal } = await import('./modal-manager.js');

        // Check if there are multiple campuses - if so, show campus selection modal
        let studentsToSend = students;
        const campuses = getCampusesFromStudents(students);

        if (campuses.length > 1) {
            console.log(`[Step 4] Multiple campuses detected (${campuses.length}) - showing selection modal`);
            const selectedCampus = await openCampusSelectionModal(campuses);

            if (selectedCampus === null) {
                // User cancelled - skip sending but don't fail
                console.log('[Step 4] User cancelled campus selection - skipping send');
                const durationSeconds = (Date.now() - startTime) / 1000;
                step4.className = 'queue-item completed';
                step4.querySelector('i').className = 'fas fa-check';
                step4.querySelector('.queue-content').innerHTML = '<i class="fas fa-check"></i> Sending List to Excel';
                timeSpan.textContent = `${formatDuration(durationSeconds)} (skipped)`;
                updateTotalTime();
                return students;
            }

            // Filter students by selected campus (empty string means all)
            if (selectedCampus !== '') {
                studentsToSend = students.filter(s => s.campus === selectedCampus);
                console.log(`[Step 4] Filtered to ${studentsToSend.length} students from campus: ${selectedCampus}`);
            } else {
                console.log(`[Step 4] Sending all ${students.length} students from all campuses`);
            }
        }

        // Check how many Excel tabs are open
        const excelTabs = await getExcelTabs();

        let targetTabId = null;

        if (excelTabs.length === 0) {
            console.log('[Step 4] No Excel tabs detected - sending to all (will be handled when tabs open)');
            // Continue with null targetTabId - will attempt to send to any matching tabs
        } else if (excelTabs.length > 1) {
            // Multiple Excel tabs - show selection modal
            console.log(`[Step 4] Multiple Excel tabs detected (${excelTabs.length}) - showing selection modal`);
            targetTabId = await openExcelInstanceModal(excelTabs);

            if (targetTabId === null) {
                // User cancelled - skip sending but don't fail
                console.log('[Step 4] User cancelled Excel instance selection - skipping send');
                const durationSeconds = (Date.now() - startTime) / 1000;
                step4.className = 'queue-item completed';
                step4.querySelector('i').className = 'fas fa-check';
                step4.querySelector('.queue-content').innerHTML = '<i class="fas fa-check"></i> Sending List to Excel';
                timeSpan.textContent = `${formatDuration(durationSeconds)} (skipped)`;
                // Update total time counter
                updateTotalTime();
                return students;
            }
        } else {
            // Only one tab - use it directly
            targetTabId = excelTabs[0].id;
            console.log(`[Step 4] Single Excel tab detected - sending to tab ${targetTabId}`);
        }

        await sendMasterListWithMissingAssignmentsToExcel(studentsToSend, targetTabId);

        const durationSeconds = (Date.now() - startTime) / 1000;
        step4.className = 'queue-item completed';
        step4.querySelector('i').className = 'fas fa-check';
        step4.querySelector('.queue-content').innerHTML = '<i class="fas fa-check"></i> Sending List to Excel';
        timeSpan.textContent = formatDuration(durationSeconds);

        console.log(`[Step 4] ✓ Complete in ${formatDuration(durationSeconds)}`);

        // Update total time counter (final update)
        updateTotalTime();

        return students;

    } catch (error) {
        console.error("[Step 4 Error]", error);
        step4.querySelector('i').className = 'fas fa-times';
        step4.querySelector('.queue-content').innerHTML = '<i class="fas fa-times"></i> Sending List to Excel';
        step4.style.color = '#ef4444';
        timeSpan.textContent = 'Error';
        return students; // Don't throw, just return students
    }
}
