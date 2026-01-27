// Canvas Integration - Handles all Canvas API calls for student data and assignments
import { STORAGE_KEYS, CANVAS_DOMAIN, GENERIC_AVATAR_URL } from '../constants/index.js';
import { getCachedData, setCachedData, hasCachedData, getCache } from '../utils/canvasCache.js';

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
    if (!student.SyStudentId) return student;

    try {
        // Ensure SyStudentId is a string for consistent cache key lookup
        const syStudentId = String(student.SyStudentId);
        console.log(`[fetchCanvasDetails] Processing student: ${student.name}, SyStudentId: ${syStudentId}`);

        // Only check cache if caching is enabled
        const cachedData = cacheEnabled ? await getCachedData(syStudentId) : null;

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

            if (!userResp.ok) {
                console.warn(`✗ Failed to fetch user data for ${student.SyStudentId}: ${userResp.status} ${userResp.statusText}`);
                return student;
            }
            userData = await userResp.json();

            const canvasUserId = userData.id;

            if (canvasUserId) {
                // Fetch ALL courses (not just active) to support Time Machine mode
                // enrollment_state can be: active, invited_or_pending, completed
                const coursesUrl = `${CANVAS_DOMAIN}/api/v1/users/${canvasUserId}/courses?include[]=enrollments&per_page=100`;
                const coursesResp = await fetch(coursesUrl, { headers: { 'Accept': 'application/json' } });

                if (coursesResp.ok) {
                    courses = await coursesResp.json();
                    console.log(`✓ Fetched ${courses.length} course(s) for ${student.name || student.SyStudentId}`);
                } else {
                    console.warn(`✗ Failed to fetch courses for ${student.SyStudentId}: ${coursesResp.status} ${coursesResp.statusText}`);
                    courses = [];
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
            const batch = cachedStudents.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;

            console.log(`[Step 2] Cached batch ${batchNumber}/${totalCachedBatches} (students ${i + 1}-${Math.min(i + BATCH_SIZE, cachedStudents.length)})`);

            const promises = batch.map(student => fetchCanvasDetails(student, cacheEnabled));
            const results = await Promise.all(promises);

            results.forEach((updatedStudent, batchIndex) => {
                const originalStudent = batch[batchIndex];
                const originalIndex = studentIndexMap.get(originalStudent);
                updatedStudents[originalIndex] = updatedStudent;
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
            const batch = uncachedStudents.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;

            console.log(`[Step 2] Uncached batch ${batchNumber}/${totalUncachedBatches} (students ${i + 1}-${Math.min(i + BATCH_SIZE, uncachedStudents.length)})`);

            const promises = batch.map(student => fetchCanvasDetails(student, cacheEnabled));
            const results = await Promise.all(promises);

            results.forEach((updatedStudent, batchIndex) => {
                const originalStudent = batch[batchIndex];
                const originalIndex = studentIndexMap.get(originalStudent);
                updatedStudents[originalIndex] = updatedStudent;
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

        const duration = ((Date.now() - startTime) / 1000).toFixed(1);
        step2.className = 'queue-item completed';
        step2.querySelector('i').className = 'fas fa-check';
        timeSpan.textContent = `${duration}s`;

        console.log(`[Step 2] ✓ Complete in ${duration}s - ${students.length} students processed`);

        if (renderCallback) {
            renderCallback(updatedStudents);
        }

        return updatedStudents;

    } catch (error) {
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
    const headers = {
        'Accept': 'application/json',
        'X-Requested-With': 'XMLHttpRequest'
    };

    try {
        const response = await fetch(url, { method: 'GET', credentials: 'include', headers });

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
 * Fetches missing assignments for a single student
 * @param {Date} referenceDate - The date to use for checking if assignments are past due
 */
async function fetchMissingAssignments(student, referenceDate = new Date()) {
    // Support both 'url' field and legacy 'Gradebook' field
    const gradebookUrl = student.url || student.Gradebook;

    if (!gradebookUrl) {
        console.log(`[Step 3] ${student.name}: No gradebook URL, skipping`);
        return { ...student, missingCount: 0, missingAssignments: [] };
    }

    const parsed = parseGradebookUrl(gradebookUrl);
    if (!parsed) {
        console.warn(`[Step 3] ${student.name}: Failed to parse gradebook URL: ${gradebookUrl}`);
        return { ...student, missingCount: 0, missingAssignments: [] };
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

        return {
            ...student,
            missingCount: result.count,
            missingAssignments: result.assignments,
            currentGrade: result.currentGrade
        };

    } catch (e) {
        console.error(`[Step 3] ${student.name}: Error fetching data:`, e);
        return { ...student, missingCount: 0, missingAssignments: [] };
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
        const settings = await chrome.storage.local.get([STORAGE_KEYS.USE_SPECIFIC_DATE, STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE]);
        const useSpecificDate = settings[STORAGE_KEYS.USE_SPECIFIC_DATE] || false;
        const specificDateStr = settings[STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE];

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
            const batch = updatedStudents.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;

            console.log(`[Step 3] Processing batch ${batchNumber}/${totalBatches} (students ${i + 1}-${Math.min(i + BATCH_SIZE, updatedStudents.length)})`);

            const promises = batch.map(student => fetchMissingAssignments(student, referenceDate));
            const results = await Promise.all(promises);

            results.forEach((updatedStudent, index) => {
                updatedStudents[i + index] = updatedStudent;
            });

            processedCount += batch.length;
            timeSpan.textContent = `${Math.round((processedCount / updatedStudents.length) * 100)}%`;

            if (i + BATCH_SIZE < updatedStudents.length) {
                await new Promise(resolve => setTimeout(resolve, BATCH_DELAY_MS));
            }
        }

        await chrome.storage.local.set({ [STORAGE_KEYS.MASTER_ENTRIES]: updatedStudents });

        const duration = ((Date.now() - startTime) / 1000).toFixed(1);
        step3.className = 'queue-item completed';
        step3.querySelector('i').className = 'fas fa-check';
        timeSpan.textContent = `${duration}s`;

        const totalMissing = updatedStudents.reduce((sum, s) => sum + (s.missingCount || 0), 0);
        console.log(`[Step 3] ✓ Complete in ${duration}s - Found ${totalMissing} total missing assignments`);

        if (renderCallback) {
            renderCallback(updatedStudents);
        }

        return updatedStudents;

    } catch (error) {
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

        await sendMasterListWithMissingAssignmentsToExcel(students);

        const duration = ((Date.now() - startTime) / 1000).toFixed(1);
        step4.className = 'queue-item completed';
        step4.querySelector('i').className = 'fas fa-check';
        step4.querySelector('.queue-content').innerHTML = '<i class="fas fa-check"></i> Sending List to Excel';
        timeSpan.textContent = `${duration}s`;

        console.log(`[Step 4] ✓ Complete in ${duration}s`);

        // Calculate and display total completion time
        const queueTotalTimeDiv = document.getElementById('queueTotalTime');
        if (queueTotalTimeDiv && queueTotalTimeDiv.dataset.processStartTime) {
            const totalDuration = ((Date.now() - parseInt(queueTotalTimeDiv.dataset.processStartTime)) / 1000).toFixed(1);
            queueTotalTimeDiv.textContent = `Total Time: ${totalDuration}s`;
            queueTotalTimeDiv.style.display = 'block';
        }

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
