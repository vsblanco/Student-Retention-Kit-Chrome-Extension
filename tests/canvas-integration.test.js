/**
 * Tests for canvas-integration.js - Canvas API analysis and data processing
 *
 * These tests verify the core logic for:
 * - Grade extraction from enrollment data
 * - Missing assignment detection and classification
 * - Next upcoming assignment finding
 * - Course grouping for batch API calls
 * - Gradebook URL parsing
 */

// ============================================
// Pure functions extracted for testing
// ============================================

/**
 * Extracts the current grade from a Canvas user enrollment object.
 * (Extracted from canvas-integration.js)
 */
function extractCurrentGrade(userObject) {
    if (!userObject || !userObject.enrollments) return '';

    const enrollment = userObject.enrollments.find(e => e.type === 'StudentEnrollment')
        || userObject.enrollments[0];

    if (!enrollment || !enrollment.grades) return '';

    const { current_score, final_score, current_grade } = enrollment.grades;
    if (current_score != null) return current_score;
    if (final_score != null) return final_score;
    if (current_grade != null) return String(current_grade).replace(/%/g, '');
    return '';
}

/**
 * Analyzes submissions to find missing assignments.
 * (Extracted from canvas-integration.js)
 */
function analyzeMissingAssignments(submissions, userObject, studentName, courseId, origin, referenceDate = new Date()) {
    const now = referenceDate;
    const currentGrade = extractCurrentGrade(userObject);
    const collectedAssignments = [];

    for (const sub of submissions) {
        const dueDate = sub.cached_due_date ? new Date(sub.cached_due_date) : null;
        if (dueDate && dueDate > now) continue;

        const scoreStr = String(sub.score || sub.grade || '').toLowerCase();
        if (scoreStr === 'complete') continue;

        const isMissing = (sub.missing === true) ||
            ((sub.workflow_state === 'unsubmitted' || sub.workflow_state === 'unsubmitted (ungraded)') && dueDate && dueDate < now) ||
            (sub.score === 0);

        if (!isMissing) continue;

        const assignmentId = sub.assignment ? sub.assignment.id : null;
        const assignmentUrl = assignmentId ? `${origin}/courses/${courseId}/assignments/${assignmentId}` : '';

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
            dueDate: dueDate ? dueDate.toLocaleDateString() : 'No Date',
            score: formattedScore,
            workflow_state: sub.workflow_state
        });
    }

    return { currentGrade, count: collectedAssignments.length, assignments: collectedAssignments };
}

/**
 * Finds the next upcoming unsubmitted assignment.
 * (Extracted from canvas-integration.js)
 */
function findNextAssignment(submissions, courseId, origin, referenceDate = new Date()) {
    const todayStart = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), referenceDate.getDate());
    const tomorrow = new Date(todayStart);
    tomorrow.setDate(tomorrow.getDate() + 1);

    const upcomingAssignments = [];

    for (const sub of submissions) {
        if (!sub.assignment) continue;

        const dueDate = sub.assignment.due_at ? new Date(sub.assignment.due_at) : null;
        if (!dueDate || dueDate < todayStart) continue;

        const submittedStates = ['submitted', 'graded', 'pending_review'];
        const isSubmitted = submittedStates.includes(sub.workflow_state)
            || (sub.submitted_at != null && sub.submitted_at !== '')
            || (sub.score != null && sub.score > 0);

        const scoreStr = String(sub.score || sub.grade || '').toLowerCase();
        if (isSubmitted || scoreStr === 'complete') continue;

        const dueDateStart = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
        let formattedDueDate;
        if (dueDateStart.getTime() === todayStart.getTime()) {
            formattedDueDate = 'Today';
        } else if (dueDateStart.getTime() === tomorrow.getTime()) {
            formattedDueDate = 'Tomorrow';
        } else {
            formattedDueDate = dueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        }

        const assignmentId = sub.assignment.id;
        upcomingAssignments.push({
            Assignment: sub.assignment.name || 'Unknown Assignment',
            DueDate: formattedDueDate,
            AssignmentLink: assignmentId ? `${origin}/courses/${courseId}/assignments/${assignmentId}` : '',
            _dueDateObj: dueDate
        });
    }

    if (upcomingAssignments.length === 0) return null;

    upcomingAssignments.sort((a, b) => a._dueDateObj - b._dueDateObj);
    const next = upcomingAssignments[0];
    delete next._dueDateObj;
    return next;
}

/**
 * Parses a Canvas gradebook URL.
 * (Extracted from canvas-integration.js)
 */
function parseGradebookUrl(url) {
    try {
        const urlObj = new URL(url);
        const match = urlObj.pathname.match(/courses\/(\d+)\/grades\/(\d+)/);
        if (match) {
            return { origin: urlObj.origin, courseId: match[1], studentId: match[2] };
        }
    } catch (e) {
        // Invalid URL
    }
    return null;
}

/**
 * Formats duration in seconds.
 * (Extracted from canvas-integration.js)
 */
function formatDuration(seconds) {
    if (seconds >= 60) {
        return `${(seconds / 60).toFixed(1)}m`;
    }
    return `${seconds.toFixed(1)}s`;
}

// ============================================
// Tests
// ============================================

describe('extractCurrentGrade', () => {
    test('extracts current_score from StudentEnrollment', () => {
        const user = {
            enrollments: [
                { type: 'StudentEnrollment', grades: { current_score: 85 } }
            ]
        };
        expect(extractCurrentGrade(user)).toBe(85);
    });

    test('falls back to final_score when current_score is null', () => {
        const user = {
            enrollments: [
                { type: 'StudentEnrollment', grades: { current_score: null, final_score: 78 } }
            ]
        };
        expect(extractCurrentGrade(user)).toBe(78);
    });

    test('falls back to current_grade string with % stripped', () => {
        const user = {
            enrollments: [
                { type: 'StudentEnrollment', grades: { current_score: null, final_score: null, current_grade: '92%' } }
            ]
        };
        expect(extractCurrentGrade(user)).toBe('92');
    });

    test('returns empty string when no grades available', () => {
        const user = { enrollments: [{ type: 'StudentEnrollment', grades: {} }] };
        expect(extractCurrentGrade(user)).toBe('');
    });

    test('returns empty string for null/undefined user', () => {
        expect(extractCurrentGrade(null)).toBe('');
        expect(extractCurrentGrade(undefined)).toBe('');
    });

    test('returns empty string when no enrollments', () => {
        expect(extractCurrentGrade({ enrollments: [] })).toBe('');
        expect(extractCurrentGrade({})).toBe('');
    });

    test('falls back to first enrollment when no StudentEnrollment', () => {
        const user = {
            enrollments: [
                { type: 'ObserverEnrollment', grades: { current_score: 70 } }
            ]
        };
        expect(extractCurrentGrade(user)).toBe(70);
    });

    test('handles zero score correctly', () => {
        const user = {
            enrollments: [
                { type: 'StudentEnrollment', grades: { current_score: 0 } }
            ]
        };
        expect(extractCurrentGrade(user)).toBe(0);
    });
});

describe('analyzeMissingAssignments', () => {
    const origin = 'https://northbridge.instructure.com';
    const courseId = '12345';
    const refDate = new Date(2026, 1, 5); // Feb 5, 2026

    test('detects assignment flagged as missing by Canvas', () => {
        const submissions = [{
            missing: true,
            assignment: { id: 1, name: 'Essay 1', points_possible: 100 },
            score: null,
            grade: null,
            workflow_state: 'unsubmitted',
            cached_due_date: '2026-02-01T00:00:00Z',
            preview_url: ''
        }];

        const result = analyzeMissingAssignments(submissions, null, 'Jane', courseId, origin, refDate);
        expect(result.count).toBe(1);
        expect(result.assignments[0].assignmentTitle).toBe('Essay 1');
    });

    test('detects unsubmitted past-due assignment', () => {
        const submissions = [{
            missing: false,
            assignment: { id: 2, name: 'Homework 3', points_possible: 50 },
            score: null,
            grade: null,
            workflow_state: 'unsubmitted',
            cached_due_date: '2026-01-15T00:00:00Z',
            preview_url: ''
        }];

        const result = analyzeMissingAssignments(submissions, null, 'Jane', courseId, origin, refDate);
        expect(result.count).toBe(1);
    });

    test('detects assignment with score of 0', () => {
        const submissions = [{
            missing: false,
            assignment: { id: 3, name: 'Quiz', points_possible: 20 },
            score: 0,
            grade: null,
            workflow_state: 'graded',
            cached_due_date: '2026-01-20T00:00:00Z',
            preview_url: ''
        }];

        const result = analyzeMissingAssignments(submissions, null, 'Jane', courseId, origin, refDate);
        expect(result.count).toBe(1);
        expect(result.assignments[0].score).toBe('0/20');
    });

    test('skips future assignments', () => {
        const submissions = [{
            missing: true,
            assignment: { id: 4, name: 'Future Essay', points_possible: 100 },
            score: null,
            workflow_state: 'unsubmitted',
            cached_due_date: '2026-03-01T00:00:00Z'
        }];

        const result = analyzeMissingAssignments(submissions, null, 'Jane', courseId, origin, refDate);
        expect(result.count).toBe(0);
    });

    test('skips completed assignments', () => {
        const submissions = [{
            missing: false,
            assignment: { id: 5, name: 'Completed Work' },
            score: 'complete',
            grade: 'complete',
            workflow_state: 'graded',
            cached_due_date: '2026-01-10T00:00:00Z'
        }];

        const result = analyzeMissingAssignments(submissions, null, 'Jane', courseId, origin, refDate);
        expect(result.count).toBe(0);
    });

    test('includes current grade from user object', () => {
        const user = {
            enrollments: [{ type: 'StudentEnrollment', grades: { current_score: 92 } }]
        };

        const result = analyzeMissingAssignments([], user, 'Jane', courseId, origin, refDate);
        expect(result.currentGrade).toBe(92);
        expect(result.count).toBe(0);
    });

    test('builds correct assignment URL', () => {
        const submissions = [{
            missing: true,
            assignment: { id: 999, name: 'Test', points_possible: 10 },
            score: null,
            workflow_state: 'unsubmitted',
            cached_due_date: '2026-01-01T00:00:00Z'
        }];

        const result = analyzeMissingAssignments(submissions, null, 'Jane', courseId, origin, refDate);
        expect(result.assignments[0].assignmentLink).toBe(
            'https://northbridge.instructure.com/courses/12345/assignments/999'
        );
    });

    test('handles empty submissions array', () => {
        const result = analyzeMissingAssignments([], null, 'Jane', courseId, origin, refDate);
        expect(result.count).toBe(0);
        expect(result.assignments).toEqual([]);
    });
});

describe('findNextAssignment', () => {
    const origin = 'https://northbridge.instructure.com';
    const courseId = '100';
    const refDate = new Date(2026, 1, 5); // Feb 5, 2026

    test('finds the nearest upcoming unsubmitted assignment', () => {
        const submissions = [
            {
                assignment: { id: 1, name: 'Far Assignment', due_at: '2026-03-01T00:00:00Z' },
                workflow_state: 'unsubmitted', submitted_at: null, score: null
            },
            {
                assignment: { id: 2, name: 'Near Assignment', due_at: '2026-02-06T00:00:00Z' },
                workflow_state: 'unsubmitted', submitted_at: null, score: null
            }
        ];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result).not.toBeNull();
        expect(result.Assignment).toBe('Near Assignment');
    });

    test('labels assignment due today as "Today"', () => {
        const submissions = [{
            assignment: { id: 1, name: 'Due Today', due_at: '2026-02-05T23:59:00Z' },
            workflow_state: 'unsubmitted', submitted_at: null, score: null
        }];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result.DueDate).toBe('Today');
    });

    test('labels assignment due tomorrow as "Tomorrow"', () => {
        const submissions = [{
            assignment: { id: 1, name: 'Due Tomorrow', due_at: '2026-02-06T23:59:00Z' },
            workflow_state: 'unsubmitted', submitted_at: null, score: null
        }];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result.DueDate).toBe('Tomorrow');
    });

    test('skips already-submitted assignments', () => {
        const submissions = [{
            assignment: { id: 1, name: 'Submitted', due_at: '2026-02-10T00:00:00Z' },
            workflow_state: 'submitted', submitted_at: '2026-02-04T00:00:00Z', score: null
        }];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result).toBeNull();
    });

    test('skips graded assignments', () => {
        const submissions = [{
            assignment: { id: 1, name: 'Graded', due_at: '2026-02-10T00:00:00Z' },
            workflow_state: 'graded', submitted_at: '2026-02-04T00:00:00Z', score: 85
        }];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result).toBeNull();
    });

    test('skips past-due assignments', () => {
        const submissions = [{
            assignment: { id: 1, name: 'Past Due', due_at: '2026-01-01T00:00:00Z' },
            workflow_state: 'unsubmitted', submitted_at: null, score: null
        }];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result).toBeNull();
    });

    test('returns null for empty submissions', () => {
        expect(findNextAssignment([], courseId, origin, refDate)).toBeNull();
    });

    test('includes correct assignment link', () => {
        const submissions = [{
            assignment: { id: 42, name: 'Essay', due_at: '2026-02-10T00:00:00Z' },
            workflow_state: 'unsubmitted', submitted_at: null, score: null
        }];

        const result = findNextAssignment(submissions, courseId, origin, refDate);
        expect(result.AssignmentLink).toBe(
            'https://northbridge.instructure.com/courses/100/assignments/42'
        );
    });
});

describe('parseGradebookUrl', () => {
    test('extracts course and student IDs', () => {
        const result = parseGradebookUrl('https://northbridge.instructure.com/courses/123/grades/456');
        expect(result).toEqual({
            origin: 'https://northbridge.instructure.com',
            courseId: '123',
            studentId: '456'
        });
    });

    test('returns null for non-gradebook URL', () => {
        expect(parseGradebookUrl('https://example.com/other')).toBeNull();
        expect(parseGradebookUrl('https://northbridge.instructure.com/courses/123')).toBeNull();
    });

    test('returns null for invalid input', () => {
        expect(parseGradebookUrl('not-a-url')).toBeNull();
        expect(parseGradebookUrl('')).toBeNull();
        expect(parseGradebookUrl(null)).toBeNull();
    });
});

describe('formatDuration', () => {
    test('formats seconds under 60', () => {
        expect(formatDuration(5.123)).toBe('5.1s');
        expect(formatDuration(0.5)).toBe('0.5s');
        expect(formatDuration(59.9)).toBe('59.9s');
    });

    test('formats seconds >= 60 as minutes', () => {
        expect(formatDuration(60)).toBe('1.0m');
        expect(formatDuration(90)).toBe('1.5m');
        expect(formatDuration(300)).toBe('5.0m');
    });
});

describe('Course grouping logic', () => {
    // This tests the grouping logic used by processStep3

    function groupStudentsByCourse(students) {
        const courseGroups = new Map();
        const noUrlIndices = [];

        for (let i = 0; i < students.length; i++) {
            const student = students[i];
            const gradebookUrl = student.url || student.Gradebook;
            if (!gradebookUrl) { noUrlIndices.push(i); continue; }

            const parsed = parseGradebookUrl(gradebookUrl);
            if (!parsed) { noUrlIndices.push(i); continue; }

            const groupKey = `${parsed.origin}|${parsed.courseId}`;
            if (!courseGroups.has(groupKey)) courseGroups.set(groupKey, []);
            courseGroups.get(groupKey).push({ student, parsed, originalIndex: i });
        }

        return { courseGroups, noUrlIndices };
    }

    test('groups students sharing the same course', () => {
        const students = [
            { name: 'Alice', url: 'https://canvas.com/courses/100/grades/1' },
            { name: 'Bob', url: 'https://canvas.com/courses/100/grades/2' },
            { name: 'Carol', url: 'https://canvas.com/courses/200/grades/3' }
        ];

        const { courseGroups, noUrlIndices } = groupStudentsByCourse(students);
        expect(courseGroups.size).toBe(2);
        expect(noUrlIndices.length).toBe(0);

        const course100 = courseGroups.get('https://canvas.com|100');
        expect(course100.length).toBe(2);

        const course200 = courseGroups.get('https://canvas.com|200');
        expect(course200.length).toBe(1);
    });

    test('separates students without URLs', () => {
        const students = [
            { name: 'Alice', url: 'https://canvas.com/courses/100/grades/1' },
            { name: 'Bob' },
            { name: 'Carol', url: '' }
        ];

        const { courseGroups, noUrlIndices } = groupStudentsByCourse(students);
        expect(courseGroups.size).toBe(1);
        expect(noUrlIndices).toEqual([1, 2]);
    });

    test('handles all students without URLs', () => {
        const students = [
            { name: 'Alice' },
            { name: 'Bob', url: 'invalid-url' }
        ];

        const { courseGroups, noUrlIndices } = groupStudentsByCourse(students);
        expect(courseGroups.size).toBe(0);
        expect(noUrlIndices.length).toBe(2);
    });

    test('preserves original indices for result mapping', () => {
        const students = [
            { name: 'Skip', url: '' },                                              // index 0
            { name: 'Alice', url: 'https://canvas.com/courses/100/grades/1' },      // index 1
            { name: 'Skip2' },                                                       // index 2
            { name: 'Bob', url: 'https://canvas.com/courses/100/grades/2' }          // index 3
        ];

        const { courseGroups } = groupStudentsByCourse(students);
        const group = courseGroups.get('https://canvas.com|100');
        expect(group[0].originalIndex).toBe(1);
        expect(group[1].originalIndex).toBe(3);
    });

    test('supports legacy Gradebook field name', () => {
        const students = [
            { name: 'Legacy', Gradebook: 'https://canvas.com/courses/300/grades/5' }
        ];

        const { courseGroups } = groupStudentsByCourse(students);
        expect(courseGroups.size).toBe(1);
    });
});
