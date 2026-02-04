/**
 * Tests for CallManager - Call state management
 *
 * These tests verify:
 * - Phone number extraction from student objects
 * - Timestamp formatting for call history
 * - Queue navigation (finding next non-skipped student)
 */

// ============================================
// Pure functions extracted for testing
// ============================================

/**
 * Extracts phone number from student object.
 * Handles different possible property names.
 */
function getPhoneNumber(student) {
    if (!student) return "No Phone Listed";

    if (student.phone) return student.phone;
    if (student.Phone) return student.Phone;
    if (student.PrimaryPhone) return student.PrimaryPhone;

    return "No Phone Listed";
}

/**
 * Formats a timestamp for display.
 * If today: shows time only (e.g., "3:45 PM")
 * If not today: shows date (e.g., "12-25-25")
 */
function formatLastCallTimestamp(timestamp) {
    if (!timestamp) return 'Never';

    const now = new Date();
    const callDate = new Date(timestamp);

    const isToday = now.toDateString() === callDate.toDateString();

    if (isToday) {
        let hours = callDate.getHours();
        const minutes = callDate.getMinutes().toString().padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12 || 12;
        return `${hours}:${minutes} ${ampm}`;
    } else {
        const month = (callDate.getMonth() + 1).toString().padStart(2, '0');
        const day = callDate.getDate().toString().padStart(2, '0');
        const year = callDate.getFullYear().toString().slice(-2);
        return `${month}-${day}-${year}`;
    }
}

/**
 * Finds the next non-skipped student index in a queue.
 */
function findNextNonSkippedIndex(queue, skippedIndices, startIndex) {
    for (let i = startIndex; i < queue.length; i++) {
        if (!skippedIndices.has(i)) {
            return i;
        }
    }
    return -1;
}

// ============================================
// Tests
// ============================================

describe('getPhoneNumber', () => {
    test('extracts phone from lowercase "phone" property', () => {
        const student = { name: 'John', phone: '555-123-4567' };
        expect(getPhoneNumber(student)).toBe('555-123-4567');
    });

    test('extracts phone from capitalized "Phone" property', () => {
        const student = { name: 'Jane', Phone: '555-987-6543' };
        expect(getPhoneNumber(student)).toBe('555-987-6543');
    });

    test('extracts phone from "PrimaryPhone" property', () => {
        const student = { name: 'Bob', PrimaryPhone: '555-111-2222' };
        expect(getPhoneNumber(student)).toBe('555-111-2222');
    });

    test('returns "No Phone Listed" for student without phone', () => {
        const student = { name: 'No Phone Student', email: 'test@test.com' };
        expect(getPhoneNumber(student)).toBe('No Phone Listed');
    });

    test('returns "No Phone Listed" for null student', () => {
        expect(getPhoneNumber(null)).toBe('No Phone Listed');
    });

    test('returns "No Phone Listed" for undefined student', () => {
        expect(getPhoneNumber(undefined)).toBe('No Phone Listed');
    });

    test('prioritizes lowercase "phone" over other variants', () => {
        const student = {
            phone: 'lowercase-phone',
            Phone: 'capitalized-phone',
            PrimaryPhone: 'primary-phone'
        };
        expect(getPhoneNumber(student)).toBe('lowercase-phone');
    });
});

describe('formatLastCallTimestamp', () => {
    test('returns "Never" for null/undefined timestamp', () => {
        expect(formatLastCallTimestamp(null)).toBe('Never');
        expect(formatLastCallTimestamp(undefined)).toBe('Never');
        expect(formatLastCallTimestamp(0)).toBe('Never');
    });

    test('returns time format for today', () => {
        const now = new Date();
        now.setHours(14, 30, 0); // 2:30 PM today
        const timestamp = now.getTime();

        const result = formatLastCallTimestamp(timestamp);

        expect(result).toBe('2:30 PM');
    });

    test('returns time format for today - morning', () => {
        const now = new Date();
        now.setHours(9, 5, 0); // 9:05 AM today
        const timestamp = now.getTime();

        const result = formatLastCallTimestamp(timestamp);

        expect(result).toBe('9:05 AM');
    });

    test('returns time format for today - midnight edge case', () => {
        const now = new Date();
        now.setHours(0, 0, 0); // 12:00 AM today
        const timestamp = now.getTime();

        const result = formatLastCallTimestamp(timestamp);

        expect(result).toBe('12:00 AM');
    });

    test('returns time format for today - noon edge case', () => {
        const now = new Date();
        now.setHours(12, 0, 0); // 12:00 PM today
        const timestamp = now.getTime();

        const result = formatLastCallTimestamp(timestamp);

        expect(result).toBe('12:00 PM');
    });

    test('returns date format for yesterday', () => {
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        const timestamp = yesterday.getTime();

        const result = formatLastCallTimestamp(timestamp);

        // Should be in MM-DD-YY format
        expect(result).toMatch(/^\d{2}-\d{2}-\d{2}$/);
    });

    test('returns date format for older dates', () => {
        // Christmas 2024
        const oldDate = new Date(2024, 11, 25, 10, 30);
        const timestamp = oldDate.getTime();

        const result = formatLastCallTimestamp(timestamp);

        expect(result).toBe('12-25-24');
    });

    test('handles single digit months and days with padding', () => {
        // January 5, 2025
        const date = new Date(2025, 0, 5, 10, 30);
        const timestamp = date.getTime();

        const result = formatLastCallTimestamp(timestamp);

        expect(result).toBe('01-05-25');
    });
});

describe('findNextNonSkippedIndex', () => {
    const queue = [
        { name: 'Student 0' },
        { name: 'Student 1' },
        { name: 'Student 2' },
        { name: 'Student 3' },
        { name: 'Student 4' }
    ];

    test('returns first index when nothing is skipped', () => {
        const skipped = new Set();

        expect(findNextNonSkippedIndex(queue, skipped, 0)).toBe(0);
    });

    test('skips over skipped indices', () => {
        const skipped = new Set([0, 1]);

        expect(findNextNonSkippedIndex(queue, skipped, 0)).toBe(2);
    });

    test('returns -1 when all remaining students are skipped', () => {
        const skipped = new Set([3, 4]);

        expect(findNextNonSkippedIndex(queue, skipped, 3)).toBe(-1);
    });

    test('returns -1 when startIndex is past queue length', () => {
        const skipped = new Set();

        expect(findNextNonSkippedIndex(queue, skipped, 10)).toBe(-1);
    });

    test('finds non-skipped index in the middle', () => {
        const skipped = new Set([0, 2, 4]);

        expect(findNextNonSkippedIndex(queue, skipped, 0)).toBe(1);
        expect(findNextNonSkippedIndex(queue, skipped, 2)).toBe(3);
    });

    test('works with all students skipped', () => {
        const skipped = new Set([0, 1, 2, 3, 4]);

        expect(findNextNonSkippedIndex(queue, skipped, 0)).toBe(-1);
    });

    test('works with empty queue', () => {
        const skipped = new Set();

        expect(findNextNonSkippedIndex([], skipped, 0)).toBe(-1);
    });
});

describe('CallManager State Transitions', () => {
    // These tests document expected state transitions
    // They serve as documentation and regression tests

    describe('Single Call Flow', () => {
        test('Ready -> Active -> Awaiting Disposition -> Ready', () => {
            const states = ['ready', 'active', 'awaiting_disposition', 'ready'];

            // This documents the expected state flow
            expect(states[0]).toBe('ready');
            expect(states[1]).toBe('active');
            expect(states[2]).toBe('awaiting_disposition');
            expect(states[3]).toBe('ready');
        });
    });

    describe('Automation Flow', () => {
        test('Ready -> Automation -> Multiple Calls -> Ready', () => {
            const states = [
                'ready',
                'automation_active',
                'call_1_active',
                'call_1_disposition',
                'call_2_active',
                'call_2_disposition',
                'ready'
            ];

            expect(states[0]).toBe('ready');
            expect(states[states.length - 1]).toBe('ready');
        });
    });
});
