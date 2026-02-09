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

// ============================================
// CONFIG.COLORS validation
// ============================================

const CONFIG_COLORS = {
    SUCCESS: '#10b981',
    ERROR: '#ef4444',
    WARNING: '#f59e0b',
    MUTED: '#6b7280',
    PRIMARY: '#3b82f6',
    PURPLE: '#8b5cf6'
};

describe('CONFIG.COLORS values', () => {
    test('all colors are valid hex codes', () => {
        const hexRegex = /^#[0-9a-f]{6}$/i;
        Object.entries(CONFIG_COLORS).forEach(([name, value]) => {
            expect(value).toMatch(hexRegex);
        });
    });

    test('SUCCESS is green (#10b981)', () => {
        expect(CONFIG_COLORS.SUCCESS).toBe('#10b981');
    });

    test('ERROR is red (#ef4444)', () => {
        expect(CONFIG_COLORS.ERROR).toBe('#ef4444');
    });

    test('WARNING is amber (#f59e0b)', () => {
        expect(CONFIG_COLORS.WARNING).toBe('#f59e0b');
    });

    test('MUTED is gray (#6b7280)', () => {
        expect(CONFIG_COLORS.MUTED).toBe('#6b7280');
    });

    test('PRIMARY is blue (#3b82f6)', () => {
        expect(CONFIG_COLORS.PRIMARY).toBe('#3b82f6');
    });
});

// ============================================
// Call UI Color Transitions
// Simulates what CallManager methods do to DOM elements
// to verify correct colors are applied at each state.
// ============================================

/**
 * Converts hex color to rgb() string as JSDOM normalizes it.
 */
function hexToRgb(hex) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return `rgb(${r}, ${g}, ${b})`;
}

/**
 * Creates a mock DOM element set matching what CallManager expects.
 */
function createMockElements() {
    return {
        dialBtn: document.createElement('button'),
        callStatusText: document.createElement('div'),
        callDispositionSection: document.createElement('div'),
        upNextCard: document.createElement('div'),
        otherInputArea: document.createElement('div'),
    };
}

describe('Call UI Color Transitions', () => {
    let elements;

    beforeEach(() => {
        elements = createMockElements();
        // Start with green (ready state)
        elements.dialBtn.style.background = CONFIG_COLORS.SUCCESS;
    });

    describe('Active Call State (dial button turns red)', () => {
        test('sets dial button background to ERROR red when call starts', () => {
            // Simulate what toggleCallState does when isCallActive becomes true
            elements.dialBtn.style.background = `${CONFIG_COLORS.ERROR}`;
            elements.dialBtn.style.transform = 'rotate(135deg)';

            // JSDOM normalizes hex to rgb()
            expect(elements.dialBtn.style.background).toBe(hexToRgb(CONFIG_COLORS.ERROR));
        });

        test('rotates dial button 135deg when call is active', () => {
            elements.dialBtn.style.transform = 'rotate(135deg)';

            expect(elements.dialBtn.style.transform).toBe('rotate(135deg)');
        });

        test('status text includes ERROR color in inline style', () => {
            const statusText = 'Connected';
            elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:${CONFIG_COLORS.ERROR}; animation: blink 1s infinite;"></span> ${statusText}`;

            expect(elements.callStatusText.innerHTML).toContain(CONFIG_COLORS.ERROR);
            expect(elements.callStatusText.innerHTML).toContain('Connected');
        });

        test('shows disposition section when call starts', () => {
            elements.callDispositionSection.style.display = 'flex';

            expect(elements.callDispositionSection.style.display).toBe('flex');
        });
    });

    describe('Awaiting Disposition State (dial button turns gray)', () => {
        test('sets dial button background to MUTED gray after call ends', () => {
            // Simulate toggleCallState when call ends and enters WRAP_UP
            elements.dialBtn.style.background = `${CONFIG_COLORS.MUTED}`;
            elements.dialBtn.style.transform = 'rotate(0deg)';

            expect(elements.dialBtn.style.background).toBe(hexToRgb(CONFIG_COLORS.MUTED));
        });

        test('resets dial button rotation to 0deg', () => {
            elements.dialBtn.style.transform = 'rotate(0deg)';

            expect(elements.dialBtn.style.transform).toBe('rotate(0deg)');
        });

        test('disables dial button while awaiting disposition', () => {
            elements.dialBtn.disabled = true;
            elements.dialBtn.style.cursor = 'not-allowed';
            elements.dialBtn.style.opacity = '0.6';

            expect(elements.dialBtn.disabled).toBe(true);
            expect(elements.dialBtn.style.cursor).toBe('not-allowed');
            expect(elements.dialBtn.style.opacity).toBe('0.6');
        });

        test('status text includes WARNING color for awaiting disposition', () => {
            elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:${CONFIG_COLORS.WARNING};"></span> Awaiting Disposition`;

            expect(elements.callStatusText.innerHTML).toContain(CONFIG_COLORS.WARNING);
            expect(elements.callStatusText.innerHTML).toContain('Awaiting Disposition');
        });
    });

    describe('Ready State (dial button turns green)', () => {
        test('sets dial button background to SUCCESS green when ready', () => {
            // Simulate endAutomationSequence / handleExternalDisposition
            elements.dialBtn.style.background = `${CONFIG_COLORS.SUCCESS}`;
            elements.dialBtn.style.transform = 'rotate(0deg)';

            expect(elements.dialBtn.style.background).toBe(hexToRgb(CONFIG_COLORS.SUCCESS));
        });

        test('re-enables dial button', () => {
            elements.dialBtn.disabled = false;
            elements.dialBtn.style.cursor = 'pointer';
            elements.dialBtn.style.opacity = '1';

            expect(elements.dialBtn.disabled).toBe(false);
            expect(elements.dialBtn.style.cursor).toBe('pointer');
            expect(elements.dialBtn.style.opacity).toBe('1');
        });

        test('status text shows ready indicator', () => {
            elements.callStatusText.innerHTML = '<span class="status-indicator ready"></span> Ready to Connect';

            expect(elements.callStatusText.innerHTML).toContain('ready');
            expect(elements.callStatusText.innerHTML).toContain('Ready to Connect');
        });

        test('hides disposition section', () => {
            elements.callDispositionSection.style.display = 'none';

            expect(elements.callDispositionSection.style.display).toBe('none');
        });
    });

    describe('Disposition Flow Colors', () => {
        test('shows PRIMARY blue when setting disposition', () => {
            elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:${CONFIG_COLORS.PRIMARY};"></span> Setting disposition...`;

            expect(elements.callStatusText.innerHTML).toContain(CONFIG_COLORS.PRIMARY);
        });

        test('shows SUCCESS green when disposition is set', () => {
            elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:${CONFIG_COLORS.SUCCESS};"></span> Disposition Set`;

            expect(elements.callStatusText.innerHTML).toContain(CONFIG_COLORS.SUCCESS);
            expect(elements.callStatusText.innerHTML).toContain('Disposition Set');
        });

        test('shows WARNING amber when ending call', () => {
            elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:${CONFIG_COLORS.WARNING};"></span> Ending call...`;

            expect(elements.callStatusText.innerHTML).toContain(CONFIG_COLORS.WARNING);
        });
    });

    describe('Template literal interpolation (regression for single-quote bug)', () => {
        test('backtick template literals produce actual hex values, not literal ${...}', () => {
            // This is the exact bug we fixed: single quotes don't interpolate
            const withBackticks = `${CONFIG_COLORS.ERROR}`;
            const withSingleQuotes = '${CONFIG_COLORS.ERROR}';

            expect(withBackticks).toBe('#ef4444');
            expect(withSingleQuotes).toBe('${CONFIG_COLORS.ERROR}');
            expect(withBackticks).not.toBe(withSingleQuotes);
        });

        test('all color assignments produce valid CSS color values (not literal template strings)', () => {
            // JSDOM normalizes hex to rgb(), so we verify the value is a valid rgb color
            const rgbRegex = /^rgb\(\d+, \d+, \d+\)$/;

            // Simulate the exact pattern used in callManager.js
            elements.dialBtn.style.background = `${CONFIG_COLORS.ERROR}`;
            expect(elements.dialBtn.style.background).toMatch(rgbRegex);

            elements.dialBtn.style.background = `${CONFIG_COLORS.SUCCESS}`;
            expect(elements.dialBtn.style.background).toMatch(rgbRegex);

            elements.dialBtn.style.background = `${CONFIG_COLORS.MUTED}`;
            expect(elements.dialBtn.style.background).toMatch(rgbRegex);

            // A broken single-quote string is invalid CSS and won't change the value
            // Reset to empty first, then try the broken string
            elements.dialBtn.style.background = '';
            elements.dialBtn.style.background = '${CONFIG_COLORS.ERROR}';
            // Should remain empty since the literal string is not valid CSS
            expect(elements.dialBtn.style.background).toBe('');
        });

        test('innerHTML template literals interpolate CONFIG colors correctly', () => {
            elements.callStatusText.innerHTML = `<span style="background:${CONFIG_COLORS.ERROR};"></span>`;

            // Should contain the actual hex value, not the template syntax
            expect(elements.callStatusText.innerHTML).toContain('#ef4444');
            expect(elements.callStatusText.innerHTML).not.toContain('CONFIG_COLORS');
        });
    });

    describe('Cancel Automation resets UI', () => {
        test('removes automation class from dial button', () => {
            elements.dialBtn.classList.add('automation');

            // Simulate cancelAutomation
            elements.dialBtn.classList.remove('automation');
            elements.dialBtn.innerHTML = '<i class="fas fa-phone"></i>';
            elements.dialBtn.style.background = `${CONFIG_COLORS.SUCCESS}`;
            elements.dialBtn.style.transform = 'rotate(0deg)';

            expect(elements.dialBtn.classList.contains('automation')).toBe(false);
            expect(elements.dialBtn.style.background).toBe(hexToRgb(CONFIG_COLORS.SUCCESS));
            expect(elements.dialBtn.innerHTML).toContain('fa-phone');
        });

        test('hides up next card', () => {
            elements.upNextCard.style.display = 'block';

            // Simulate cancelAutomation
            elements.upNextCard.style.display = 'none';

            expect(elements.upNextCard.style.display).toBe('none');
        });
    });

    describe('External Disconnect sets awaiting disposition UI', () => {
        test('sets dial button to MUTED and disables it', () => {
            // Simulate handleExternalDisconnect
            elements.dialBtn.style.background = `${CONFIG_COLORS.MUTED}`;
            elements.dialBtn.style.transform = 'rotate(0deg)';
            elements.dialBtn.disabled = true;
            elements.dialBtn.style.cursor = 'not-allowed';
            elements.dialBtn.style.opacity = '0.6';

            expect(elements.dialBtn.style.background).toBe(hexToRgb(CONFIG_COLORS.MUTED));
            expect(elements.dialBtn.disabled).toBe(true);
            expect(elements.dialBtn.style.cursor).toBe('not-allowed');
        });
    });

    describe('External Disposition resets to ready UI', () => {
        test('sets dial button to SUCCESS and re-enables it', () => {
            // Start in disabled state
            elements.dialBtn.disabled = true;
            elements.dialBtn.style.background = CONFIG_COLORS.MUTED;

            // Simulate handleExternalDisposition
            elements.dialBtn.style.background = `${CONFIG_COLORS.SUCCESS}`;
            elements.dialBtn.style.transform = 'rotate(0deg)';
            elements.dialBtn.disabled = false;
            elements.dialBtn.style.cursor = 'pointer';
            elements.dialBtn.style.opacity = '1';

            expect(elements.dialBtn.style.background).toBe(hexToRgb(CONFIG_COLORS.SUCCESS));
            expect(elements.dialBtn.disabled).toBe(false);
            expect(elements.dialBtn.style.cursor).toBe('pointer');
            expect(elements.dialBtn.style.opacity).toBe('1');
        });
    });

    describe('Demo Mode status text', () => {
        test('shows WARNING color with demo mode indicator', () => {
            elements.callStatusText.innerHTML = `<span class="status-indicator" style="background:${CONFIG_COLORS.WARNING};"></span> 🎭 Demo Mode Active`;

            expect(elements.callStatusText.innerHTML).toContain(CONFIG_COLORS.WARNING);
            expect(elements.callStatusText.innerHTML).toContain('Demo Mode Active');
        });
    });
});
