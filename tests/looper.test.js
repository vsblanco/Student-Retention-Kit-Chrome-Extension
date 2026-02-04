/**
 * Tests for looper.js - The submission checking engine
 *
 * These tests verify the core logic for:
 * - URL parsing (extracting course/student IDs from Canvas URLs)
 * - Pagination link parsing
 * - Batch preparation
 * - Filter logic
 */

// ============================================
// Constants needed for tests
// ============================================

const CANVAS_SUBDOMAIN = "northbridge";
const LEGACY_CANVAS_SUBDOMAINS = ["nuc"];

// ============================================
// Pure functions extracted for testing
// ============================================

/**
 * Normalizes a Canvas URL to use the current subdomain.
 * Provides backwards compatibility when schools rebrand.
 */
function normalizeCanvasUrl(url) {
    if (!url) return url;

    for (const legacySubdomain of LEGACY_CANVAS_SUBDOMAINS) {
        const legacyPattern = new RegExp(`https://${legacySubdomain}\\.instructure\\.com`, 'gi');
        if (legacyPattern.test(url)) {
            return url.replace(legacyPattern, `https://${CANVAS_SUBDOMAIN}.instructure.com`);
        }
    }

    return url;
}

/**
 * Parses course ID and student ID from a Canvas gradebook URL.
 */
function parseIdsFromUrl(url) {
    try {
        const normalizedUrl = normalizeCanvasUrl(url);
        const urlObj = new URL(normalizedUrl);
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
        // Invalid URL
    }
    return null;
}

/**
 * Extracts the "next" page URL from a Link header for pagination.
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
 * Regex for advanced filter queries like '>=5' or '<10'
 */
const ADVANCED_FILTER_REGEX = /^\s*([><]=?|=)\s*(\d+)\s*$/;

/**
 * Prepares batches of students grouped by course.
 */
function prepareBatches(entries, batchSize = 30) {
    const courses = {};
    const skippedStudents = [];

    entries.forEach(entry => {
        const gradebookUrl = entry.url || entry.Gradebook;
        const parsed = parseIdsFromUrl(gradebookUrl);
        if (parsed) {
            entry.parsed = parsed;
            if (!courses[parsed.courseId]) {
                courses[parsed.courseId] = [];
            }
            courses[parsed.courseId].push(entry);
        } else {
            skippedStudents.push(entry);
        }
    });

    const batches = [];
    Object.values(courses).forEach(courseEntries => {
        for (let i = 0; i < courseEntries.length; i += batchSize) {
            batches.push(courseEntries.slice(i, i + batchSize));
        }
    });

    return { batches, skippedStudents };
}

// ============================================
// Tests
// ============================================

describe('normalizeCanvasUrl', () => {
    test('converts legacy subdomain to current subdomain', () => {
        const legacyUrl = 'https://nuc.instructure.com/courses/123/grades/456';
        const expected = 'https://northbridge.instructure.com/courses/123/grades/456';

        expect(normalizeCanvasUrl(legacyUrl)).toBe(expected);
    });

    test('leaves current subdomain URLs unchanged', () => {
        const currentUrl = 'https://northbridge.instructure.com/courses/123/grades/456';

        expect(normalizeCanvasUrl(currentUrl)).toBe(currentUrl);
    });

    test('handles null/undefined gracefully', () => {
        expect(normalizeCanvasUrl(null)).toBeNull();
        expect(normalizeCanvasUrl(undefined)).toBeUndefined();
        expect(normalizeCanvasUrl('')).toBe('');
    });

    test('is case insensitive for legacy subdomains', () => {
        const upperUrl = 'https://NUC.instructure.com/courses/123/grades/456';
        const expected = 'https://northbridge.instructure.com/courses/123/grades/456';

        expect(normalizeCanvasUrl(upperUrl)).toBe(expected);
    });
});

describe('parseIdsFromUrl', () => {
    test('extracts courseId and studentId from valid URL', () => {
        const url = 'https://northbridge.instructure.com/courses/12345/grades/67890';
        const result = parseIdsFromUrl(url);

        expect(result).toEqual({
            origin: 'https://northbridge.instructure.com',
            courseId: '12345',
            studentId: '67890'
        });
    });

    test('handles legacy subdomain URLs', () => {
        const legacyUrl = 'https://nuc.instructure.com/courses/111/grades/222';
        const result = parseIdsFromUrl(legacyUrl);

        expect(result).toEqual({
            origin: 'https://northbridge.instructure.com',
            courseId: '111',
            studentId: '222'
        });
    });

    test('returns null for invalid URLs', () => {
        expect(parseIdsFromUrl('not-a-url')).toBeNull();
        expect(parseIdsFromUrl('https://example.com/other/path')).toBeNull();
        expect(parseIdsFromUrl('')).toBeNull();
        expect(parseIdsFromUrl(null)).toBeNull();
    });

    test('returns null for Canvas URLs without grades path', () => {
        expect(parseIdsFromUrl('https://northbridge.instructure.com/courses/123')).toBeNull();
        expect(parseIdsFromUrl('https://northbridge.instructure.com/courses/123/assignments')).toBeNull();
    });

    test('handles URLs with query parameters', () => {
        const url = 'https://northbridge.instructure.com/courses/123/grades/456?focus=true';
        const result = parseIdsFromUrl(url);

        expect(result.courseId).toBe('123');
        expect(result.studentId).toBe('456');
    });

    test('handles URLs with fragments', () => {
        const url = 'https://northbridge.instructure.com/courses/123/grades/456#section';
        const result = parseIdsFromUrl(url);

        expect(result.courseId).toBe('123');
        expect(result.studentId).toBe('456');
    });
});

describe('getNextPageUrl', () => {
    test('extracts next URL from Link header', () => {
        const linkHeader = '<https://api.example.com?page=2>; rel="next", <https://api.example.com?page=1>; rel="prev"';

        expect(getNextPageUrl(linkHeader)).toBe('https://api.example.com?page=2');
    });

    test('returns null when no next link exists', () => {
        const linkHeader = '<https://api.example.com?page=1>; rel="prev", <https://api.example.com?page=1>; rel="first"';

        expect(getNextPageUrl(linkHeader)).toBeNull();
    });

    test('returns null for empty/null header', () => {
        expect(getNextPageUrl(null)).toBeNull();
        expect(getNextPageUrl('')).toBeNull();
        expect(getNextPageUrl(undefined)).toBeNull();
    });

    test('handles Link header with only next link', () => {
        const linkHeader = '<https://api.example.com?page=2>; rel="next"';

        expect(getNextPageUrl(linkHeader)).toBe('https://api.example.com?page=2');
    });

    test('handles complex URLs in Link header', () => {
        const linkHeader = '<https://canvas.instructure.com/api/v1/courses/123/students/submissions?student_ids[]=1&page=2&per_page=100>; rel="next"';

        expect(getNextPageUrl(linkHeader)).toBe('https://canvas.instructure.com/api/v1/courses/123/students/submissions?student_ids[]=1&page=2&per_page=100');
    });
});

describe('ADVANCED_FILTER_REGEX', () => {
    test('matches >= operator', () => {
        const match = '>=5'.match(ADVANCED_FILTER_REGEX);
        expect(match[1]).toBe('>=');
        expect(match[2]).toBe('5');
    });

    test('matches <= operator', () => {
        const match = '<=10'.match(ADVANCED_FILTER_REGEX);
        expect(match[1]).toBe('<=');
        expect(match[2]).toBe('10');
    });

    test('matches > operator', () => {
        const match = '>3'.match(ADVANCED_FILTER_REGEX);
        expect(match[1]).toBe('>');
        expect(match[2]).toBe('3');
    });

    test('matches < operator', () => {
        const match = '<7'.match(ADVANCED_FILTER_REGEX);
        expect(match[1]).toBe('<');
        expect(match[2]).toBe('7');
    });

    test('matches = operator', () => {
        const match = '=5'.match(ADVANCED_FILTER_REGEX);
        expect(match[1]).toBe('=');
        expect(match[2]).toBe('5');
    });

    test('handles whitespace', () => {
        const match = '  >= 5  '.match(ADVANCED_FILTER_REGEX);
        expect(match[1]).toBe('>=');
        expect(match[2]).toBe('5');
    });

    test('does not match invalid formats', () => {
        expect('all'.match(ADVANCED_FILTER_REGEX)).toBeNull();
        expect('five'.match(ADVANCED_FILTER_REGEX)).toBeNull();
        expect('>='.match(ADVANCED_FILTER_REGEX)).toBeNull();
        expect('5'.match(ADVANCED_FILTER_REGEX)).toBeNull();
    });
});

describe('prepareBatches', () => {
    test('groups students by course', () => {
        const entries = [
            { name: 'Student A', url: 'https://northbridge.instructure.com/courses/100/grades/1' },
            { name: 'Student B', url: 'https://northbridge.instructure.com/courses/100/grades/2' },
            { name: 'Student C', url: 'https://northbridge.instructure.com/courses/200/grades/3' }
        ];

        const { batches } = prepareBatches(entries);

        // Should have 2 batches (one per course)
        expect(batches.length).toBe(2);

        // Find the batch for course 100
        const course100Batch = batches.find(b => b[0].parsed.courseId === '100');
        expect(course100Batch.length).toBe(2);
    });

    test('splits large courses into multiple batches', () => {
        // Create 35 students in same course (batch size is 30)
        const entries = Array.from({ length: 35 }, (_, i) => ({
            name: `Student ${i}`,
            url: `https://northbridge.instructure.com/courses/100/grades/${i}`
        }));

        const { batches } = prepareBatches(entries, 30);

        // Should split into 2 batches (30 + 5)
        expect(batches.length).toBe(2);
        expect(batches[0].length).toBe(30);
        expect(batches[1].length).toBe(5);
    });

    test('handles legacy Gradebook field name', () => {
        const entries = [
            { name: 'Student A', Gradebook: 'https://northbridge.instructure.com/courses/100/grades/1' }
        ];

        const { batches } = prepareBatches(entries);

        expect(batches.length).toBe(1);
        expect(batches[0][0].parsed.courseId).toBe('100');
    });

    test('skips students with invalid URLs', () => {
        const entries = [
            { name: 'Valid Student', url: 'https://northbridge.instructure.com/courses/100/grades/1' },
            { name: 'Invalid Student', url: 'not-a-valid-url' },
            { name: 'Missing URL Student' }
        ];

        const { batches, skippedStudents } = prepareBatches(entries);

        expect(batches.length).toBe(1);
        expect(batches[0].length).toBe(1);
        expect(skippedStudents.length).toBe(2);
    });

    test('returns empty batches for empty input', () => {
        const { batches, skippedStudents } = prepareBatches([]);

        expect(batches.length).toBe(0);
        expect(skippedStudents.length).toBe(0);
    });
});
