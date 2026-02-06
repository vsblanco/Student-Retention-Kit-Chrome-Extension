/**
 * Tests for canvasCache.js - Canvas API caching layer
 *
 * These tests verify the core logic for:
 * - Cache data filtering (storing only essential fields)
 * - Cache expiration calculation from course end dates
 * - Staged (batched) cache writes
 * - Cache hit/miss and expiration checking
 */

// ============================================
// Pure functions extracted for testing
// ============================================

/**
 * Filters user data to only cache essential fields.
 * (Extracted from canvasCache.js)
 */
function filterUserData(userData) {
    if (!userData) return {};
    return {
        id: userData.id,
        name: userData.name,
        sortable_name: userData.sortable_name,
        avatar_url: userData.avatar_url,
        created_at: userData.created_at
    };
}

/**
 * Filters course data to only cache essential fields.
 * (Extracted from canvasCache.js)
 */
function filterCourses(courses) {
    if (!courses || !Array.isArray(courses)) return [];
    return courses.map(course => ({
        id: course.id,
        name: course.name,
        start_at: course.start_at,
        end_at: course.end_at,
        enrollments: course.enrollments ? course.enrollments.map(enrollment => ({
            type: enrollment.type,
            grades: enrollment.grades ? {
                current_score: enrollment.grades.current_score
            } : null
        })) : []
    }));
}

/**
 * Calculates cache expiration date from course end dates.
 * (Extracted from canvasCache.js setCachedData/stageCacheData)
 */
function calculateExpiration(filteredCourses) {
    let latestEndDate = null;

    if (filteredCourses && filteredCourses.length > 0) {
        for (const course of filteredCourses) {
            if (course.end_at) {
                const endDate = new Date(course.end_at);
                if (!latestEndDate || endDate > latestEndDate) {
                    latestEndDate = endDate;
                }
            }
        }
    }

    if (!latestEndDate) {
        latestEndDate = new Date();
        latestEndDate.setDate(latestEndDate.getDate() + 30);
    }

    return latestEndDate;
}

/**
 * Checks if a cache entry is expired.
 * (Extracted from canvasCache.js getCachedData/hasCachedData)
 */
function isCacheExpired(entry) {
    if (!entry || !entry.expiresAt) return true;
    return new Date() > new Date(entry.expiresAt);
}

/**
 * Simulates the staged cache write pattern.
 * (Extracted from canvasCache.js stageCacheData + flushPendingCacheWrites)
 */
function createCacheStager() {
    const pendingWrites = {};

    return {
        stage(syStudentId, userData, courses) {
            if (!syStudentId) return;
            const filteredUserData = filterUserData(userData);
            const filteredCourses = filterCourses(courses);
            const expiration = calculateExpiration(filteredCourses);

            pendingWrites[syStudentId] = {
                userData: filteredUserData,
                courses: filteredCourses,
                expiresAt: expiration.toISOString(),
                cachedAt: new Date().toISOString()
            };
        },

        flush(existingCache = {}) {
            const keys = Object.keys(pendingWrites);
            if (keys.length === 0) return { cache: existingCache, count: 0 };

            const cache = { ...existingCache };
            for (const key of keys) {
                cache[key] = pendingWrites[key];
                delete pendingWrites[key];
            }
            return { cache, count: keys.length };
        },

        getPendingCount() {
            return Object.keys(pendingWrites).length;
        }
    };
}

// ============================================
// Tests
// ============================================

describe('filterUserData', () => {
    test('retains only essential fields', () => {
        const fullUserData = {
            id: 12345,
            name: 'Jane Doe',
            sortable_name: 'Doe, Jane',
            avatar_url: 'https://canvas.com/avatar.png',
            created_at: '2025-08-01T00:00:00Z',
            // Fields that should be stripped:
            email: 'jane@example.com',
            login_id: 'jdoe',
            sis_user_id: '99999',
            locale: 'en',
            permissions: { can_update: true }
        };

        const filtered = filterUserData(fullUserData);

        expect(filtered).toEqual({
            id: 12345,
            name: 'Jane Doe',
            sortable_name: 'Doe, Jane',
            avatar_url: 'https://canvas.com/avatar.png',
            created_at: '2025-08-01T00:00:00Z'
        });

        // Verify stripped fields are gone
        expect(filtered.email).toBeUndefined();
        expect(filtered.login_id).toBeUndefined();
        expect(filtered.permissions).toBeUndefined();
    });

    test('handles null/undefined input', () => {
        expect(filterUserData(null)).toEqual({});
        expect(filterUserData(undefined)).toEqual({});
    });
});

describe('filterCourses', () => {
    test('retains only essential course fields', () => {
        const fullCourses = [{
            id: 100,
            name: 'Math 101',
            start_at: '2026-01-10T00:00:00Z',
            end_at: '2026-03-15T00:00:00Z',
            enrollments: [{
                type: 'StudentEnrollment',
                enrollment_state: 'active',
                grades: { current_score: 85, final_score: 80, html_url: '/grades' }
            }],
            // Fields that should be stripped:
            course_code: 'MTH101',
            workflow_state: 'available',
            teachers: [{ id: 1, name: 'Prof. Smith' }]
        }];

        const filtered = filterCourses(fullCourses);

        expect(filtered).toEqual([{
            id: 100,
            name: 'Math 101',
            start_at: '2026-01-10T00:00:00Z',
            end_at: '2026-03-15T00:00:00Z',
            enrollments: [{
                type: 'StudentEnrollment',
                grades: { current_score: 85 }
            }]
        }]);

        // Verify stripped fields
        expect(filtered[0].course_code).toBeUndefined();
        expect(filtered[0].teachers).toBeUndefined();
        expect(filtered[0].enrollments[0].enrollment_state).toBeUndefined();
    });

    test('handles null/undefined/empty input', () => {
        expect(filterCourses(null)).toEqual([]);
        expect(filterCourses(undefined)).toEqual([]);
        expect(filterCourses([])).toEqual([]);
    });

    test('handles courses without enrollments', () => {
        const courses = [{ id: 1, name: 'Test', start_at: null, end_at: null }];
        const filtered = filterCourses(courses);
        expect(filtered[0].enrollments).toEqual([]);
    });
});

describe('calculateExpiration', () => {
    test('uses latest course end date', () => {
        const courses = [
            { end_at: '2026-02-01T00:00:00Z' },
            { end_at: '2026-04-15T00:00:00Z' },
            { end_at: '2026-03-01T00:00:00Z' }
        ];

        const exp = calculateExpiration(courses);
        expect(exp.toISOString()).toBe('2026-04-15T00:00:00.000Z');
    });

    test('falls back to 30 days from now when no end dates', () => {
        const courses = [{ end_at: null }, { end_at: null }];
        const exp = calculateExpiration(courses);

        const thirtyDaysFromNow = new Date();
        thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);

        // Should be within 1 second of 30 days from now
        expect(Math.abs(exp.getTime() - thirtyDaysFromNow.getTime())).toBeLessThan(1000);
    });

    test('falls back to 30 days for empty courses', () => {
        const exp = calculateExpiration([]);
        const thirtyDaysFromNow = new Date();
        thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);

        expect(Math.abs(exp.getTime() - thirtyDaysFromNow.getTime())).toBeLessThan(1000);
    });

    test('ignores courses without end_at', () => {
        const courses = [
            { end_at: null },
            { end_at: '2026-06-01T00:00:00Z' },
            { end_at: null }
        ];

        const exp = calculateExpiration(courses);
        expect(exp.toISOString()).toBe('2026-06-01T00:00:00.000Z');
    });
});

describe('isCacheExpired', () => {
    test('returns true for expired entry', () => {
        const entry = { expiresAt: '2020-01-01T00:00:00Z' };
        expect(isCacheExpired(entry)).toBe(true);
    });

    test('returns false for valid entry', () => {
        const futureDate = new Date();
        futureDate.setFullYear(futureDate.getFullYear() + 1);
        const entry = { expiresAt: futureDate.toISOString() };
        expect(isCacheExpired(entry)).toBe(false);
    });

    test('returns true for null/undefined entry', () => {
        expect(isCacheExpired(null)).toBe(true);
        expect(isCacheExpired(undefined)).toBe(true);
    });

    test('returns true for entry without expiresAt', () => {
        expect(isCacheExpired({})).toBe(true);
        expect(isCacheExpired({ userData: {} })).toBe(true);
    });
});

describe('Staged cache writes (batching)', () => {
    test('stages multiple entries without writing', () => {
        const stager = createCacheStager();

        stager.stage('S001', { id: 1, name: 'Alice' }, []);
        stager.stage('S002', { id: 2, name: 'Bob' }, []);

        expect(stager.getPendingCount()).toBe(2);
    });

    test('flush writes all pending entries to cache', () => {
        const stager = createCacheStager();

        stager.stage('S001', { id: 1, name: 'Alice' }, []);
        stager.stage('S002', { id: 2, name: 'Bob' }, []);

        const { cache, count } = stager.flush({});

        expect(count).toBe(2);
        expect(cache['S001']).toBeDefined();
        expect(cache['S002']).toBeDefined();
        expect(cache['S001'].userData.name).toBe('Alice');
        expect(cache['S002'].userData.name).toBe('Bob');
    });

    test('flush clears pending writes', () => {
        const stager = createCacheStager();
        stager.stage('S001', { id: 1, name: 'Alice' }, []);

        stager.flush({});
        expect(stager.getPendingCount()).toBe(0);

        // Second flush should be a no-op
        const { count } = stager.flush({});
        expect(count).toBe(0);
    });

    test('flush merges with existing cache', () => {
        const stager = createCacheStager();
        stager.stage('S002', { id: 2, name: 'Bob' }, []);

        const existingCache = {
            'S001': { userData: { name: 'Alice' }, courses: [], expiresAt: '2027-01-01T00:00:00Z' }
        };

        const { cache } = stager.flush(existingCache);

        // Both old and new entries should be present
        expect(cache['S001'].userData.name).toBe('Alice');
        expect(cache['S002'].userData.name).toBe('Bob');
    });

    test('later stage overwrites earlier stage for same student', () => {
        const stager = createCacheStager();

        stager.stage('S001', { id: 1, name: 'Alice v1' }, []);
        stager.stage('S001', { id: 1, name: 'Alice v2' }, []);

        expect(stager.getPendingCount()).toBe(1);

        const { cache } = stager.flush({});
        expect(cache['S001'].userData.name).toBe('Alice v2');
    });

    test('ignores null/empty syStudentId', () => {
        const stager = createCacheStager();

        stager.stage(null, { id: 1 }, []);
        stager.stage('', { id: 2 }, []);
        stager.stage(undefined, { id: 3 }, []);

        expect(stager.getPendingCount()).toBe(0);
    });

    test('staged entries include expiration and cached timestamp', () => {
        const stager = createCacheStager();
        stager.stage('S001', { id: 1 }, [{ end_at: '2026-06-01T00:00:00Z' }]);

        const { cache } = stager.flush({});
        const entry = cache['S001'];

        expect(entry.expiresAt).toBe('2026-06-01T00:00:00.000Z');
        expect(entry.cachedAt).toBeDefined();
        expect(new Date(entry.cachedAt).getTime()).not.toBeNaN();
    });
});
