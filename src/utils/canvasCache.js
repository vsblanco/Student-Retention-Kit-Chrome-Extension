/**
 * Canvas API Response Cache Module
 *
 * Caches Canvas API responses (user data and courses) to reduce API calls.
 * Uses course end_at dates as cache expiration timestamps.
 *
 * Version: 1.0
 * Date: 2025-12-18
 */

const CANVAS_CACHE_KEY = 'canvasApiCache';

/**
 * Cache entry structure:
 * {
 *   [SyStudentId]: {
 *     userData: {...},           // Canvas user profile data
 *     courses: [...],            // Array of course objects
 *     expiresAt: Date ISO string, // Latest course end_at date
 *     cachedAt: Date ISO string  // When this was cached
 *   }
 * }
 */

/**
 * Retrieves the entire cache from Chrome storage
 * @returns {Promise<Object>} The cache object
 */
export async function getCache() {
    try {
        const result = await chrome.storage.local.get(CANVAS_CACHE_KEY);
        return result[CANVAS_CACHE_KEY] || {};
    } catch (error) {
        console.error('Error retrieving cache:', error);
        return {};
    }
}

/**
 * Saves the cache to Chrome storage
 * @param {Object} cache - The cache object to save
 * @returns {Promise<void>}
 */
async function saveCache(cache) {
    try {
        await chrome.storage.local.set({ [CANVAS_CACHE_KEY]: cache });
    } catch (error) {
        console.error('Error saving cache:', error);
    }
}

/**
 * Gets cached data for a specific student
 * @param {string} syStudentId - The SyStudentId to look up
 * @returns {Promise<Object|null>} The cached data or null if not found/expired
 */
export async function getCachedData(syStudentId) {
    if (!syStudentId) return null;

    const cache = await getCache();
    const entry = cache[syStudentId];

    if (!entry) return null;

    // Check if cache has expired
    const now = new Date();
    const expiresAt = new Date(entry.expiresAt);

    if (now > expiresAt) {
        // Cache expired, remove it
        await removeCachedData(syStudentId);
        return null;
    }

    return {
        userData: entry.userData,
        courses: entry.courses,
        expiresAt: entry.expiresAt,
        cachedAt: entry.cachedAt
    };
}

/**
 * Filters user data to only include fields we actually use
 * This significantly reduces cache size
 * @param {Object} userData - Full Canvas user object
 * @returns {Object} Filtered user data with only necessary fields
 */
function filterUserData(userData) {
    if (!userData) return null;

    return {
        id: userData.id,
        name: userData.name,
        sortable_name: userData.sortable_name,
        avatar_url: userData.avatar_url,
        created_at: userData.created_at
    };
}

/**
 * Filters course data to only include fields we actually use
 * This significantly reduces cache size
 * @param {Array} courses - Full Canvas courses array
 * @returns {Array} Filtered courses with only necessary fields
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
 * Prepares a cache entry for a student without writing to storage.
 * Use with flushPendingCacheWrites() for batched writes.
 * @param {string} syStudentId - The SyStudentId
 * @param {Object} userData - The Canvas user profile data
 * @param {Array} courses - The array of course objects
 */
const _pendingCacheWrites = {};

export function stageCacheData(syStudentId, userData, courses) {
    if (!syStudentId) return;

    const filteredUserData = filterUserData(userData);
    const filteredCourses = filterCourses(courses);

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

    _pendingCacheWrites[syStudentId] = {
        userData: filteredUserData,
        courses: filteredCourses,
        expiresAt: latestEndDate.toISOString(),
        cachedAt: new Date().toISOString()
    };
}

/**
 * Flushes all pending cache writes to Chrome storage in a single operation.
 * @returns {Promise<number>} Number of entries written
 */
export async function flushPendingCacheWrites() {
    const keys = Object.keys(_pendingCacheWrites);
    if (keys.length === 0) return 0;

    const cache = await getCache();
    for (const key of keys) {
        cache[key] = _pendingCacheWrites[key];
        delete _pendingCacheWrites[key];
    }
    await saveCache(cache);
    return keys.length;
}

/**
 * Caches Canvas API data for a student
 * @param {string} syStudentId - The SyStudentId
 * @param {Object} userData - The Canvas user profile data
 * @param {Array} courses - The array of course objects
 * @returns {Promise<void>}
 */
export async function setCachedData(syStudentId, userData, courses) {
    if (!syStudentId) return;

    // Filter data to only cache what we need
    const filteredUserData = filterUserData(userData);
    const filteredCourses = filterCourses(courses);

    // Determine expiration date from courses
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

    // If no end dates found, set expiration to 30 days from now as fallback
    if (!latestEndDate) {
        latestEndDate = new Date();
        latestEndDate.setDate(latestEndDate.getDate() + 30);
    }

    const cache = await getCache();

    cache[syStudentId] = {
        userData: filteredUserData,
        courses: filteredCourses,
        expiresAt: latestEndDate.toISOString(),
        cachedAt: new Date().toISOString()
    };

    await saveCache(cache);
}

/**
 * Removes cached data for a specific student
 * @param {string} syStudentId - The SyStudentId
 * @returns {Promise<void>}
 */
export async function removeCachedData(syStudentId) {
    if (!syStudentId) return;

    const cache = await getCache();
    delete cache[syStudentId];
    await saveCache(cache);
}

/**
 * Clears all cached Canvas API data
 * @returns {Promise<void>}
 */
export async function clearAllCache() {
    await chrome.storage.local.remove(CANVAS_CACHE_KEY);
}

/**
 * Checks if a student has valid cached data (without retrieving the full cache)
 * @param {string} syStudentId - The SyStudentId to check
 * @returns {Promise<boolean>} True if valid cache exists, false otherwise
 */
export async function hasCachedData(syStudentId) {
    if (!syStudentId) return false;

    const cache = await getCache();
    const entry = cache[syStudentId];

    if (!entry) return false;

    const now = new Date();
    const expiresAt = new Date(entry.expiresAt);

    return now <= expiresAt;
}

/**
 * Gets cache statistics (total entries, expired entries, etc.)
 * @returns {Promise<Object>} Statistics about the cache
 */
export async function getCacheStats() {
    const cache = await getCache();
    const entries = Object.entries(cache);
    const now = new Date();

    let totalEntries = entries.length;
    let expiredEntries = 0;
    let validEntries = 0;

    for (const [syStudentId, entry] of entries) {
        const expiresAt = new Date(entry.expiresAt);
        if (now > expiresAt) {
            expiredEntries++;
        } else {
            validEntries++;
        }
    }

    return {
        totalEntries,
        validEntries,
        expiredEntries
    };
}

/**
 * Cleans up expired cache entries
 * @returns {Promise<number>} Number of entries removed
 */
export async function cleanupExpiredCache() {
    const cache = await getCache();
    const entries = Object.entries(cache);
    const now = new Date();
    let removedCount = 0;

    for (const [syStudentId, entry] of entries) {
        const expiresAt = new Date(entry.expiresAt);
        if (now > expiresAt) {
            delete cache[syStudentId];
            removedCount++;
        }
    }

    if (removedCount > 0) {
        await saveCache(cache);
    }

    return removedCount;
}
