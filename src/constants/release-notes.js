/**
 * Release Notes Configuration
 *
 * This file contains the release notes for the extension.
 * To add new release notes:
 * 1. Add a new entry with the version number as the key
 * 2. Fill in the title, date, and updates array
 * 3. The modal will automatically show when users update to this version
 *
 * Example:
 * '10.4': {
 *     title: 'New Feature Release',
 *     date: 'February 1, 2026',
 *     updates: [
 *         'Added new feature X',
 *         'Fixed bug Y',
 *         'Improved performance of Z'
 *     ]
 * }
 */

export const RELEASE_NOTES = {
    // Add new version entries here (newest first)
    // The key should match the version in manifest.json

    '11.0': {
        title: 'Latest Updates',
        date: 'January 30, 2026',
        updates: [
            'Added numerous bug fixes and stababilty',
            'Added support for Power Automate integration',
            'Included a non api feature for users without API permissions.',
            'Adjusted UI for a more professional and polished look',
            'Improved data management handling with 6k students',
            'Included a Campus filter in case you import multiple Campus populations',
            'Added safety guards for the extension to disable if Chrome crashes'
        ]
    }

    // Example of how to add future versions:
    // '10.4': {
    //     title: 'Performance Improvements',
    //     date: 'February 15, 2026',
    //     updates: [
    //         'Faster master list loading',
    //         'Improved Canvas API caching',
    //         'Bug fixes and stability improvements'
    //     ]
    // }
};

/**
 * Gets the release notes for a specific version
 * @param {string} version - The version number to get notes for
 * @returns {Object|null} The release notes object or null if not found
 */
export function getReleaseNotes(version) {
    return RELEASE_NOTES[version] || null;
}

/**
 * Checks if release notes exist for a version
 * @param {string} version - The version number to check
 * @returns {boolean} True if release notes exist for this version
 */
export function hasReleaseNotes(version) {
    return version in RELEASE_NOTES;
}

/**
 * Gets all versions that have release notes (sorted newest first)
 * @returns {string[]} Array of version strings
 */
export function getAllVersionsWithNotes() {
    return Object.keys(RELEASE_NOTES).sort((a, b) => {
        // Sort by version number (descending)
        const partsA = a.split('.').map(Number);
        const partsB = b.split('.').map(Number);

        for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
            const numA = partsA[i] || 0;
            const numB = partsB[i] || 0;
            if (numA !== numB) return numB - numA;
        }
        return 0;
    });
}

