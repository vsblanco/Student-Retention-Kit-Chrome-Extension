/**
 * Tests for field-utils.js - trimCommonPrefix
 *
 * Verifies the common prefix detection and trimming logic used
 * for cleaner campus name display across the extension.
 */

// Copy of trimCommonPrefix from src/constants/field-utils.js
function trimCommonPrefix(names) {
    const identity = new Map(names?.map(n => [n, n]) || []);
    if (!names || names.length <= 1) {
        return { trimmedNames: identity, prefix: '' };
    }

    let prefix = names[0];
    for (let i = 1; i < names.length; i++) {
        while (prefix && !names[i].startsWith(prefix)) {
            prefix = prefix.slice(0, -1);
        }
    }

    prefix = prefix.replace(/[\s\-–—:]+$/, '');

    if (prefix.length < 3) {
        return { trimmedNames: identity, prefix: '' };
    }

    const cutsMidWord = names.some(name => {
        if (name.length <= prefix.length) return false;
        return /[a-zA-Z0-9]/.test(name[prefix.length]);
    });
    if (cutsMidWord) {
        return { trimmedNames: identity, prefix: '' };
    }

    const trimmedNames = new Map();
    for (const name of names) {
        let trimmed = name.substring(prefix.length);
        trimmed = trimmed.replace(/^[\s\-–—:]+/, '').trim();
        trimmedNames.set(name, trimmed || name);
    }

    return { trimmedNames, prefix };
}

// ============================================
// Tests
// ============================================

describe('trimCommonPrefix', () => {
    test('trims "Northbridge" prefix with separator variations', () => {
        const campuses = [
            'Northbridge - South Miami',
            'Northbridge Hialeah - South Florida',
            'Northbridge - Kissimmee',
            'Northbridge - Pembroke'
        ];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('Northbridge');
        expect(trimmedNames.get('Northbridge - South Miami')).toBe('South Miami');
        expect(trimmedNames.get('Northbridge Hialeah - South Florida')).toBe('Hialeah - South Florida');
        expect(trimmedNames.get('Northbridge - Kissimmee')).toBe('Kissimmee');
        expect(trimmedNames.get('Northbridge - Pembroke')).toBe('Pembroke');
    });

    test('trims consistent "Prefix - " pattern', () => {
        const campuses = [
            'Northbridge - South Miami',
            'Northbridge - Kissimmee',
            'Northbridge - Pembroke'
        ];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('Northbridge');
        expect(trimmedNames.get('Northbridge - South Miami')).toBe('South Miami');
        expect(trimmedNames.get('Northbridge - Kissimmee')).toBe('Kissimmee');
        expect(trimmedNames.get('Northbridge - Pembroke')).toBe('Pembroke');
    });

    test('returns identity map for a single campus', () => {
        const campuses = ['Northbridge - South Miami'];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('');
        expect(trimmedNames.get('Northbridge - South Miami')).toBe('Northbridge - South Miami');
    });

    test('returns identity map for empty array', () => {
        const { trimmedNames, prefix } = trimCommonPrefix([]);

        expect(prefix).toBe('');
        expect(trimmedNames.size).toBe(0);
    });

    test('returns identity map for null/undefined', () => {
        expect(trimCommonPrefix(null).prefix).toBe('');
        expect(trimCommonPrefix(undefined).prefix).toBe('');
    });

    test('does not trim when no common prefix exists', () => {
        const campuses = ['Alpha Campus', 'Beta Campus', 'Gamma Campus'];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('');
        expect(trimmedNames.get('Alpha Campus')).toBe('Alpha Campus');
        expect(trimmedNames.get('Beta Campus')).toBe('Beta Campus');
        expect(trimmedNames.get('Gamma Campus')).toBe('Gamma Campus');
    });

    test('does not trim prefix shorter than 3 characters', () => {
        const campuses = ['AB Downtown', 'AB Uptown'];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('');
        expect(trimmedNames.get('AB Downtown')).toBe('AB Downtown');
    });

    test('handles prefix with colon separator', () => {
        const campuses = [
            'University: Main Campus',
            'University: North Campus',
            'University: South Campus'
        ];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('University');
        expect(trimmedNames.get('University: Main Campus')).toBe('Main Campus');
        expect(trimmedNames.get('University: North Campus')).toBe('North Campus');
    });

    test('handles prefix with em dash separator', () => {
        const campuses = [
            'Academy – Downtown',
            'Academy – Midtown',
            'Academy – Uptown'
        ];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('Academy');
        expect(trimmedNames.get('Academy – Downtown')).toBe('Downtown');
        expect(trimmedNames.get('Academy – Midtown')).toBe('Midtown');
    });

    test('does not trim partial words', () => {
        const campuses = ['Northeast Campus', 'Northwest Campus'];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('');
        expect(trimmedNames.get('Northeast Campus')).toBe('Northeast Campus');
        expect(trimmedNames.get('Northwest Campus')).toBe('Northwest Campus');
    });

    test('falls back to original name if trimming would produce empty string', () => {
        const campuses = ['Northbridge', 'Northbridge - Miami'];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('Northbridge');
        // 'Northbridge' after prefix removal is '' → falls back to original
        expect(trimmedNames.get('Northbridge')).toBe('Northbridge');
        expect(trimmedNames.get('Northbridge - Miami')).toBe('Miami');
    });

    test('preserves all original keys in the map', () => {
        const campuses = [
            'Northbridge - A',
            'Northbridge - B',
            'Northbridge - C'
        ];

        const { trimmedNames } = trimCommonPrefix(campuses);

        expect(trimmedNames.size).toBe(3);
        for (const campus of campuses) {
            expect(trimmedNames.has(campus)).toBe(true);
        }
    });

    test('handles multi-word prefix', () => {
        const campuses = [
            'New York University - Arts',
            'New York University - Science',
            'New York University - Law'
        ];

        const { trimmedNames, prefix } = trimCommonPrefix(campuses);

        expect(prefix).toBe('New York University');
        expect(trimmedNames.get('New York University - Arts')).toBe('Arts');
        expect(trimmedNames.get('New York University - Science')).toBe('Science');
        expect(trimmedNames.get('New York University - Law')).toBe('Law');
    });
});
