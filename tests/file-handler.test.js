/**
 * Tests for file-handler.js - File upload parsing and field matching
 *
 * These tests verify the core logic for:
 * - Field name normalization
 * - Alias-based column matching (pre-computed lookup)
 * - Column index finding with normalized headers
 * - Phone number cleaning
 * - Field value retrieval with alias resolution
 */

// ============================================
// Constants extracted from constants/index.js
// ============================================

function normalizeFieldName(fieldName) {
    if (!fieldName) return '';
    return String(fieldName)
        .toLowerCase()
        .replace(/[\s\-_]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

const FIELD_ALIASES = {
    name: ['Student Name', 'StudentName', 'Full Name', 'Name'],
    StudentNumber: ['Student ID', 'StudentId', 'SIS ID', 'SIS User ID', 'Student Number', 'SyStudentId'],
    phone: ['Phone', 'Phone Number', 'PrimaryPhone', 'Mobile', 'Cell'],
    otherPhone: ['Other Phone', 'OtherPhone', 'Secondary Phone', 'Alt Phone'],
    lda: ['LDA', 'Last Day Attended', 'LastDayAttended', 'Last Day of Attendance'],
    grade: ['Grade', 'Current Grade', 'CurrentGrade', 'Score'],
    campus: ['Campus', 'Location', 'Site'],
    programVersion: ['Program Version', 'ProgramVersion', 'Program'],
    url: ['Gradebook', 'Grade Book', 'GradeBook', 'Canvas URL', 'URL']
};

// ============================================
// Pre-computed alias lookup (mirrors file-handler.js optimization)
// ============================================

const _normalizedAliasToField = {};
const _normalizedFieldNames = {};
for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    const normalizedField = normalizeFieldName(field);
    _normalizedFieldNames[field] = normalizedField;
    _normalizedAliasToField[normalizedField] = field;
    for (const alias of aliases) {
        _normalizedAliasToField[normalizeFieldName(alias)] = field;
    }
}

// ============================================
// Pure functions extracted for testing
// ============================================

/**
 * Finds the column index for a field using pre-computed alias lookup.
 * (Extracted from file-handler.js)
 */
function findColumnIndex(normalizedHeaders, rawHeaders, fieldName) {
    const normalizedFieldName = _normalizedFieldNames[fieldName] || normalizeFieldName(fieldName);

    for (let i = 0; i < normalizedHeaders.length; i++) {
        const nh = normalizedHeaders[i];
        if (nh === normalizedFieldName) return i;
        if (_normalizedAliasToField[nh] === fieldName) return i;
    }
    return -1;
}

/**
 * Finds a field value in an object using normalized matching and aliases.
 * (Extracted from file-handler.js)
 */
function getFieldWithAlias(obj, fieldName, defaultValue = null) {
    if (!obj || !fieldName) return defaultValue;

    if (fieldName in obj) {
        const value = obj[fieldName];
        return value !== null && value !== undefined ? value : defaultValue;
    }

    const normalizedFieldName = _normalizedFieldNames[fieldName] || normalizeFieldName(fieldName);

    for (const key in obj) {
        const normalizedKey = normalizeFieldName(key);
        if (normalizedKey === normalizedFieldName || _normalizedAliasToField[normalizedKey] === fieldName) {
            const value = obj[key];
            return value !== null && value !== undefined ? value : defaultValue;
        }
    }

    return defaultValue;
}

/**
 * Cleans phone numbers by removing trailing non-numeric text.
 * (Extracted from file-handler.js)
 */
function cleanPhoneNumber(value) {
    if (!value) return '';
    let cleaned = String(value).trim();
    cleaned = cleaned.replace(/(\d)\s+[a-zA-Z].*$/, '$1');
    return cleaned;
}

// ============================================
// Tests
// ============================================

describe('normalizeFieldName', () => {
    test('normalizes spaces, hyphens, and underscores', () => {
        expect(normalizeFieldName('Student Name')).toBe('studentname');
        expect(normalizeFieldName('student-name')).toBe('studentname');
        expect(normalizeFieldName('student_name')).toBe('studentname');
    });

    test('converts to lowercase', () => {
        expect(normalizeFieldName('StudentNumber')).toBe('studentnumber');
        expect(normalizeFieldName('GRADE')).toBe('grade');
    });

    test('removes special characters', () => {
        expect(normalizeFieldName('Phone #')).toBe('phone');
        expect(normalizeFieldName('Grade (%)')).toBe('grade');
    });

    test('handles empty/null/undefined', () => {
        expect(normalizeFieldName('')).toBe('');
        expect(normalizeFieldName(null)).toBe('');
        expect(normalizeFieldName(undefined)).toBe('');
    });

    test('handles numeric input', () => {
        expect(normalizeFieldName(12345)).toBe('12345');
    });
});

describe('Pre-computed alias lookup', () => {
    test('maps canonical field names to themselves', () => {
        expect(_normalizedAliasToField['name']).toBe('name');
        expect(_normalizedAliasToField['studentnumber']).toBe('StudentNumber');
        expect(_normalizedAliasToField['phone']).toBe('phone');
    });

    test('maps aliases to canonical field names', () => {
        expect(_normalizedAliasToField['studentname']).toBe('name');
        expect(_normalizedAliasToField['fullname']).toBe('name');
        expect(_normalizedAliasToField['studentid']).toBe('StudentNumber');
        expect(_normalizedAliasToField['sisid']).toBe('StudentNumber');
        expect(_normalizedAliasToField['sisuserid']).toBe('StudentNumber');
        expect(_normalizedAliasToField['phonenumber']).toBe('phone');
        expect(_normalizedAliasToField['primaryphone']).toBe('phone');
    });

    test('LDA aliases resolve correctly', () => {
        expect(_normalizedAliasToField['lda']).toBe('lda');
        expect(_normalizedAliasToField['lastdayattended']).toBe('lda');
        expect(_normalizedAliasToField['lastdayofattendance']).toBe('lda');
    });

    test('URL/gradebook aliases resolve correctly', () => {
        expect(_normalizedAliasToField['gradebook']).toBe('url');
        expect(_normalizedAliasToField['gradebook']).toBe('url');
        expect(_normalizedAliasToField['canvasurl']).toBe('url');
    });
});

describe('findColumnIndex', () => {
    const rawHeaders = ['Student Name', 'SIS ID', 'Grade', 'LDA', 'Campus', 'Unknown Column'];
    const normalizedHeaders = rawHeaders.map(h => normalizeFieldName(h));

    test('finds direct field name match', () => {
        expect(findColumnIndex(normalizedHeaders, rawHeaders, 'grade')).toBe(2);
        expect(findColumnIndex(normalizedHeaders, rawHeaders, 'campus')).toBe(4);
    });

    test('finds alias match', () => {
        // 'Student Name' is an alias for 'name'
        expect(findColumnIndex(normalizedHeaders, rawHeaders, 'name')).toBe(0);
        // 'SIS ID' is an alias for 'StudentNumber'
        expect(findColumnIndex(normalizedHeaders, rawHeaders, 'StudentNumber')).toBe(1);
        // 'LDA' is an alias for 'lda'
        expect(findColumnIndex(normalizedHeaders, rawHeaders, 'lda')).toBe(3);
    });

    test('returns -1 for unknown field', () => {
        expect(findColumnIndex(normalizedHeaders, rawHeaders, 'nonExistentField')).toBe(-1);
    });

    test('handles empty headers', () => {
        expect(findColumnIndex([], [], 'name')).toBe(-1);
    });

    test('is case-insensitive through normalization', () => {
        const headers = ['STUDENT NAME', 'sis id', 'Grade'];
        const normalized = headers.map(h => normalizeFieldName(h));
        expect(findColumnIndex(normalized, headers, 'name')).toBe(0);
        expect(findColumnIndex(normalized, headers, 'StudentNumber')).toBe(1);
    });
});

describe('getFieldWithAlias', () => {
    test('returns direct property value', () => {
        const student = { name: 'John Doe', grade: '85' };
        expect(getFieldWithAlias(student, 'name')).toBe('John Doe');
        expect(getFieldWithAlias(student, 'grade')).toBe('85');
    });

    test('returns value via alias', () => {
        const student = { 'Student Name': 'Jane Doe', 'SIS ID': '12345' };
        expect(getFieldWithAlias(student, 'name')).toBe('Jane Doe');
        expect(getFieldWithAlias(student, 'StudentNumber')).toBe('12345');
    });

    test('returns default when field not found', () => {
        const student = { name: 'John' };
        expect(getFieldWithAlias(student, 'phone')).toBeNull();
        expect(getFieldWithAlias(student, 'phone', 'N/A')).toBe('N/A');
    });

    test('returns default for null/undefined values', () => {
        const student = { name: null, grade: undefined };
        expect(getFieldWithAlias(student, 'name', 'default')).toBe('default');
        expect(getFieldWithAlias(student, 'grade', 'default')).toBe('default');
    });

    test('handles null/undefined object', () => {
        expect(getFieldWithAlias(null, 'name')).toBeNull();
        expect(getFieldWithAlias(undefined, 'name')).toBeNull();
    });

    test('handles null/undefined fieldName', () => {
        expect(getFieldWithAlias({ name: 'John' }, null)).toBeNull();
        expect(getFieldWithAlias({ name: 'John' }, undefined)).toBeNull();
    });

    test('prefers direct property over alias', () => {
        // If object has both 'name' and 'Student Name', direct match wins
        const student = { name: 'Direct', 'Student Name': 'Alias' };
        expect(getFieldWithAlias(student, 'name')).toBe('Direct');
    });
});

describe('cleanPhoneNumber', () => {
    test('removes trailing text after digits', () => {
        expect(cleanPhoneNumber('555-123-4567 ext 100')).toBe('555-123-4567');
        expect(cleanPhoneNumber('5551234567 Home')).toBe('5551234567');
    });

    test('preserves clean phone numbers', () => {
        expect(cleanPhoneNumber('555-123-4567')).toBe('555-123-4567');
        expect(cleanPhoneNumber('(555) 123-4567')).toBe('(555) 123-4567');
    });

    test('handles empty/null input', () => {
        expect(cleanPhoneNumber('')).toBe('');
        expect(cleanPhoneNumber(null)).toBe('');
        expect(cleanPhoneNumber(undefined)).toBe('');
    });

    test('trims whitespace', () => {
        expect(cleanPhoneNumber('  555-123-4567  ')).toBe('555-123-4567');
    });

    test('handles numeric input', () => {
        expect(cleanPhoneNumber(5551234567)).toBe('5551234567');
    });
});
