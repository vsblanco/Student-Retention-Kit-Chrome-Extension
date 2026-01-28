// Storage Utility - Handles nested storage paths for organized settings
// Version: 1.0

import { STORAGE_KEYS, DEFAULT_SETTINGS } from '../constants/index.js';

/**
 * Maps old flat storage keys to new nested paths for migration.
 */
const LEGACY_KEY_MAP = {
    // General settings
    'autoUpdateMasterList': STORAGE_KEYS.AUTO_UPDATE_MASTER_LIST,
    'tutorialCompleted': STORAGE_KEYS.TUTORIAL_COMPLETED,

    // Excel Add-in settings
    'sendMasterListToExcel': STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL,
    'reformatNameEnabled': STORAGE_KEYS.REFORMAT_NAME_ENABLED,
    'syncActiveStudent': STORAGE_KEYS.SYNC_ACTIVE_STUDENT,
    'autoSideloadManifest': STORAGE_KEYS.AUTO_SIDELOAD_MANIFEST,

    // Excel Add-in Highlight settings
    'highlightStudentRowEnabled': STORAGE_KEYS.HIGHLIGHT_STUDENT_ROW_ENABLED,
    'highlightRowColor': STORAGE_KEYS.HIGHLIGHT_ROW_COLOR,
    'highlightStartCol': STORAGE_KEYS.HIGHLIGHT_START_COL,
    'highlightEndCol': STORAGE_KEYS.HIGHLIGHT_END_COL,
    'highlightEditColumn': STORAGE_KEYS.HIGHLIGHT_EDIT_COLUMN,
    'highlightEditText': STORAGE_KEYS.HIGHLIGHT_EDIT_TEXT,
    'highlightTargetSheet': STORAGE_KEYS.HIGHLIGHT_TARGET_SHEET,

    // Power Automate settings
    'powerAutomateUrl': STORAGE_KEYS.POWER_AUTOMATE_URL,

    // Canvas settings
    'embedInCanvas': STORAGE_KEYS.EMBED_IN_CANVAS,
    'highlightColor': STORAGE_KEYS.HIGHLIGHT_COLOR,

    // Five9 settings
    'debugMode': STORAGE_KEYS.CALL_DEMO,

    // Submission Checker settings
    'looperDaysOutFilter': STORAGE_KEYS.LOOPER_DAYS_OUT_FILTER,
    'scanFilterIncludeFailing': STORAGE_KEYS.SCAN_FILTER_INCLUDE_FAILING,

    // State
    'extensionState': STORAGE_KEYS.EXTENSION_STATE,
    'srkSessionId': STORAGE_KEYS.SRK_SESSION_ID,

    // Timestamps
    'lastCallTimestamp': STORAGE_KEYS.LAST_CALL_TIMESTAMP,
    'lastUpdated': STORAGE_KEYS.LAST_UPDATED,
    'masterListSourceTimestamp': STORAGE_KEYS.MASTER_LIST_SOURCE_TIMESTAMP,
    'referenceDate': STORAGE_KEYS.REFERENCE_DATE
};

/**
 * Checks if a key is a nested path (contains dots).
 * @param {string} key - The storage key
 * @returns {boolean} True if the key is a nested path
 */
function isNestedPath(key) {
    return key && key.includes('.');
}

/**
 * Gets the root key from a nested path.
 * @param {string} path - The nested path (e.g., 'settings.excelAddIn.highlight.rowColor')
 * @returns {string} The root key (e.g., 'settings')
 */
function getRootKey(path) {
    return path.split('.')[0];
}

/**
 * Gets a nested value from an object using a dot-separated path.
 * @param {object} obj - The object to get the value from
 * @param {string} path - The dot-separated path
 * @returns {*} The value at the path, or undefined if not found
 */
function getNestedValue(obj, path) {
    if (!obj || !path) return undefined;
    const parts = path.split('.');
    let current = obj;
    for (const part of parts) {
        if (current === undefined || current === null) return undefined;
        current = current[part];
    }
    return current;
}

/**
 * Sets a nested value in an object using a dot-separated path.
 * Creates intermediate objects as needed.
 * @param {object} obj - The object to set the value in
 * @param {string} path - The dot-separated path
 * @param {*} value - The value to set
 */
function setNestedValue(obj, path, value) {
    const parts = path.split('.');
    let current = obj;
    for (let i = 0; i < parts.length - 1; i++) {
        const part = parts[i];
        if (current[part] === undefined || current[part] === null) {
            current[part] = {};
        }
        current = current[part];
    }
    current[parts[parts.length - 1]] = value;
}

/**
 * Reads values from storage, handling nested paths.
 * @param {string|string[]} keys - Storage key(s) to read (can be nested paths)
 * @returns {Promise<object>} Object with values keyed by the original paths
 */
export async function storageGet(keys) {
    const keyArray = Array.isArray(keys) ? keys : [keys];

    // Group keys by their root
    const rootKeys = new Set();
    const flatKeys = [];

    for (const key of keyArray) {
        if (isNestedPath(key)) {
            rootKeys.add(getRootKey(key));
        } else {
            flatKeys.push(key);
        }
    }

    // Fetch all root keys and flat keys
    const allKeys = [...rootKeys, ...flatKeys];
    const rawData = await chrome.storage.local.get(allKeys);

    // Build result object with values at original paths
    const result = {};

    for (const key of keyArray) {
        if (isNestedPath(key)) {
            const rootKey = getRootKey(key);
            const pathWithinRoot = key.substring(rootKey.length + 1);
            const value = getNestedValue(rawData[rootKey], pathWithinRoot);
            result[key] = value;
        } else {
            result[key] = rawData[key];
        }
    }

    return result;
}

/**
 * Writes values to storage, handling nested paths.
 * @param {object} data - Object with keys as storage paths and values to set
 * @returns {Promise<void>}
 */
export async function storageSet(data) {
    // Group by root key to batch nested updates
    const rootUpdates = {};
    const flatUpdates = {};

    for (const [key, value] of Object.entries(data)) {
        if (isNestedPath(key)) {
            const rootKey = getRootKey(key);
            if (!rootUpdates[rootKey]) {
                rootUpdates[rootKey] = [];
            }
            rootUpdates[rootKey].push({ path: key, value });
        } else {
            flatUpdates[key] = value;
        }
    }

    // For nested updates, fetch current root objects, update them, and write back
    if (Object.keys(rootUpdates).length > 0) {
        const rootKeys = Object.keys(rootUpdates);
        const currentRoots = await chrome.storage.local.get(rootKeys);

        for (const rootKey of rootKeys) {
            const rootObj = currentRoots[rootKey] || {};

            for (const { path, value } of rootUpdates[rootKey]) {
                const pathWithinRoot = path.substring(rootKey.length + 1);
                setNestedValue(rootObj, pathWithinRoot, value);
            }

            flatUpdates[rootKey] = rootObj;
        }
    }

    // Write all updates
    if (Object.keys(flatUpdates).length > 0) {
        await chrome.storage.local.set(flatUpdates);
    }
}

/**
 * Removes values from storage, handling nested paths.
 * For nested paths, sets the value to undefined within the parent object.
 * @param {string|string[]} keys - Storage key(s) to remove
 * @returns {Promise<void>}
 */
export async function storageRemove(keys) {
    const keyArray = Array.isArray(keys) ? keys : [keys];

    const flatKeys = [];
    const nestedPaths = {};

    for (const key of keyArray) {
        if (isNestedPath(key)) {
            const rootKey = getRootKey(key);
            if (!nestedPaths[rootKey]) {
                nestedPaths[rootKey] = [];
            }
            nestedPaths[rootKey].push(key);
        } else {
            flatKeys.push(key);
        }
    }

    // Remove flat keys
    if (flatKeys.length > 0) {
        await chrome.storage.local.remove(flatKeys);
    }

    // For nested paths, we need to update the parent object
    if (Object.keys(nestedPaths).length > 0) {
        const rootKeys = Object.keys(nestedPaths);
        const currentRoots = await chrome.storage.local.get(rootKeys);

        const updates = {};
        for (const rootKey of rootKeys) {
            const rootObj = currentRoots[rootKey];
            if (rootObj) {
                for (const path of nestedPaths[rootKey]) {
                    const pathWithinRoot = path.substring(rootKey.length + 1);
                    setNestedValue(rootObj, pathWithinRoot, undefined);
                }
                updates[rootKey] = rootObj;
            }
        }

        if (Object.keys(updates).length > 0) {
            await chrome.storage.local.set(updates);
        }
    }
}

/**
 * Gets a single value from storage with a default fallback.
 * @param {string} key - Storage key (can be nested path)
 * @param {*} defaultValue - Default value if key doesn't exist
 * @returns {Promise<*>} The stored value or default
 */
export async function storageGetValue(key, defaultValue = undefined) {
    const result = await storageGet([key]);
    const value = result[key];
    if (value === undefined) {
        // Check DEFAULT_SETTINGS
        const defaultFromSettings = DEFAULT_SETTINGS[key];
        return defaultFromSettings !== undefined ? defaultFromSettings : defaultValue;
    }
    return value;
}

/**
 * Sets a single value in storage.
 * @param {string} key - Storage key (can be nested path)
 * @param {*} value - Value to set
 * @returns {Promise<void>}
 */
export async function storageSetValue(key, value) {
    await storageSet({ [key]: value });
}

/**
 * Migrates old flat storage keys to new nested structure.
 * Should be called once on extension startup.
 * @returns {Promise<boolean>} True if migration was performed
 */
export async function migrateStorage() {
    // Check if migration has already been done
    const migrationCheck = await chrome.storage.local.get(['_storageMigrationV1']);
    if (migrationCheck._storageMigrationV1) {
        console.log('[Storage] Migration already completed');
        return false;
    }

    console.log('[Storage] Starting storage migration...');

    // Get all current storage
    const allData = await chrome.storage.local.get(null);

    // Build migration data
    const migrationData = {};
    let migratedCount = 0;

    for (const [oldKey, newPath] of Object.entries(LEGACY_KEY_MAP)) {
        if (allData[oldKey] !== undefined) {
            migrationData[newPath] = allData[oldKey];
            migratedCount++;
            console.log(`[Storage] Migrating: ${oldKey} -> ${newPath}`);
        }
    }

    if (migratedCount > 0) {
        // Write migrated data
        await storageSet(migrationData);

        // Optionally remove old keys (keep them for now for backwards compatibility)
        // await chrome.storage.local.remove(Object.keys(LEGACY_KEY_MAP));

        console.log(`[Storage] Migration complete: ${migratedCount} keys migrated`);
    } else {
        console.log('[Storage] No legacy keys found to migrate');
    }

    // Mark migration as complete
    await chrome.storage.local.set({ _storageMigrationV1: Date.now() });

    return migratedCount > 0;
}

/**
 * Gets the nested structure for debugging/inspection.
 * @returns {Promise<object>} The full storage structure
 */
export async function getStorageStructure() {
    const data = await chrome.storage.local.get(null);
    return data;
}

/**
 * Clears the migration flag to allow re-migration (for testing).
 * @returns {Promise<void>}
 */
export async function resetMigration() {
    await chrome.storage.local.remove(['_storageMigrationV1']);
    console.log('[Storage] Migration flag reset');
}

// Export legacy key map for reference
export { LEGACY_KEY_MAP };

// ============================================================================
// SESSION STORAGE FUNCTIONS
// For temporary state that should not persist across browser restarts
// Used primarily for EXTENSION_STATE to prevent "stuck on" state after crash
// ============================================================================

/**
 * Reads values from session storage, handling nested paths.
 * @param {string|string[]} keys - Storage key(s) to read (can be nested paths)
 * @returns {Promise<object>} Object with values keyed by the original paths
 */
export async function sessionGet(keys) {
    const keyArray = Array.isArray(keys) ? keys : [keys];

    // Group keys by their root
    const rootKeys = new Set();
    const flatKeys = [];

    for (const key of keyArray) {
        if (isNestedPath(key)) {
            rootKeys.add(getRootKey(key));
        } else {
            flatKeys.push(key);
        }
    }

    // Fetch all root keys and flat keys from session storage
    const allKeys = [...rootKeys, ...flatKeys];
    const rawData = await chrome.storage.session.get(allKeys);

    // Build result object with values at original paths
    const result = {};

    for (const key of keyArray) {
        if (isNestedPath(key)) {
            const rootKey = getRootKey(key);
            const pathWithinRoot = key.substring(rootKey.length + 1);
            const value = getNestedValue(rawData[rootKey], pathWithinRoot);
            result[key] = value;
        } else {
            result[key] = rawData[key];
        }
    }

    return result;
}

/**
 * Writes values to session storage, handling nested paths.
 * @param {object} data - Object with keys as storage paths and values to set
 * @returns {Promise<void>}
 */
export async function sessionSet(data) {
    // Group by root key to batch nested updates
    const rootUpdates = {};
    const flatUpdates = {};

    for (const [key, value] of Object.entries(data)) {
        if (isNestedPath(key)) {
            const rootKey = getRootKey(key);
            if (!rootUpdates[rootKey]) {
                rootUpdates[rootKey] = [];
            }
            rootUpdates[rootKey].push({ path: key, value });
        } else {
            flatUpdates[key] = value;
        }
    }

    // For nested updates, fetch current root objects, update them, and write back
    if (Object.keys(rootUpdates).length > 0) {
        const rootKeys = Object.keys(rootUpdates);
        const currentRoots = await chrome.storage.session.get(rootKeys);

        for (const rootKey of rootKeys) {
            const rootObj = currentRoots[rootKey] || {};

            for (const { path, value } of rootUpdates[rootKey]) {
                const pathWithinRoot = path.substring(rootKey.length + 1);
                setNestedValue(rootObj, pathWithinRoot, value);
            }

            flatUpdates[rootKey] = rootObj;
        }
    }

    // Write all updates to session storage
    if (Object.keys(flatUpdates).length > 0) {
        await chrome.storage.session.set(flatUpdates);
    }
}

/**
 * Gets a single value from session storage with a default fallback.
 * @param {string} key - Storage key (can be nested path)
 * @param {*} defaultValue - Default value if key doesn't exist
 * @returns {Promise<*>} The stored value or default
 */
export async function sessionGetValue(key, defaultValue = undefined) {
    const result = await sessionGet([key]);
    const value = result[key];
    if (value === undefined) {
        return defaultValue;
    }
    return value;
}
