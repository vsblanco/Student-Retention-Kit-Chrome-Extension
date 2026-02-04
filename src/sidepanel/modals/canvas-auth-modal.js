// Canvas Auth Error Modal - Handles Canvas authorization error dialogs
import { STORAGE_KEYS } from '../../constants/index.js';
import { storageGet, storageSet } from '../../utils/storage.js';
import { updateToggleUI, isToggleEnabled } from '../../utils/ui-helpers.js';
import { elements } from '../ui-manager.js';

// Canvas Auth Error Modal state
let canvasAuthErrorResolve = null;
let canvasAuthErrorShown = false;

/**
 * Updates the non-API toggle UI in the Canvas Auth Error modal
 * @param {boolean} isEnabled - Whether non-API course fetch is enabled
 */
function updateCanvasAuthNonApiToggleUI(isEnabled) {
    updateToggleUI(elements.canvasAuthNonApiToggle, isEnabled);
}

/**
 * Toggles the non-API course fetch setting in the Canvas Auth Error modal
 */
export function toggleCanvasAuthNonApi() {
    updateCanvasAuthNonApiToggleUI(!isToggleEnabled(elements.canvasAuthNonApiToggle));
}

/**
 * Opens the Canvas Auth Error modal and returns a promise that resolves with the user's choice
 * @returns {Promise<'continue'|'shutdown'>} Promise that resolves with 'continue' or 'shutdown'
 */
export async function openCanvasAuthErrorModal() {
    return new Promise(async (resolve) => {
        // Prevent multiple modals from stacking
        if (canvasAuthErrorShown) {
            resolve('continue'); // Default to continue if already shown
            return;
        }

        if (!elements.canvasAuthErrorModal) {
            resolve('continue');
            return;
        }

        canvasAuthErrorShown = true;
        canvasAuthErrorResolve = resolve;

        // Load current non-API setting and update toggle UI
        const settings = await storageGet([STORAGE_KEYS.NON_API_COURSE_FETCH]);
        const nonApiFetch = settings[STORAGE_KEYS.NON_API_COURSE_FETCH] || false;
        updateCanvasAuthNonApiToggleUI(nonApiFetch);

        // Show modal
        elements.canvasAuthErrorModal.style.display = 'flex';
    });
}

/**
 * Closes the Canvas Auth Error modal
 * @param {'continue'|'shutdown'} choice - The user's choice
 */
export async function closeCanvasAuthErrorModal(choice = 'continue') {
    // Save the non-API toggle setting if user chose to continue
    if (choice === 'continue') {
        const isNonApiEnabled = isToggleEnabled(elements.canvasAuthNonApiToggle);
        await storageSet({ [STORAGE_KEYS.NON_API_COURSE_FETCH]: isNonApiEnabled });
        if (isNonApiEnabled) {
            console.log('[Canvas Auth Error] Non-API course fetch enabled by user');
        }
    }

    if (elements.canvasAuthErrorModal) {
        elements.canvasAuthErrorModal.style.display = 'none';
    }

    canvasAuthErrorShown = false;

    // Resolve the promise with the user's choice
    if (canvasAuthErrorResolve) {
        canvasAuthErrorResolve(choice);
        canvasAuthErrorResolve = null;
    }
}

/**
 * Checks if an error response indicates a Canvas authorization error
 * @param {Response} response - The fetch response object
 * @returns {boolean} True if it's an authorization error
 */
export function isCanvasAuthError(response) {
    return response && (response.status === 401 || response.status === 403);
}

/**
 * Checks if a JSON error body indicates a Canvas authorization error
 * @param {Object} errorBody - The parsed JSON error body
 * @returns {boolean} True if it's an authorization error
 */
export function isCanvasAuthErrorBody(errorBody) {
    if (!errorBody) return false;

    // Check for "unauthorized" status
    if (errorBody.status === 'unauthorized') return true;

    // Check for authorization error messages
    if (errorBody.errors && Array.isArray(errorBody.errors)) {
        return errorBody.errors.some(err =>
            err.message && (
                err.message.toLowerCase().includes('unauthorized') ||
                err.message.toLowerCase().includes('not authorized')
            )
        );
    }

    return false;
}
