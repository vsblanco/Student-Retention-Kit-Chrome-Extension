// Canvas Login Modal - Pre-check for Canvas session before Step 2
import { elements } from '../ui-manager.js';

let canvasLoginResolve = null;

/**
 * Checks if the user is logged into Canvas by making a lightweight API call.
 * @param {string} canvasDomain - The Canvas domain to check (e.g. https://northbridge.instructure.com)
 * @returns {Promise<boolean>} True if logged in, false otherwise
 */
export async function isLoggedIntoCanvas(canvasDomain) {
    try {
        const response = await fetch(`${canvasDomain}/api/v1/users/self`, {
            method: 'GET',
            credentials: 'include',
            headers: { 'Accept': 'application/json' }
        });
        return response.ok;
    } catch {
        return false;
    }
}

/**
 * Opens the Canvas login modal and waits for the user to click Resume.
 * @param {string} canvasDomain - Used to build the login link
 * @returns {Promise<void>} Resolves when the user clicks Resume
 */
function openCanvasLoginModal(canvasDomain) {
    return new Promise((resolve) => {
        canvasLoginResolve = resolve;

        // Set the login link href
        if (elements.canvasLoginLink) {
            elements.canvasLoginLink.href = `${canvasDomain}/login/`;
        }

        if (elements.canvasLoginModal) {
            elements.canvasLoginModal.style.display = 'flex';
        }
    });
}

/**
 * Closes the Canvas login modal and resolves the pending promise.
 */
export function closeCanvasLoginModal() {
    if (elements.canvasLoginModal) {
        elements.canvasLoginModal.style.display = 'none';
    }

    if (canvasLoginResolve) {
        canvasLoginResolve();
        canvasLoginResolve = null;
    }
}

/**
 * Ensures the user is logged into Canvas before proceeding.
 * Shows the login modal in a loop until a successful login check passes.
 * @param {string} canvasDomain - The Canvas domain to check
 */
export async function ensureCanvasLogin(canvasDomain) {
    while (true) {
        const loggedIn = await isLoggedIntoCanvas(canvasDomain);
        if (loggedIn) return;

        console.warn('[Canvas] Not logged in — showing login modal');
        await openCanvasLoginModal(canvasDomain);

        // User clicked Resume — loop back to re-check
    }
}
