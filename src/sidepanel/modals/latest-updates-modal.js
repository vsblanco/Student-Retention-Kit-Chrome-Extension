// Latest Updates Modal - Shows release notes for new versions
import { STORAGE_KEYS } from '../../constants/index.js';
import { storageGetValue, storageSet } from '../../utils/storage.js';
import { hasReleaseNotes, getLatestReleaseNotes } from '../../constants/release-notes.js';
import { elements } from '../ui-manager.js';

/**
 * Checks if the latest updates modal should be shown
 * Returns true if:
 * 1. The current version has release notes defined
 * 2. The user hasn't seen the updates for this version yet
 * @returns {Promise<boolean>} Whether the modal should be shown
 */
export async function shouldShowLatestUpdatesModal() {
    try {
        // Get the current version from the manifest
        const manifest = chrome.runtime.getManifest();
        const currentVersion = manifest.version;

        // Check if there are release notes for this version
        if (!hasReleaseNotes(currentVersion)) {
            return false;
        }

        // Get the last seen version from storage
        const lastSeenVersion = await storageGetValue(STORAGE_KEYS.LAST_SEEN_VERSION, null);

        // Show modal if user hasn't seen this version's updates
        return lastSeenVersion !== currentVersion;
    } catch (error) {
        console.error('Error checking if latest updates modal should show:', error);
        return false;
    }
}

/**
 * Opens the latest updates modal and populates it with release notes
 */
export function openLatestUpdatesModal() {
    if (!elements.latestUpdatesModal) return;

    // Get the latest release notes (version comes from release-notes.js, not manifest)
    const latest = getLatestReleaseNotes();

    if (!latest) {
        console.warn('No release notes found');
        return;
    }

    const { version, notes: releaseNotes } = latest;

    // Update the modal title
    if (elements.latestUpdatesTitle) {
        elements.latestUpdatesTitle.textContent = releaseNotes.title || "What's New";
    }

    // Update the version display (uses version from release notes)
    if (elements.latestUpdatesVersion) {
        elements.latestUpdatesVersion.textContent = `Version ${version}`;
    }

    // Update the date display
    if (elements.latestUpdatesDate) {
        if (releaseNotes.date) {
            elements.latestUpdatesDate.textContent = `Last Updated: ${releaseNotes.date}`;
            elements.latestUpdatesDate.style.display = '';
        } else {
            elements.latestUpdatesDate.style.display = 'none';
        }
    }

    // Populate the updates list
    if (elements.latestUpdatesList) {
        elements.latestUpdatesList.innerHTML = '';

        if (releaseNotes.updates && Array.isArray(releaseNotes.updates)) {
            releaseNotes.updates.forEach(update => {
                const li = document.createElement('li');
                li.textContent = update;
                elements.latestUpdatesList.appendChild(li);
            });
        }
    }

    // Show the modal
    elements.latestUpdatesModal.style.display = 'flex';
}

/**
 * Closes the latest updates modal and marks the version as seen
 */
export async function closeLatestUpdatesModal() {
    if (!elements.latestUpdatesModal) return;

    // Hide the modal
    elements.latestUpdatesModal.style.display = 'none';

    // Mark this version as seen
    const manifest = chrome.runtime.getManifest();
    const currentVersion = manifest.version;

    await storageSet({ [STORAGE_KEYS.LAST_SEEN_VERSION]: currentVersion });
    console.log(`Marked version ${currentVersion} as seen`);
}
