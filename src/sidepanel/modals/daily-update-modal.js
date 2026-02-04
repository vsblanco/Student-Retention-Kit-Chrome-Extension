// Daily Update Modal - Reminds users to update their master list daily
import { STORAGE_KEYS } from '../../constants/index.js';
import { storageGetValue } from '../../utils/storage.js';
import { elements } from '../ui-manager.js';

/**
 * Checks if the daily update modal should be shown
 * Returns true if the modal should be shown, false otherwise
 */
export async function shouldShowDailyUpdateModal() {
    const lastUpdated = await storageGetValue(STORAGE_KEYS.LAST_UPDATED);

    // If there's no master list yet, don't show the modal
    if (!lastUpdated) {
        return false;
    }

    const now = new Date();
    const todayDateString = now.toLocaleDateString('en-US');

    // Check if master list was updated today
    const lastUpdatedDate = new Date(lastUpdated);
    const lastUpdatedDateString = lastUpdatedDate.toLocaleDateString('en-US');

    if (todayDateString === lastUpdatedDateString) {
        // Master list already updated today
        return false;
    }

    // Show modal if:
    // 1. Master list exists
    // 2. Master list hasn't been updated today
    return true;
}

/**
 * Opens the daily update modal
 */
export function openDailyUpdateModal() {
    if (!elements.dailyUpdateModal) return;
    elements.dailyUpdateModal.style.display = 'flex';
}

/**
 * Closes the daily update modal
 */
export async function closeDailyUpdateModal() {
    if (!elements.dailyUpdateModal) return;
    elements.dailyUpdateModal.style.display = 'none';
}
