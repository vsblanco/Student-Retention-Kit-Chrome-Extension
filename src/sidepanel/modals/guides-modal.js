// Guides Modal - Handles the guides/help dialog
import { elements } from '../ui-manager.js';

/**
 * Opens the guides modal
 */
export function openGuidesModal() {
    if (!elements.guidesModal) return;
    elements.guidesModal.style.display = 'flex';
}

/**
 * Closes the guides modal
 */
export function closeGuidesModal() {
    if (!elements.guidesModal) return;
    elements.guidesModal.style.display = 'none';
}
