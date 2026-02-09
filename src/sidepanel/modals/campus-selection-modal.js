// Campus Selection Modal - Allows selecting which campus data to send
import { elements } from '../ui-manager.js';
import { trimCommonPrefix } from '../../constants/field-utils.js';

// Campus Selection Modal state
let campusSelectionResolve = null;

/**
 * Gets unique campuses from student array
 * @param {Array} students - Array of student objects
 * @returns {Array} Array of unique campus names, sorted alphabetically
 */
export function getCampusesFromStudents(students) {
    if (!students || students.length === 0) return [];

    const campuses = [...new Set(
        students
            .map(s => s.campus)
            .filter(c => c && c.trim() !== '')
    )].sort();

    return campuses;
}

/**
 * Opens the Campus Selection modal and returns a promise that resolves with the selected campus
 * @param {Array} campuses - Array of campus names
 * @param {string} [customMessage] - Optional custom message to display in the modal
 * @returns {Promise<string|null>} Promise that resolves with selected campus (empty string for all) or null if cancelled
 */
export function openCampusSelectionModal(campuses, customMessage = null) {
    return new Promise((resolve) => {
        if (!elements.campusSelectionModal || !elements.campusSelectionList) {
            resolve(''); // Default to all campuses if modal not available
            return;
        }

        // Store resolve function for later
        campusSelectionResolve = resolve;

        // Update message if custom message provided
        if (elements.campusSelectionMessage) {
            elements.campusSelectionMessage.textContent = customMessage ||
                'Multiple campuses detected. Select which campus data to send:';
        }

        // Clear existing buttons
        elements.campusSelectionList.innerHTML = '';

        // Create "All Campuses" button first
        const allButton = document.createElement('button');
        allButton.className = 'btn-secondary';
        allButton.style.cssText = 'width: 100%; text-align: left; padding: 12px 15px; display: flex; align-items: center; gap: 10px;';
        allButton.innerHTML = `
            <i class="fas fa-globe" style="color: #6366f1; font-size: 1.2em;"></i>
            <span style="flex: 1;">All Campuses</span>
            <span style="font-size: 0.85em; color: var(--text-secondary);">(Send all data)</span>
        `;
        allButton.addEventListener('click', () => {
            closeCampusSelectionModal('');
        });
        elements.campusSelectionList.appendChild(allButton);

        // Trim common prefix for cleaner display names
        const { trimmedNames } = trimCommonPrefix(campuses);

        // Create a button for each campus
        campuses.forEach(campus => {
            const displayName = trimmedNames.get(campus) || campus;
            const button = document.createElement('button');
            button.className = 'btn-secondary';
            button.style.cssText = 'width: 100%; text-align: left; padding: 12px 15px; display: flex; align-items: center; gap: 10px;';

            button.innerHTML = `
                <i class="fas fa-building" style="color: #6366f1; font-size: 1.2em;"></i>
                <span style="flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${displayName}</span>
            `;
            button.title = campus;

            button.addEventListener('click', () => {
                closeCampusSelectionModal(campus);
            });

            elements.campusSelectionList.appendChild(button);
        });

        // Show modal
        elements.campusSelectionModal.style.display = 'flex';
    });
}

/**
 * Closes the Campus Selection modal
 * @param {string|null} selectedCampus - The selected campus or null if cancelled
 */
export function closeCampusSelectionModal(selectedCampus = null) {
    if (elements.campusSelectionModal) {
        elements.campusSelectionModal.style.display = 'none';
    }

    // Resolve the promise with the selected campus
    if (campusSelectionResolve) {
        campusSelectionResolve(selectedCampus);
        campusSelectionResolve = null;
    }
}
