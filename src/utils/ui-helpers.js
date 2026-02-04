// UI Helper Utilities - Standardized functions for common UI patterns

/**
 * Updates a toggle element's visual state
 * @param {HTMLElement} element - The toggle icon element (Font Awesome)
 * @param {boolean} isEnabled - Whether the toggle should be on or off
 */
export function updateToggleUI(element, isEnabled) {
    if (!element) return;
    element.className = isEnabled ? 'fas fa-toggle-on' : 'fas fa-toggle-off';
    element.style.color = isEnabled ? 'var(--primary-color)' : 'gray';
}

/**
 * Checks if a toggle element is currently enabled
 * @param {HTMLElement} element - The toggle icon element
 * @returns {boolean} True if toggle is on
 */
export function isToggleEnabled(element) {
    if (!element) return false;
    return element.classList.contains('fa-toggle-on');
}

/**
 * Toggles a toggle element and returns the new state
 * @param {HTMLElement} element - The toggle icon element
 * @returns {boolean} The new state after toggling
 */
export function toggleAndGetState(element) {
    const newState = !isToggleEnabled(element);
    updateToggleUI(element, newState);
    return newState;
}

/**
 * Sets an element's enabled/disabled visual state (opacity + pointer-events)
 * @param {HTMLElement} element - The element to enable/disable
 * @param {boolean} isEnabled - Whether the element should be enabled
 * @param {string} disabledOpacity - Opacity when disabled (default '0.5')
 */
export function setElementEnabled(element, isEnabled, disabledOpacity = '0.5') {
    if (!element) return;
    element.style.opacity = isEnabled ? '1' : disabledOpacity;
    element.style.pointerEvents = isEnabled ? 'auto' : 'none';
}

/**
 * Updates a step element's icon (for progress indicators)
 * @param {HTMLElement} stepElement - The step container element
 * @param {'spinner' | 'check' | 'error' | 'circle'} status - The status to display
 */
export function updateStepIcon(stepElement, status) {
    if (!stepElement) return;
    const icon = stepElement.querySelector('i');
    if (!icon) return;

    const iconClasses = {
        spinner: 'fas fa-spinner fa-spin',
        check: 'fas fa-check',
        error: 'fas fa-times',
        circle: 'fas fa-circle',
        pending: 'far fa-circle'
    };

    icon.className = iconClasses[status] || iconClasses.circle;
}
