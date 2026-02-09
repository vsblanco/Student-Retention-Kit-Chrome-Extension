/**
 * Tests for Call Interface Styling and CSS Consistency
 *
 * These tests verify:
 * - CSS color values match CONFIG.COLORS definitions
 * - Call button sizing and structure in HTML
 * - Canvas login modal button sizing (box-sizing fix)
 * - Disposition grid layout structure
 */

const fs = require('fs');
const path = require('path');

// ============================================
// Load source files for validation
// ============================================

const cssContent = fs.readFileSync(
    path.resolve(__dirname, '../src/sidepanel/sidepanel.css'),
    'utf8'
);

const htmlContent = fs.readFileSync(
    path.resolve(__dirname, '../src/sidepanel/sidepanel.html'),
    'utf8'
);

const configContent = fs.readFileSync(
    path.resolve(__dirname, '../src/constants/config.js'),
    'utf8'
);

const callManagerContent = fs.readFileSync(
    path.resolve(__dirname, '../src/sidepanel/callManager.js'),
    'utf8'
);

// ============================================
// CSS / CONFIG Color Consistency
// ============================================

describe('CSS and CONFIG color consistency', () => {
    test('dial button CSS background matches CONFIG SUCCESS (#10b981)', () => {
        // .control-btn.dial should use the same green as CONFIG.COLORS.SUCCESS
        expect(cssContent).toMatch(/\.control-btn\.dial\s*\{[^}]*background:\s*#10b981/);
    });

    test('hangup button CSS background matches CONFIG ERROR (#ef4444)', () => {
        // .control-btn.hangup should use the same red as CONFIG.COLORS.ERROR
        expect(cssContent).toMatch(/\.control-btn\.hangup\s*\{[^}]*background:\s*#ef4444/);
    });

    test('automation mode CSS background matches CONFIG MUTED (#6b7280)', () => {
        // .control-btn.dial.automation should use the muted gray
        expect(cssContent).toMatch(/\.control-btn\.dial\.automation\s*\{[^}]*background:\s*#6b7280/);
    });

    test('status-indicator.ready matches CONFIG SUCCESS (#10b981)', () => {
        expect(cssContent).toMatch(/\.status-indicator\.ready\s*\{\s*background:\s*#10b981/);
    });

    test('CONFIG.COLORS.SUCCESS is defined as #10b981', () => {
        expect(configContent).toContain("SUCCESS: '#10b981'");
    });

    test('CONFIG.COLORS.ERROR is defined as #ef4444', () => {
        expect(configContent).toContain("ERROR: '#ef4444'");
    });

    test('CONFIG.COLORS.MUTED is defined as #6b7280', () => {
        expect(configContent).toContain("MUTED: '#6b7280'");
    });

    test('CONFIG.COLORS.WARNING is defined as #f59e0b', () => {
        expect(configContent).toContain("WARNING: '#f59e0b'");
    });

    test('CONFIG.COLORS.PRIMARY is defined as #3b82f6', () => {
        expect(configContent).toContain("PRIMARY: '#3b82f6'");
    });
});

// ============================================
// Call Button Sizing (CSS)
// ============================================

describe('Call button sizing in CSS', () => {
    test('dial button is 64x64px', () => {
        const dialBlock = cssContent.match(/\.control-btn\.dial\s*\{[^}]+\}/);
        expect(dialBlock).not.toBeNull();
        expect(dialBlock[0]).toContain('width: 64px');
        expect(dialBlock[0]).toContain('height: 64px');
    });

    test('hangup button is 64x64px (same as dial)', () => {
        const hangupBlock = cssContent.match(/\.control-btn\.hangup\s*\{[^}]+\}/);
        expect(hangupBlock).not.toBeNull();
        expect(hangupBlock[0]).toContain('width: 64px');
        expect(hangupBlock[0]).toContain('height: 64px');
    });

    test('base control button is 50x50px', () => {
        const baseBlock = cssContent.match(/\.control-btn\s*\{[^}]+\}/);
        expect(baseBlock).not.toBeNull();
        expect(baseBlock[0]).toContain('width: 50px');
        expect(baseBlock[0]).toContain('height: 50px');
    });

    test('all control buttons are circular (border-radius: 50%)', () => {
        const baseBlock = cssContent.match(/\.control-btn\s*\{[^}]+\}/);
        expect(baseBlock[0]).toContain('border-radius: 50%');
    });

    test('dial button has green box-shadow', () => {
        const dialBlock = cssContent.match(/\.control-btn\.dial\s*\{[^}]+\}/);
        expect(dialBlock[0]).toMatch(/box-shadow:.*rgba\(16,\s*185,\s*129/);
    });

    test('hangup button has red box-shadow', () => {
        const hangupBlock = cssContent.match(/\.control-btn\.hangup\s*\{[^}]+\}/);
        expect(hangupBlock[0]).toMatch(/box-shadow:.*rgba\(239,\s*68,\s*68/);
    });
});

// ============================================
// Call Interface HTML Structure
// ============================================

describe('Call interface HTML structure', () => {
    test('dial button exists with correct id and classes', () => {
        expect(htmlContent).toMatch(/id="dialBtn"/);
        expect(htmlContent).toMatch(/class="control-btn dial"/);
    });

    test('hangup button exists with correct id and classes', () => {
        expect(htmlContent).toMatch(/id="hangupBtn"/);
        expect(htmlContent).toMatch(/class="control-btn hangup"/);
    });

    test('hangup button is hidden by default', () => {
        expect(htmlContent).toMatch(/id="hangupBtn"[^>]*style="display:none;"/);
    });

    test('dial button contains phone icon', () => {
        // Match the dial button and verify it has the phone icon
        expect(htmlContent).toMatch(/id="dialBtn"[\s\S]*?fa-phone/);
    });

    test('hangup button contains phone-slash icon', () => {
        expect(htmlContent).toMatch(/id="hangupBtn"[\s\S]*?fa-phone-slash/);
    });

    test('call controls row contains mute, dial, hangup, and speaker buttons', () => {
        const controlsRow = htmlContent.match(/class="call-controls-row"[\s\S]*?<\/div>/);
        expect(controlsRow).not.toBeNull();
        expect(controlsRow[0]).toContain('mute');
        expect(controlsRow[0]).toContain('dialBtn');
        expect(controlsRow[0]).toContain('hangupBtn');
        expect(controlsRow[0]).toContain('speaker');
    });

    test('disposition section exists and is hidden by default', () => {
        expect(htmlContent).toMatch(/id="callDispositionSection"[^>]*style="display:none/);
    });

    test('has 6 disposition buttons', () => {
        const dispositionButtons = htmlContent.match(/class="disposition-btn"/g);
        expect(dispositionButtons).not.toBeNull();
        expect(dispositionButtons.length).toBe(6);
    });
});

// ============================================
// Canvas Login Modal Styling
// ============================================

describe('Canvas login modal styling', () => {
    test('Open Canvas Login button has box-sizing: border-box', () => {
        expect(htmlContent).toMatch(/id="canvasLoginLink"[^>]*box-sizing:\s*border-box/);
    });

    test('Open Canvas Login button has border-radius matching Resume button (0.75rem)', () => {
        expect(htmlContent).toMatch(/id="canvasLoginLink"[^>]*border-radius:\s*0\.75rem/);
    });

    test('Open Canvas Login button has btn-primary class', () => {
        expect(htmlContent).toMatch(/id="canvasLoginLink"[^>]*class="btn-primary"/);
    });

    test('Resume button has full width', () => {
        expect(htmlContent).toMatch(/id="canvasLoginResumeBtn"[^>]*width:\s*100%/);
    });

    test('Canvas logo is used in the modal header', () => {
        // The modal should include the canvas-logo.png
        const modalSection = htmlContent.match(/id="canvasLoginModal"[\s\S]*?<\/div>\s*<\/div>\s*<\/div>/);
        expect(modalSection).not.toBeNull();
        expect(modalSection[0]).toContain('canvas-logo.png');
    });

    test('btn-primary CSS has width: 100%', () => {
        const btnPrimaryBlock = cssContent.match(/\.btn-primary\s*\{[^}]+\}/);
        expect(btnPrimaryBlock).not.toBeNull();
        expect(btnPrimaryBlock[0]).toContain('width: 100%');
    });

    test('btn-primary and btn-secondary share the same border-radius', () => {
        const btnPrimaryBlock = cssContent.match(/\.btn-primary\s*\{[^}]+\}/);
        const btnSecondaryBlock = cssContent.match(/\.btn-secondary\s*\{[^}]+\}/);

        const primaryRadius = btnPrimaryBlock[0].match(/border-radius:\s*([^;]+)/);
        const secondaryRadius = btnSecondaryBlock[0].match(/border-radius:\s*([^;]+)/);

        expect(primaryRadius[1].trim()).toBe(secondaryRadius[1].trim());
    });
});

// ============================================
// CallManager Template Literal Regression
// ============================================

describe('callManager.js uses backtick template literals (not single quotes)', () => {
    test('no single-quoted CONFIG.COLORS in style.background assignments', () => {
        // This was the bug: '${CONFIG.COLORS.ERROR}' instead of `${CONFIG.COLORS.ERROR}`
        const singleQuotedPattern = /\.style\.background\s*=\s*'\$\{CONFIG\.COLORS\./g;
        const matches = callManagerContent.match(singleQuotedPattern);

        expect(matches).toBeNull();
    });

    test('no single-quoted CONFIG.COLORS in innerHTML assignments', () => {
        const singleQuotedPattern = /\.innerHTML\s*=\s*'[^']*\$\{CONFIG\.COLORS\./g;
        const matches = callManagerContent.match(singleQuotedPattern);

        expect(matches).toBeNull();
    });

    test('uses backtick template literals for style.background with CONFIG.COLORS', () => {
        const backtickPattern = /\.style\.background\s*=\s*`\$\{CONFIG\.COLORS\./g;
        const matches = callManagerContent.match(backtickPattern);

        expect(matches).not.toBeNull();
        expect(matches.length).toBeGreaterThan(0);
    });

    test('CONFIG is imported from constants/config.js', () => {
        expect(callManagerContent).toMatch(/import\s*\{\s*CONFIG\s*\}\s*from\s*['"]\.\.\/constants\/config/);
    });
});
