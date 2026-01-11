/**
 * Tutorial Configuration
 *
 * This file contains all tutorial pages and their content.
 * Customize the header and body for each page to guide new users through the extension.
 */

/**
 * Tutorial Pages
 * Each page contains:
 * - id: Unique identifier for the page
 * - header: Title shown at the top of the tutorial page
 * - body: HTML content for the tutorial page (supports basic HTML tags)
 * - showSkip: Whether to show the skip button on this page
 * - showPrevious: Whether to show the previous button on this page
 * - showNext: Whether to show the next button on this page
 * - nextLabel: Custom label for the next button (default: "Next")
 */
export const TUTORIAL_PAGES = [
    {
        id: 'what-is-srk',
        header: 'What is Student Retention Kit?',
        body: `
            <div class="tutorial-content" style="text-align: center;">
                <p style="font-size: 1.1em; line-height: 1.8; margin: 20px 0;">
                    The Student Retention Kit is a tool designed to help educators identify and support at-risk students.
                    Its goal is to make your workflow as efficiently as possible, so that you can focus on what's most important.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: false,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'what-can-i-do',
        header: 'What can I do with this?',
        body: `
            <div class="tutorial-content">
                <p style="margin-bottom: 20px;">There's a variety of features bundled in this kit. They include:</p>
                <ul style="line-height: 2; margin-left: 20px; font-size: 1.05em;">
                    <li>Importing external reports onto your sheets</li>
                    <li>Automatic LDA creation</li>
                    <li>Sending personalized emails to students</li>
                    <li>Real-time student submission feedback</li>
                    <li>Student communication tracking</li>
					<li>Calling Students via Five9</li>
                </ul>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'initial-setup',
        header: 'Initial Setup',
        body: `
            <div class="tutorial-content">
                <p style="margin-bottom: 20px;">Before we continue further, let's make sure your workbook is set up correctly.</p>
                <h3 style="margin: 20px 0 15px 0; color: var(--primary-color);">Checklist:</h3>
                <div class="tutorial-checklist">
                    <div class="checklist-item excel-connection-item">
                        <span>Excel Workbook</span>
                        <span class="connection-status" id="tutorialExcelStatus">Waiting...</span>
                    </div>
                    <div class="checklist-item sheet-item" id="masterListItem">
                        <span>Master List Sheet</span>
                        <button class="btn-secondary tutorial-create-btn" style="padding: 5px 15px; font-size: 0.9em;">Create</button>
                    </div>
                    <div class="checklist-item sheet-item" id="studentHistoryItem">
                        <span>Student History Sheet</span>
                        <button class="btn-secondary tutorial-create-btn" style="padding: 5px 15px; font-size: 0.9em;">Create</button>
                    </div>
                    <div class="checklist-item sheet-item" id="missingAssignmentsItem">
                        <span>Missing Assignments Sheet</span>
                        <button class="btn-secondary tutorial-create-btn" style="padding: 5px 15px; font-size: 0.9em;">Create</button>
                    </div>
                </div>
                <p style="margin-top: 20px; font-size: 0.9em; color: var(--text-secondary);">
                    <em>Note: Open your Excel workbook with the Student Retention Kit Add-in to connect.</em>
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'master-list',
        header: 'Master List',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    This sheet contains the complete student population of your campus. It serves as the target for your imports and the source for your LDA.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    You can update your Master List by clicking the <strong>Update Master List</strong> button in the Data tab. Then simply upload your student population report.
                </p>
                <p style="margin-top: 30px; text-align: center;">
                    <strong>You're all set! Click "Finish" to start using the extension.</strong>
                </p>
            </div>
        `,
        showSkip: false,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Finish'
    }
    // Add more tutorial pages here as needed
];

/**
 * Tutorial Settings
 */
export const TUTORIAL_SETTINGS = {
    // Whether to show the tutorial automatically for new users
    showForNewUsers: true,

    // Whether to grey out other tabs during tutorial
    greyOutTabs: true,

    // Animation duration for page transitions (in ms)
    transitionDuration: 300
};
