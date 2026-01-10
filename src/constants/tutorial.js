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
        id: 'welcome',
        header: 'Welcome to Student Retention Kit',
        body: `
            <div class="tutorial-welcome">
                <h2>👋 Hello!</h2>
                <p>Welcome to the Student Retention Kit Chrome Extension. This powerful tool helps you track student submissions, manage outreach, and improve student retention.</p>
                <p>This quick tutorial will walk you through the main features and get you started.</p>
                <p><strong>Ready to begin?</strong></p>
            </div>
        `,
        showSkip: true,
        showPrevious: false,
        showNext: true,
        nextLabel: 'Get Started'
    }
    // Add more tutorial pages here as needed
    // Example:
    // {
    //     id: 'checker-overview',
    //     header: 'Checker Tab',
    //     body: `
    //         <div class="tutorial-content">
    //             <p>The Checker tab allows you to scan for student submissions...</p>
    //         </div>
    //     `,
    //     showSkip: false,
    //     showPrevious: true,
    //     showNext: true,
    //     nextLabel: 'Next'
    // }
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
