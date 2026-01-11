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
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'student-view',
        header: 'Student View',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    Student View is an interactive panel that displays information on your Excel sheet in an organized manner.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    It's also the place where you can submit comments to store in the Student History.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'student-history',
        header: 'Student History',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    This sheet maintains a record of student interactions and communications. New entries can be added through the Student View panel.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    There are two methods for making a comment:
                </p>
                <ul style="line-height: 2; margin-left: 20px; margin-top: 10px; font-size: 1.05em;">
                    <li>Manually create a comment from the Student View panel</li>
                    <li>Type in the Outreach Column, which will automatically generate one</li>
                </ul>
                <p style="margin-top: 20px; line-height: 1.8;">
                    <strong>Comments are automatically timestamped</strong> and can be organized using tags. This step helps organize the history sheet. Some tags serve more special purposes.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'tags-part1',
        header: 'Tags (Part 1)',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    Tags help you categorize and organize student interactions. Here are the available tags:
                </p>
                <h3 style="margin: 20px 0 10px 0; color: var(--primary-color); font-size: 1.1em;">General Tags:</h3>
                <ul style="line-height: 2; margin-left: 20px; font-size: 1.05em;">
                    <li><strong>Urgent</strong> - Reserved for urgent attention</li>
                    <li><strong>Note</strong> - A pinned note for general information</li>
                    <li><strong>Outreach</strong> - Sourced from the Outreach Column</li>
                    <li><strong>Quote</strong> - Contains quoted text</li>
                </ul>
                <h3 style="margin: 20px 0 10px 0; color: var(--primary-color); font-size: 1.1em;">Special Tags:</h3>
                <p style="line-height: 1.8; margin-top: 10px;">
                    Three tags have special functionality: <strong>LDA</strong>, <strong>DNC</strong>, and <strong>Contacted</strong>. These tags trigger specific automations and visual indicators in the system.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'tags-part2',
        header: 'Tags (Part 2): Special Tags',
        body: `
            <div class="tutorial-content">
                <h3 style="margin: 0 0 15px 0; color: var(--primary-color); font-size: 1.1em;">LDA Tag</h3>
                <p style="line-height: 1.8; font-size: 1.05em;">
                    Used as a follow-up tag. If a student says they will submit on the weekend, you can add an LDA tag for Saturday. When the LDA sheet for that day is created, you'll see a <strong>special indication highlighting their planned submission date</strong>. This helps you better keep track of when students are submitting.
                </p>
                <h3 style="margin: 20px 0 15px 0; color: var(--primary-color); font-size: 1.1em;">DNC (Do Not Contact)</h3>
                <p style="line-height: 1.8; font-size: 1.05em;">
                    If a student wishes to stop communication, insert this tag. The student will be <strong>crossed out on the LDA sheet</strong> and filtered out when sending emails. Subtags include DNC - Phone, DNC - Other Phone, and DNC - Email for specific contact preferences.
                </p>
                <h3 style="margin: 20px 0 15px 0; color: var(--primary-color); font-size: 1.1em;">Contacted</h3>
                <p style="line-height: 1.8; font-size: 1.05em;">
                    Shows if the student has been contacted that day. Special keywords in the Outreach column (like "will engage," "answered," "will submit," "come to class") will trigger this tag and <strong>auto-highlight the row in yellow</strong> to indicate contact was made.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'missing-assignments',
        header: 'Missing Assignments',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    This sheet contains a list of students' missing assignments.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    The report is generated using the <strong>Student Retention Kit—Chrome Extension</strong>.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'importing-data',
        header: 'Importing Data',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    Importing Data is handled through the <strong>Student Retention Kit—Chrome Extension</strong>.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    All it needs is a student population report with either a <strong>SyStudentId</strong> or <strong>Student SIS</strong> column in it.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    It will automatically organize the data and create a canvas for their grade book information. The system will automatically import the data into the Master List sheet for your viewing pleasure.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'creating-lda',
        header: 'Creating the LDA Sheet',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    You can create the LDA sheet by clicking on the <strong>Create LDA</strong> button on the ribbon.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'personalized-emails',
        header: 'Sending Personalized Emails',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    You can automatically send personalized emails to each student by granting consent to send emails on your behalf.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    You will see what parameters you can choose from to help personalize your emails. You can also create your own parameters for further customization.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    Alternatively, if you have a <strong>Power Automate Premium</strong> license, you can configure an HTTP request to create more advanced automations.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'calling-via-five9',
        header: 'Calling via Five9',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em;">
                    When you have a Five9 tab open and a student selected, you can call from the Student Retention Kit Side Panel.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    This feature acts as an assistant, allowing you to contact students more quickly without the need to copy and paste their phone numbers manually.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    Additionally, you can run a <strong>call sequence automation</strong> by selecting multiple students. When you initiate this automation, each time you finish a call, the system will automatically dial the next student in line.
                </p>
                <p style="margin-top: 20px; line-height: 1.8;">
                    This process helps you efficiently reach out to all your students in a timely manner.
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
