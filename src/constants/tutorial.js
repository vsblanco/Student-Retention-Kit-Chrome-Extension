/**
 * Tutorial Configuration
 *
 * This file contains all tutorial pages and their content.
 * Customize the header and body for each page to guide new users through the extension.
 */

/**
 * Helper function to get the next Saturday from today
 * @returns {Object} Object containing date string and relative text
 */
function getNextSaturday() {
    const today = new Date();
    const dayOfWeek = today.getDay(); // 0 = Sunday, 6 = Saturday

    // Calculate days until next Saturday
    let daysUntilSaturday;
    if (dayOfWeek === 6) {
        // Today is Saturday, get next Saturday
        daysUntilSaturday = 7;
    } else {
        // Get upcoming Saturday
        daysUntilSaturday = (6 - dayOfWeek + 7) % 7;
        if (daysUntilSaturday === 0) daysUntilSaturday = 7;
    }

    // Calculate the next Saturday date
    const nextSaturday = new Date(today);
    nextSaturday.setDate(today.getDate() + daysUntilSaturday);

    // Format as MM-DD-YY
    const month = String(nextSaturday.getMonth() + 1).padStart(2, '0');
    const day = String(nextSaturday.getDate()).padStart(2, '0');
    const year = String(nextSaturday.getFullYear()).slice(-2);
    const dateString = `${month}-${day}-${year}`;

    // Determine relative text
    const relativeText = daysUntilSaturday <= 7 ? 'this Saturday' : 'next Saturday';

    return { dateString, relativeText };
}

/**
 * Tutorial Pages
 * Each page contains:
 * - id: Unique identifier for the page
 * - header: Title shown at the top of the tutorial page
 * - body: HTML content for the tutorial page (supports basic HTML tags) or a function that returns HTML
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
        id: 'tags',
        header: 'Tags',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em; margin-bottom: 20px;">
                    Tags help you categorize and organize student interactions. Here are the available tags:
                </p>
                <div style="line-height: 2.5; font-size: 1.05em;">
                    <div style="margin-bottom: 15px;">
                        <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fee2e2; color: #991b1b; margin-right: 10px;">Urgent</span>
                        <span>Reserved for urgent attention</span>
                    </div>
                    <div style="margin-bottom: 15px;">
                        <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #374151; color: #e5e7eb; margin-right: 10px;">Note</span>
                        <span>A pinned note for general information</span>
                    </div>
                    <div style="margin-bottom: 15px;">
                        <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #e5e7eb; color: #1f2937; margin-right: 10px;">Outreach</span>
                        <span>Sourced from the Outreach Column</span>
                    </div>
                    <div style="margin-bottom: 15px;">
                        <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #eff6ff; color: #1e40af; margin-right: 10px;">Quote</span>
                        <span>Contains quoted text</span>
                    </div>
                </div>
                <p style="line-height: 1.8; margin-top: 25px; font-size: 1.05em;">
                    Three tags have special functionality: <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fed7aa; color: #9a3412;">LDA</span>, <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fecaca; color: #000000;">DNC</span>, and <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fef08a; color: #854d0e;">Contacted</span>. These tags trigger specific automations and visual indicators in the system.
                </p>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'special-tags',
        header: 'Special Tags',
        body: () => {
            const { dateString, relativeText } = getNextSaturday();
            return `
                <div class="tutorial-content">
                    <div style="margin-bottom: 25px;">
                        <h3 style="margin: 0 0 10px 0; color: var(--primary-color); font-size: 1.1em;">
                            <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fed7aa; color: #9a3412; margin-right: 8px;">LDA</span>
                        </h3>
                        <p style="line-height: 1.8; font-size: 1.05em;">
                            Used as a follow-up tag. If a student says they will submit on the weekend, you can add an <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fed7aa; color: #9a3412;">LDA ${dateString}</span> for Saturday. When the LDA sheet for that day is created, you'll see a <strong>special indication highlighting their planned submission date</strong>. This helps you better keep track of when students are submitting.
                        </p>
                        <p style="line-height: 1.8; font-size: 0.95em; margin-top: 10px; font-style: italic; color: var(--text-secondary);">
                            Example: "Student plans to submit ${relativeText}"
                        </p>
                    </div>
                    <div style="margin-bottom: 25px;">
                        <h3 style="margin: 0 0 10px 0; color: var(--primary-color); font-size: 1.1em;">
                            <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fecaca; color: #000000; margin-right: 8px;">DNC</span>
                            (Do Not Contact)
                        </h3>
                        <p style="line-height: 1.8; font-size: 1.05em;">
                            If a student wishes to stop communication, insert this tag. The student will be <strong>crossed out on the LDA sheet</strong> and filtered out when sending emails. Subtags include <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fecaca; color: #000000; font-size: 0.9em;">DNC - Phone</span>, <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fecaca; color: #000000; font-size: 0.9em;">DNC - Other Phone</span>, and <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fecaca; color: #000000; font-size: 0.9em;">DNC - Email</span> for specific contact preferences.
                        </p>
                    </div>
                    <div>
                        <h3 style="margin: 0 0 10px 0; color: var(--primary-color); font-size: 1.1em;">
                            <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fef08a; color: #854d0e; margin-right: 8px;">Contacted</span>
                        </h3>
                        <p style="line-height: 1.8; font-size: 1.05em;">
                            Shows if the student has been contacted that day. Special keywords in the Outreach column (like "will engage," "answered," "will submit," "come to class") will trigger this tag and <strong>auto-highlight the row in yellow</strong> to indicate contact was made.
                        </p>
                    </div>
                </div>
            `;
        },
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
                <p style="margin-top: 20px; line-height: 1.8;">
                    You can use the <strong>Special Parameter: Missing Assignments List</strong> in sending personalized emails to students. This creates clickable hyperlinks that lead you straight to the assignments.
                </p>
                <h3 style="margin: 25px 0 15px 0; color: var(--primary-color); font-size: 1.1em;">Example Email:</h3>
                <div style="background: rgba(255, 255, 255, 0.4); border: 1px solid rgba(0, 0, 0, 0.05); border-radius: 0.75rem; padding: 15px; margin-bottom: 20px; font-size: 0.95em; line-height: 1.8;">
                    <p style="margin: 0 0 10px 0;">Hello Jane,</p>
                    <p style="margin: 0 0 10px 0;">I hope you are having a great day. I noticed it's been 5 days since you last engaged with your course.</p>
                    <p style="margin: 0 0 5px 0;">You are currently missing:</p>
                    <ul style="margin: 5px 0 10px 20px; padding: 0;">
                        <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">Discussion Post 2.0</span></li>
                        <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">Mid Term</span></li>
                        <li style="margin-bottom: 5px;"><span style="color: var(--primary-color); text-decoration: underline; cursor: pointer;">Mind Tap 2.2 - Human Anatomy</span></li>
                    </ul>
                </div>
            </div>
        `,
        showSkip: true,
        showPrevious: true,
        showNext: true,
        nextLabel: 'Next'
    },
    {
        id: 'submission-checker',
        header: 'Submission Checker',
        body: `
            <div class="tutorial-content">
                <p style="line-height: 1.8; font-size: 1.05em; margin-bottom: 20px;">
                    This tool automatically communicates with the Canvas API to track when a student submits an assignment. You will see their name, timestamp, and assignment title with a link to quickly view their submission when you click on their name.
                </p>
                <h3 style="margin: 20px 0 15px 0; color: var(--primary-color); font-size: 1.1em;">Example Submission:</h3>
                <div style="background: rgba(255, 255, 255, 0.4); border: 1px solid rgba(0, 0, 0, 0.05); border-radius: 0.75rem; padding: 12px; margin-bottom: 20px;">
                    <div style="display: flex; align-items: center; width:100%;">
                        <div style="width: 6px; height: 32px; border-radius: 4px; margin-right: 12px; background-color: #10b981;"></div>
                        <div style="flex-grow:1; display:flex; justify-content:space-between; align-items:center;">
                            <div style="display:flex; flex-direction:column;">
                                <span style="font-weight:500; color:var(--primary-color); cursor:pointer;">Doe, John</span>
                                <span style="font-size:0.8em; color:var(--text-secondary);">Discussion Post 2.0</span>
                            </div>
                            <span style="font-size: 0.8em; padding: 4px 10px; border-radius: 12px; background-color: rgba(0, 0, 0, 0.06); color: var(--text-secondary);">2:45 PM</span>
                        </div>
                    </div>
                </div>
                <p style="line-height: 1.8; font-size: 1.05em; margin-bottom: 15px;">
                    If you have your Excel workbook open, it will automatically look for the student to highlight in green, indicating they submitted. It will also input in the outreach column: <strong>'Submitted {Assignment Title}'</strong>
                </p>
                <p style="line-height: 1.8; font-size: 1.05em;">
                    You can filter which students you want to check based on their days out and grades.
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
                <p style="line-height: 1.8; font-size: 1.05em; margin-bottom: 20px;">
                    You can create the LDA sheet by clicking on the <strong>Create LDA</strong> button on the ribbon.
                </p>
                <h3 style="margin: 20px 0 15px 0; color: var(--primary-color); font-size: 1.1em;">Customizable Settings:</h3>
                <div style="line-height: 2; font-size: 1.05em;">
                    <div style="margin-bottom: 12px;">
                        <strong>Days Out Filter</strong> - Choose how many days out to include students who haven't submitted
                    </div>
                    <div style="margin-bottom: 12px;">
                        <strong>Include Failing List</strong> - Optionally add students who are currently failing to the LDA sheet
                    </div>
                    <div style="margin-bottom: 12px;">
                        <strong>Include LDA Tag Indicators</strong> - Display <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fed7aa; color: #9a3412; font-size: 0.95em;">LDA</span> tag markers showing students' planned submission dates
                    </div>
                    <div style="margin-bottom: 12px;">
                        <strong>Include DNC Tag Indicators</strong> - Show <span style="display: inline-block; padding: 0.5px 6px; font-weight: 600; border-radius: 9999px; background-color: #fecaca; color: #000000; font-size: 0.95em;">DNC</span> tag markers to identify students who requested no contact
                    </div>
                </div>
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
