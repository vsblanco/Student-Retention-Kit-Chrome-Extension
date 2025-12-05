# **Submission Checker for Canvas**

**Version:** 9.0

The **Submission Checker** is a powerful Chrome Extension designed to automate and streamline the process of checking for student submissions within the Canvas Learning Management System. 

**New in Version 9.0:** The extension now operates using a **Headless API Engine**. Instead of physically opening tabs, it communicates directly with the Canvas API in the background. This makes checking significantly faster, more reliable, and allows for robust reporting on missing assignments.

## **🚀 Key Features**

* **⚡ Headless API Checking:** Checks student data in the background using the Canvas API. No more opening dozens of tabs—the process is silent and fast.
* **Two Operation Modes:**
    * **Submission Check:** Identifies students who have submitted work matching a specific keyword (e.g., today's date).
    * **Missing Assignments:** Scans gradebooks to generate a comprehensive report of missing work, zeros, and unsubmitted past-due items.
* **Batch Processing:** Intelligently groups students by course and fetches data in batches (30 students at a time) to minimize network requests and maximize speed.
* **Detailed Reporting:** Generates downloadable reports (CSV or JSON) containing student grades, missing assignment counts, and direct links to submissions.
* **Smart "Missing" Logic:** Automatically detects missing work based on:
    * Official Canvas "Missing" flags.
    * Unsubmitted items that are past the due date.
    * **Manual Zeros:** Scores entered as "0" with no submission (even if the due date is in the future).
* **External Integrations:** Connect to **Power Automate** and **Pusher** to trigger workflows or update dashboards in real-time.
* **Embed in Canvas:** Optionally injects visual helpers (a search bar and "Missing" pills) directly into the Canvas gradebook UI for manual browsing.

## **📥 Installation**

Since this is a custom extension, it needs to be loaded into Chrome in Developer Mode.

1.  Download or clone this repository to your local machine.
2.  Open Google Chrome and navigate to `chrome://extensions`.
3.  Enable **"Developer mode"** using the toggle in the top-right corner.
4.  Click on the **"Load unpacked"** button.
5.  Select the folder where you saved the extension files.
6.  The "Submission Checker" will now appear in your extensions list.

## **🛠 How to Use**

### **1. Prepare Your Master List**

The extension relies on a "Master List" of students formatted as a JSON array.

**JSON Structure:**
```json
[
  {
    "StudentName": "Doe, Jane",
    "GradeBook": "[https://canvas.example.com/courses/123/grades/456](https://canvas.example.com/courses/123/grades/456)",
    "DaysOut": 5,
    "LDA": "2023-10-27T10:00:00Z",
    "Grade": "95%"
  },
  {
    "StudentName": "Smith, John",
    "GradeBook": "[https://canvas.example.com/courses/123/grades/457](https://canvas.example.com/courses/123/grades/457)",
    "DaysOut": 12,
    "LDA": "2023-10-16T11:30:00Z",
    "Grade": "88%"
  }
]