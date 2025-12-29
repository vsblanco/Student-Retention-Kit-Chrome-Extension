// content/excel_connector.js

// Prevent multiple injections
if (window.hasSRKConnectorRun) {
  // Script already loaded, stop here.
  console.log("SRK Connector already active.");
} else {
  window.hasSRKConnectorRun = true;

  console.log("%c SRK: Excel Connector Script LOADED", "background: #222; color: #bada55; font-size: 14px");

  /**
   * Normalizes a field name for comparison by:
   * - Converting to lowercase
   * - Removing spaces, hyphens, underscores
   * - Removing special characters
   *
   * @param {string} fieldName - The field name to normalize
   * @returns {string} The normalized field name
   */
  function normalizeFieldName(fieldName) {
    if (!fieldName) return '';
    return String(fieldName)
        .toLowerCase()
        .replace(/[\s\-_]/g, '') // Remove spaces, hyphens, underscores
        .replace(/[^a-z0-9]/g, ''); // Remove any remaining special characters
  }

  /**
   * Field aliases for semantic field matching.
   * NOTE: Case/space variations are handled automatically by normalizeFieldName().
   * Only specify true semantic aliases here.
   */
  const FIELD_ALIASES = {
    name: ['studentname', 'student'],
    phone: ['primaryphone', 'phonenumber', 'mobile', 'cell', 'cellphone', 'contact', 'telephone', 'otherphone'],
    grade: ['gradelevel', 'level'],
    StudentNumber: ['studentid', 'sisid'],
    SyStudentId: ['studentsis'],
    daysOut: ['dayssincepriorlda', 'daysinactive', 'days']
  };

  /**
   * Finds a field value in an object using normalized matching and aliases.
   * Matches fields case-insensitively and ignoring spaces/special characters.
   *
   * @param {Object} obj - The object to search in
   * @param {String} fieldName - The internal field name
   * @param {*} defaultValue - Default value if field not found
   * @returns {*} The field value or defaultValue
   */
  function getFieldWithAlias(obj, fieldName, defaultValue = null) {
    if (!obj || !fieldName) return defaultValue;

    // Normalize the target field name
    const normalizedFieldName = normalizeFieldName(fieldName);

    // Get aliases for this field
    const aliases = FIELD_ALIASES[fieldName] || [];
    const normalizedAliases = aliases.map(alias => normalizeFieldName(alias));

    // Search through object keys
    for (const key in obj) {
      const normalizedKey = normalizeFieldName(key);

      // Check direct match
      if (normalizedKey === normalizedFieldName) {
        const value = obj[key];
        return value !== null && value !== undefined ? value : defaultValue;
      }

      // Check alias matches
      if (normalizedAliases.includes(normalizedKey)) {
        const value = obj[key];
        return value !== null && value !== undefined ? value : defaultValue;
      }
    }

    return defaultValue;
  }

  // Notify extension that connector is active
  chrome.runtime.sendMessage({
      type: "SRK_CONNECTOR_ACTIVE",
      timestamp: Date.now()
  }).catch(() => {
      // Extension might not be ready yet, that's ok
  });

  window.addEventListener("message", (event) => {
      if (!event.data || !event.data.type) return;

      // Check for the Ping from Office Add-in
      if (event.data.type === "SRK_CHECK_EXTENSION") {
          console.log("%c SRK Connector: Ping Received! Ponging Sender...", "color: green; font-weight: bold");

          // Reply specifically to the window that sent the message
          if (event.source) {
              event.source.postMessage({ type: "SRK_EXTENSION_INSTALLED" }, "*");
          }

          // Notify extension that Office Add-in is connected
          chrome.runtime.sendMessage({
              type: "SRK_OFFICE_ADDIN_CONNECTED",
              timestamp: Date.now()
          }).catch(() => {
              // Extension might not be ready, that's ok
          });
      }

      // Handle Master List Request
      else if (event.data.type === "SRK_REQUEST_MASTER_LIST") {
          console.log("%c SRK Connector: Master List Request Received", "color: blue; font-weight: bold");
          console.log("Request timestamp:", event.data.timestamp);

          // Check setting to determine if we should accept the data
          checkIfShouldAcceptMasterList(event.source);
      }

      // Handle Master List Data
      else if (event.data.type === "SRK_MASTER_LIST_DATA") {
          console.log("%c SRK Connector: Master List Data Received!", "color: green; font-weight: bold");
          handleMasterListData(event.data.data);
      }

      // Handle Selected Students
      else if (event.data.type === "SRK_SELECTED_STUDENTS") {
          console.log("%c SRK Connector: Selected Students Received!", "color: purple; font-weight: bold");
          handleSelectedStudents(event.data.data);
      }

      // Handle Office User Info
      else if (event.data.type === "SRK_OFFICE_USER_INFO") {
          console.log("%c SRK Connector: Office User Info Received!", "color: cyan; font-weight: bold");
          handleOfficeUserInfo(event.data.data);
      }
  });

  /**
   * Checks if we should accept the master list update based on user settings
   * @param {MessageEventSource} source - The event source to send response to
   */
  function checkIfShouldAcceptMasterList(source) {
      chrome.storage.local.get(['autoUpdateMasterList', 'lastUpdated'], (result) => {
          const setting = result.autoUpdateMasterList || 'always';
          let wantsData = false;

          if (setting === 'never') {
              console.log("%c Auto-update is disabled. Rejecting data.", "color: orange");
              wantsData = false;
          } else if (setting === 'always') {
              console.log("%c Auto-update is set to always. Accepting data.", "color: green");
              wantsData = true;
          } else if (setting === 'once-daily') {
              // Check if last update was today
              const lastUpdated = result.lastUpdated;
              const isToday = checkIfUpdatedToday(lastUpdated);

              if (isToday) {
                  console.log("%c Already updated today. Rejecting data.", "color: orange");
                  wantsData = false;
              } else {
                  console.log("%c Not updated today. Accepting data.", "color: green");
                  wantsData = true;
              }
          }

          // Send response
          if (source) {
              source.postMessage({
                  type: "SRK_MASTER_LIST_RESPONSE",
                  wantsData: wantsData
              }, "*");
          }
      });
  }

  /**
   * Checks if the last update timestamp was today
   * @param {string} lastUpdated - The last updated timestamp
   * @returns {boolean} True if last update was today
   */
  function checkIfUpdatedToday(lastUpdated) {
      if (!lastUpdated) return false;

      try {
          const lastUpdateDate = new Date(lastUpdated);
          const today = new Date();

          return lastUpdateDate.getDate() === today.getDate() &&
                 lastUpdateDate.getMonth() === today.getMonth() &&
                 lastUpdateDate.getFullYear() === today.getFullYear();
      } catch (error) {
          console.error("Error parsing last updated date:", error);
          return false;
      }
  }

  /**
   * Handles incoming Master List data from the Office Add-in
   * Transforms the data from add-in format to extension format and stores it
   * Uses FIELD_ALIASES to handle different capitalizations and field name variations
   * Dynamically includes all fields from source data, applying special logic only where needed
   */
  function handleMasterListData(data) {
      try {
          console.log(`Processing Master List with ${data.totalStudents} students`);
          console.log("Data timestamp:", data.timestamp);

          // Transform students from add-in format to extension format
          const transformedStudents = data.students.map(student => {
              // Start by including all fields from the source student object
              const transformedStudent = {};

              // Copy all fields from source to preserve all Excel columns
              for (const key in student) {
                  transformedStudent[key] = student[key];
              }

              // Use getFieldWithAlias to ensure our standard field names are present
              // This handles cases where fields come with different names (aliases)

              // Name - ensure it exists and has a fallback
              transformedStudent.name = getFieldWithAlias(student, 'name', 'Unknown') || 'Unknown';

              // Phone - use aliases to find it
              const phone = getFieldWithAlias(student, 'phone');
              if (phone !== null && phone !== undefined) {
                  transformedStudent.phone = phone;
              }

              // StudentNumber - use aliases to find it
              const StudentNumber = getFieldWithAlias(student, 'StudentNumber');
              if (StudentNumber !== null && StudentNumber !== undefined) {
                  transformedStudent.StudentNumber = StudentNumber;
              }

              // SyStudentId - use aliases to find it
              const SyStudentId = getFieldWithAlias(student, 'SyStudentId');
              if (SyStudentId !== null && SyStudentId !== undefined) {
                  transformedStudent.SyStudentId = SyStudentId;
              }

              // SPECIAL LOGIC FIELDS - These need transformation or initialization

              // Grade - convert to string
              const gradeValue = getFieldWithAlias(student, 'grade');
              transformedStudent.grade = gradeValue !== undefined && gradeValue !== null ? String(gradeValue) : null;

              // Days Out - convert to integer and normalize field name to 'daysout'
              const daysOutValue = getFieldWithAlias(student, 'daysOut');
              transformedStudent.daysout = parseInt(daysOutValue) || 0;

              // Missing Count - initialize if not present
              if (!('missingCount' in transformedStudent)) {
                  transformedStudent.missingCount = 0;
              }

              // Assignments - initialize if not present
              if (!('assignments' in transformedStudent)) {
                  transformedStudent.assignments = [];
              }

              return transformedStudent;
          });

          const lastUpdated = new Date().toLocaleString('en-US', {
              year: 'numeric',
              month: 'numeric',
              day: 'numeric',
              hour: '2-digit',
              minute: '2-digit'
          });

          // Store the transformed data using chrome storage
          chrome.storage.local.set({
              masterEntries: transformedStudents,
              lastUpdated: lastUpdated,
              masterListSourceTimestamp: data.timestamp
          }, () => {
              console.log(`%c ✓ Master List Updated Successfully!`, "color: green; font-weight: bold");
              console.log(`   Students: ${transformedStudents.length}`);
              console.log(`   Updated: ${lastUpdated}`);

              // Notify the extension that master list has been updated
              chrome.runtime.sendMessage({
                  type: "SRK_MASTER_LIST_UPDATED",
                  timestamp: Date.now(),
                  studentCount: transformedStudents.length,
                  sourceTimestamp: data.timestamp
              }).catch(() => {
                  // Extension might not be ready, that's ok
              });
          });

      } catch (error) {
          console.error("%c Error processing Master List data:", "color: red; font-weight: bold", error);

          // Notify extension of error
          chrome.runtime.sendMessage({
              type: "SRK_MASTER_LIST_ERROR",
              error: error.message,
              timestamp: Date.now()
          }).catch(() => {});
      }
  }

  /**
   * Handles incoming Selected Students data from the Office Add-in
   * Transforms and sends student data to the extension
   * @param {Object} data - The selected students data
   */
  function handleSelectedStudents(data) {
      try {
          console.log(`Processing ${data.count} selected student(s)`);
          console.log("Selection timestamp:", data.timestamp);

          if (!data.students || data.students.length === 0) {
              console.log("No students selected");
              return;
          }

          // Check if sync active student is enabled
          chrome.storage.local.get(['syncActiveStudent'], (result) => {
              const syncEnabled = result.syncActiveStudent !== undefined ? result.syncActiveStudent : true;

              if (!syncEnabled) {
                  console.log("%c Sync Active Student is disabled. Skipping student sync.", "color: orange; font-weight: bold");
                  return;
              }

              // Transform all students from add-in format to extension format
              const transformedStudents = data.students.map(student => ({
                  name: student.name || 'Unknown',
                  phone: student.phone || student.otherPhone || null,
                  SyStudentId: student.syStudentId || null,
                  // Set defaults for fields not provided by the Office add-in
                  grade: null,
                  StudentNumber: null,
                  daysout: 0,
                  missingCount: 0,
                  url: null,
                  assignments: []
              }));

              if (data.count === 1) {
                  console.log(`Setting active student: ${transformedStudents[0].name}`);
              } else {
                  console.log(`Setting up automation mode with ${data.count} students`);
              }

              // Send to extension (works for both single and multiple)
              chrome.runtime.sendMessage({
                  type: "SRK_SELECTED_STUDENTS",
                  students: transformedStudents,
                  count: data.count,
                  timestamp: Date.now(),
                  sourceTimestamp: data.timestamp
              }).catch(() => {
                  // Extension might not be ready, that's ok
              });
          });

      } catch (error) {
          console.error("%c Error processing Selected Students data:", "color: red; font-weight: bold", error);
      }
  }

  /**
   * Handles incoming Office User Info from the Office Add-in
   * Stores the authenticated user's name, email, and other profile data
   * @param {Object} data - The user info data
   */
  function handleOfficeUserInfo(data) {
      try {
          console.log(`Processing Office User Info`);
          console.log(`  Name: ${data.name}`);
          console.log(`  Email: ${data.email}`);
          console.log("User info timestamp:", data.timestamp);

          // Validate required fields
          if (!data.name && !data.email) {
              console.warn("Office user info missing both name and email");
              return;
          }

          // Store the user info
          const userInfo = {
              name: data.name || null,
              email: data.email || null,
              userId: data.userId || null,
              jobTitle: data.jobTitle || null,
              department: data.department || null,
              officeLocation: data.officeLocation || null,
              lastUpdated: new Date().toISOString(),
              sourceTimestamp: data.timestamp || null
          };

          chrome.storage.local.set({
              officeUserInfo: userInfo
          }, () => {
              console.log(`%c ✓ Office User Info Stored Successfully!`, "color: cyan; font-weight: bold");
              console.log(`   Name: ${userInfo.name}`);
              console.log(`   Email: ${userInfo.email}`);

              // Notify the extension that user info has been updated
              chrome.runtime.sendMessage({
                  type: "SRK_OFFICE_USER_INFO",
                  userInfo: userInfo,
                  timestamp: Date.now()
              }).catch(() => {
                  // Extension might not be ready, that's ok
              });
          });

      } catch (error) {
          console.error("%c Error processing Office User Info:", "color: red; font-weight: bold", error);

          // Notify extension of error
          chrome.runtime.sendMessage({
              type: "SRK_OFFICE_USER_INFO_ERROR",
              error: error.message,
              timestamp: Date.now()
          }).catch(() => {});
      }
  }

  // Listen for messages from extension (e.g., highlight student row requests)
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
      if (request.action === "postToPage" && request.message) {
          console.log("%c SRK Connector: Forwarding message to page", "color: blue; font-weight: bold", request.message);
          window.postMessage(request.message, "*");
          sendResponse({ success: true });
      }
      return true; // Keep channel open for async response
  });

  // Periodically announce presence to extension
  setInterval(() => {
      chrome.runtime.sendMessage({
          type: "SRK_CONNECTOR_HEARTBEAT",
          timestamp: Date.now()
      }).catch(() => {
          // Silently fail if extension is not available
      });
  }, 5000); // Every 5 seconds
}