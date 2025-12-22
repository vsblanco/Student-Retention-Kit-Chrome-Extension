// content/excel_connector.js

// Prevent multiple injections
if (window.hasSRKConnectorRun) {
  // Script already loaded, stop here.
  console.log("SRK Connector already active.");
} else {
  window.hasSRKConnectorRun = true;

  console.log("%c SRK: Excel Connector Script LOADED", "background: #222; color: #bada55; font-size: 14px");

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

          // Always accept data for now (as per requirements)
          if (event.source) {
              event.source.postMessage({
                  type: "SRK_MASTER_LIST_RESPONSE",
                  wantsData: true
              }, "*");
          }
      }

      // Handle Master List Data
      else if (event.data.type === "SRK_MASTER_LIST_DATA") {
          console.log("%c SRK Connector: Master List Data Received!", "color: green; font-weight: bold");
          handleMasterListData(event.data.data);
      }
  });

  /**
   * Handles incoming Master List data from the Office Add-in
   * Transforms the data from add-in format to extension format and stores it
   */
  function handleMasterListData(data) {
      try {
          console.log(`Processing Master List with ${data.totalStudents} students`);
          console.log("Data timestamp:", data.timestamp);

          // Transform students from add-in format to extension format
          const transformedStudents = data.students.map(student => ({
              name: student.studentName || 'Unknown',
              phone: student.primaryPhone || student.otherPhone || null,
              grade: student.grade !== undefined && student.grade !== null ? String(student.grade) : null,
              StudentNumber: student.studentNumber || null,
              SyStudentId: student.studentId || null,
              daysout: parseInt(student.daysOut) || 0,
              missingCount: 0,
              url: student.gradeBook || null,
              assignments: [],
              // Additional fields that might be useful
              lastLda: student.lastLda || null,
              studentEmail: student.studentEmail || null,
              personalEmail: student.personalEmail || null,
              assigned: student.assigned || null,
              outreach: student.outreach || null
          }));

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