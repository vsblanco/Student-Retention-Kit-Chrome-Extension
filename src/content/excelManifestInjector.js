// excelManifestInjector.js
// Content script that auto-sideloads the Office Add-in manifest into Excel's localStorage

console.log('%c [SRK] Excel Manifest Injector Script LOADED', 'background: #FF9800; color: white; font-size: 14px; font-weight: bold; padding: 2px 4px;');

// Embedded manifest XML - avoids fetch issues in SharePoint iframes
const MANIFEST_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>a8b1e479-1b3d-4e9e-9a1c-2f8e1c8b4a0e</Id>
  <Version>2.1.0.1</Version>
  <ProviderName>Victor Blanco</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Student Retention Add-in"/>
  <Description DefaultValue="An add-in for tracking student retention tasks."/>

  <IconUrl DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/icon-32.png"/>

  <SupportUrl DefaultValue="https://github.com/vsblanco/Student-Retention-Add-in"/>
  <AppDomains>
    <AppDomain>https://vsblanco.github.io</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/react/dist/index.html"/>
    <RequestedWidth>450</RequestedWidth>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <!-- Requirements for autoload functionality -->
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>

          <FunctionFile resid="Commands.Url"/>

          <!-- Runtimes for shared runtime and autoload -->
          <Runtimes>
            <Runtime resid="Commands.Url" lifetime="long" />
          </Runtimes>

          <!-- LaunchEvent: Load add-in automatically when document opens -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewDocument" FunctionName="onDocumentOpen" />
              <LaunchEvent Type="OnDocumentOpened" FunctionName="onDocumentOpen" />
            </LaunchEvents>
            <SourceLocation resid="Commands.Url" />
          </ExtensionPoint>

          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Victor.RetentionTab">
              <Label resid="RetentionTab.Label"/>
              <Group id="SystemGroup">
                <Label resid="SystemGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="SettingsButton">
                  <Label resid="SettingsButton.Label"/>
                  <Supertip>
                    <Title resid="SettingsButton.Label"/>
                    <Description resid="SettingsButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="SettingsIcon.16x16"/>
                    <bt:Image size="32" resid="SettingsIcon.32x32"/>
                    <bt:Image size="80" resid="SettingsIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>SettingsPane</TaskpaneId>
                    <SourceLocation resid="Settings.Url"/>
                  </Action>
                </Control>
              </Group>
              <Group id="ReportGroup">
                <Label resid="ReportGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="CreateLdaButton">
                  <Label resid="CreateLdaButton.Label"/>
                  <Supertip>
                    <Title resid="CreateLdaButton.Label"/>
                    <Description resid="CreateLdaButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="CreateLdaIcon.16x16"/>
                    <bt:Image size="32" resid="CreateLdaIcon.32x32"/>
                    <bt:Image size="80" resid="CreateLdaIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>CreateLdaPane</TaskpaneId>
                    <SourceLocation resid="CreateLdaDialog.Url"/>
                  </Action>
                </Control>
                 <Control xsi:type="Button" id="PersonalizedEmailButton">
                  <Label resid="PersonalizedEmailButton.Label"/>
                  <Supertip>
                    <Title resid="PersonalizedEmailButton.Label"/>
                    <Description resid="PersonalizedEmailButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="EmailIcon.16x16"/>
                    <bt:Image size="32" resid="EmailIcon.32x32"/>
                    <bt:Image size="80" resid="EmailIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>PersonalizedEmailPane</TaskpaneId>
                    <SourceLocation resid="PersonalizedEmail.Url"/>
                  </Action>
                </Control>
              </Group>
              <Group id="StudentViewGroup">
                <Label resid="StudentViewGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="DetailsIcon.16x16"/>
                  <bt:Image size="32" resid="DetailsIcon.32x32"/>
                  <bt:Image size="80" resid="DetailsIcon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="StudentViewButton">
                  <Label resid="StudentViewButton.Label"/>
                  <Supertip>
                    <Title resid="StudentViewButton.Label"/>
                    <Description resid="StudentViewButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="DetailsIcon.16x16"/>
                    <bt:Image size="32" resid="DetailsIcon.32x32"/>
                    <bt:Image size="80" resid="DetailsIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>StudentViewPane</TaskpaneId>
                    <SourceLocation resid="StudentView.Url"/>
                  </Action>
                </Control>
              </Group>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/icon-80.png"/>
        <bt:Image id="CreateLdaIcon.16x16" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/create-lda-icon.png"/>
        <bt:Image id="CreateLdaIcon.32x32" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/create-lda-icon.png"/>
        <bt:Image id="CreateLdaIcon.80x80" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/create-lda-icon.png"/>
        <bt:Image id="DetailsIcon.16x16" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/details-icon.png"/>
        <bt:Image id="DetailsIcon.32x32" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/details-icon.png"/>
        <bt:Image id="DetailsIcon.80x80" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/details-icon.png"/>
        <bt:Image id="SettingsIcon.16x16" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/settings-icon.png"/>
        <bt:Image id="SettingsIcon.32x32" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/settings-icon.png"/>
        <bt:Image id="SettingsIcon.80x80" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/settings-icon.png"/>
        <bt:Image id="EmailIcon.16x16" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/email-icon.png"/>
        <bt:Image id="EmailIcon.32x32" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/email-icon.png"/>
        <bt:Image id="EmailIcon.80x80" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/assets/email-icon.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/background-services/commands.html"/>
        <bt:Url id="StudentView.Url" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/react/dist/index.html?page=student-view"/>
        <bt:Url id="Settings.Url" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/react/dist/index.html?page=settings"/>
        <bt:Url id="CreateLdaDialog.Url" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/react/dist/index.html?page=create-lda"/>
        <bt:Url id="WelcomeDialog.Url" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/welcome-dialog.html"/>
        <bt:Url id="PersonalizedEmail.Url" DefaultValue="https://vsblanco.github.io/Student-Retention-Add-in/react/dist/index.html?page=personalized-email"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with the Retention Add-in!"/>
        <bt:String id="RetentionTab.Label" DefaultValue="Retention"/>
        <bt:String id="SystemGroup.Label" DefaultValue="System"/>
        <bt:String id="ReportGroup.Label" DefaultValue="Report"/>
        <bt:String id="StudentViewGroup.Label" DefaultValue="Student View"/>
        <bt:String id="StudentViewButton.Label" DefaultValue="Student View"/>
        <bt:String id="CreateLdaButton.Label" DefaultValue="Create LDA"/>
        <bt:String id="SettingsButton.Label" DefaultValue="Settings"/>
        <bt:String id="PersonalizedEmailButton.Label" DefaultValue="Send Emails"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your add-in loaded successfully. Go to the 'Retention' tab to get started."/>
        <bt:String id="StudentViewButton.Tooltip" DefaultValue="Click to show the student details pane."/>
        <bt:String id="CreateLdaButton.Tooltip" DefaultValue="Creates a new LDA sheet for the current date."/>
        <bt:String id="SettingsButton.Tooltip" DefaultValue="Click to configure add-in settings."/>
        <bt:String id="PersonalizedEmailButton.Tooltip" DefaultValue="Opens a task pane to send a personalized email to the selected student."/>
      </bt:LongStrings>
    </Resources>

    <!--
      WebApplicationInfo: Required for Microsoft SSO (Single Sign-On)
      IMPORTANT: Must be the last child element of VersionOverrides

      Configuration:
      - Id: Azure AD Application (Client) ID
      - Resource: Application ID URI (must match Azure AD configuration)
      - Scopes: Required OAuth scopes (openid and profile are mandatory for SSO)
    -->
    <WebApplicationInfo>
      <Id>71f37f39-a330-413a-be61-0baa5ce03ea3</Id>
      <Resource>api://vsblanco.github.io/71f37f39-a330-413a-be61-0baa5ce03ea3</Resource>
      <Scopes>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
        <Scope>User.Read</Scope>
        <Scope>User.ReadBasic.All</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
</OfficeApp>`;

const ADDIN_ID = 'a8b1e479-1b3d-4e9e-9a1c-2f8e1c8b4a0e';

/**
 * Stores the session ID in chrome.storage for future use
 * @param {string} sessionId - The session ID to store
 */
async function storeSessionId(sessionId) {
    return new Promise((resolve) => {
        chrome.storage.local.set({ 'srkSessionId': sessionId }, () => {
            console.log('%c💾 SESSION ID STORED', 'background: #2196F3; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;', sessionId);
            resolve();
        });
    });
}

/**
 * Retrieves the stored session ID from chrome.storage
 * Checks nested path first, then falls back to legacy flat key
 * @returns {Promise<string|null>} The stored session ID or null if not found
 */
async function getStoredSessionId() {
    return new Promise((resolve) => {
        chrome.storage.local.get(['state', 'srkSessionId'], (result) => {
            // Check nested path first (state.srkSessionId)
            const sessionId = result.state?.srkSessionId || result.srkSessionId || null;
            resolve(sessionId);
        });
    });
}

/**
 * Finds existing Office session IDs in localStorage or uses stored session ID
 * This ensures the add-in works for all users, not just those with a specific session ID
 * @returns {Promise<string|null>} The session ID to use, or null if none found
 */
async function findOrGenerateSessionId() {
    console.log('%c🔍 SEARCHING FOR SESSION ID...', 'background: #9C27B0; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px; font-size: 12px;');

    // Strategy 1: Look for existing __OSF_UPLOADFILE keys to find the session ID (most reliable)
    console.log('%c📂 Strategy 1: Checking __OSF_UPLOADFILE keys...', 'color: #9C27B0; font-weight: bold;');
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith('__OSF_UPLOADFILE.MyAddins.16.')) {
            // Extract session ID from key like "__OSF_UPLOADFILE.MyAddins.16.3735224676"
            const parts = key.split('.');
            if (parts.length >= 4) {
                const sessionId = parts[3];
                console.log('%c✅ FOUND SESSION ID (Strategy 1)', 'background: #4CAF50; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;', sessionId);
                await storeSessionId(sessionId); // Store for future use
                return sessionId;
            }
        }
    }

    // Strategy 2: Look for WAC (Web Application Companion) keys with various patterns
    console.log('%c📂 Strategy 2: Checking WAC keys...', 'color: #9C27B0; font-weight: bold;');
    // These keys are created by Office and contain the session ID
    // Patterns include:
    //   - ack3_WAC_Excel_{SESSION_ID}_0/10/18
    //   - ak0_WAC_Excel_{SESSION_ID}_0
    //   - ak4_WAC_Excel_{SESSION_ID}_0/10/18
    //   - Flyout_WAC_Excel_{SESSION_ID}__0_ExpirationTime/StoreDisabled
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.includes('_WAC_Excel_')) {
            // Match session ID with flexible pattern to handle both single (_) and double (__) underscores
            // Examples: "ack3_WAC_Excel_3735224676_8" or "Flyout_WAC_Excel_3735224676__0_ExpirationTime"
            const match = key.match(/_WAC_Excel_(\d+)(_|__)/);
            if (match && match[1]) {
                const sessionId = match[1];
                console.log('%c✅ FOUND SESSION ID (Strategy 2)', 'background: #4CAF50; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;', sessionId);
                console.log('%c   From key:', 'color: #666;', key);
                await storeSessionId(sessionId); // Store for future use
                return sessionId;
            }
        }
    }

    // Strategy 3: Try to use previously stored session ID
    console.log('%c📂 Strategy 3: Checking stored session ID...', 'color: #9C27B0; font-weight: bold;');
    const storedSessionId = await getStoredSessionId();
    if (storedSessionId) {
        console.log('%c✅ USING STORED SESSION ID (Strategy 3)', 'background: #FF9800; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;', storedSessionId);
        console.log('%c   ⚠️  Using last known session ID - may need update if Office session changed', 'color: #FF9800; font-style: italic;');
        return storedSessionId;
    }

    // No session ID found anywhere - do not sideload
    console.log('%c❌ NO SESSION ID FOUND - SKIPPING SIDELOAD', 'background: #f44336; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px; font-size: 12px;');
    console.log('%c   ℹ️  No Office session ID detected in localStorage or storage', 'color: #f44336; font-weight: bold;');
    console.log('%c   ℹ️  The add-in will not be sideloaded automatically', 'color: #f44336;');
    console.log('%c   ℹ️  Please ensure you have opened Excel Online at least once', 'color: #f44336;');
    return null;
}

// Main injection function
function injectManifest(SESSION_ID) {
    console.log('%c🚀 INJECTING MANIFEST...', 'background: #4CAF50; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px; font-size: 12px;');

    try {
        // Check if already sideloaded
        const manifestKey = `__OSF_UPLOADFILE.Manifest.16.${ADDIN_ID}`;
        const alreadyExists = localStorage.getItem(manifestKey) !== null;

        if (alreadyExists) {
            console.log('%c   📝 Manifest already exists, updating...', 'color: #2196F3; font-weight: bold;');
        } else {
            console.log('%c   🆕 First-time sideload detected', 'color: #4CAF50; font-weight: bold;');
        }

        console.log('%c   🔑 Using session ID:', 'color: #9C27B0; font-weight: bold;', SESSION_ID);

        // Write manifest
        const manifestValue = JSON.stringify({
            data: MANIFEST_XML,
            createdOn: Date.now(),
            refreshRate: 3
        });
        localStorage.setItem(manifestKey, manifestValue);

        // Update MyAddins
        const myAddinsKey = `__OSF_UPLOADFILE.MyAddins.16.${SESSION_ID}`;
        const existingMyAddins = localStorage.getItem(myAddinsKey);
        let myAddinsValue;

        if (existingMyAddins) {
            try {
                const parsed = JSON.parse(existingMyAddins);
                const addinIds = parsed.data || [];
                if (!addinIds.includes(ADDIN_ID)) {
                    addinIds.push(ADDIN_ID);
                }
                myAddinsValue = JSON.stringify({
                    data: addinIds,
                    createdOn: parsed.createdOn || Date.now(),
                    refreshRate: 3
                });
            } catch (e) {
                myAddinsValue = JSON.stringify({
                    data: [ADDIN_ID],
                    createdOn: Date.now(),
                    refreshRate: 3
                });
            }
        } else {
            myAddinsValue = JSON.stringify({
                data: [ADDIN_ID],
                createdOn: Date.now(),
                refreshRate: 3
            });
        }
        localStorage.setItem(myAddinsKey, myAddinsValue);

        // Update AddinCommandsMyAddins
        const commandsKey = `__OSF_UPLOADFILE.AddinCommandsMyAddins.16.${SESSION_ID}`;
        const existingCommands = localStorage.getItem(commandsKey);
        let commandsValue;

        if (existingCommands) {
            try {
                const parsed = JSON.parse(existingCommands);
                const addinIds = parsed.data || [];
                if (!addinIds.includes(ADDIN_ID)) {
                    addinIds.push(ADDIN_ID);
                }
                commandsValue = JSON.stringify({
                    data: addinIds,
                    createdOn: parsed.createdOn || Date.now(),
                    refreshRate: 3
                });
            } catch (e) {
                commandsValue = JSON.stringify({
                    data: [ADDIN_ID],
                    createdOn: Date.now(),
                    refreshRate: 3
                });
            }
        } else {
            commandsValue = JSON.stringify({
                data: [ADDIN_ID],
                createdOn: Date.now(),
                refreshRate: 3
            });
        }
        localStorage.setItem(commandsKey, commandsValue);

        // CRITICAL: Set the acknowledgment flag that tells Office the add-ins list has been loaded
        const ackKey = `ack3_WAC_Excel_${SESSION_ID}_8`;
        localStorage.setItem(ackKey, 'true');

        console.log('%c✅ INJECTION SUCCESS!', 'background: #4CAF50; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px; font-size: 12px;');
        console.log('%c   📦 Keys written to localStorage:', 'color: #4CAF50; font-weight: bold;');
        console.log('%c      •', 'color: #4CAF50;', manifestKey);
        console.log('%c      •', 'color: #4CAF50;', myAddinsKey);
        console.log('%c      •', 'color: #4CAF50;', commandsKey);
        console.log('%c      •', 'color: #4CAF50;', ackKey, '= true (Office acknowledgment flag)');
        console.log('%c   🎉 The Student Retention Add-in should load automatically!', 'color: #4CAF50; font-weight: bold; font-size: 11px;');
        console.log('%c   💡 If not visible, try refreshing Excel (F5)', 'color: #666; font-style: italic;');

        // Notify the background script
        chrome.runtime.sendMessage({
            type: 'SRK_MANIFEST_INJECTED',
            addinId: ADDIN_ID,
            timestamp: Date.now()
        }).catch(() => {
            // Extension might not be ready, that's ok
        });

        return true;

    } catch (error) {
        console.error('%c❌ INJECTION FAILED!', 'background: #f44336; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px; font-size: 12px;');
        console.error('%c   Error details:', 'color: #f44336; font-weight: bold;', error);
        return false;
    }
}

// Helper to get value from nested storage structure
function getSettingValue(data, key, defaultValue) {
    // Check nested path first
    if (key === 'autoSideloadManifest') {
        if (data.settings?.excelAddIn?.autoSideloadManifest !== undefined) {
            return data.settings.excelAddIn.autoSideloadManifest;
        }
    } else if (key === 'srkSessionId') {
        if (data.state?.srkSessionId !== undefined) {
            return data.state.srkSessionId;
        }
    }
    // Fall back to legacy flat key
    return data[key] !== undefined ? data[key] : defaultValue;
}

// Check if auto-sideload is enabled and run
chrome.storage.local.get(['settings', 'autoSideloadManifest'], async (result) => {
    const autoSideloadEnabled = getSettingValue(result, 'autoSideloadManifest', true);

    if (!autoSideloadEnabled) {
        console.log('%c⚙️  AUTO-SIDELOAD DISABLED', 'background: #757575; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;');
        console.log('%c   Auto-sideload is disabled in extension settings', 'color: #757575;');
        return;
    }

    console.log('%c⚙️  AUTO-SIDELOAD ENABLED', 'background: #2196F3; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;');
    console.log('%c   Proceeding with automatic sideload...', 'color: #2196F3;');

    // Find or retrieve the session ID
    const SESSION_ID = await findOrGenerateSessionId();

    // Check if we have a valid session ID before injecting
    if (!SESSION_ID) {
        console.log('%c⛔ SIDELOAD ABORTED', 'background: #f44336; color: white; font-weight: bold; padding: 4px 8px; border-radius: 3px;');
        console.log('%c   Cannot proceed - no session ID found', 'color: #f44336; font-weight: bold;');
        return;
    }

    // Simply inject the manifest - the ack key should make Office recognize it
    injectManifest(SESSION_ID);
});
