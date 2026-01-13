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
 * Gets a stable session ID for this user
 * First checks chrome.storage for a saved ID, then scans for Office's ack key, then generates new
 * @returns {Promise<string>} The session ID to use
 */
async function getStableSessionId() {
    return new Promise((resolve) => {
        // First, check if we already have a saved session ID in chrome.storage
        chrome.storage.local.get(['srkOfficeSessionId'], (result) => {
            if (result.srkOfficeSessionId) {
                console.log('[SRK Auto-Sideloader] Using stored session ID:', result.srkOfficeSessionId);
                resolve(result.srkOfficeSessionId);
                return;
            }

            // No stored ID, scan for Office's own ack key to find the REAL active session ID
            // Office creates keys like: ack3_WAC_Excel_3735224676_0
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                if (key && key.startsWith('ack3_WAC_Excel_')) {
                    // Extract session ID from key like "ack3_WAC_Excel_3735224676_0"
                    const parts = key.substring('ack3_WAC_Excel_'.length).split('_');
                    if (parts.length > 0 && parts[0]) {
                        const sessionId = parts[0];
                        console.log('[SRK Auto-Sideloader] Found Office session ID from ack key:', sessionId);
                        // Save it for future use
                        chrome.storage.local.set({ srkOfficeSessionId: sessionId });
                        resolve(sessionId);
                        return;
                    }
                }
            }

            // Fallback: try to find one in __OSF_UPLOADFILE.MyAddins keys
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                if (key && key.startsWith('__OSF_UPLOADFILE.MyAddins.16.')) {
                    const parts = key.split('.');
                    if (parts.length >= 4) {
                        const sessionId = parts[3];
                        console.log('[SRK Auto-Sideloader] Found existing session ID from MyAddins:', sessionId);
                        chrome.storage.local.set({ srkOfficeSessionId: sessionId });
                        resolve(sessionId);
                        return;
                    }
                }
            }

            // No existing session found, generate a new one
            // Excel uses what appears to be a random number
            const newSessionId = Math.floor(Math.random() * 9999999999).toString();
            console.log('[SRK Auto-Sideloader] Generated new session ID:', newSessionId);
            // Save it for future use
            chrome.storage.local.set({ srkOfficeSessionId: newSessionId });
            resolve(newSessionId);
        });
    });
}

// Main injection function
async function injectManifest() {
    console.log('%c [SRK Auto-Sideloader] Injecting manifest into localStorage...', 'background: #4CAF50; color: white; font-weight: bold; padding: 2px 4px;');

    try {
        // Get a stable session ID for this user (saved in chrome.storage)
        const SESSION_ID = await getStableSessionId();

        // Check if already sideloaded
        const manifestKey = `__OSF_UPLOADFILE.Manifest.16.${ADDIN_ID}`;
        const alreadyExists = localStorage.getItem(manifestKey) !== null;

        if (alreadyExists) {
            console.log('[SRK Auto-Sideloader] Manifest already exists, updating...');
        } else {
            console.log('[SRK Auto-Sideloader] First-time sideload detected');
        }

        console.log('%c [SRK Auto-Sideloader] 🔑 SESSION ID: ' + SESSION_ID, 'background: #9C27B0; color: white; font-weight: bold; padding: 4px 8px; font-size: 12px;');

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

        console.log('%c [SRK Auto-Sideloader] ✓ SUCCESS!', 'background: #4CAF50; color: white; font-weight: bold; padding: 2px 4px;');
        console.log('[SRK Auto-Sideloader] Manifest injected. Keys written:');
        console.log('  •', manifestKey);
        console.log('  •', myAddinsKey);
        console.log('  •', commandsKey);
        console.log('  •', ackKey, '= true (Office acknowledgment flag)');
        console.log('[SRK Auto-Sideloader] The Student Retention Add-in should load automatically!');
        console.log('[SRK Auto-Sideloader] If not visible, try refreshing Excel (F5)');

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
        console.error('%c [SRK Auto-Sideloader] FAILED:', 'background: #f44336; color: white; font-weight: bold; padding: 2px 4px;', error);
        return false;
    }
}

// Check if auto-sideload is enabled and run
chrome.storage.local.get(['autoSideloadManifest'], (result) => {
    const autoSideloadEnabled = result.autoSideloadManifest !== undefined
        ? result.autoSideloadManifest
        : true; // Default to enabled

    if (!autoSideloadEnabled) {
        console.log('[SRK] Auto-sideload is disabled in settings');
        return;
    }

    console.log('[SRK] Auto-sideload is enabled, proceeding...');

    // Simply inject the manifest - the ack key should make Office recognize it
    injectManifest();
});
