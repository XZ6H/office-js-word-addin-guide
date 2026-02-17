# Getting Started with Office JS Word Add-ins

## Prerequisites

- **Node.js** 16.x or later (LTS recommended)
- **npm** 7.x or later
- **Visual Studio Code** with Office Add-in extension
- **Office 2016** or later / Microsoft 365 subscription
- **Git** for version control

## Environment Setup

### 1. Install Yeoman and Office Generator

```bash
npm install -g yo generator-office
```

### 2. Create a New Project

```bash
# Create project directory
mkdir my-word-addin
cd my-word-addin

# Generate Office Add-in
yo office --projectType task-pane --name "MyWordAddin" --host word --ts true
```

Available project types:
- `task-pane` — Side panel add-in (**recommended for most scenarios**)
- `commands` — Ribbon commands without UI
- `excel-functions` — Custom Excel functions (not applicable for Word)

### 3. Project Structure

```
my-word-addin/
├── manifest.xml           # Add-in configuration
├── package.json           # Dependencies and scripts
├── tsconfig.json          # TypeScript configuration
├── webpack.config.js      # Build configuration
├── src/
│   ├── taskpane/
│   │   ├── taskpane.ts    # Main task pane logic
│   │   ├── taskpane.html  # Task pane UI
│   │   └── taskpane.css   # Task pane styles
│   └── commands/
│       └── commands.ts    # Ribbon command handlers
└── assets/                # Icons and images
```

## Understanding the Manifest

The `manifest.xml` file configures your add-in:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <!-- Basic add-in info -->
  <Id>12345678-1234-1234-1234-123456789012</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Company</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="My Word Add-in"/>
  <Description DefaultValue="Description of your add-in"/>

  <!-- Icon URLs (must be HTTPS) -->
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-64.png"/>

  <!-- Support page -->
  <SupportUrl DefaultValue="https://www.example.com/help"/>

  <!-- App domains -->
  <AppDomains>
    <AppDomain>https://www.example.com</AppDomain>
  </AppDomains>

  <!-- Hosting configuration -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>

  <!-- Default settings -->
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>

  <!-- Permissions -->
  <Permissions>ReadWriteDocument</Permissions>

  <!-- Version overrides for modern features -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <!-- Resources section -->
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded succesfully. Go to the HOME tab and click the 'Show Taskpane' button to get started."/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Commands Group"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## Development Workflow

### 1. Install Dependencies

```bash
npm install
```

### 2. Start Development Server

```bash
npm run dev-server
```

This starts a local HTTPS server (required for Office add-ins).

### 3. Sideload the Add-in

#### Windows (Word Desktop):
1. Open Word
2. Go to **Insert** → **Add-ins** → **My Add-ins** → **Manage My Add-ins** → **Upload My Add-in**
3. Browse to `manifest.xml` in your project folder

#### macOS:
1. Open Word
2. Go to **Insert** → **Add-ins** → **My Add-ins** → **Manage** → **My Add-ins** → **Add from File**
3. Select `manifest.xml`

#### Web (Word Online):
1. Go to [Office.com](https://www.office.com) and open Word
2. **Insert** → **Office Add-ins** → **Manage My Add-ins** → **Upload My Add-in**
3. Upload `manifest.xml`

### 4. Enable Developer Mode

For easier sideloading during development:

**Windows Registry (optional for easier development):**
```
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer]
"EnableLiveReload"=dword:00000001
```

## Next Steps

- Learn [Architecture Best Practices](./02-architecture-best-practices.md)
- Master [Performance Optimization](./03-performance-optimization.md)
- Explore [API Deep Dive](./04-api-deep-dive.md)
- Try [Common Recipes](./05-recipes/)
