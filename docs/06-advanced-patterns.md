# Advanced Patterns

This guide covers advanced Office JS Word Add-in patterns including ribbon customization, Dialog API usage, and Single Sign-On (SSO) integration.

## Ribbon Customization

### Adding Custom Ribbon Buttons

Extend your manifest.xml to add ribbon commands:

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <OfficeTab id="TabHome">
    <Group id="CustomGroup">
      <Label resid="CustomGroup.Label"/>
      <Icon>
        <bt:Image size="16" resid="Icon.16x16"/>
        <bt:Image size="32" resid="Icon.32x32"/>
      </Icon>
      
      <!-- Action button -->
      <Control xsi:type="Button" id="ActionButton">
        <Label resid="ActionButton.Label"/>
        <Supertip>
          <Title resid="ActionButton.Title"/>
          <Description resid="ActionButton.Description"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
          <FunctionName>runAction</FunctionName>
        </Action>
      </Control>
      
      <!-- Menu button with dropdown -->
      <Control xsi:type="Menu" id="MenuButton">
        <Label resid="MenuButton.Label"/>
        <Supertip>
          <Title resid="MenuButton.Title"/>
          <Description resid="MenuButton.Description"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
        </Icon>
        <Items>
          <Item id="MenuItem1">
            <Label resid="MenuItem1.Label"/>
            <Supertip>
              <Title resid="MenuItem1.Title"/>
              <Description resid="MenuItem1.Description"/>
            </Supertip>
            <Icon>
              <bt:Image size="16" resid="Icon.16x16"/>
            </Icon>
            <Action xsi:type="ExecuteFunction">
              <FunctionName>runMenuItem1</FunctionName>
            </Action>
          </Item>
        </Items>
      </Control>
    </Group>
  </OfficeTab>
</ExtensionPoint>
```

### Command Handlers

```typescript
// commands.ts
Office.onReady(() => {
  // Register command handlers
  Office.actions.associate("runAction", runAction);
  Office.actions.associate("runMenuItem1", runMenuItem1);
});

async function runAction(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();
      
      console.log("Selected text:", range.text);
    });
  } catch (error) {
    console.error("Action failed:", error);
  } finally {
    event.completed();
  }
}

async function runMenuItem1(event: Office.AddinCommands.Event): Promise<void> {
  try {
    // Implementation here
    console.log("Menu item 1 executed");
  } finally {
    event.completed();
  }
}
```

## Dialog API

The Dialog API enables custom HTML dialogs for complex interactions.

### Opening a Dialog

```typescript
async function openCustomDialog(): Promise<void> {
  const dialogOptions: Office.DialogOptions = {
    width: 50,     // Percentage of screen
    height: 60,
    displayInIframe: true  // For desktop compatibility
  };
  
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/dialog.html",
    dialogOptions,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        
        // Listen for messages from dialog
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args) => {
            console.log("Message from dialog:", args.message);
            dialog.close();
          }
        );
        
        // Handle dialog closed
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (args) => {
            console.log("Dialog event:", args.error);
          }
        );
      } else {
        console.error("Failed to open dialog:", result.error);
      }
    }
  );
}
```

### Dialog HTML

```html
<!-- dialog.html -->
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Custom Dialog</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>
<body>
  <h2>Enter Information</h2>
  <input type="text" id="userInput" placeholder="Type something...">
  <button id="submitBtn">Submit</button>
  
  <script>
    Office.onReady(() => {
      document.getElementById("submitBtn").addEventListener("click", () => {
        const input = document.getElementById("userInput").value;
        Office.context.ui.messageParent(input);
      });
    });
  </script>
</body>
</html>
```

## Single Sign-On (SSO)

### SSO Overview

Office Add-ins support SSO with Microsoft Identity Platform, allowing seamless authentication without separate login prompts.

### SSO Configuration

**1. Register application in Azure AD:**
```
- App registration → New registration
- Supported account types: Accounts in any organizational directory
- Redirect URIs: https://localhost:3000/taskpane.html
- Add Office.js API permission: access_as_user
```

**2. Update manifest.xml:**
```xml
<WebApplicationInfo>
  <Id>YOUR_APP_CLIENT_ID</Id>
  <Resource>api://localhost:3000/YOUR_APP_CLIENT_ID</Resource>
  <Scopes>
    <Scope>openid</Scope>
    <Scope>profile</Scope>
    <Scope>User.Read</Scope>
  </Scopes>
</WebApplicationInfo>
```

**3. Implement SSO:**

```typescript
interface SSOConfig {
  fallbackToDialog?: boolean;
  consentRequired?: boolean;
}

async function getAccessTokenSSO(config: SSOConfig = {}): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessToken(
      {
        allowConsentPrompt: config.consentRequired ?? true,
        allowSignInPrompt: true
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          if (config.fallbackToDialog && 
              result.error.code === 13003) {
            // Consent required, fall back to dialog
            resolve(fallbackToDialog());
          } else {
            reject(new Error(`SSO failed: ${result.error.message}`));
          }
        }
      }
    );
  });
}

async function fallbackToDialog(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/login.html",
      { width: 30, height: 40 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (args) => {
              dialog.close();
              resolve(args.message);
            }
          );
        } else {
          reject(result.error);
        }
      }
    );
  });
}
```

## Error Handling Patterns

### Global Error Handler

```typescript
function setupGlobalErrorHandler(): void {
  window.onerror = (message, source, lineno, colno, error) => {
    console.error("Global error:", { message, source, lineno, error });
    
    // Log to telemetry
    logTelemetry("error", {
      message: String(message),
      stack: error?.stack
    });
    
    return true;
  };
  
  window.addEventListener("unhandledrejection", (event) => {
    console.error("Unhandled promise rejection:", event.reason);
    logTelemetry("unhandled_rejection", { reason: event.reason });
  });
}
```

### Context.sync() Error Handling

```typescript
async function safeSync(context: Word.RequestContext): Promise<void> {
  try {
    await context.sync();
  } catch (error) {
    if (OfficeExtension.ErrorCodes.generalException) {
      console.error("General Office JS error:", error);
    }
    throw error;
  }
}
```

## Performance Patterns

### Caching Document State

```typescript
class DocumentCache {
  private cache: Map<string, unknown> = new Map();
  private lastSync: number = 0;
  private readonly CACHE_TTL = 5000; // 5 seconds

  get<T>(key: string): T | undefined {
    if (Date.now() - this.lastSync > this.CACHE_TTL) {
      this.cache.clear();
      return undefined;
    }
    return this.cache.get(key) as T | undefined;
  }

  set<T>(key: string, value: T): void {
    this.cache.set(key, value);
    this.lastSync = Date.now();
  }

  invalidate(): void {
    this.cache.clear();
    this.lastSync = 0;
  }
}

const documentCache = new DocumentCache();
```

## Best Practices Summary

1. **Fallback gracefully** — Always have non-SSO fallback
2. **Handle dialog errors** — Network issues can break dialogs
3. **Limit dialog size** — Keep dialogs under 80% screen size
4. **Secure token storage** — Never store tokens in localStorage
5. **Monitor performance** — Track command execution time
