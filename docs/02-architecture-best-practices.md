# Architecture Best Practices for Word Add-ins

## Separation of Concerns

### Task Pane vs. Document Interaction

**Task Pane (UI Layer):**
- User interface and event handling
- State management
- API calls to backend services
- No direct document manipulation

**Document Service (Business Logic):**
- All Word API interactions
- Document manipulation logic
- Error handling for Office JS operations
- Performance optimizations

```typescript
// src/services/DocumentService.ts
export class DocumentService {
  async insertText(text: string, location: 'start' | 'end' | 'selection'): Promise<void> {
    return Word.run(async (context) => {
      let range: Word.Range;

      switch (location) {
        case 'start':
          range = context.document.body.getRange(Word.RangeLocation.start);
          break;
        case 'end':
          range = context.document.body.getRange(Word.RangeLocation.end);
          break;
        case 'selection':
          range = context.document.getSelection();
          break;
      }

      range.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  async formatSelection(fontName: string, fontSize: number): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.name = fontName;
      selection.font.size = fontSize;
      await context.sync();
    });
  }
}
```

```typescript
// src/taskpane/taskpane.ts
import { DocumentService } from '../services/DocumentService';

const documentService = new DocumentService();

// UI Event Handlers
async function handleInsertText() {
  const textInput = document.getElementById('text-input') as HTMLInputElement;
  const locationSelect = document.getElementById('location-select') as HTMLSelectElement;

  try {
    await documentService.insertText(textInput.value, locationSelect.value as any);
    showNotification('Text inserted successfully');
  } catch (error) {
    showError('Failed to insert text', error);
  }
}
```

## State Management

### Centralized State Store

```typescript
// src/store/AddinState.ts
interface AddinState {
  isLoading: boolean;
  currentDocument: {
    name: string;
    selection: string;
  } | null;
  userSettings: {
    defaultFont: string;
    defaultSize: number;
  };
}

class StateStore {
  private state: AddinState = {
    isLoading: false,
    currentDocument: null,
    userSettings: {
      defaultFont: 'Calibri',
      defaultSize: 11
    }
  };

  private listeners: Set<(state: AddinState) => void> = new Set();

  getState(): AddinState {
    return { ...this.state };
  }

  setState(partial: Partial<AddinState>): void {
    this.state = { ...this.state, ...partial };
    this.notifyListeners();
  }

  subscribe(listener: (state: AddinState) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private notifyListeners(): void {
    this.listeners.forEach(listener => listener(this.getState()));
  }
}

export const store = new StateStore();
```

## Error Handling Strategies

### Layered Error Handling

```typescript
// src/utils/ErrorHandler.ts
export class OfficeError extends Error {
  constructor(
    message: string,
    public code: string,
    public isRecoverable: boolean = false
  ) {
    super(message);
    this.name = 'OfficeError';
  }
}

export async function withErrorHandling<T>(
  operation: () => Promise<T>,
  errorMessage: string
): Promise<T | null> {
  try {
    return await operation();
  } catch (error) {
    console.error(`${errorMessage}:`, error);

    if (error instanceof OfficeError) {
      showUserNotification(error.message, error.isRecoverable ? 'warning' : 'error');
    } else {
      showUserNotification(errorMessage, 'error');
    }

    return null;
  }
}

// Usage in DocumentService
async function safeDocumentOperation() {
  return withErrorHandling(
    () => this.performComplexOperation(),
    'Failed to perform document operation'
  );
}
```

## Security Best Practices

### Input Validation

```typescript
// src/utils/Validation.ts
export function sanitizeInput(input: string): string {
  // Remove potential XSS vectors
  return input
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

export function validateRange(range: Word.Range): boolean {
  return range && !range.isEmpty;
}

export function validateContentControlId(id: string): boolean {
  return /^[a-zA-Z0-9_-]+$/.test(id);
}
```

### Secure API Calls

```typescript
// src/services/ApiService.ts
export class ApiService {
  private apiKey: string;

  constructor() {
    // Load from secure storage, never hardcode
    this.apiKey = this.loadApiKey();
  }

  async fetchData(endpoint: string): Promise<any> {
    const response = await fetch(`/api/${endpoint}`, {
      headers: {
        'Authorization': `Bearer ${this.apiKey}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`API call failed: ${response.status}`);
    }

    return response.json();
  }

  private loadApiKey(): string {
    // In production, use OfficeRuntime.Storage or secure cookie
    return OfficeRuntime.Storage.getItem('apiKey') || '';
  }
}
```

## Lifecycle Management

### Add-in Initialization

```typescript
// src/utils/Initializer.ts
export async function initializeAddin(): Promise<void> {
  try {
    // Check Office.js is ready
    await Office.onReady();

    // Initialize UI
    initializeUI();

    // Load user settings
    await loadUserSettings();

    // Register event handlers
    registerEventHandlers();

    console.log('Add-in initialized successfully');
  } catch (error) {
    console.error('Failed to initialize add-in:', error);
    showFatalError('Unable to initialize add-in. Please restart Word.');
  }
}

function initializeUI(): void {
  // Set up UI components
  const insertButton = document.getElementById('insert-button');
  if (insertButton) {
    insertButton.addEventListener('click', handleInsertClick);
  }
}

function registerEventHandlers(): void {
  // Listen for document changes
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    handleSelectionChange
  );
}
```

### Cleanup on Shutdown

```typescript
// src/utils/Cleanup.ts
export function registerCleanupHandlers(): void {
  window.addEventListener('beforeunload', () => {
    // Save any pending state
    savePendingChanges();

    // Remove event handlers
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged
    );

    // Clean up resources
    cleanupResources();
  });
}
```

## Module Organization

```
src/
├── services/           # Business logic, API calls, Office JS interactions
│   ├── DocumentService.ts
│   ├── TableService.ts
│   └── ApiService.ts
├── components/         # UI components
│   ├── Button.ts
│   ├── Input.ts
│   └── Notification.ts
├── store/             # State management
│   └── AddinState.ts
├── utils/             # Utilities
│   ├── ErrorHandler.ts
│   ├── Validation.ts
│   └── Formatting.ts
├── types/             # TypeScript type definitions
│   └── index.ts
└── taskpane/
    ├── taskpane.ts    # Entry point
    ├── taskpane.html
    └── taskpane.css
```

## Key Principles Summary

1. **Separation of Concerns**: UI and document logic are separate
2. **Centralized State**: Single source of truth for application state
3. **Layered Error Handling**: Catch errors at appropriate levels
4. **Input Validation**: Never trust user input
5. **Secure Storage**: Don't hardcode secrets
6. **Proper Cleanup**: Remove handlers and save state on exit
7. **Module Organization**: Clear folder structure for maintainability
