# Data Binding Patterns

Two-way data binding between Word content controls and external data sources including Custom XML Parts, Excel, SharePoint, and APIs.

## Overview

Office JS supports data binding for:
- **Content controls** — Bind to Custom XML Parts
- **Document sections** — Bind to named ranges
- **Tables** — Bind row/column data
- **External sources** — Excel, SharePoint, REST APIs

## References

- [Office.Bindings API](https://learn.microsoft.com/en-us/javascript/api/office/office.bindings?view=word-js-preview)
- [Custom XML Parts](https://learn.microsoft.com/en-us/archive/msdn-magazine/2013/april/microsoft-office-exploring-the-javascript-api-for-office-data-binding-and-custom-xml-parts)
- [Content Control Binding Sample](https://github.com/OfficeDev/Word-Add-in-Content-Control-Binding)

## Binding to Content Controls

### Creating Bindings

```typescript
/**
 * Creates a binding between a content control and a named binding
 * Based on Microsoft OfficeDev Content Control Binding sample
 */
export async function createContentControlBinding(
  contentControlTag: string,
  bindingId: string
): Promise<Office.Binding> {
  return new Promise((resolve, reject) => {
    Office.context.document.bindings.addFromNamedItemAsync(
      contentControlTag,
      Office.BindingType.Text,
      { id: bindingId },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(`Binding failed: ${result.error.message}`));
        }
      }
    );
  });
}

/**
 * Retrieves data from a binding
 */
export async function getBindingData(bindingId: string): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.bindings.getByIdAsync(
      bindingId,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          result.value.getDataAsync(
            { coercionType: Office.CoercionType.Text },
            (dataResult) => {
              if (dataResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve(dataResult.value as string);
              } else {
                reject(new Error(`Get data failed: ${dataResult.error.message}`));
              }
            }
          );
        } else {
          reject(new Error(`Get binding failed: ${result.error.message}`));
        }
      }
    );
  });
}

/**
 * Sets data to a binding (updates content control)
 */
export async function setBindingData(bindingId: string, data: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.bindings.getByIdAsync(
      bindingId,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          result.value.setDataAsync(
            data,
            { coercionType: Office.CoercionType.Text },
            (setResult) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
              } else {
                reject(new Error(`Set data failed: ${setResult.error.message}`));
              }
            }
          );
        } else {
          reject(new Error(`Get binding failed: ${result.error.message}`));
        }
      }
    );
  });
}
```

### Binding Change Events

```typescript
/**
 * Sets up automatic synchronization when binding data changes
 */
export function setupBindingChangeHandler(
  bindingId: string,
  onChange: (newData: string) => void
): void {
  Office.context.document.bindings.getByIdAsync(
    bindingId,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        result.value.addHandlerAsync(
          Office.EventType.BindingDataChanged,
          () => {
            getBindingData(bindingId)
              .then(onChange)
              .catch(console.error);
          }
        );
      }
    }
  );
}
```

## Custom XML Parts

### Creating Custom XML Parts

```typescript
/**
 * Custom XML Part for structured data binding
 * Based on Microsoft documentation patterns
 */
export interface CustomXmlPart {
  namespace: string;
  rootElement: string;
  data: { [key: string]: string };
}

export async function addCustomXmlPart(xmlContent: string): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.customXmlParts.addAsync(
      xmlContent,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value.id);
        } else {
          reject(new Error(`Failed to add XML part: ${result.error.message}`));
        }
      }
    );
  });
}

/**
 * Generates XML for a Custom XML Part
 */
export function generateCustomXml(
  namespace: string,
  rootElement: string,
  data: { [key: string]: string }
): string {
  const entries = Object.entries(data)
    .map(([key, value]) => `  <${key}>${escapeXml(value)}</${key}>`)
    .join('\n');
  
  return `<?xml version="1.0" encoding="UTF-8"?>
<${rootElement} xmlns="${namespace}">
${entries}
</${rootElement}>`;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
```

### Binding Content Controls to XML

```typescript
/**
 * Binds content controls to Custom XML Part nodes
 * The content control's tag should match the XML element name
 */
export async function bindControlsToXml(xmlNamespace: string): Promise<void> {
  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag, title");
    }
    await context.sync();
    
    // Note: Direct XML binding via Office JS is limited
    // Typically you would:
    // 1. Load XML data
    // 2. Map to content controls by tag
    // 3. Update control values
    
    for (const control of contentControls.items) {
      if (control.tag) {
        // This is where you would bind to XML node
        // For now, we just mark the control
        control.title = `Bound to: ${control.tag}`;
      }
    }
    
    await context.sync();
  });
}
```

## External Data Source Integration

### Excel Integration

```typescript
export interface ExcelDataRange {
  sheetName: string;
  startCell: string;
  endCell: string;
}

/**
 * Fetches data from Excel file
 * Requires user to select Excel file
 */
export async function importFromExcel(
  file: File
): Promise<{ headers: string[]; rows: string[][] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        // In real implementation, use xlsx library
        // const data = new Uint8Array(e.target?.result as ArrayBuffer);
        // const workbook = XLSX.read(data, { type: 'array' });
        
        // Mock for demonstration
        resolve({
          headers: ['Name', 'Value'],
          rows: [['Sample', 'Data']]
        });
      } catch (error) {
        reject(error);
      }
    };
    
    reader.onerror = () => reject(new Error('Failed to read Excel file'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Binds Excel data to document content controls
 */
export async function bindExcelDataToDocument(
  excelData: { headers: string[]; rows: string[][] }
): Promise<void> {
  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    // Create a map from headers to values (first row)
    const dataMap: { [key: string]: string } = {};
    for (let i = 0; i < excelData.headers.length; i++) {
      const key = excelData.headers[i].toLowerCase().replace(/\s+/g, '-');
      dataMap[key] = excelData.rows[0]?.[i] || '';
    }
    
    for (const control of contentControls.items) {
      control.load("tag");
    }
    await context.sync();
    
    // Bind data to controls by matching tags
    for (const control of contentControls.items) {
      if (control.tag && dataMap[control.tag] !== undefined) {
        control.insertText(dataMap[control.tag], Word.InsertLocation.replace);
      }
    }
    
    await context.sync();
  });
}
```

### REST API Integration

```typescript
export interface ApiBindingConfig {
  endpoint: string;
  method?: 'GET' | 'POST';
  headers?: { [key: string]: string };
  mapping: { [controlTag: string]: string }; // controlTag -> API response field
}

/**
 * Fetches data from API and binds to document
 */
export async function bindFromApi(config: ApiBindingConfig): Promise<void> {
  try {
    const response = await fetch(config.endpoint, {
      method: config.method || 'GET',
      headers: {
        'Content-Type': 'application/json',
        ...config.headers
      }
    });
    
    if (!response.ok) {
      throw new Error(`API error: ${response.status}`);
    }
    
    const data = await response.json();
    
    await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();
      
      for (const control of contentControls.items) {
        control.load("tag");
      }
      await context.sync();
      
      // Map API data to content controls
      for (const [controlTag, fieldPath] of Object.entries(config.mapping)) {
        const control = contentControls.items.find(c => c.tag === controlTag);
        
        if (control) {
          const value = getNestedValue(data, fieldPath);
          if (value !== undefined) {
            control.insertText(String(value), Word.InsertLocation.replace);
          }
        }
      }
      
      await context.sync();
    });
  } catch (error) {
    console.error('API binding failed:', error);
    throw error;
  }
}

function getNestedValue(obj: any, path: string): any {
  return path.split('.').reduce((current, key) => current?.[key], obj);
}

// Example usage:
// await bindFromApi({
//   endpoint: 'https://api.example.com/customer/123',
//   mapping: {
//     'customer-name': 'name',
//     'customer-email': 'contact.email',
//     'contract-value': 'contract.amount'
//   }
// });
```

### SharePoint Integration

```typescript
export interface SharePointListBinding {
  siteUrl: string;
  listName: string;
  itemId: number;
  fields: string[];
}

/**
 * Binds SharePoint list item to document
 * Requires SharePoint REST API access
 */
export async function bindFromSharePoint(
  config: SharePointListBinding,
  accessToken: string
): Promise<void> {
  const url = `${config.siteUrl}/_api/web/lists/getbytitle('${config.listName}')/items(${config.itemId})`;
  
  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json;odata=verbose'
    }
  });
  
  if (!response.ok) {
    throw new Error(`SharePoint error: ${response.status}`);
  }
  
  const data = await response.json();
  const item = data.d;
  
  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag");
    }
    await context.sync();
    
    // Map SharePoint fields to content controls
    for (const field of config.fields) {
      const control = contentControls.items.find(c => c.tag === field);
      if (control && item[field] !== undefined) {
        control.insertText(String(item[field]), Word.InsertLocation.replace);
      }
    }
    
    await context.sync();
  });
}
```

## Two-Way Synchronization

```typescript
export interface SyncConfig {
  source: 'document' | 'api' | 'sharepoint';
  direction: 'toDocument' | 'fromDocument' | 'bidirectional';
  fieldMapping: { [documentField: string]: string };
}

/**
 * Synchronizes data between document and external source
 */
export async function synchronizeData(config: SyncConfig): Promise<void> {
  if (config.direction === 'toDocument' || config.direction === 'bidirectional') {
    // Pull from external source to document
    console.log('Syncing to document...');
    // Implementation depends on source type
  }
  
  if (config.direction === 'fromDocument' || config.direction === 'bidirectional') {
    // Push from document to external source
    console.log('Syncing from document...');
    
    const documentData = await extractDocumentData(
      Object.keys(config.fieldMapping)
    );
    
    // Send to external source
    console.log('Data to sync:', documentData);
  }
}

async function extractDocumentData(tags: string[]): Promise<{ [tag: string]: string }> {
  return await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag, text");
    }
    await context.sync();
    
    const data: { [tag: string]: string } = {};
    
    for (const tag of tags) {
      const control = contentControls.items.find(c => c.tag === tag);
      if (control) {
        data[tag] = control.text;
      }
    }
    
    return data;
  });
}
```

## Best Practices

1. **Handle binding errors** — Bindings may fail if named items don't exist
2. **Use content control tags** — Consistent naming for mapping
3. **Batch API calls** — Minimize network requests
4. **Validate data types** — Ensure API data matches expected format
5. **Cache when appropriate** — Don't re-fetch static data

## Limitations

- Direct XML data binding is limited in Office JS
- Two-way sync requires manual implementation
- Binding IDs must be unique per document
- Some operations require specific Word versions
