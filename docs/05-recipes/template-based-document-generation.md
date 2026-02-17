# Template-Based Document Generation

Enterprise-grade document generation using Word templates with content controls, data binding, and OOXML for complex formatting.

## Overview

Template-based document generation involves:
1. Creating Word templates with content control placeholders
2. Binding data to those placeholders programmatically
3. Handling repeating sections for lists and tables
4. Using OOXML for complex formatting when needed

## References

- [Word.Template API (Preview)](https://learn.microsoft.com/en-us/javascript/api/word/word.template?view=word-js-preview)
- [Office Open XML in Word Add-ins](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/word/create-better-add-ins-for-word-with-office-open-xml.md)
- [Content Control Binding Sample](https://github.com/OfficeDev/Word-Add-in-Content-Control-Binding)

## Creating a Document Template

### Template Structure with Content Controls

```typescript
/**
 * Creates a template structure with content controls for document generation
 * Based on Microsoft OfficeDev patterns
 */
export interface TemplateField {
  tag: string;
  title: string;
  type: Word.ContentControlType;
  placeholder?: string;
}

export async function createDocumentTemplate(fields: TemplateField[]): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    
    // Clear document for template creation
    body.clear();
    await context.sync();
    
    // Add header
    const header = body.insertParagraph("Document Template", Word.InsertLocation.start);
    header.style = "Heading 1";
    header.font.bold = true;
    header.font.size = 16;
    
    body.insertParagraph("", Word.InsertLocation.end);
    
    // Create content controls for each field
    for (const field of fields) {
      // Label paragraph
      const label = body.insertParagraph(`${field.title}:`, Word.InsertLocation.end);
      label.font.bold = true;
      
      // Content control
      const range = body.insertParagraph("", Word.InsertLocation.end).getRange();
      const control = range.insertContentControl();
      control.tag = field.tag;
      control.title = field.title;
      control.type = field.type;
      
      if (field.placeholder) {
        control.placeholderText = field.placeholder;
      }
      
      body.insertParagraph("", Word.InsertLocation.end);
    }
    
    await context.sync();
  });
}

// Example usage:
// await createDocumentTemplate([
//   { tag: "company-name", title: "Company Name", type: Word.ContentControlType.richText },
//   { tag: "contract-date", title: "Date", type: Word.ContentControlType.datePicker },
//   { tag: "priority", title: "Priority", type: Word.ContentControlType.dropDownList }
// ]);
```

## Populating Template with Data

### Single Record Population

```typescript
export interface TemplateData {
  [fieldTag: string]: string | string[];
}

/**
 * Populates a template document with data
 * Based on OfficeDev Content Control Binding patterns
 */
export async function populateTemplate(data: TemplateData): Promise<string[]> {
  return await Word.run(async (context) => {
    const filled: string[] = [];
    const skipped: string[] = [];
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    // Load all tags first (batch operation)
    for (const control of contentControls.items) {
      control.load("tag, type");
    }
    await context.sync();
    
    // Populate each control
    for (const control of contentControls.items) {
      const value = data[control.tag];
      
      if (value !== undefined) {
        if (Array.isArray(value)) {
          // Handle arrays (comma-separated or special handling)
          control.insertText(value.join(", "), Word.InsertLocation.replace);
        } else {
          control.insertText(value, Word.InsertLocation.replace);
        }
        filled.push(control.tag);
      } else {
        skipped.push(control.tag);
      }
    }
    
    await context.sync();
    
    if (skipped.length > 0) {
      console.warn("Fields without data:", skipped);
    }
    
    return filled;
  });
}
```

## Repeating Sections (Tables as Lists)

Since Office JS doesn't natively support repeating content control sections, use tables as an alternative:

```typescript
export interface LineItem {
  description: string;
  quantity: number;
  unitPrice: number;
}

/**
 * Creates a dynamic table for repeating data
 * Pattern from Microsoft OOXML documentation
 */
export async function createDynamicLineItemsTable(
  headers: string[],
  items: LineItem[]
): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert table
    const table = body.insertTable(items.length + 1, headers.length, Word.InsertLocation.end);
    
    // Add headers
    for (let i = 0; i < headers.length; i++) {
      const cell = table.getCell(0, i);
      cell.value = headers[i];
      cell.body.font.bold = true;
    }
    
    // Add data rows
    for (let row = 0; row < items.length; row++) {
      const item = items[row];
      table.getCell(row + 1, 0).value = item.description;
      table.getCell(row + 1, 1).value = item.quantity.toString();
      table.getCell(row + 1, 2).value = `$${item.unitPrice.toFixed(2)}`;
      table.getCell(row + 1, 3).value = `$${(item.quantity * item.unitPrice).toFixed(2)}`;
    }
    
    // Style the table
    table.style = "Table Grid";
    
    await context.sync();
  });
}
```

## Using OOXML for Complex Templates

When standard APIs aren't sufficient, use OOXML coercion:

```typescript
/**
 * Inserts complex formatted content using OOXML
 * Based on Microsoft guidance for Office Open XML
 */
export async function insertFormattedContentWithOOXML(ooxmlContent: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      ooxmlContent,
      { coercionType: Office.CoercionType.Ooxml },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(`OOXML insertion failed: ${result.error.message}`));
        }
      }
    );
  });
}

/**
 * Generates OOXML for a styled paragraph with content control
 */
export function generateContentControlOOXML(
  tag: string,
  title: string,
  defaultText: string
): string {
  return `
    <w:sdt>
      <w:sdtPr>
        <w:alias w:val="${title}"/>
        <w:tag w:val="${tag}"/>
        <w:sdtContent w:val="${defaultText}"/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t>${defaultText}</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
  `;
}
```

## Complete Document Generation Workflow

```typescript
export interface DocumentGenerationConfig {
  templatePath?: string;
  outputPath?: string;
  data: TemplateData;
  lineItems?: LineItem[];
}

/**
 * Complete workflow for generating documents from templates
 */
export async function generateDocument(config: DocumentGenerationConfig): Promise<void> {
  await Word.run(async (context) => {
    // 1. Populate main template fields
    await populateTemplate(config.data);
    
    // 2. Add line items if provided
    if (config.lineItems && config.lineItems.length > 0) {
      await createDynamicLineItemsTable(
        ["Description", "Quantity", "Unit Price", "Total"],
        config.lineItems
      );
    }
    
    // 3. Add timestamp
    const body = context.document.body;
    body.insertParagraph("", Word.InsertLocation.end);
    const timestamp = body.insertParagraph(
      `Generated: ${new Date().toLocaleString()}`,
      Word.InsertLocation.end
    );
    timestamp.font.size = 10;
    timestamp.font.color = "gray";
    
    await context.sync();
  });
}

// Example usage:
// await generateDocument({
//   data: {
//     "company-name": "Acme Corp",
//     "contract-date": "2024-01-15",
//     "priority": "High"
//   },
//   lineItems: [
//     { description: "Consulting", quantity: 10, unitPrice: 150 },
//     { description: "Development", quantity: 20, unitPrice: 100 }
//   ]
// });
```

## Best Practices

1. **Validate before generation** — Check all required fields have data
2. **Use content control types appropriately** — Date pickers for dates, dropdowns for enums
3. **Handle missing fields gracefully** — Don't fail if optional fields are empty
4. **Batch operations** — Load all content controls, then sync once
5. **Use OOXML sparingly** — Prefer Word JS APIs for simple operations

## Limitations

- Word.Template API is in preview (WordApi BETA)
- Repeating sections require custom table implementation
- OOXML validation errors are hard to debug
- Content controls must have unique tags
