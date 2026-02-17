# Mail Merge

Mail merge enables personalized document generation from data sources like databases, spreadsheets, or APIs.

## Overview

A mail merge typically involves:
1. Template document with placeholders
2. Data source with recipient information
3. Merging data into copies of the template

## Template Preparation

### Creating a Template with Placeholders

```typescript
async function createMergeTemplate(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert merge fields using content controls
    const greeting = body.insertParagraph("Dear {{FIRST_NAME}} {{LAST_NAME}},", Word.InsertLocation.end);
    greeting.insertContentControl().tag = "greeting";
    
    body.insertParagraph("", Word.InsertLocation.end);
    body.insertParagraph("Your account {{ACCOUNT_NUMBER}} is ready.", Word.InsertLocation.end);
    body.insertParagraph("Company: {{COMPANY}}", Word.InsertLocation.end);
    body.insertParagraph("Address: {{ADDRESS}}", Word.InsertLocation.end);
    
    await context.sync();
  });
}
```

### Using Content Controls as Merge Fields

```typescript
async function setupMergeFields(): Promise<void> {
  await Word.run(async (context) => {
    const fields = [
      { tag: "first-name", display: "First Name" },
      { tag: "last-name", display: "Last Name" },
      { tag: "company", display: "Company" },
      { tag: "address", display: "Address" },
      { tag: "city", display: "City" },
      { tag: "postal-code", display: "Postal Code" }
    ];
    
    for (const field of fields) {
      const range = context.document.getSelection();
      const control = range.insertContentControl();
      control.tag = field.tag;
      control.title = field.display;
      control.placeholderText = `[${field.display}]`;
    }
    
    await context.sync();
  });
}
```

## Data Merge Operation

### Single Record Merge

```typescript
interface MergeData {
  [field: string]: string;
}

async function mergeSingleRecord(data: MergeData): Promise<void> {
  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag");
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag && data[control.tag]) {
        control.insertText(data[control.tag], Word.InsertLocation.replace);
      }
    }
    
    await context.sync();
  });
}
```

### Batch Mail Merge

```typescript
interface Recipient {
  id: string;
  data: MergeData;
}

async function batchMailMerge(
  recipients: Recipient[],
  templateId: string
): Promise<string[]> {
  const generatedDocuments: string[] = [];
  
  for (const recipient of recipients) {
    try {
      // In real implementation, you'd:
      // 1. Open template document
      // 2. Merge data
      // 3. Save as new document
      // 4. Track document ID
      
      await Word.run(async (context) => {
        // Merge logic here
        await mergeSingleRecord(recipient.data);
        
        // Save document
        const savedDoc = await saveMergedDocument(recipient.id);
        generatedDocuments.push(savedDoc);
      });
    } catch (error) {
      console.error(`Failed to merge for recipient ${recipient.id}:`, error);
    }
  }
  
  return generatedDocuments;
}

async function saveMergedDocument(recipientId: string): Promise<string> {
  // Implementation depends on your storage strategy
  // Could save to SharePoint, OneDrive, or local storage
  return `document-${recipientId}.docx`;
}
```

## Data Source Integration

### From Excel

```typescript
async function readDataFromExcel(file: File): Promise<Recipient[]> {
  const data: Recipient[] = [];
  
  // Use FileReader to read Excel
  const reader = new FileReader();
  
  return new Promise((resolve, reject) => {
    reader.onload = (e) => {
      try {
        const result = e.target?.result;
        // Parse Excel data (you'd use xlsx library here)
        // const workbook = XLSX.read(result, { type: 'array' });
        // const sheet = workbook.Sheets[workbook.SheetNames[0]];
        // const json = XLSX.utils.sheet_to_json(sheet);
        
        // Mock for example
        resolve(data);
      } catch (error) {
        reject(error);
      }
    };
    
    reader.readAsArrayBuffer(file);
  });
}
```

### From API

```typescript
async function fetchMergeData(apiUrl: string): Promise<Recipient[]> {
  try {
    const response = await fetch(apiUrl, {
      headers: {
        'Authorization': 'Bearer YOUR_TOKEN',
        'Content-Type': 'application/json'
      }
    });
    
    if (!response.ok) {
      throw new Error(`API returned ${response.status}`);
    }
    
    const data = await response.json();
    
    return data.map((item: any, index: number) => ({
      id: item.id || `recipient-${index}`,
      data: {
        "first-name": item.firstName,
        "last-name": item.lastName,
        "company": item.company,
        "address": item.address,
        "email": item.email
      }
    }));
  } catch (error) {
    console.error("Failed to fetch merge data:", error);
    throw error;
  }
}
```

## Conditional Logic

### Conditional Content

```typescript
async function mergeWithConditions(data: MergeData): Promise<void> {
  await Word.run(async (context) => {
    // Example: Only show VIP section for VIP customers
    if (data.customerType === "VIP") {
      const vipSection = context.document.body
        .insertParagraph("VIP Benefits: Priority Support", Word.InsertLocation.end);
      vipSection.font.color = "gold";
    }
    
    // Conditional formatting
    const balance = parseFloat(data.accountBalance || "0");
    const balancePara = context.document.body
      .insertParagraph(`Balance: $${balance}`, Word.InsertLocation.end);
    
    if (balance < 0) {
      balancePara.font.color = "red";
    } else if (balance > 10000) {
      balancePara.font.color = "green";
    }
    
    await context.sync();
  });
}
```

## Best Practices

1. **Always validate data** — Check for missing fields before merging
2. **Handle large datasets** — Process in batches to avoid timeout
3. **Preview before merge** — Show user a sample of first record
4. **Log failures** — Track which records failed and why
5. **Sanitize inputs** — Prevent XSS and formatting issues

## Common Patterns

| Pattern | Use Case | Implementation |
|---------|----------|----------------|
| Table merge | Line items | Insert rows dynamically |
| Conditional paragraphs | VIP sections | Check field value, insert if true |
| Image merge | Profile photos | Use picture content controls |
| Multi-page | Letters | Automatic pagination |
