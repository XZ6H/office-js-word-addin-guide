# Office JS Word API Deep Dive

## Document Manipulation

### Working with the Document Body

```typescript
async function documentOperations(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;

    // Insert text at different locations
    body.insertText('Start of document\\n', Word.InsertLocation.start);
    body.insertText('End of document', Word.InsertLocation.end);
    
    // Insert HTML
    body.insertHtml('<p>HTML paragraph</p>', Word.InsertLocation.end);
    
    // Insert Office Open XML
    body.insertOoxml('<w:p><w:r><w:t>OOXML text</w:t></w:r></w:p>', Word.InsertLocation.end);
    
    // Insert file
    body.insertFileFromBase64(base64Content, Word.InsertLocation.end);
    
    await context.sync();
  });
}
```

### Search and Replace

```typescript
async function searchAndReplace(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;

    // Simple search
    const searchResults = body.search('target text');
    searchResults.load('items');
    await context.sync();

    // Replace all occurrences
    searchResults.items.forEach(item => {
      item.insertText('replacement text', Word.InsertLocation.replace);
    });

    // Search with wildcards
    const wildcardResults = body.search('[0-9]{3}-[0-9]{4}', {
      matchWildcards: true
    });

    await context.sync();
  });
}
```

## Range Operations

### Creating and Manipulating Ranges

```typescript
async function rangeOperations(): Promise<void> {
  await Word.run(async (context) => {
    // Get different range types
    const selection = context.document.getSelection();
    const body = context.document.body;
    const entireDoc = body.getRange();
    
    // Expand range
    const expandedRange = selection.expandTo(body);
    
    // Get range from different locations
    const startRange = body.getRange(Word.RangeLocation.start);
    const endRange = body.getRange(Word.RangeLocation.end);
    const wholeRange = body.getRange(Word.RangeLocation.whole);
    
    // Compare ranges
    const comparison = selection.compareLocationWith(body);
    await context.sync();
    
    // comparison.value will be:
    // 'Equal', 'Contains', 'Inside', 'AdjacentBefore', 'AdjacentAfter', 'Before', 'After', 'Unknown'
  });
}
```

### Range Formatting

```typescript
async function formatRange(): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Font formatting
    range.font.name = 'Calibri';
    range.font.size = 11;
    range.font.bold = true;
    range.font.italic = false;
    range.font.underline = Word.UnderlineType.single;
    range.font.color = 'blue';
    range.font.highlightColor = 'yellow';
    
    // Paragraph formatting
    range.paragraphFormat.alignment = Word.Alignment.center;
    range.paragraphFormat.lineSpacing = 18;
    range.paragraphFormat.spaceBefore = 12;
    range.paragraphFormat.spaceAfter = 12;
    
    await context.sync();
  });
}
```

## Table Operations

### Creating Tables

```typescript
async function createTable(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    
    // Create 3x3 table
    const table = body.insertTable(3, 3, Word.InsertLocation.end);
    
    // Set header row
    table.getCell(0, 0).insertText('Name', Word.InsertLocation.replace);
    table.getCell(0, 1).insertText('Age', Word.InsertLocation.replace);
    table.getCell(0, 2).insertText('City', Word.InsertLocation.replace);
    
    // Add data
    table.getCell(1, 0).insertText('John', Word.InsertLocation.replace);
    table.getCell(1, 1).insertText('30', Word.InsertLocation.replace);
    table.getCell(1, 2).insertText('NYC', Word.InsertLocation.replace);
    
    // Format table
    table.style = 'Table Grid';
    table.getRow(0).font.bold = true;
    
    await context.sync();
  });
}
```

### Table Manipulation

```typescript
async function manipulateTable(): Promise<void> {
  await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();
    
    if (tables.items.length > 0) {
      const table = tables.items[0];
      
      // Add row
      table.addRows(Word.InsertLocation.end, 1, table.getRow(0).values);
      
      // Add column
      table.addColumns(Word.InsertLocation.end, 1, ['New Col']);
      
      // Merge cells
      const cell1 = table.getCell(0, 0);
      const cell2 = table.getCell(0, 1);
      cell1.merge(cell2);
      
      // Set column width
      table.columns.getItemAt(0).width = 150;
      
      await context.sync();
    }
  });
}
```

## Content Controls

### Creating Content Controls

```typescript
async function createContentControls(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert text and wrap in content control
    const range = body.insertText('Editable content', Word.InsertLocation.end);
    const contentControl = range.insertContentControl();
    
    // Configure content control
    contentControl.title = 'User Input';
    contentControl.tag = 'user-input-1';
    contentControl.placeholderText = 'Enter your text here';
    contentControl.appearance = Word.ContentControlAppearance.boundingBox;
    
    // Set lock settings
    contentControl.cannotEdit = false;
    contentControl.cannotDelete = true;
    
    await context.sync();
  });
}
```

### Working with Existing Content Controls

```typescript
async function manageContentControls(): Promise<void> {
  await Word.run(async (context) => {
    // Get all content controls
    const contentControls = context.document.contentControls;
    contentControls.load('items');
    await context.sync();
    
    // Iterate and modify
    contentControls.items.forEach(cc => {
      if (cc.tag === 'user-input-1') {
        cc.insertText('Updated content', Word.InsertLocation.replace);
      }
    });
    
    // Get specific content control by tag
    const targetCC = contentControls.getByTag('target-tag');
    targetCC.load('items');
    await context.sync();
    
    if (targetCC.items.length > 0) {
      targetCC.items[0].color = 'blue';
    }
    
    await context.sync();
  });
}
```

## Custom XML Parts

### Reading and Writing Custom XML

```typescript
async function customXmlOperations(): Promise<void> {
  await Word.run(async (context) => {
    // Get custom XML parts
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load('items');
    await context.sync();
    
    // Add new XML part
    const xmlString = `
      <root xmlns="http://schemas.myapp.com/data">
        <metadata>
          <version>1.0</version>
          <created>2024-01-01</created>
        </metadata>
      </root>
    `;
    
    const newPart = customXmlParts.add(xmlString);
    await context.sync();
    
    // Query XML
    const xpathResults = newPart.getXml('/root/metadata/version');
    await context.sync();
    
    console.log(xpathResults.value);
    
    // Update XML
    newPart.setXml('/root/metadata/version', '<version>2.0</version>');
    await context.sync();
  });
}
```

## Event Handling

### Document Selection Changed

```typescript
function registerSelectionHandler(): void {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    handleSelectionChange
  );
}

async function handleSelectionChange(event: Office.DocumentSelectionChangedEventArgs): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    
    console.log('Selection changed:', selection.text);
    updateUI(selection.text);
  });
}

function unregisterHandler(): void {
  Office.context.document.removeHandlerAsync(
    Office.EventType.DocumentSelectionChanged
  );
}
```

## Advanced Range Techniques

### Splitting Ranges

```typescript
async function splitRange(): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Split by delimiter
    const delimiter = '\\n';
    const subRanges = range.split(delimiter);
    
    subRanges.load('items');
    await context.sync();
    
    subRanges.items.forEach((subRange, index) => {
      subRange.font.color = index % 2 === 0 ? 'blue' : 'green';
    });
    
    await context.sync();
  });
}
```

### Intersecting Ranges

```typescript
async function intersectRanges(): Promise<void> {
  await Word.run(async (context) => {
    const range1 = context.document.body.paragraphs.getFirst().getRange();
    const range2 = context.document.body.paragraphs.getLast().getRange();
    
    // Get intersection
    const intersection = range1.intersectWith(range2);
    intersection.load('isEmpty');
    await context.sync();
    
    if (!intersection.isEmpty) {
      // Ranges overlap
      intersection.font.highlightColor = 'yellow';
      await context.sync();
    }
  });
}
```

## API Quick Reference

| Operation | Method |
|-----------|--------|
| Insert text | `range.insertText(text, location)` |
| Insert HTML | `range.insertHtml(html, location)` |
| Insert table | `range.insertTable(rows, cols, location)` |
| Search | `range.search(text, options)` |
| Select | `range.select()` |
| Delete | `range.delete()` |
| Load properties | `object.load(properties)` |
| Sync | `context.sync()` |

