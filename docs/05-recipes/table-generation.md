# Recipe: Table Generation

## Overview

Dynamically generate and populate tables from data sources with formatting and styling.

## Implementation

```typescript
// src/recipes/table-generation.ts
export interface TableColumn {
  name: string;
  width?: number;
}

export interface TableData {
  headers: TableColumn[];
  rows: string[][];
}

export class TableGenerationService {
  async createTable(
    data: TableData,
    options: {
      style?: string;
      headerBold?: boolean;
      autoFit?: boolean;
    } = {}
  ): Promise<Word.Table> {
    return Word.run(async (context) => {
      const body = context.document.body;
      const rowCount = data.rows.length + 1; // +1 for header
      const colCount = data.headers.length;
      
      // Insert table
      const table = body.insertTable(rowCount, colCount, Word.InsertLocation.end);
      
      // Set header row
      data.headers.forEach((col, index) => {
        const cell = table.getCell(0, index);
        cell.insertText(col.name, Word.InsertLocation.replace);
        
        if (options.headerBold !== false) {
          cell.font.bold = true;
        }
        
        if (col.width) {
          table.columns.getItemAt(index).width = col.width;
        }
      });
      
      // Set data rows
      data.rows.forEach((row, rowIndex) => {
        row.forEach((cellText, colIndex) => {
          const cell = table.getCell(rowIndex + 1, colIndex);
          cell.insertText(cellText, Word.InsertLocation.replace);
        });
      });
      
      // Apply style
      if (options.style) {
        table.style = options.style;
      } else {
        table.style = 'Table Grid';
      }
      
      // Auto-fit if requested
      if (options.autoFit) {
        table.load('width');
        await context.sync();
        
        const availableWidth = context.document.body.width;
        if (table.width > availableWidth) {
          // Adjust column widths proportionally
          const scaleFactor = availableWidth / table.width;
          table.columns.load('items/width');
          await context.sync();
          
          table.columns.items.forEach(col => {
            col.width = col.width * scaleFactor;
          });
        }
      }
      
      await context.sync();
      return table;
    });
  }

  async createTableFromJSON(jsonData: any[]): Promise<Word.Table> {
    if (jsonData.length === 0) {
      throw new Error('Cannot create table from empty data');
    }
    
    const headers = Object.keys(jsonData[0]).map(key => ({ name: key }));
    const rows = jsonData.map(obj => headers.map(h => String(obj[h.name] ?? '')));
    
    return this.createTable({ headers, rows });
  }

  async formatTable(
    table: Word.Table,
    formatting: {
      headerBackground?: string;
      alternateRowColors?: boolean;
      rowColor1?: string;
      rowColor2?: string;
    }
  ): Promise<void> {
    return Word.run(async (context) => {
      if (formatting.headerBackground) {
        const headerRow = table.rows.getItemAt(0);
        headerRow.shadingColor = formatting.headerBackground;
      }
      
      if (formatting.alternateRowColors) {
        table.rows.load('items');
        await context.sync();
        
        table.rows.items.forEach((row, index) => {
          if (index > 0) { // Skip header
            row.shadingColor = index % 2 === 0 
              ? (formatting.rowColor1 ?? '#F2F2F2')
              : (formatting.rowColor2 ?? '#FFFFFF');
          }
        });
      }
      
      await context.sync();
    });
  }
}
```

## Usage Examples

```typescript
const service = new TableGenerationService();

// Basic table
const tableData: TableData = {
  headers: [
    { name: 'Product', width: 150 },
    { name: 'Price', width: 100 },
    { name: 'Quantity', width: 100 }
  ],
  rows: [
    ['Widget A', '$10.00', '5'],
    ['Widget B', '$20.00', '3'],
    ['Widget C', '$15.00', '8']
  ]
};

const table = await service.createTable(tableData, {
  style: 'Table Grid',
  headerBold: true
});

// Format with alternating colors
await service.formatTable(table, {
  headerBackground: '#4472C4',
  alternateRowColors: true,
  rowColor1: '#E7E6E6',
  rowColor2: '#FFFFFF'
});

// Create from JSON
const jsonData = [
  { name: 'John', age: 30, city: 'NYC' },
  { name: 'Jane', age: 25, city: 'LA' }
];
const jsonTable = await service.createTableFromJSON(jsonData);
```

## Best Practices

- Set explicit column widths for consistent appearance
- Use table styles for consistent formatting across documents
- Load table data before inserting to calculate optimal widths
- Consider maximum table size (Word has limits on rows/columns)
