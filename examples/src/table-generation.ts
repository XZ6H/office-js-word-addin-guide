/**
 * Table Generation Operations
 * 
 * Examples from docs/05-recipes/table-generation.md
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

export interface TableData {
  headers: string[];
  rows: string[][];
}

/**
 * Create a table at the current selection
 */
export async function createTable(
  data: TableData,
  style?: string
): Promise<void> {
  await safeWordRun(async (context) => {
    const selection = context.document.getSelection();
    const table = selection.insertTable(
      data.rows.length + 1, // +1 for header
      data.headers.length
    );
    
    // Add headers
    for (let col = 0; col < data.headers.length; col++) {
      table.getCell(0, col).value = data.headers[col];
    }
    
    // Add data rows
    for (let row = 0; row < data.rows.length; row++) {
      for (let col = 0; col < data.rows[row].length; col++) {
        table.getCell(row + 1, col).value = data.rows[row][col];
      }
    }
    
    // Apply style if provided
    if (style) {
      table.style = style;
    }
    
    await context.sync();
  });
}

/**
 * Insert a row at the end of a table
 */
export async function addTableRow(
  tableIndex: number,
  rowData: string[]
): Promise<void> {
  await safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    
    if (tableIndex >= tables.items.length) {
      throw new Error(`Table index ${tableIndex} not found`);
    }
    
    const table = tables.items[tableIndex];
    const newRow = table.addRow();
    
    for (let col = 0; col < rowData.length; col++) {
      newRow.getCell(col).value = rowData[col];
    }
    
    await context.sync();
  });
}

/**
 * Format table headers
 */
export async function formatTableHeaders(tableIndex: number): Promise<void> {
  await safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    
    if (tableIndex >= tables.items.length) {
      throw new Error(`Table index ${tableIndex} not found`);
    }
    
    const table = tables.items[tableIndex];
    const headerRow = table.getRow(0);
    headerRow.load("font");
    await context.sync();
    
    headerRow.font.bold = true;
    headerRow.font.color = "white";
    
    await context.sync();
  });
}

/**
 * Auto-fit table columns
 */
export async function autoFitTable(tableIndex: number): Promise<void> {
  await safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    
    if (tableIndex >= tables.items.length) {
      throw new Error(`Table index ${tableIndex} not found`);
    }
    
    const table = tables.items[tableIndex];
    table.autoFitColumns();
    
    await context.sync();
  });
}
