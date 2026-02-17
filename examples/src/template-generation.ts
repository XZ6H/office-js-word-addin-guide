/**
 * Template-Based Document Generation
 * 
 * Examples from docs/05-recipes/template-based-document-generation.md
 * Based on Microsoft Office JS patterns
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

export interface TemplateField {
  tag: string;
  title: string;
  type: Word.ContentControlType;
  placeholder?: string;
}

export interface TemplateData {
  [fieldTag: string]: string | string[];
}

export interface LineItem {
  description: string;
  quantity: number;
  unitPrice: number;
}

/**
 * Creates a document template with content control placeholders
 */
export async function createDocumentTemplate(fields: TemplateField[]): Promise<void> {
  await safeWordRun(async (context) => {
    const body = context.document.body;
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
      const label = body.insertParagraph(`${field.title}:`, Word.InsertLocation.end);
      label.font.bold = true;
      
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

/**
 * Populates a template document with data
 */
export async function populateTemplate(data: TemplateData): Promise<string[]> {
  return await safeWordRun(async (context) => {
    const filled: string[] = [];
    const skipped: string[] = [];
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag, type");
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      const value = data[control.tag];
      
      if (value !== undefined) {
        if (Array.isArray(value)) {
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

/**
 * Creates a dynamic table for repeating data
 */
export async function createDynamicLineItemsTable(
  headers: string[],
  items: LineItem[]
): Promise<void> {
  await safeWordRun(async (context) => {
    const body = context.document.body;
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
    
    table.style = "Table Grid";
    await context.sync();
  });
}

/**
 * Inserts complex formatted content using OOXML coercion
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
