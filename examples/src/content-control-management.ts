/**
 * Content Control Management
 * 
 * Examples from docs/05-recipes/content-control-management.md
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

export interface FormData {
  [tag: string]: string;
}

/**
 * Create a rich text content control
 */
export async function createRichTextControl(
  title: string,
  tag: string,
  placeholder?: string
): Promise<number> {
  return await safeWordRun(async (context) => {
    const range = context.document.getSelection();
    const contentControl = range.insertContentControl();
    contentControl.title = title;
    contentControl.tag = tag;
    if (placeholder) {
      contentControl.placeholderText = placeholder;
    }
    await context.sync();
    return contentControl.id;
  });
}

/**
 * Create a dropdown list content control
 */
export async function createDropdownControl(
  title: string,
  tag: string,
  options: { displayText: string; value: string }[],
  defaultIndex: number = 0
): Promise<number> {
  return await safeWordRun(async (context) => {
    const range = context.document.getSelection();
    const contentControl = range.insertContentControl();
    contentControl.type = Word.ContentControlType.dropDownList;
    contentControl.title = title;
    contentControl.tag = tag;
    contentControl.dropdownListValues = options.map((opt, idx) => ({
      ...opt,
      selected: idx === defaultIndex
    }));
    await context.sync();
    return contentControl.id;
  });
}

/**
 * Fill form template with data
 */
export async function fillFormTemplate(data: FormData): Promise<string[]> {
  return await safeWordRun(async (context) => {
    const filled: string[] = [];
    const missing: string[] = [];
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag");
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag && data[control.tag] !== undefined) {
        control.insertText(data[control.tag], Word.InsertLocation.replace);
        filled.push(control.tag);
      } else if (control.tag) {
        missing.push(control.tag);
      }
    }
    
    await context.sync();
    
    if (missing.length > 0) {
      console.warn("Missing data for fields:", missing);
    }
    
    return filled;
  });
}

/**
 * Get all content control values
 */
export async function getControlValues(): Promise<FormData> {
  return await safeWordRun(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load(["tag", "text"]);
    }
    await context.sync();
    
    const values: FormData = {};
    for (const control of contentControls.items) {
      if (control.tag) {
        values[control.tag] = control.text;
      }
    }
    
    return values;
  });
}

/**
 * Find content control by tag
 */
export async function findControlByTag(
  tag: string
): Promise<{ id: number; text: string } | null> {
  return await safeWordRun(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load(["tag", "id", "text"]);
    }
    await context.sync();
    
    const found = contentControls.items.find(c => c.tag === tag);
    
    return found ? { id: found.id, text: found.text } : null;
  });
}
