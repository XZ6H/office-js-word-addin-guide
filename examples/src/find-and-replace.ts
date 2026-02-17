/**
 * Find and Replace Operations
 * 
 * Examples from docs/05-recipes/find-and-replace.md
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

/**
 * Simple text replacement in the document body
 */
export async function replaceText(searchText: string, replaceText: string): Promise<number> {
  return await safeWordRun(async (context) => {
    const body = context.document.body;
    const replaceResult = body.replace(searchText, replaceText);
    replaceResult.load("count");
    await context.sync();
    return replaceResult.count;
  });
}

/**
 * Replace text in current selection only
 */
export async function replaceInSelection(
  searchText: string,
  replaceWith: string
): Promise<number> {
  return await safeWordRun(async (context) => {
    const selection = context.document.getSelection();
    const replaceResult = selection.replace(searchText, replaceWith);
    replaceResult.load("count");
    await context.sync();
    return replaceResult.count;
  });
}

/**
 * Replace across the entire document with options
 */
export async function replaceAll(
  searchText: string,
  replaceWith: string,
  matchCase: boolean = false
): Promise<number> {
  return await safeWordRun(async (context) => {
    const body = context.document.body;
    const replaceResult = body.replace(
      searchText,
      replaceWith,
      matchCase ? Word.ReplaceMode.replace : Word.ReplaceMode.replace
    );
    replaceResult.load("count");
    await context.sync();
    return replaceResult.count;
  });
}

/**
 * Replace and format the replacement text
 */
export async function replaceAndFormat(
  searchText: string,
  replaceWith: string,
  formatOptions: {
    bold?: boolean;
    italic?: boolean;
    color?: string;
  }
): Promise<void> {
  await safeWordRun(async (context) => {
    const body = context.document.body;
    
    // Perform replacement
    const replaceResult = body.replace(searchText, replaceWith);
    await context.sync();
    
    // Apply formatting to the inserted text
    const insertedText = body.search(replaceWith);
    insertedText.load("font");
    await context.sync();
    
    if (formatOptions.bold !== undefined) {
      insertedText.font.bold = formatOptions.bold;
    }
    if (formatOptions.italic !== undefined) {
      insertedText.font.italic = formatOptions.italic;
    }
    if (formatOptions.color) {
      insertedText.font.color = formatOptions.color;
    }
    
    await context.sync();
  });
}

/**
 * Replace using wildcards (patterns)
 */
export async function replaceWithWildcards(
  pattern: string,
  replacement: string
): Promise<number> {
  return await safeWordRun(async (context) => {
    const body = context.document.body;
    const replaceResult = body.replace(pattern, replacement);
    replaceResult.load("count");
    await context.sync();
    return replaceResult.count;
  });
}
