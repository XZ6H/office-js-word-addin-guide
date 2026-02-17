/**
 * Document Utilities - Core helper functions for Office JS Word Add-ins
 */

import Word from "@types/office-js";

/**
 * Safely execute Word.run with error handling
 */
export async function safeWordRun<T>(
  fn: (context: Word.RequestContext) => Promise<T>
): Promise<T> {
  try {
    return await Word.run(fn);
  } catch (error) {
    console.error("Word.run error:", error);
    throw new DocumentError("Failed to execute Word operation", error);
  }
}

/**
 * Custom error class for document operations
 */
export class DocumentError extends Error {
  constructor(
    message: string,
    public readonly cause?: unknown
  ) {
    super(message);
    this.name = "DocumentError";
  }
}

/**
 * Batch load properties from an array of objects
 */
export async function batchLoad<T extends { load: (props: string | string[]) => void }>(
  items: T[],
  properties: string | string[]
): Promise<void> {
  for (const item of items) {
    item.load(properties);
  }
}

/**
 * Check if add-in is running in desktop application
 */
export function isDesktop(): boolean {
  return Office.context.platform === Office.PlatformType.PC ||
         Office.context.platform === Office.PlatformType.Mac;
}

/**
 * Check if add-in is running in web browser
 */
export function isWeb(): boolean {
  return Office.context.platform === Office.PlatformType.OfficeOnline;
}
