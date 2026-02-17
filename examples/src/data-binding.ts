/**
 * Data Binding Patterns
 * 
 * Examples from docs/05-recipes/data-binding-patterns.md
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

export interface ApiBindingConfig {
  endpoint: string;
  method?: 'GET' | 'POST';
  headers?: { [key: string]: string };
  mapping: { [controlTag: string]: string };
}

/**
 * Creates a binding between a content control and a named binding
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
    
    await safeWordRun(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();
      
      for (const control of contentControls.items) {
        control.load("tag");
      }
      await context.sync();
      
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

/**
 * Extracts data from document content controls
 */
export async function extractDocumentData(tags: string[]): Promise<{ [tag: string]: string }> {
  return await safeWordRun(async (context) => {
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
