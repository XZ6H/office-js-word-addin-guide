/**
 * Document Validation
 * 
 * Examples from docs/05-recipes/validation-checking.md
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

export interface ValidationResult {
  passed: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Validate that all content controls have values
 */
export async function validateRequiredFields(): Promise<ValidationResult> {
  return await safeWordRun(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load(["tag", "title", "text"]);
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (!control.text || control.text.trim() === "") {
        result.errors.push(`Required field "${control.title || control.tag}" is empty`);
        result.passed = false;
      }
    }
    
    return result;
  });
}

/**
 * Check for placeholder text patterns
 */
export async function validateNoPlaceholders(): Promise<ValidationResult> {
  return await safeWordRun(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const body = context.document.body;
    body.load("text");
    await context.sync();
    
    const placeholderPatterns = [
      /\{\{.*?\}\}/g,
      /\[.*?\]/g,
      /XXX+/g,
      /TODO|FIXME|PLACEHOLDER/gi
    ];
    
    for (const pattern of placeholderPatterns) {
      const matches = body.text.match(pattern);
      if (matches) {
        result.errors.push(`Found placeholders: ${matches.join(", ")}`);
        result.passed = false;
      }
    }
    
    return result;
  });
}

/**
 * Validate document word count
 */
export interface WordCountValidation {
  minWords?: number;
  maxWords?: number;
}

export async function validateWordCount(
  config: WordCountValidation
): Promise<ValidationResult> {
  return await safeWordRun(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const body = context.document.body;
    const range = body.getRange();
    context.load(range, "text");
    await context.sync();
    
    const words = range.text.trim().split(/\s+/).length;
    
    if (config.minWords && words < config.minWords) {
      result.errors.push(`Document has ${words} words, minimum required is ${config.minWords}`);
      result.passed = false;
    }
    
    if (config.maxWords && words > config.maxWords) {
      result.errors.push(`Document has ${words} words, maximum allowed is ${config.maxWords}`);
      result.passed = false;
    }
    
    return result;
  });
}

/**
 * Validate email format in content controls
 */
export async function validateEmails(): Promise<ValidationResult> {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  
  return await safeWordRun(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load(["tag", "text"]);
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag?.includes("email") && control.text) {
        if (!emailRegex.test(control.text)) {
          result.errors.push(`Invalid email format: ${control.text}`);
          result.passed = false;
        }
      }
    }
    
    return result;
  });
}

/**
 * Run all validations
 */
export async function runAllValidations(): Promise<ValidationResult[]> {
  const validations = [
    { name: "Required Fields", fn: validateRequiredFields },
    { name: "Placeholders", fn: validateNoPlaceholders },
    { name: "Word Count", fn: () => validateWordCount({ minWords: 100 }) },
    { name: "Emails", fn: validateEmails }
  ];
  
  const results: ValidationResult[] = [];
  
  for (const validation of validations) {
    try {
      const result = await validation.fn();
      results.push(result);
    } catch (error) {
      results.push({
        passed: false,
        errors: [`Validation error: ${error}`],
        warnings: []
      });
    }
  }
  
  return results;
}
