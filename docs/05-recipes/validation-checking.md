# Validation Checking

Document validation ensures content meets quality standards, compliance requirements, or business rules before publishing or distribution.

## Overview

Validation checks can verify:
- Content completeness (all required fields filled)
- Format consistency (dates, currency, phone numbers)
- Compliance (word counts, required sections)
- Quality (grammar, broken links, placeholders)

## Basic Validation

### Checking for Empty Fields

```typescript
interface ValidationResult {
  passed: boolean;
  errors: string[];
  warnings: string[];
}

async function validateRequiredFields(): Promise<ValidationResult> {
  return await Word.run(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load(["tag", "title", "text"])
    ;
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
```

### Checking for Placeholders

```typescript
async function validateNoPlaceholders(): Promise<ValidationResult> {
  return await Word.run(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const body = context.document.body;
    body.load("text");
    await context.sync();
    
    const placeholderPatterns = [
      /\{\{.*?\}\}/g,  // {{placeholder}}
      /\[.*?\]/g,       // [placeholder]
      /XXX+/g,          // XXX placeholders
      /TODO|FIXME|PLACEHOLDER/gi  // Common markers
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
```

## Advanced Validation

### Word Count Validation

```typescript
interface WordCountValidation {
  minWords?: number;
  maxWords?: number;
}

async function validateWordCount(config: WordCountValidation): Promise<ValidationResult> {
  return await Word.run(async (context) => {
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
```

### Structural Validation

```typescript
async function validateDocumentStructure(): Promise<ValidationResult> {
  return await Word.run(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    
    // Check for minimum sections
    if (sections.items.length < 1) {
      result.errors.push("Document has no sections");
      result.passed = false;
    }
    
    // Check for title (first paragraph style)
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.load("style");
    await context.sync();
    
    const titleStyles = ["Title", "Heading 1"];
    if (!titleStyles.includes(firstParagraph.style)) {
      result.warnings.push("First paragraph should use Title or Heading 1 style");
    }
    
    return result;
  });
}
```

## Custom Validation Rules

### Email Validation

```typescript
function validateEmail(email: string): boolean {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

async function validateEmailsInDocument(): Promise<ValidationResult> {
  return await Word.run(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    // Find content controls tagged as email
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load(["tag", "text"]);
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag?.includes("email") && control.text) {
        if (!validateEmail(control.text)) {
          result.errors.push(`Invalid email format: ${control.text}`);
          result.passed = false;
        }
      }
    }
    
    return result;
  });
}
```

### Date Format Validation

```typescript
function validateDateFormat(dateStr: string, format: string): boolean {
  // Simple ISO date validation
  const isoDateRegex = /^\d{4}-\d{2}-\d{2}$/;
  return isoDateRegex.test(dateStr);
}

async function validateDates(): Promise<ValidationResult> {
  return await Word.run(async (context) => {
    const result: ValidationResult = {
      passed: true,
      errors: [],
      warnings: []
    };
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag?.includes("date")) {
        control.load(["tag", "text"]);
      }
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag?.includes("date") && control.text) {
        if (!validateDateFormat(control.text, "ISO")) {
          result.warnings.push(`Date "${control.text}" may not be in correct format`);
        }
      }
    }
    
    return result;
  });
}
```

## Running All Validations

```typescript
async function runAllValidations(): Promise<ValidationResult[]> {
  const validations = [
    { name: "Required Fields", fn: validateRequiredFields },
    { name: "Placeholders", fn: validateNoPlaceholders },
    { name: "Word Count", fn: () => validateWordCount({ minWords: 100 }) },
    { name: "Structure", fn: validateDocumentStructure },
    { name: "Emails", fn: validateEmailsInDocument }
  ];
  
  const results: ValidationResult[] = [];
  
  for (const validation of validations) {
    try {
      const result = await validation.fn();
      result.passed 
        ? console.log(`✅ ${validation.name} passed`)
        : console.log(`❌ ${validation.name} failed`);
      results.push(result);
    } catch (error) {
      console.error(`Error in ${validation.name}:`, error);
      results.push({
        passed: false,
        errors: [`Validation error: ${error}`],
        warnings: []
      });
    }
  }
  
  return results;
}

function formatValidationReport(results: ValidationResult[]): string {
  let report = "# Validation Report\n\n";
  
  for (const result of results) {
    const status = result.passed ? "✅ PASS" : "❌ FAIL";
    report += `## ${status}\n\n`;
    
    if (result.errors.length > 0) {
      report += "### Errors\n";
      for (const error of result.errors) {
        report += `- ${error}\n`;
      }
      report += "\n";
    }
    
    if (result.warnings.length > 0) {
      report += "### Warnings\n";
      for (const warning of result.warnings) {
        report += `- ${warning}\n`;
      }
      report += "\n";
    }
  }
  
  return report;
}
```

## Best Practices

1. **Categorize issues** — Separate errors (blocking) from warnings (advisory)
2. **Provide location info** — Tell user where issues are found
3. **Batch validations** — Run all checks before showing results
4. **Allow customization** — Let users configure which rules to enforce
5. **Show progress** — Validation can take time on large documents

## User Interface

### Displaying Results

```typescript
function displayValidationResults(results: ValidationResult[]): void {
  const container = document.getElementById("validation-results");
  if (!container) return;
  
  container.innerHTML = "";
  
  const allPassed = results.every(r => r.passed);
  
  const summary = document.createElement("div");
  summary.className = allPassed ? "validation-success" : "validation-fail";
  summary.textContent = allPassed 
    ? "✅ All validations passed!"
    : `❌ ${results.filter(r => !r.passed).length} validation(s) failed`;
  container.appendChild(summary);
  
  for (const result of results) {
    if (!result.passed || result.warnings.length > 0) {
      const section = document.createElement("div");
      section.className = "validation-section";
      
      result.errors.forEach(error => {
        const item = document.createElement("div");
        item.className = "validation-error";
        item.textContent = `❌ ${error}`;
        section.appendChild(item);
      });
      
      result.warnings.forEach(warning => {
        const item = document.createElement("div");
        item.className = "validation-warning";
        item.textContent = `⚠️ ${warning}`;
        section.appendChild(item);
      });
      
      container.appendChild(section);
    }
  }
}
```
