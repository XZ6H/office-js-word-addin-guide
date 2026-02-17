# Recipe: Find and Replace

## Overview

Implement powerful find and replace functionality with support for wildcards, formatting preservation, and batch operations.

## Basic Implementation

```typescript
// src/recipes/find-and-replace.ts
export interface FindReplaceOptions {
  matchCase?: boolean;
  matchWildcards?: boolean;
  matchWholeWord?: boolean;
  preserveFormatting?: boolean;
}

export class FindAndReplaceService {
  async findAndReplace(
    searchText: string,
    replaceText: string,
    options: FindReplaceOptions = {}
  ): Promise<number> {
    return Word.run(async (context) => {
      const body = context.document.body;
      
      const searchResults = body.search(searchText, {
        matchCase: options.matchCase ?? false,
        matchWildcards: options.matchWildcards ?? false,
        matchWholeWord: options.matchWholeWord ?? false
      });
      
      searchResults.load('items');
      await context.sync();
      
      let count = 0;
      
      for (const item of searchResults.items) {
        if (options.preserveFormatting) {
          // Save formatting before replacement
          item.load('font/name, font/size, font/bold, font/italic');
          await context.sync();
          
          const fontName = item.font.name;
          const fontSize = item.font.size;
          const isBold = item.font.bold;
          const isItalic = item.font.italic;
          
          // Replace and restore formatting
          item.insertText(replaceText, Word.InsertLocation.replace);
          item.font.name = fontName;
          item.font.size = fontSize;
          item.font.bold = isBold;
          item.font.italic = isItalic;
        } else {
          item.insertText(replaceText, Word.InsertLocation.replace);
        }
        count++;
      }
      
      await context.sync();
      return count;
    });
  }

  async findAndReplaceWithWildcard(
    pattern: string,
    replaceTemplate: string
  ): Promise<number> {
    return Word.run(async (context) => {
      const body = context.document.body;
      
      // Example: Find phone numbers like (123) 456-7890
      const searchResults = body.search(pattern, {
        matchWildcards: true
      });
      
      searchResults.load('items');
      await context.sync();
      
      searchResults.items.forEach(item => {
        item.insertText(replaceTemplate, Word.InsertLocation.replace);
      });
      
      await context.sync();
      return searchResults.items.length;
    });
  }
}
```

## Usage Examples

```typescript
const service = new FindAndReplaceService();

// Simple replacement
const count1 = await service.findAndReplace('old text', 'new text');
console.log(`Replaced ${count1} occurrences`);

// Case-sensitive with formatting preservation
const count2 = await service.findAndReplace(
  'Company Name',
  'New Company Name',
  { matchCase: true, preserveFormatting: true }
);

// Wildcard pattern (phone numbers)
const count3 = await service.findAndReplaceWithWildcard(
  '([0-9]{3}) [0-9]{3}-[0-9]{4}',
  '[PHONE NUMBER REDACTED]'
);
```

## Performance Tips

- Process large documents in chunks if replacing thousands of items
- Use `preserveFormatting: false` when formatting doesn't matter (faster)
- Batch replacements by collecting all changes before single `context.sync()`

## Error Handling

```typescript
async function safeFindAndReplace(): Promise<void> {
  try {
    const count = await service.findAndReplace('text', 'replacement');
    showNotification(`Successfully replaced ${count} occurrences`);
  } catch (error) {
    if (error.code === 'SearchNotFound') {
      showNotification('No matches found');
    } else {
      showError('Failed to perform replacement', error);
    }
  }
}
```
