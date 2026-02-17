# Recipe: Document Assembly

## Overview

Assemble documents from templates by replacing placeholders with dynamic content.

## Implementation

```typescript
// src/recipes/document-assembly.ts
export interface PlaceholderValue {
  key: string;
  value: string;
  preserveFormatting?: boolean;
}

export class DocumentAssemblyService {
  private placeholderPattern: RegExp = /\\{\\{(\\w+)\\}\\}/g;

  async assembleDocument(
    template: string,
    values: PlaceholderValue[],
    options: {
      clearPlaceholders?: boolean;
      placeholderWrapper?: string;
    } = {}
  ): Promise<void> {
    return Word.run(async (context) => {
      const body = context.document.body;
      const wrapper = options.placeholderWrapper ?? '\\{\\{\\}\\}';
      
      for (const placeholder of values) {
        const pattern = wrapper.replace('{}', placeholder.key);
        const searchResults = body.search(pattern, { matchWildcards: false });
        searchResults.load('items');
        await context.sync();
        
        for (const item of searchResults.items) {
          if (placeholder.preserveFormatting) {
            item.load('font/name, font/size, font/bold, font/italic');
            await context.sync();
            
            const formatting = {
              name: item.font.name,
              size: item.font.size,
              bold: item.font.bold,
              italic: item.font.italic
            };
            
            item.insertText(placeholder.value, Word.InsertLocation.replace);
            item.font.name = formatting.name;
            item.font.size = formatting.size;
            item.font.bold = formatting.bold;
            item.font.italic = formatting.italic;
          } else {
            item.insertText(placeholder.value, Word.InsertLocation.replace);
          }
        }
      }
      
      // Clear remaining placeholders if requested
      if (options.clearPlaceholders) {
        const remainingPattern = wrapper.replace('{}', '*');
        const remaining = body.search(remainingPattern, { matchWildcards: true });
        remaining.load('items');
        await context.sync();
        
        remaining.items.forEach(item => {
          item.insertText('', Word.InsertLocation.replace);
        });
      }
      
      await context.sync();
    });
  }

  async extractPlaceholders(): Promise<string[]> {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      
      const matches = body.text.match(this.placeholderPattern) || [];
      const unique = new Set(matches.map(m => m.replace(/[\\{\\}]/g, '')));
      return Array.from(unique);
    });
  }

  async createTemplateFromContent(
    content: string,
    placeholders: Record<string, string>
  ): Promise<void> {
    return Word.run(async (context) => {
      const body = context.document.body;
      
      // Clear existing content
      body.clear();
      
      // Replace placeholders in content string
      let processedContent = content;
      Object.entries(placeholders).forEach(([key, description]) => {
        processedContent = processedContent.replace(
          new RegExp(`\\\\{\\\\{${key}\\\\}\\\\}`, 'g'),
          `{{${key}}}`
        );
      });
      
      body.insertText(processedContent, Word.InsertLocation.start);
      await context.sync();
    });
  }
}
```

## Usage Examples

```typescript
const service = new DocumentAssemblyService();

// Define placeholder values
const values: PlaceholderValue[] = [
  { key: 'CLIENT_NAME', value: 'Acme Corporation' },
  { key: 'PROJECT_NAME', value: 'Website Redesign' },
  { key: 'START_DATE', value: 'January 1, 2024' },
  { key: 'END_DATE', value: 'March 31, 2024' },
  { key: 'TOTAL_COST', value: '$25,000' }
];

// Assemble document
await service.assembleDocument(template, values, {
  clearPlaceholders: true,
  preserveFormatting: true
});

// Extract placeholders from existing document
const placeholders = await service.extractPlaceholders();
console.log('Found placeholders:', placeholders);
```

## Template Example

```
CONTRACT AGREEMENT

This agreement is between {{CLIENT_NAME}} and our company for the project
{{PROJECT_NAME}} starting on {{START_DATE}} and ending on {{END_DATE}}.

Total project cost: {{TOTAL_COST}}
```
