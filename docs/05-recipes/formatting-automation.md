# Recipe: Formatting Automation

## Overview

Apply consistent formatting across documents with styles, themes, and automated formatting rules.

## Implementation

```typescript
// src/recipes/formatting-automation.ts
export interface FormattingRule {
  selector: 'heading' | 'paragraph' | 'table' | 'list';
  condition?: (item: any) => boolean;
  format: Partial<Word.Font> & Partial<Word.ParagraphFormat>;
}

export class FormattingAutomationService {
  async applyStylesFromTemplate(): Promise<void> {
    return Word.run(async (context) => {
      const body = context.document.body;
      
      // Apply heading styles
      const headings = body.search('^13[!^13]@^13', { matchWildcards: true });
      headings.load('items');
      await context.sync();
      
      headings.items.forEach(heading => {
        heading.load('text');
      });
      await context.sync();
      
      headings.items.forEach(heading => {
        if (heading.text.length < 100) {
          heading.style = 'Heading 1';
        }
      });
      
      await context.sync();
    });
  }

  async applyConsistentFormatting(rules: FormattingRule[]): Promise<void> {
    return Word.run(async (context) => {
      for (const rule of rules) {
        switch (rule.selector) {
          case 'heading':
            await this.formatHeadings(context, rule);
            break;
          case 'paragraph':
            await this.formatParagraphs(context, rule);
            break;
          case 'table':
            await this.formatTables(context, rule);
            break;
          case 'list':
            await this.formatLists(context, rule);
            break;
        }
      }
    });
  }

  private async formatHeadings(context: Word.RequestContext, rule: FormattingRule): Promise<void> {
    const headings = context.document.body.search('^13[!^13]@^13', { matchWildcards: true });
    headings.load('items');
    await context.sync();
    
    for (const heading of headings.items) {
      if (!rule.condition || rule.condition(heading)) {
        Object.assign(heading.font, rule.format);
        Object.assign(heading.paragraphFormat, rule.format);
      }
    }
    
    await context.sync();
  }

  private async formatParagraphs(context: Word.RequestContext, rule: FormattingRule): Promise<void> {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    
    for (const paragraph of paragraphs.items) {
      if (!rule.condition || rule.condition(paragraph)) {
        Object.assign(paragraph.font, rule.format);
        Object.assign(paragraph.paragraphFormat, rule.format);
      }
    }
    
    await context.sync();
  }

  private async formatTables(context: Word.RequestContext, rule: FormattingRule): Promise<void> {
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();
    
    for (const table of tables.items) {
      if (!rule.condition || rule.condition(table)) {
        table.style = rule.format.styleName || 'Table Grid';
      }
    }
    
    await context.sync();
  }

  private async formatLists(context: Word.RequestContext, rule: FormattingRule): Promise<void> {
    const lists = context.document.body.lists;
    lists.load('items');
    await context.sync();
    
    for (const list of lists.items) {
      if (!rule.condition || rule.condition(list)) {
        // Apply list formatting
      }
    }
    
    await context.sync();
  }

  async createTheme(
    themeName: string,
    colors: {
      primary: string;
      secondary: string;
      accent: string;
      text: string;
    }
  ): Promise<void> {
    return Word.run(async (context) => {
      // Store theme in custom XML
      const customXml = context.document.customXmlParts;
      const themeXml = `
        <theme xmlns="http://schemas.myapp.com/theme" name="${themeName}">
          <colors>
            <primary>${colors.primary}</primary>
            <secondary>${colors.secondary}</secondary>
            <accent>${colors.accent}</accent>
            <text>${colors.text}</text>
          </colors>
        </theme>
      `;
      
      customXml.add(themeXml);
      await context.sync();
    });
  }
}
```

## Usage Examples

```typescript
const service = new FormattingAutomationService();

// Apply comprehensive formatting rules
const rules: FormattingRule[] = [
  {
    selector: 'heading',
    condition: (heading) => heading.text.length < 50,
    format: {
      name: 'Calibri',
      size: 16,
      bold: true,
      color: '#2E75B6'
    }
  },
  {
    selector: 'paragraph',
    condition: (para) => para.text.startsWith('Note:'),
    format: {
      italic: true,
      color: '#666666'
    }
  },
  {
    selector: 'table',
    format: {
      styleName: 'Table Grid'
    }
  }
];

await service.applyConsistentFormatting(rules);

// Create and apply theme
await service.createTheme('Corporate', {
  primary: '#4472C4',
  secondary: '#ED7D31',
  accent: '#A5A5A5',
  text: '#333333'
});
```

## Performance Tips

- Batch formatting operations for better performance
- Use built-in styles instead of manual formatting when possible
- Load properties selectively to minimize sync calls
