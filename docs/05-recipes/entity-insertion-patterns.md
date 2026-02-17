# Entity Insertion Patterns

Inserting fixed entities, reusable content blocks, and standard clauses using content controls for enterprise document automation.

## Overview

Entity insertion patterns enable:
- **Standard clause libraries** — Legal paragraphs, terms & conditions
- **Reusable content blocks** — Company boilerplate, signatures
- **Dynamic content insertion** — Based on document context
- **Version-controlled entities** — Track changes to standard text

## References

- [Word.ContentControl API](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol)
- [Building Block Gallery Content Controls](https://learn.microsoft.com/en-us/office/dev/add-ins/word/content-controls#building-block-gallery-content-controls)

## Fixed Entity Library

### Defining Standard Entities

```typescript
/**
 * Standard entity definition for reusable content
 */
export interface StandardEntity {
  id: string;
  name: string;
  category: 'legal' | 'boilerplate' | 'signature' | 'header' | 'footer';
  content: string;
  version: string;
  lastUpdated: Date;
}

/**
 * Entity library for document automation
 * Based on Microsoft patterns for reusable content
 */
export class EntityLibrary {
  private entities: Map<string, StandardEntity> = new Map();

  constructor() {
    this.loadDefaultEntities();
  }

  private loadDefaultEntities(): void {
    const defaults: StandardEntity[] = [
      {
        id: 'confidentiality-clause',
        name: 'Confidentiality Clause',
        category: 'legal',
        content: 'Both parties agree to maintain the confidentiality of all proprietary information disclosed during the term of this agreement. This obligation shall survive termination for a period of five (5) years.',
        version: '1.0',
        lastUpdated: new Date('2024-01-01')
      },
      {
        id: 'governing-law',
        name: 'Governing Law',
        category: 'legal',
        content: 'This Agreement shall be governed by and construed in accordance with the laws of the State of [STATE], without regard to its conflict of law provisions.',
        version: '2.1',
        lastUpdated: new Date('2024-02-15')
      },
      {
        id: 'company-signature',
        name: 'Company Signature Block',
        category: 'signature',
        content: `ACME CORPORATION

By: _______________________
Name: [AUTHORIZED_NAME]
Title: [AUTHORIZED_TITLE]
Date: _____________________`,
        version: '1.0',
        lastUpdated: new Date('2024-01-01')
      },
      {
        id: 'limitation-liability',
        name: 'Limitation of Liability',
        category: 'legal',
        content: 'IN NO EVENT SHALL EITHER PARTY BE LIABLE FOR ANY INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES, INCLUDING WITHOUT LIMITATION, LOSS OF PROFITS, DATA, USE, GOODWILL, OR OTHER INTANGIBLE LOSSES.',
        version: '1.2',
        lastUpdated: new Date('2024-03-01')
      }
    ];

    for (const entity of defaults) {
      this.entities.set(entity.id, entity);
    }
  }

  getEntity(id: string): StandardEntity | undefined {
    return this.entities.get(id);
  }

  getByCategory(category: StandardEntity['category']): StandardEntity[] {
    return Array.from(this.entities.values())
      .filter(e => e.category === category);
  }

  search(query: string): StandardEntity[] {
    const lowerQuery = query.toLowerCase();
    return Array.from(this.entities.values())
      .filter(e => 
        e.name.toLowerCase().includes(lowerQuery) ||
        e.content.toLowerCase().includes(lowerQuery)
      );
  }
}

export const entityLibrary = new EntityLibrary();
```

## Inserting Fixed Entities

### Basic Entity Insertion

```typescript
/**
 * Inserts a standard entity at the current selection
 */
export async function insertEntity(entityId: string): Promise<void> {
  const entity = entityLibrary.getEntity(entityId);
  
  if (!entity) {
    throw new Error(`Entity not found: ${entityId}`);
  }

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Insert entity content
    const inserted = range.insertText(entity.content, Word.InsertLocation.replace);
    
    // Style based on category
    switch (entity.category) {
      case 'legal':
        inserted.font.italic = true;
        break;
      case 'signature':
        inserted.font.size = 11;
        break;
    }
    
    await context.sync();
  });
}
```

### Entity with Content Controls

```typescript
export interface EntityWithFields {
  entityId: string;
  fieldValues: { [placeholder: string]: string };
}

/**
 * Inserts an entity with replaceable field placeholders
 * Pattern: [FIELD_NAME] gets replaced with actual value
 */
export async function insertEntityWithFields(
  entityWithFields: EntityWithFields
): Promise<void> {
  const entity = entityLibrary.getEntity(entityWithFields.entityId);
  
  if (!entity) {
    throw new Error(`Entity not found: ${entityWithFields.entityId}`);
  }

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Replace placeholders with values
    let content = entity.content;
    for (const [placeholder, value] of Object.entries(entityWithFields.fieldValues)) {
      content = content.replace(
        new RegExp(`\\[${placeholder}\\]`, 'g'),
        value
      );
    }
    
    // Insert processed content
    const inserted = range.insertText(content, Word.InsertLocation.replace);
    
    await context.sync();
  });
}

// Example:
// await insertEntityWithFields({
//   entityId: 'company-signature',
//   fieldValues: {
//     AUTHORIZED_NAME: 'John Doe',
//     AUTHORIZED_TITLE: 'CEO'
//   }
// });
```

## Content Control-Based Entity Insertion

### Creating Insertion Points

```typescript
/**
 * Creates content controls that act as insertion points for entities
 * These can be replaced with actual content later
 */
export async function createEntityInsertionPoint(
  entityType: string,
  label: string
): Promise<number> {
  return await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Create content control as placeholder
    const control = range.insertContentControl();
    control.tag = `entity:${entityType}`;
    control.title = label;
    control.placeholderText = `[Insert ${label}]`;
    control.appearance = Word.ContentControlAppearance.boundingBox;
    control.color = "blue";
    
    await context.sync();
    return control.id;
  });
}

/**
 * Replaces all entity insertion points with actual content
 */
export async function resolveEntityInsertionPoints(): Promise<string[]> {
  return await Word.run(async (context) => {
    const resolved: string[] = [];
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag");
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      // Check if this is an entity insertion point
      if (control.tag?.startsWith("entity:")) {
        const entityType = control.tag.replace("entity:", "");
        
        // Find matching entity
        const entity = entityLibrary.getEntity(entityType);
        
        if (entity) {
          // Replace content control content
          control.insertText(entity.content, Word.InsertLocation.replace);
          control.tag = `resolved:${entityType}`;
          resolved.push(entityType);
        } else {
          console.warn(`No entity found for type: ${entityType}`);
        }
      }
    }
    
    await context.sync();
    return resolved;
  });
}
```

## Building Block Gallery Content Controls

For native Word building blocks (AutoText, Quick Parts):

```typescript
/**
 * Creates a building block gallery content control
 * Allows users to pick from Word's built-in gallery
 */
export async function createBuildingBlockGallery(
  galleryType: 'autoText' | 'quickParts' | 'coverPages' | 'equations',
  title: string
): Promise<number> {
  return await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    const control = range.insertContentControl();
    control.type = Word.ContentControlType.buildingBlockGallery;
    control.title = title;
    control.placeholderText = `Select ${galleryType}...`;
    
    // Note: The galleryType property may vary based on Word version
    // This creates a content control that shows the building block gallery
    
    await context.sync();
    return control.id;
  });
}
```

## Conditional Entity Insertion

```typescript
export interface ConditionalEntity {
  entityId: string;
  condition: (documentContext: DocumentContext) => boolean;
}

export interface DocumentContext {
  documentType?: string;
  region?: string;
  industry?: string;
  hasSensitiveData?: boolean;
}

/**
 * Inserts entities based on document context
 * Useful for compliance and region-specific clauses
 */
export async function insertConditionalEntities(
  context: DocumentContext,
  conditionalEntities: ConditionalEntity[]
): Promise<string[]> {
  const inserted: string[] = [];
  
  for (const item of conditionalEntities) {
    if (item.condition(context)) {
      await insertEntity(item.entityId);
      inserted.push(item.entityId);
    }
  }
  
  return inserted;
}

// Example usage:
// const conditionalEntities: ConditionalEntity[] = [
//   {
//     entityId: 'gdpr-clause',
//     condition: (ctx) => ctx.region === 'EU'
//   },
//   {
//     entityId: 'data-processing-agreement',
//     condition: (ctx) => ctx.hasSensitiveData === true
//   }
// ];
// await insertConditionalEntities(context, conditionalEntities);
```

## Version Tracking

```typescript
export interface EntityVersionInfo {
  entityId: string;
  version: string;
  insertedAt: Date;
  insertedBy?: string;
}

/**
 * Tracks entity insertion for audit purposes
 */
export async function insertEntityWithTracking(
  entityId: string,
  userId?: string
): Promise<EntityVersionInfo> {
  const entity = entityLibrary.getEntity(entityId);
  
  if (!entity) {
    throw new Error(`Entity not found: ${entityId}`);
  }

  await insertEntity(entityId);
  
  return {
    entityId,
    version: entity.version,
    insertedAt: new Date(),
    insertedBy: userId
  };
}
```

## Best Practices

1. **Categorize entities** — Group by type (legal, boilerplate, etc.)
2. **Version control** — Track entity versions for compliance
3. **Use content controls** — Enable later editing and replacement
4. **Validate placeholders** — Check all [FIELDS] are replaced
5. **Document dependencies** — Some entities may require others

## Common Patterns

| Pattern | Use Case | Implementation |
|---------|----------|----------------|
| Standard clauses | Legal documents | Entity library with IDs |
| Conditional insertion | Compliance | Context-based conditions |
| Field replacement | Customization | Placeholder substitution |
| Version tracking | Audit trails | Metadata on insertion |
| Building blocks | Quick Parts | Native Word gallery |
