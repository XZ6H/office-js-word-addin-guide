/**
 * Entity Insertion Patterns
 * 
 * Examples from docs/05-recipes/entity-insertion-patterns.md
 */

import Word from "@types/office-js";
import { safeWordRun } from "./utils";

export interface StandardEntity {
  id: string;
  name: string;
  category: 'legal' | 'boilerplate' | 'signature' | 'header' | 'footer';
  content: string;
  version: string;
  lastUpdated: Date;
}

export interface EntityWithFields {
  entityId: string;
  fieldValues: { [placeholder: string]: string };
}

/**
 * Entity library for document automation
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
        content: 'Both parties agree to maintain the confidentiality of all proprietary information disclosed during the term of this agreement.',
        version: '1.0',
        lastUpdated: new Date('2024-01-01')
      },
      {
        id: 'governing-law',
        name: 'Governing Law',
        category: 'legal',
        content: 'This Agreement shall be governed by and construed in accordance with the laws of the State of [STATE].',
        version: '2.1',
        lastUpdated: new Date('2024-02-15')
      },
      {
        id: 'company-signature',
        name: 'Company Signature Block',
        category: 'signature',
        content: 'ACME CORPORATION\n\nBy: _______________________\nName: [AUTHORIZED_NAME]\nTitle: [AUTHORIZED_TITLE]',
        version: '1.0',
        lastUpdated: new Date('2024-01-01')
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
}

export const entityLibrary = new EntityLibrary();

/**
 * Inserts a standard entity at the current selection
 */
export async function insertEntity(entityId: string): Promise<void> {
  const entity = entityLibrary.getEntity(entityId);
  
  if (!entity) {
    throw new Error(`Entity not found: ${entityId}`);
  }

  await safeWordRun(async (context) => {
    const range = context.document.getSelection();
    const inserted = range.insertText(entity.content, Word.InsertLocation.replace);
    
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

/**
 * Inserts an entity with replaceable field placeholders
 */
export async function insertEntityWithFields(
  entityWithFields: EntityWithFields
): Promise<void> {
  const entity = entityLibrary.getEntity(entityWithFields.entityId);
  
  if (!entity) {
    throw new Error(`Entity not found: ${entityWithFields.entityId}`);
  }

  await safeWordRun(async (context) => {
    const range = context.document.getSelection();
    
    let content = entity.content;
    for (const [placeholder, value] of Object.entries(entityWithFields.fieldValues)) {
      content = content.replace(
        new RegExp(`\\[${placeholder}\\]`, 'g'),
        value
      );
    }
    
    range.insertText(content, Word.InsertLocation.replace);
    await context.sync();
  });
}

/**
 * Creates content controls that act as insertion points for entities
 */
export async function createEntityInsertionPoint(
  entityType: string,
  label: string
): Promise<number> {
  return await safeWordRun(async (context) => {
    const range = context.document.getSelection();
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
  return await safeWordRun(async (context) => {
    const resolved: string[] = [];
    
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    for (const control of contentControls.items) {
      control.load("tag");
    }
    await context.sync();
    
    for (const control of contentControls.items) {
      if (control.tag?.startsWith("entity:")) {
        const entityType = control.tag.replace("entity:", "");
        const entity = entityLibrary.getEntity(entityType);
        
        if (entity) {
          control.insertText(entity.content, Word.InsertLocation.replace);
          control.tag = `resolved:${entityType}`;
          resolved.push(entityType);
        }
      }
    }
    
    await context.sync();
    return resolved;
  });
}
