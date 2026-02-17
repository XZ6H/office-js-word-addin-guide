# Performance Optimization

## The Golden Rule: Batch Before Sync

Every `context.sync()` is a round-trip to the document. Minimize these calls by batching all operations first.

### ❌ Bad: Multiple Syncs
```typescript
// DON'T DO THIS - 5 round trips!
await Word.run(async (context) => {
  const body = context.document.body;
  
  body.insertText('First', Word.InsertLocation.end);
  await context.sync(); // Round trip 1
  
  body.insertText('Second', Word.InsertLocation.end);
  await context.sync(); // Round trip 2
  
  body.insertText('Third', Word.InsertLocation.end);
  await context.sync(); // Round trip 3
  
  body.font.name = 'Arial';
  await context.sync(); // Round trip 4
  
  body.font.size = 12;
  await context.sync(); // Round trip 5
});
```

### ✅ Good: Single Sync
```typescript
// DO THIS - 1 round trip!
await Word.run(async (context) => {
  const body = context.document.body;
  
  body.insertText('First', Word.InsertLocation.end);
  body.insertText('Second', Word.InsertLocation.end);
  body.insertText('Third', Word.InsertLocation.end);
  body.font.name = 'Arial';
  body.font.size = 12;
  
  await context.sync(); // Single round trip
});
```

## Load Only What You Need

### ❌ Bad: Loading Everything
```typescript
// Loads ALL properties - slow!
const range = context.document.getSelection();
range.load();
await context.sync();
```

### ✅ Good: Selective Loading
```typescript
// Load only required properties - fast!
const range = context.document.getSelection();
range.load('text, font/name, font/size, font/bold');
await context.sync();

console.log(range.text);
console.log(range.font.name);
```

## Working with Ranges Efficiently

### Range Caching Pattern
```typescript
export class RangeManager {
  private cachedRanges: Map<string, Word.Range> = new Map();

  async getOrCreateRange(
    context: Word.RequestContext,
    key: string,
    createFn: () => Word.Range
  ): Promise<Word.Range> {
    if (!this.cachedRanges.has(key)) {
      const range = createFn();
      range.load('text');
      this.cachedRanges.set(key, range);
    }
    return this.cachedRanges.get(key)!;
  }

  clearCache(): void {
    this.cachedRanges.clear();
  }
}
```

### Efficient Range Traversal
```typescript
// Process paragraphs in batches
async function processParagraphsEfficiently(
  context: Word.RequestContext,
  batchSize: number = 50
): Promise<void> {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load('items');
  await context.sync();

  const totalParagraphs = paragraphs.items.length;

  for (let i = 0; i < totalParagraphs; i += batchSize) {
    const batch = paragraphs.items.slice(i, i + batchSize);

    // Load properties for this batch
    batch.forEach(p => p.load('text, font/name'));
    await context.sync();

    // Process batch
    batch.forEach(p => {
      if (p.text.includes('TODO')) {
        p.font.highlightColor = 'yellow';
      }
    });

    // Sync batch changes
    await context.sync();
  }
}
```

## Memory Management

### Clean Up References
```typescript
async function processWithCleanup(): Promise<void> {
  await Word.run(async (context) => {
    try {
      const tempRange = context.document.body.insertText(
        'Temp text',
        Word.InsertLocation.end
      );
      
      // Do work...
      
      // Clean up if no longer needed
      tempRange.delete();
      await context.sync();
    } finally {
      // Ensure cleanup happens
    }
  });
}
```

### Large Document Handling
```typescript
export class LargeDocumentProcessor {
  private chunkSize: number = 100;

  async processLargeDocument(
    onProgress: (current: number, total: number) => void
  ): Promise<void> {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load('items/$count');
      await context.sync();

      const total = paragraphs.items.length;

      for (let i = 0; i < total; i += this.chunkSize) {
        const chunk = paragraphs.items.slice(i, Math.min(i + this.chunkSize, total));
        
        // Process chunk
        await this.processChunk(context, chunk);
        
        // Report progress
        onProgress(i + chunk.length, total);
        
        // Allow UI to update
        await new Promise(resolve => setTimeout(resolve, 0));
      }
    });
  }

  private async processChunk(
    context: Word.RequestContext,
    paragraphs: Word.Paragraph[]
  ): Promise<void> {
    paragraphs.forEach(p => {
      p.load('text, font/name, font/size');
    });
    await context.sync();

    // Apply changes
    paragraphs.forEach(p => {
      // Transformations...
    });
    await context.sync();
  }
}
```

## Caching Strategies

### Document Property Cache
```typescript
export class DocumentCache {
  private cache: Map<string, any> = new Map();
  private cacheExpiry: Map<string, number> = new Map();
  private readonly CACHE_TTL = 30000; // 30 seconds

  async getCachedProperty<T>(
    key: string,
    fetchFn: () => Promise<T>
  ): Promise<T> {
    const now = Date.now();
    const expiry = this.cacheExpiry.get(key);

    if (this.cache.has(key) && expiry && now < expiry) {
      return this.cache.get(key);
    }

    const value = await fetchFn();
    this.cache.set(key, value);
    this.cacheExpiry.set(key, now + this.CACHE_TTL);
    return value;
  }

  invalidate(key?: string): void {
    if (key) {
      this.cache.delete(key);
      this.cacheExpiry.delete(key);
    } else {
      this.cache.clear();
      this.cacheExpiry.clear();
    }
  }
}
```

## Async/Await Patterns

### Parallel Operations (When Safe)
```typescript
async function loadMultipleRanges(
  context: Word.RequestContext
): Promise<void> {
  const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
  const footer = context.document.sections.getFirst().getFooter(Word.HeaderFooterType.primary);
  const body = context.document.body;

  // Queue all loads
  header.load('text');
  footer.load('text');
  body.load('text');

  // Single sync for all
  await context.sync();

  // Now all are available
  console.log(header.text);
  console.log(footer.text);
  console.log(body.text);
}
```

### Sequential When Dependencies Exist
```typescript
async function dependentOperations(
  context: Word.RequestContext
): Promise<void> {
  // First operation
  const range = context.document.getSelection();
  range.load('text');
  await context.sync();

  // Use result for second operation
  if (range.text.length > 0) {
    const nextRange = range.getRange(Word.RangeLocation.end);
    nextRange.insertText('Appendix', Word.InsertLocation.after);
    await context.sync();
  }
}
```

## Performance Benchmarks

| Operation | Target Time |
|-----------|-------------|
| Initial load | < 1s |
| Simple insert | < 100ms |
| Format 1 page | < 200ms |
| Format 10 pages | < 1s |
| Large doc (100 pages) | < 5s |
| Search entire document | < 2s |

## Optimization Checklist

- [ ] Batch all operations before single `context.sync()`
- [ ] Use selective `.load()` with specific properties
- [ ] Process large documents in chunks
- [ ] Cache frequently accessed properties
- [ ] Clean up temporary ranges
- [ ] Use progress callbacks for long operations
- [ ] Minimize round trips to the document
- [ ] Load `$count` instead of full collections when possible
