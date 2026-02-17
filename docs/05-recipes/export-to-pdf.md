# Export to PDF

Exporting Word documents to PDF programmatically is a common requirement for document workflows, archiving, and sharing.

## Overview

The Office JS API provides the `File` API for document export. For PDF specifically, you use the `PDF` file type.

## Basic PDF Export

```typescript
async function exportToPDF(): Promise<void> {
  await Word.run(async (context) => {
    // Get the document as PDF
    const pdfFile = context.document.saveAsPDF();
    pdfFile.load("content");
    await context.sync();
    
    // Convert to base64 for download or upload
    const base64Content = pdfFile.content;
    
    // Trigger download
    const blob = base64ToBlob(base64Content, "application/pdf");
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement("a");
    a.href = url;
    a.download = "document.pdf";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
}

function base64ToBlob(base64: string, contentType: string): Blob {
  const byteCharacters = atob(base64);
  const byteNumbers = new Array(byteCharacters.length);
  
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: contentType });
}
```

## Export with Options

```typescript
interface PDFExportOptions {
  filename?: string;
  includeMarkup?: boolean;
  includeComments?: boolean;
  optimizeFor?: 'screen' | 'print' | 'archive';
}

async function exportToPDFWithOptions(options: PDFExportOptions): Promise<void> {
  await Word.run(async (context) => {
    const filename = options.filename || "document.pdf";
    
    // Note: Some options require specific Word versions
    // Basic export works across all versions
    const pdfFile = context.document.saveAsPDF();
    pdfFile.load("content");
    await context.sync();
    
    // Download with custom filename
    downloadPDF(pdfFile.content, filename);
  });
}

function downloadPDF(base64Content: string, filename: string): void {
  const blob = base64ToBlob(base64Content, "application/pdf");
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
```

## Server-Side Export

For production scenarios, you might want to export on the server:

```typescript
// Client-side: Get document content
async function getDocumentForExport(): Promise<string> {
  return await Word.run(async (context) => {
    const doc = context.document;
    const body = doc.body;
    body.load("text");
    await context.sync();
    return body.text;
  });
}

// Server-side: Convert to PDF (pseudo-code)
// You would use a library like Puppeteer, Playwright, or a PDF service
async function serverSideExport(documentContent: string): Promise<Buffer> {
  // Implementation depends on your server stack
  // Example using a hypothetical PDF service:
  const response = await fetch('https://api.pdfservice.com/convert', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ content: documentContent, format: 'pdf' })
  });
  
  return await response.arrayBuffer();
}
```

## Best Practices

1. **Always provide filename** — Don't rely on default names
2. **Handle large documents** — Consider chunking for very large exports
3. **Show progress** — Export can take time; show a spinner
4. **Error handling** — Wrap in try-catch with user-friendly messages
5. **Memory management** — Revoke object URLs after download

## Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| Export fails silently | Document too large | Split into sections |
| Corrupted PDF | Encoding issues | Ensure proper base64 handling |
| Slow export | Complex formatting | Simplify document first |
| Download blocked | Popup blocker | Use direct download or notify user |
