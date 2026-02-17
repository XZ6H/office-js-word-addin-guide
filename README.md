# Office JS Word Add-ins - Complete Guide

A comprehensive documentation project for building enterprise-grade Word Add-ins using Office.js.

## 📚 Documentation Structure

### Core Guides

1. **[Getting Started](docs/01-getting-started.md)** — Environment setup, yo office generator, manifest structure, sideloading
2. **[Architecture Best Practices](docs/02-architecture-best-practices.md)** — Separation of concerns, state management, error handling, security
3. **[Performance Optimization](docs/03-performance-optimization.md)** — Batch operations, minimizing context.sync(), memory management
4. **[API Deep Dive](docs/04-api-deep-dive.md)** — Document manipulation, ranges, tables, content controls, custom XML
5. **[Advanced Patterns](docs/06-advanced-patterns.md)** — Ribbon customization, Dialog API, SSO

### Common Recipes

| Recipe | Description |
|--------|-------------|
| [Find and Replace](docs/05-recipes/find-and-replace.md) | Text search and replacement patterns |
| [Table Generation](docs/05-recipes/table-generation.md) | Creating and manipulating tables |
| [Document Assembly](docs/05-recipes/document-assembly.md) | Combining document sections |
| [Formatting Automation](docs/05-recipes/formatting-automation.md) | Styles, fonts, and paragraph formatting |
| [Content Control Management](docs/05-recipes/content-control-management.md) | Forms and structured documents |
| [Export to PDF](docs/05-recipes/export-to-pdf.md) | Document export workflows |
| [Mail Merge](docs/05-recipes/mail-merge.md) | Bulk document generation |
| [Validation Checking](docs/05-recipes/validation-checking.md) | Document quality assurance |

### Advanced Recipes

| Recipe | Description |
|--------|-------------|
| [Template-Based Document Generation](docs/05-recipes/template-based-document-generation.md) | Enterprise document templates with content controls and OOXML |
| [Entity Insertion Patterns](docs/05-recipes/entity-insertion-patterns.md) | Fixed entities, standard clauses, reusable content blocks |
| [Data Binding Patterns](docs/05-recipes/data-binding-patterns.md) | Two-way binding with Excel, SharePoint, REST APIs |

## 🚀 Quick Start

```bash
# Install Yeoman and Office generator
npm install -g yo generator-office

# Generate new project
yo office --projectType task-pane --name "MyWordAddin" --host word --ts true

# Install dependencies
cd MyWordAddin && npm install

# Start development
npm run dev-server
```

## 💻 TypeScript Examples

Working examples in `examples/src/`:
- `utils.ts` — Core utilities and error handling
- `find-and-replace.ts` — Text replacement operations
- `table-generation.ts` — Dynamic table creation
- `content-control-management.ts` — Form handling
- `validation-checking.ts` — Document validation

Run tests:
```bash
cd examples && npm test
```

## 📖 Key Features

- **Production-ready TypeScript** with proper typing
- **Comprehensive error handling** in all examples
- **Performance optimized** — batch operations before sync
- **Enterprise patterns** — SSO, Dialog API, data binding
- **Real-world recipes** — Based on actual Microsoft documentation

## 🔗 References

- [Microsoft Office JS Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/word/)
- [OfficeDev GitHub Samples](https://github.com/OfficeDev/Office-Add-in-samples)
- [Word JavaScript API Reference](https://learn.microsoft.com/en-us/javascript/api/word)

## 📄 License

MIT — See LICENSE for details
