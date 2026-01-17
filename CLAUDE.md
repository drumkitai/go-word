# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

go-word is a high-performance Golang library for creating, reading, and modifying Word documents (.docx). It follows the Office Open XML (OOXML) specifications and provides a clean, fluent API with excellent performance (average 2.62ms processing time, 21x faster than Python).

**Repository:** https://github.com/drumkitai/go-word

## Development Commands

### Building
```bash
# Build all packages
go build ./...

# Build specific package
go build ./pkg/document
go build ./pkg/style
go build ./pkg/markdown
```

### Testing
```bash
# Run all tests
go test ./...

# Run tests for specific packages
go test ./pkg/document ./pkg/style ./test

# Run tests with coverage
go test -cover ./...

# Run short tests (skip long-running tests)
go test -short ./...

# Verbose test output
go test -v ./pkg/...

# Run specific test
go test -run TestFunctionName ./pkg/document
```

### Code Quality
```bash
# Format code (always run before committing)
go fmt ./...

# Vet code for issues
go vet ./...

# Run linters (if golangci-lint is installed)
golangci-lint run
```

### Running Examples
```bash
# Run basic example
go run ./examples/basic/

# Run any specific example
go run ./examples/style_demo/
go run ./examples/table/
go run ./examples/template_inheritance_demo/
go run ./examples/markdown_conversion/
```

## Architecture Overview

### Core Packages

**`pkg/document/`** - Core document manipulation
- `document.go`: Main Document type, paragraph operations, document lifecycle (New/Open/Save)
- `table.go`: Complete table operations including cell merging, styling, and iteration
- `template_engine.go`: Template rendering engine with variable substitution, conditionals, loops, and template inheritance
- `image.go`: Image embedding with size/position control and floating images
- `header_footer.go`: Header/footer management with template variable support
- `page.go`: Page settings (size, margins, orientation)
- `toc.go`: Table of contents generation
- `footnotes.go`: Footnotes and endnotes
- `numbering.go`: List numbering (ordered/unordered)
- `field.go`: Complex field operations

**`pkg/style/`** - Style management system
- `style.go`: Core style types and XML structures
- `predefined.go`: 18 predefined styles (Normal, Heading1-9, Title, Subtitle, etc.)
- `api.go`: StyleManager for centralized style handling with inheritance support

**`pkg/markdown/`** - Markdown ↔ Word conversion
- `converter.go`: Main converter using goldmark parser
- `renderer.go`: Goldmark AST → Word document renderer
- `exporter.go`: Word document → Markdown exporter
- `math.go`: Math formula support (LaTeX syntax)

### Key Design Patterns

**Document Structure:**
- `Document` contains a `Body` with `Elements []interface{}`
- Elements can be `*Paragraph` or `*Table` (implement `BodyElement` interface)
- XML serialization uses custom `MarshalXML` to maintain OOXML element ordering

**Fluent Interface:**
```go
doc.AddParagraph("Text").
    SetAlignment(document.AlignCenter).
    SetSpacing(&document.SpacingConfig{LineSpacing: 1.5}).
    SetStyle("Heading1")
```

**Template System:**
- Variable substitution: `{{variable}}`
- Conditionals: `{{#if condition}}...{{/if}}`
- Loops: `{{#each list}}...{{/each}}`
- Template inheritance: `{{extends "base"}}` with `{{#block name}}...{{/block}}`
- Image placeholders: `{{#image imageName}}`

**OOXML Compliance:**
- Uses standard namespace prefixes: `w:` (main), `r:` (relationships), `a:` (drawing)
- Measurements in twips (1 point = 20 twips)
- Strict element ordering enforced via custom XML marshaling

### Document Relationships
- `relationships`: Main document relationships (images, hyperlinks)
- `documentRelationships`: Document-level relationships (headers, footers)
- `contentTypes`: Content type definitions for ZIP package
- `parts`: In-memory storage for document parts during construction

### Style Inheritance
The style system supports inheritance through `BasedOn` relationships:
- StyleManager resolves style chains
- Custom styles can inherit from predefined styles
- Paragraph and character styles are managed separately

## Important Code Conventions

### Language Requirements
- All code, comments, and documentation should be in English
- Follow standard Go documentation conventions
- Use clear, descriptive variable and function names

### Error Handling
- Use custom error types from `errors.go`
- Wrap errors with context: `WrapError()` or `WrapErrorWithContext()`
- Return meaningful error messages

### XML Struct Tags
OOXML serialization requires proper XML tags:
```go
type Paragraph struct {
    XMLName    xml.Name             `xml:"w:p"`
    Properties *ParagraphProperties `xml:"w:pPr,omitempty"`
    Runs       []Run                `xml:"w:r"`
}
```

### Test Organization
- Unit tests in `pkg/*/` directories alongside source files
- Integration tests in `test/` directory
- Test output files go to `test_output/` directory
- Always clean up test files: `defer os.RemoveAll("test_output")`

### Element Ordering
OOXML requires specific XML element ordering. When this matters:
- Implement custom `MarshalXML` methods (see `Body.MarshalXML`)
- `SectionProperties` must be last in `<w:body>`
- Style properties have strict ordering (see `ParagraphProperties`, `RunProperties`)

## Common Operations

### Creating Documents
```go
doc := document.New()
```

### Opening Existing Documents
```go
doc, err := document.Open("path/to/file.docx")
// Also: OpenFromMemory(data []byte)
```

### Adding Content
```go
// Paragraphs
para := doc.AddParagraph("text")
doc.AddHeadingParagraph("title", 1) // level 1-9

// Tables
table := doc.AddTable(&document.TableConfig{Rows: 3, Columns: 3})
table.SetCellText(row, col, "content")

// Images
doc.AddImage("path.png", &document.ImageConfig{Width: 200, Height: 150})

// Page breaks
doc.AddPageBreak()
```

### Template Rendering
```go
engine := document.NewTemplateEngine()
engine.LoadTemplate("name", templateContent)
data := document.NewTemplateData()
data.SetVariable("key", "value")
doc, err := engine.RenderTemplateToDocument("name", data)
```

### Markdown Conversion
```go
converter := markdown.NewConverter(markdown.DefaultOptions())
doc, err := converter.ConvertString(markdownText, nil)
```

## Critical Notes

> **📖 For comprehensive best practices, error handling, performance optimization, and solutions to common problems, see [BEST_PRACTICES.md](BEST_PRACTICES.md)**

### Compatibility
- **Minimum Go version: 1.19** (specified in go.mod)
- Avoid Go 1.22+ syntax (like range-over-int) for backward compatibility

### Dependencies
- Minimal dependencies policy: only goldmark for Markdown support
- Standard library preferred for all other operations

### Performance
- Zero-allocation optimizations where possible
- Lazy loading of document parts
- Streaming approaches for large documents

### Known Issues to Avoid
1. Never modify `Body.Elements` slice directly - use provided methods
2. Always check table cell bounds before accessing
3. Preserve XML element ordering in custom marshaling
4. Clean up test files in test functions
5. Never skip `go fmt` before committing

### Diagnostic Errors
The project currently has multiple example files with `main` redeclared errors:
- `examples/markdown_demo/soft_linebreak_demo.go`
- `examples/markdown_demo/math_formula_demo.go`
- `examples/markdown_demo/table_and_tasklist_demo.go`

These are examples meant to be run individually, not simultaneously. Run them with `go run examples/markdown_demo/soft_linebreak_demo.go` (one at a time).
