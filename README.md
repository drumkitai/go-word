<div align="center">  
  <h2>go-word | Golang Word Document Library</h1>
</div>

<div align="center">
  
[![Go Version](https://img.shields.io/badge/Go-1.19+-00ADD8?style=flat&logo=go)](https://golang.org)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

</div>

go-word is a library for Microsoft Word Document manipulation that supports document parsing, creation, and 
modification. This library follows the latest Office Open XML (OOXML) specifications and focuses on supporting 
modern Word document files (.docx).

This project is forked from [wordZero](https://github.com/zerx-lab/wordZero) (MIT License) and translated to English
for use in Drumkit's codebase for document processing pipelines (currently used to convert word documents to markdown).

### Core Features

- 🚀 **Complete Document Operations**: Create, read, and modify Word documents
- 🖨️ **Markdown Conversion**: Convert to and from markdown
- 🎨 **Rich Style System**: 18 predefined styles with custom style and inheritance support
- 📝 **Text Formatting**: Full support for fonts, sizes, colors, bold, italic, and more
- 📐 **Paragraph Format**: Alignment, spacing, indentation, and other paragraph properties
- 🏷️ **Heading Navigation**: Complete support for Heading1-9 styles, recognizable by Word navigation pane
- 📊 **Table Functionality**: Complete table creation, editing, styling, and iterator support
- 📄 **Page Settings**: Page size, margins, headers/footers, and professional layout features
- 🔧 **Advanced Features**: Table of contents generation, footnotes/endnotes, list numbering, template engine, etc.
- 🎯 **Template Inheritance**: Support for base templates and block override mechanisms for template reuse and extension
- 📝 **Header/Footer Templates**: Support for template variables in headers and footers for dynamic content replacement
- ⚡ **Excellent Performance**: Zero-dependency pure Go implementation, average 2.62ms processing speed, 3.7x faster than JavaScript, 21x faster than Python
- 🔧 **Easy to Use**: Clean API design with fluent interface support

## Related Projects

### Excel Document Operations - Excelize

If you need to work with Excel documents, we recommend [**Excelize**](https://github.com/qax-os/excelize) —— the most popular Go library for Excel operations:

- ⭐ **19.2k+ GitHub Stars** - The most popular Excel processing library in the Go ecosystem
- 📊 **Complete Excel Support** - Supports all modern Excel formats including XLAM/XLSM/XLSX/XLTM/XLTX
- 🎯 **Feature Rich** - Charts, pivot tables, images, streaming APIs, and more
- 🚀 **High Performance** - Streaming read/write APIs optimized for large datasets
- 🔧 **Easy Integration** - Perfect complement to go-word for complete Office document processing solutions

## Installation

```bash
go get github.com/drumkitai/go-word
```

### Version Notes

We recommend using a pinned version installation:

```bash
# Install specific version
go get github.com/drumkitai/go-word@v1.0.1

# Install latest version
go get github.com/drumkitai/go-word@latest
```

### Word to Markdown Feature Example

```go
package main

import (
    "fmt"
    "log"
    "github.com/drumkitai/go-word/pkg/document"
    "github.com/drumkitai/go-word/pkg/markdown"
)

func main() {
    // Create Word to Markdown exporter
    exporter := markdown.NewExporter(markdown.DefaultExportOptions())
    
    // Convert Word document to Markdown file
    err := exporter.ExportToFile("document.docx", "output.md", nil)
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("Word document successfully converted to Markdown!")
    
    // Alternative: Export from an already opened document
    doc, err := document.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    markdownText, err := exporter.ExportToString(doc, nil)
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("Markdown content:")
    fmt.Println(markdownText)
}
```

## Quick Start (Creating a word document)

```go
package main

import (
    "log"
    "github.com/drumkitai/go-word/pkg/document"
    "github.com/drumkitai/go-word/pkg/style"
)

func main() {
    // Create new document
    doc := document.New()
    
    // Add title
    title := doc.AddParagraph("Hello World")
    title.SetStyle(style.StyleHeading1)
    
    // Add body paragraph
    para := doc.AddParagraph("This is a document example created using go-word.")
    para.SetFontFamily("Arial")
    para.SetFontSize(12)
    para.SetColor("333333")
    
    // Create table
    tableConfig := &document.TableConfig{
        Rows:    3,
        Columns: 3,
    }
    table := doc.AddTable(tableConfig)
    table.SetCellText(0, 0, "Header1")
    table.SetCellText(0, 1, "Header2")
    table.SetCellText(0, 2, "Header3")
    
    if err := doc.Save("example.docx"); err != nil {
        log.Fatal(err)
    }
}
```

### Template Inheritance Feature Example

```go
// Create base template
engine := document.NewTemplateEngine()
baseTemplate := `{{companyName}} Work Report

{{#block "summary"}}
Default summary content
{{/block}}

{{#block "content"}}
Default main content
{{/block}}`

engine.LoadTemplate("base_report", baseTemplate)

// Create extended template, override specific blocks
salesTemplate := `{{extends "base_report"}}

{{#block "summary"}}
Sales Performance Summary: Achieved {{achievement}}% this month
{{/block}}

{{#block "content"}}
Sales Details:
- Total Sales: {{totalSales}}
- New Customers: {{newCustomers}}
{{/block}}`

engine.LoadTemplate("sales_report", salesTemplate)

// Render template
data := document.NewTemplateData()
data.SetVariable("companyName", "Drumkit")
data.SetVariable("achievement", "125")
data.SetVariable("totalSales", "1,000,000")
data.SetVariable("newCustomers", "47")

doc, _ := engine.RenderTemplateToDocument("sales_report", data)
doc.Save("sales_report.docx")
```

### Image Placeholder Template Feature Example

```go
package main

import (
    "log"
    "github.com/drumkitai/go-word/pkg/document"
)

func main() {
    // Create template with image placeholders
    engine := document.NewTemplateEngine()
    template := `Company: {{companyName}}

{{#image companyLogo}}

Project Report: {{projectName}}

Status: {{#if isCompleted}}Completed{{else}}In Progress{{/if}}

{{#image statusChart}}

Team Members:
{{#each teamMembers}}
- {{name}}: {{role}}
{{/each}}`

    engine.LoadTemplate("project_report", template)

    // Prepare template data
    data := document.NewTemplateData()
    data.SetVariable("companyName", "Drumkit")
    data.SetVariable("projectName", "Document Processing System")
    data.SetCondition("isCompleted", true)
    
    // Set team members list
    data.SetList("teamMembers", []interface{}{
        map[string]interface{}{"name": "Alice", "role": "Lead Developer"},
        map[string]interface{}{"name": "Bob", "role": "Frontend Developer"},
    })
    
    // Configure and set images
    logoConfig := &document.ImageConfig{
        Width:     100,
        Height:    50,
        Alignment: document.AlignCenter,
    }
    data.SetImage("companyLogo", "assets/logo.png", logoConfig)
    
    chartConfig := &document.ImageConfig{
        Width:       200,
        Height:      150,
        Alignment:   document.AlignCenter,
        AltText:     "Project Status Chart",
        Title:       "Current Project Status",
    }
    data.SetImage("statusChart", "assets/chart.png", chartConfig)
    
    // Render template to document
    doc, err := engine.RenderTemplateToDocument("project_report", data)
    if err != nil {
        log.Fatal(err)
    }
    
    // Save document
    err = doc.Save("project_report.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Markdown to Word Feature Example

```go
package main

import (
    "log"
    "github.com/drumkitai/go-word/pkg/markdown"
)

func main() {
    // Create Markdown converter
    converter := markdown.NewConverter(markdown.DefaultOptions())
    
    // Markdown content
    markdownText := `# go-word Markdown Conversion Example

Welcome to go-word's **Markdown to Word** conversion feature!

## Supported Syntax

### Text Formatting
- **Bold text**
- *Italic text*
- ` + "`Inline code`" + `

### Lists
1. Ordered list item 1
2. Ordered list item 2

- Unordered list item A
- Unordered list item B

### Quotes and Code

> This is blockquote content
> Supporting multiple lines

` + "```" + `go
// Code block example
func main() {
    fmt.Println("Hello, World!")
}
` + "```" + `

---

Conversion complete!`

    // Convert to Word document
    doc, err := converter.ConvertString(markdownText, nil)
    if err != nil {
        log.Fatal(err)
    }
    
    // Save Word document
    err = doc.Save("markdown_example.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    // File conversion
    err = converter.ConvertFile("input.md", "output.docx", nil)
    if err != nil {
        log.Fatal(err)
    }
}
```

### Document Pagination and Paragraph Deletion Example

```go
package main

import (
    "log"
    "github.com/drumkitai/go-word/pkg/document"
)

func main() {
    doc := document.New()
    
    // Add first page content
    doc.AddHeadingParagraph("Chapter 1: Introduction", 1)
    doc.AddParagraph("This is the content of chapter 1.")
    
    // Add page break to start a new page
    doc.AddPageBreak()
    
    // Add second page content
    doc.AddHeadingParagraph("Chapter 2: Main Content", 1)
    tempPara := doc.AddParagraph("This is a temporary paragraph.")
    doc.AddParagraph("This is the content of chapter 2.")
    
    // Delete temporary paragraph
    doc.RemoveParagraph(tempPara)
    
    // You can also delete by index
    // doc.RemoveParagraphAt(1)  // Delete second paragraph
    
    // Save document
    if err := doc.Save("example.docx"); err != nil {
        log.Fatal(err)
    }
}
```

## Main Features

### ✅ Implemented Features
- **Document Operations**: Create, read, save, parse DOCX documents
- **Text Formatting**: Fonts, sizes, colors, bold, italic, etc.
- **Style System**: 18 predefined styles + custom style support
- **Paragraph Format**: Alignment, spacing, indentation, complete support
- **Paragraph Management**: Paragraph deletion, deletion by index, element removal
- **Document Pagination**: Page break insertion for multi-page document structure
- **Table Functionality**: Complete table operations, styling, cell iterators
- **Page Settings**: Page size, margins, headers/footers, etc.
- **Advanced Features**: Table of contents generation, footnotes/endnotes, list numbering, template engine (with template inheritance)
- **Image Features**: Image insertion, size adjustment, position setting
- **Markdown to Word**: High-quality Markdown to Word conversion based on goldmark

## Project Structure

```
go-word/
├── pkg/                    # Core library code
│   ├── document/          # Document operation features
│   └── style/             # Style management system
├── test/                  # Integration tests
```

</div>

## Contributing

Issues and Pull Requests are welcome! Please ensure before submitting code:

1. Code follows Go coding standards
2. Add necessary test cases
3. Update relevant documentation
4. Ensure all tests pass

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Attribution

This project is derived from [wordZero](https://github.com/zerx-lab/wordZero) by ZeroHawkeye, 
with modifications made initially by [Winston-Hsiao](https://github.com/Winston-Hsiao) and other future contributors.

---

**More Resources**
- 📝 [Changelog](CHANGELOG.md)

Links to Original Project documentation (WordZero—which this repo is forked from):
- 📖 [Complete Documentation](https://github.com/zerx-lab/wordZero/wiki)
- 🔧 [API Reference](https://github.com/zerx-lab/wordZero/wiki/en-API-Reference)
- 💡 [Best Practices](https://github.com/zerx-lab/wordZero/wiki/en-Best-Practices)

Note: go-word's functionality/features may be out of date relative to WordZero's project state.