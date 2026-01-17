# Document Package API Documentation

This document records all available public methods and functionality in the `pkg/document` package.

## Core Types

### Document
- [Document](document.go) - Core structure of a Word document
- [Body](document.go) - Document body
- [Paragraph](document.go) - Paragraph structure
- [Table](table.go) - Table structure

## Document Operation Methods

### Document Creation and Loading
- [`New()`](document.go#L232) - Create a new Word document
- [`Open(filename string)`](document.go#L269) - Open an existing Word document ✨ **Major Improvement**

#### Major Upgrade in Document Parsing Functionality ✨
The `Open` method now supports complete document structure parsing, including:

**Dynamic Element Parsing Support**:
- **Paragraph Parsing** (`<w:p>`): Complete parsing of paragraph content, attributes, runs, and formatting
- **Table Parsing** (`<w:tbl>`): Support for table structure, grids, rows, columns, and cell content
- **Section Properties Parsing** (`<w:sectPr>`): Page settings, margins, columns, and other properties
- **Extensible Design**: New parsing architecture can easily add more element types

**Parser Features**:
- **Stream Parsing**: Uses XML stream parser for high memory efficiency, suitable for large documents
- **Structure Preservation**: Complete preservation of original order and hierarchy of document elements
- **Error Recovery**: Intelligently skips unknown or corrupted elements to ensure stable parsing
- **Deep Parsing**: Supports nested structures (e.g., paragraphs in tables, runs in paragraphs, etc.)

**Parsed Content Includes**:
- Paragraph text content and all formatting properties (font, size, color, style, etc.)
- Complete table structure (row/column definitions, cell content, table properties)
- Page setting information (page size, orientation, margins, etc.)
- Style references and property inheritance relationships

### Document Saving and Export
- [`Save(filename string)`](document.go#L337) - Save document to file
- [`ToBytes()`](document.go#L1107) - Convert document to byte array

### Document Content Operations
- [`AddParagraph(text string)`](document.go#L420) - Add simple paragraph
- [`AddFormattedParagraph(text string, format *TextFormat)`](document.go#L459) - Add formatted paragraph
- [`AddHeadingParagraph(text string, level int)`](document.go#L682) - Add heading paragraph
- [`AddHeadingParagraphWithBookmark(text string, level int, bookmarkName string)`](document.go#L747) - Add heading paragraph with bookmark ✨ **New Feature**
- [`AddPageBreak()`](document.go#L1185) - Add page break

#### Page Break Functionality ✨

go-word provides multiple ways to add page breaks (page breaks):

**Method 1: Document-Level Page Break**
```go
doc := document.New()
doc.AddParagraph("First page content")
doc.AddPageBreak()  // Add page break
doc.AddParagraph("Second page content")
```

**Method 2: In-Paragraph Page Break**
```go
para := doc.AddParagraph("First page content")
para.AddPageBreak()  // Add page break within paragraph
para.AddFormattedText("Second page content", nil)
```

**Method 3: Page Break Before Paragraph**
```go
para := doc.AddParagraph("Chapter 2 Title")
para.SetPageBreakBefore(true)  // Set auto page break before paragraph
```

**Page Break Features**:
- **Independent Page Break**: `Document.AddPageBreak()` creates independent page break paragraphs
- **In-Paragraph Page Break**: `Paragraph.AddPageBreak()` adds page break within current paragraph
- **Page Break Before**: `Paragraph.SetPageBreakBefore(true)` sets auto page break before paragraph
- **Table Page Break Control**: Support for table page break control settings

#### Heading Paragraph Bookmark Functionality ✨
The `AddHeadingParagraphWithBookmark` method now supports adding bookmarks to heading paragraphs:

**Bookmark Features**:
- **Auto Bookmark Generation**: Create unique bookmark identifiers for heading paragraphs
- **Flexible Naming**: Support custom bookmark names or leave empty to skip bookmark creation
- **TOC Compatible**: Generated bookmarks perfectly compatible with TOC functionality, supporting navigation and hyperlinks
- **Word Standard**: Complies with Microsoft Word bookmark format specifications

**Bookmark Generation Rules**:
- Bookmark IDs auto-generated as `bookmark_{element_index}_{bookmark_name}` format
- Bookmark start marker inserted before paragraph
- Bookmark end marker inserted after paragraph
- Support empty bookmark names to skip bookmark creation

### Style Management
- [`GetStyleManager()`](document.go#L791) - Get style manager

### Page Settings ✨ New Features
- [`SetPageSettings(settings *PageSettings)`](page.go) - Set complete page properties
- [`GetPageSettings()`](page.go) - Get current page settings
- [`SetPageSize(size PageSize)`](page.go) - Set page size
- [`SetCustomPageSize(width, height float64)`](page.go) - Set custom page size (millimeters)
- [`SetPageOrientation(orientation PageOrientation)`](page.go) - Set page orientation
- [`SetPageMargins(top, right, bottom, left float64)`](page.go) - Set page margins (millimeters)
- [`SetHeaderFooterDistance(header, footer float64)`](page.go) - Set header/footer distance (millimeters)
- [`SetGutterWidth(width float64)`](page.go) - Set gutter width (millimeters)
- [`DefaultPageSettings()`](page.go) - Get default page settings (A4 portrait)

### Header and Footer Operations ✨ New Features
- [`AddHeader(headerType HeaderFooterType, text string)`](header_footer.go) - Add header
- [`AddFooter(footerType HeaderFooterType, text string)`](header_footer.go) - Add footer
- [`AddHeaderWithPageNumber(headerType HeaderFooterType, text string, showPageNum bool)`](header_footer.go) - Add header with page number
- [`AddFooterWithPageNumber(footerType HeaderFooterType, text string, showPageNum bool)`](header_footer.go) - Add footer with page number
- [`SetDifferentFirstPage(different bool)`](header_footer.go) - Set different first page

### Table of Contents Functionality ✨ New Features
- [`GenerateTOC(config *TOCConfig)`](toc.go) - Generate table of contents
- [`UpdateTOC()`](toc.go) - Update table of contents
- [`AddHeadingWithBookmark(text string, level int, bookmarkName string)`](toc.go) - Add heading with bookmark
- [`AutoGenerateTOC(config *TOCConfig)`](toc.go) - Auto-generate table of contents
- [`GetHeadingCount()`](toc.go) - Get heading count
- [`ListHeadings()`](toc.go) - List all headings
- [`SetTOCStyle(level int, style *TextFormat)`](toc.go) - Set table of contents style

### Footnotes and Endnotes Functionality ✨ New Features
- [`AddFootnote(text string, footnoteText string)`](footnotes.go) - Add footnote
- [`AddEndnote(text string, endnoteText string)`](footnotes.go) - Add endnote
- [`AddFootnoteToRun(run *Run, footnoteText string)`](footnotes.go) - Add footnote to run
- [`SetFootnoteConfig(config *FootnoteConfig)`](footnotes.go) - Set footnote configuration
- [`GetFootnoteCount()`](footnotes.go) - Get footnote count
- [`GetEndnoteCount()`](footnotes.go) - Get endnote count
- [`RemoveFootnote(footnoteID string)`](footnotes.go) - Remove footnote
- [`RemoveEndnote(endnoteID string)`](footnotes.go) - Remove endnote

### List and Numbering Functionality ✨ New Features
- [`AddListItem(text string, config *ListConfig)`](numbering.go) - Add list item
- [`AddBulletList(text string, level int, bulletType BulletType)`](numbering.go) - Add unordered list
- [`AddNumberedList(text string, level int, numType ListType)`](numbering.go) - Add ordered list
- [`CreateMultiLevelList(items []ListItem)`](numbering.go) - Create multi-level list
- [`RestartNumbering(numID string)`](numbering.go) - Restart numbering

### Structured Document Tags ✨ New Features
- [`CreateTOCSDT(title string, maxLevel int)`](sdt.go) - Create table of contents SDT structure

### Template Functionality ✨ New Features

#### Template Renderer (Recommended) ✨
- [`NewTemplateRenderer()`](template_engine.go) - Create new template renderer (recommended)
- [`SetLogging(enabled bool)`](template_engine.go) - Set logging
- [`LoadTemplateFromFile(name, filePath string)`](template_engine.go) - Load template from DOCX file
- [`RenderTemplate(templateName string, data *TemplateData)`](template_engine.go) - Render template (most recommended method)
- [`AnalyzeTemplate(templateName string)`](template_engine.go) - Analyze template structure

#### Template Engine (Low-Level API)
- [`NewTemplateEngine()`](template.go) - Create new template engine
- [`LoadTemplate(name, content string)`](template.go) - Load template from string
- [`LoadTemplateFromDocument(name string, doc *Document)`](template.go) - Create template from existing document
- [`GetTemplate(name string)`](template.go) - Get cached template
- [`RenderTemplateToDocument(templateName string, data *TemplateData)`](template.go) - Render template to new document (recommended method)
- [`RenderToDocument(templateName string, data *TemplateData)`](template.go) - Render template to new document (traditional method)
- [`ValidateTemplate(template *Template)`](template.go) - Validate template syntax
- [`ClearCache()`](template.go) - Clear template cache
- [`RemoveTemplate(name string)`](template.go) - Remove specific template

#### Template Engine Features ✨
**Variable Substitution**: Support `{{variable_name}}` syntax for dynamic content replacement
**Conditional Statements**: Support `{{#if condition}}...{{/if}}` syntax for conditional rendering
**Loop Statements**: Support `{{#each list}}...{{/each}}` syntax for list rendering
**Template Inheritance**: Support `{{extends "base_template"}}` syntax and `{{#block "block_name"}}...{{/block}}` block override mechanism for true template inheritance
  - **Block Definition**: Define rewritable content blocks in base template
  - **Block Override**: Selectively override specific blocks in child templates, unmodified blocks retain parent template default content
  - **Multi-Level Inheritance**: Support multi-level template inheritance relationships
  - **Complete Preservation**: Unmodified blocks completely preserve parent template's default content and formatting
**Loop with Conditions**: Perfect support for conditional expressions inside loops, such as `{{#each items}}{{#if isActive}}...{{/if}}{{/each}}`
**Data Type Support**: Support multiple data types including strings, numbers, booleans, and objects
**Struct Binding**: Support automatic template data generation from Go structs
**Template Analysis**: ✨ **New Feature** Auto-analyze template structure and extract variable, list, condition, and table information
  - **Structure Analysis**: Identify all variables, lists, and conditions used in template
  - **Table Analysis**: Specifically analyze template syntax and loop structure in tables
  - **Dependency Check**: Check template data dependencies
  - **Sample Data Generation**: Auto-generate sample data structures based on analysis results
**Logging**: ✨ **New Feature** Complete logging system supporting detailed records of template loading, rendering, and analysis processes
**Data Validation**: ✨ **New Feature** Auto-validate template data completeness and format correctness
**DOCX Template Support**: ✨ **New Feature** Load templates directly from existing DOCX files
**Header/Footer Template Support**: ✨ **New Feature** Complete support for template variables in headers and footers
  - **Variable Recognition**: Auto-identify `{{variable_name}}` syntax in headers and footers
  - **Variable Replacement**: Auto-replace template variables in headers and footers during rendering
  - **Conditional Statements**: Support conditional rendering in headers and footers
  - **Template Analysis**: `AnalyzeTemplate` auto-analyzes variables in headers and footers

### Template Data Operations
- [`NewTemplateData()`](template.go) - Create new template data
- [`SetVariable(name string, value interface{})`](template.go) - Set variable
- [`SetList(name string, list []interface{})`](template.go) - Set list
- [`SetCondition(name string, value bool)`](template.go) - Set condition
- [`SetVariables(variables map[string]interface{})`](template.go) - Batch set variables
- [`GetVariable(name string)`](template.go) - Get variable
- [`GetList(name string)`](template.go) - Get list
- [`GetCondition(name string)`](template.go) - Get condition
- [`Merge(other *TemplateData)`](template.go) - Merge template data
- [`Clear()`](template.go) - Clear template data
- [`FromStruct(data interface{})`](template.go) - Generate template data from struct

### Template Inheritance Detailed Usage Guide ✨ **New Feature**

Template inheritance is an advanced feature of the go-word template engine that allows creating extended templates based on existing templates, achieving template reuse and extension through block definition and override mechanisms.

#### Basic Syntax

**1. Base Template Block Definition**
```go
// Define base template with rewritable blocks
baseTemplate := `{{companyName}} Report

{{#block "header"}}
Default header content
Date: {{reportDate}}
{{/block}}

{{#block "summary"}}
Default summary content
{{/block}}

{{#block "main_content"}}
Default main content
{{/block}}

{{#block "footer"}}
Reporter: {{reporterName}}
{{/block}}`

engine.LoadTemplate("base_report", baseTemplate)
```

**2. Child Template Inheritance and Block Override**
```go
// Create child template inheriting from base template
childTemplate := `{{extends "base_report"}}

{{#block "summary"}}
Sales Performance Summary
This month's sales target has been achieved {{achievementRate}}%
{{/block}}

{{#block "main_content"}}
Detailed sales data:
- Total sales: {{totalSales}}
- New customers: {{newCustomers}}
- Completed orders: {{orders}}
{{/block}}`

engine.LoadTemplate("sales_report", childTemplate)
```

#### Inheritance Features

**Block Override Strategy**:
- Overridden block completely replaces corresponding block in parent template
- Unmodified blocks retain parent template's default content
- Support selective override for maximum flexibility

**Multi-Level Inheritance**:
```go
// Level 1: Base template
baseTemplate := `{{#block "document"}}Base document{{/block}}`

// Level 2: Business template
businessTemplate := `{{extends "base"}}
{{#block "document"}}
{{#block "business_header"}}Business header{{/block}}
{{#block "business_content"}}Business content{{/block}}
{{/block}}`

// Level 3: Specific business template
salesTemplate := `{{extends "business"}}
{{#block "business_header"}}Sales Report{{/block}}
{{#block "business_content"}}Sales data analysis{{/block}}`
```

#### Practical Application Example

```go
func demonstrateTemplateInheritance() {
    engine := document.NewTemplateEngine()

    // Base report template
    baseTemplate := `{{companyName}} Work Report
Report Date: {{reportDate}}

{{#block "summary"}}
Default summary content
{{/block}}

{{#block "main_content"}}
Default main content
{{/block}}

{{#block "conclusion"}}
Default conclusion
{{/block}}

{{#block "signature"}}
Reporter: {{reporterName}}
Department: {{department}}
{{/block}}`

    engine.LoadTemplate("base_report", baseTemplate)

    // Sales report template (override some blocks)
    salesTemplate := `{{extends "base_report"}}

{{#block "summary"}}
Sales Performance Summary
This month's sales target has been achieved {{achievementRate}}%
{{/block}}

{{#block "main_content"}}
Sales Data Analysis
- Total sales: {{totalSales}}
- New customers: {{newCustomers}}
- Completed orders: {{orders}}

{{#each channels}}
- {{name}}: {{sales}} ({{percentage}}%)
{{/each}}
{{/block}}`

    engine.LoadTemplate("sales_report", salesTemplate)

    // Prepare data and render
    data := document.NewTemplateData()
    data.SetVariable("companyName", "Drumkit")
    data.SetVariable("reportDate", "January 16, 2026")
    data.SetVariable("reporterName", "John Smith")
    data.SetVariable("department", "Sales Department")
    data.SetVariable("achievementRate", "125")
    data.SetVariable("totalSales", "1,850,000")
    data.SetVariable("newCustomers", "45")
    data.SetVariable("orders", "183")

    channels := []interface{}{
        map[string]interface{}{"name": "Online E-commerce", "sales": "742,000", "percentage": "40.1"},
        map[string]interface{}{"name": "Direct Sales Team", "sales": "555,000", "percentage": "30.0"},
    }
    data.SetList("channels", channels)

    // Render and save (recommended method)
    doc, _ := engine.RenderTemplateToDocument("sales_report", data)
    doc.Save("sales_report.docx")
}
```

#### Advantages and Application Scenarios

**Main Advantages**:
- **Code Reuse**: Avoid duplicating the same template structure
- **Maintainability**: Modifying base template automatically affects all child templates
- **Flexibility**: Selectively override only needed parts, keeping other default content
- **Extensibility**: Support multi-level inheritance to build complex template hierarchy

**Typical Application Scenarios**:
- **Enterprise Report System**: Base report template + department-specific templates
- **Document Standardization**: Unified format for different types of documents (contracts, invoices, notices, etc.)
- **Multi-Language Documents**: Documents with same structure in different languages
- **Brand Consistency**: Maintain consistency of enterprise brand elements

### Image Operation Functionality ✨ New Features
- [`AddImageFromFile(filePath string, config *ImageConfig)`](image.go) - Add image from file
- [`AddImageFromData(imageData []byte, fileName string, format ImageFormat, width, height int, config *ImageConfig)`](image.go) - Add image from data
- [`ResizeImage(imageInfo *ImageInfo, size *ImageSize)`](image.go) - Resize image
- [`SetImagePosition(imageInfo *ImageInfo, position ImagePosition, offsetX, offsetY float64)`](image.go) - Set image position
- [`SetImageWrapText(imageInfo *ImageInfo, wrapText ImageWrapText)`](image.go) - Set image text wrapping
- [`SetImageAltText(imageInfo *ImageInfo, altText string)`](image.go) - Set image alternative text
- [`SetImageTitle(imageInfo *ImageInfo, title string)`](image.go) - Set image title

## Paragraph Operation Methods

### Paragraph Formatting Settings
- [`SetAlignment(alignment AlignmentType)`](document.go) - Set paragraph alignment
- [`SetSpacing(config *SpacingConfig)`](document.go) - Set paragraph spacing
- [`SetStyle(styleID string)`](document.go) - Set paragraph style
- [`SetIndentation(firstLineCm, leftCm, rightCm float64)`](document.go) - Set paragraph indentation ✨ **Improved**
- [`SetKeepWithNext(keep bool)`](document.go) - Set keep with next paragraph ✨ **New**
- [`SetKeepLines(keep bool)`](document.go) - Set keep all paragraph lines together ✨ **New**
- [`SetPageBreakBefore(pageBreak bool)`](document.go) - Set page break before ✨ **New**
- [`SetWidowControl(control bool)`](document.go) - Set widow control ✨ **New**
- [`SetOutlineLevel(level int)`](document.go) - Set outline level ✨ **New**
- [`SetParagraphFormat(config *ParagraphFormatConfig)`](document.go) - Set all paragraph format properties at once ✨ **New**

#### Paragraph Format Advanced Features ✨ **New Features**

go-word now supports complete paragraph format customization, providing the same advanced paragraph control options as Microsoft Word.

**Page Break Control Features**:
- **SetKeepWithNext** - Ensure paragraph stays together with next paragraph on same page, preventing titles from appearing alone at page bottom
- **SetKeepLines** - Prevent paragraph from being split across pages, maintaining paragraph integrity
- **SetPageBreakBefore** - Force page break before paragraph, commonly used for chapter starts

**Widow Control**:
- **SetWidowControl** - Prevent first or last paragraph line from appearing alone at page top or bottom, improving typesetting quality

**Outline Level**:
- **SetOutlineLevel** - Set paragraph outline level (0-8) for document navigation pane display and TOC generation

**Comprehensive Format Setting**:
- **SetParagraphFormat** - Use `ParagraphFormatConfig` struct to set all paragraph properties at once
  - Basic formatting: alignment, style
  - Spacing settings: line spacing, paragraph before/after spacing, first line indent
  - Indentation settings: first line indent, left/right indent (support hanging indent)
  - Page break control: keep with next, keep lines, page break before, widow control
  - Outline level: 0-8 level settings

**Usage Examples**:

```go
// Method 1: Set using individual methods
title := doc.AddParagraph("Chapter 1 Overview")
title.SetAlignment(document.AlignCenter)
title.SetKeepWithNext(true)
title.SetPageBreakBefore(true)
title.SetOutlineLevel(0)

// Method 2: Set all at once using SetParagraphFormat
para := doc.AddParagraph("Important Content")
para.SetParagraphFormat(&document.ParagraphFormatConfig{
    Alignment:       document.AlignJustify,
    Style:           "Normal",
    LineSpacing:     1.5,
    BeforePara:      12,
    AfterPara:       6,
    FirstLineCm:     0.5,
    KeepWithNext:    true,
    KeepLines:       true,
    WidowControl:    true,
    OutlineLevel:    0,
})
```

**Application Scenarios**:
- **Document Structuring** - Use outline levels to create clear document hierarchy
- **Professional Typography** - Use page break control to ensure title and content association
- **Content Protection** - Use keep lines to prevent important paragraphs from being split
- **Chapter Management** - Use page break before to achieve chapter page independence

### Paragraph Content Operations
- [`AddFormattedText(text string, format *TextFormat)`](document.go) - Add formatted text
- [`AddPageBreak()`](document.go) - Add page break to paragraph ✨ **New**
- [`ElementType()`](document.go) - Get paragraph element type

## Document Body Operation Methods

### Element Query
- [`GetParagraphs()`](document.go) - Get all paragraphs
- [`GetTables()`](document.go) - Get all tables

### Element Addition
- [`AddElement(element interface{})`](document.go) - Add element to document body

## Table Operation Methods

### Table Creation
- [`CreateTable(config *TableConfig)`](table.go#L161) - Create new table (✨ New: includes single-line border style by default)
- [`AddTable(config *TableConfig)`](table.go#L257) - Add table to document

### Row Operations
- [`InsertRow(position int, data []string)`](table.go#L271) - Insert row at specified position
- [`AppendRow(data []string)`](table.go#L329) - Append row at table end
- [`DeleteRow(rowIndex int)`](table.go#L334) - Delete specified row
- [`DeleteRows(startIndex, endIndex int)`](table.go#L351) - Delete multiple rows
- [`GetRowCount()`](table.go#L562) - Get row count

### Column Operations
- [`InsertColumn(position int, data []string, width int)`](table.go#L369) - Insert column at specified position
- [`AppendColumn(data []string, width int)`](table.go#L438) - Append column at table end
- [`DeleteColumn(colIndex int)`](table.go#L447) - Delete specified column
- [`DeleteColumns(startIndex, endIndex int)`](table.go#L474) - Delete multiple columns
- [`GetColumnCount()`](table.go#L567) - Get column count

### Cell Operations
- [`GetCell(row, col int)`](table.go#L502) - Get specified cell
- [`SetCellText(row, col int, text string)`](table.go#L515) - Set cell text
- [`GetCellText(row, col int)`](table.go#L623) - Get cell text (upgraded: returns complete content of all paragraphs and Runs in cell, with paragraphs separated by `\n`)
    - Old behavior: only returned first Run of first paragraph, causing multi-line/soft line break content loss
    - New behavior: traverse all paragraphs and their Runs, concatenate text; empty paragraphs skip content but still produce paragraph line break (except end)
    - Note: If future needs require preserving Word `<w:br/>` (manual soft line break within same paragraph), parsing logic needs extension; currently only separates by paragraphs
- [`SetCellFormat(row, col int, format *CellFormat)`](table.go#L691) - Set cell format
- [`GetCellFormat(row, col int)`](table.go#L1238) - Get cell format

### Cell Text Formatting
- [`SetCellFormattedText(row, col int, text string, format *TextFormat)`](table.go#L780) - Set formatted text
- [`AddCellFormattedText(row, col int, text string, format *TextFormat)`](table.go#L833) - Add formatted text

### Cell Merging
- [`MergeCellsHorizontal(row, startCol, endCol int)`](table.go#L887) - Merge cells horizontally
- [`MergeCellsVertical(startRow, endRow, col int)`](table.go#L924) - Merge cells vertically
- [`MergeCellsRange(startRow, endRow, startCol, endCol int)`](table.go#L971) - Merge cells in range
- [`UnmergeCells(row, col int)`](table.go#L1004) - Unmerge cells
- [`IsCellMerged(row, col int)`](table.go#L1074) - Check if cell is merged
- [`GetMergedCellInfo(row, col int)`](table.go#L1098) - Get merged cell information

### Cell Special Properties
- [`SetCellPadding(row, col int, padding int)`](table.go#L1189) - Set cell padding
- [`SetCellTextDirection(row, col int, direction CellTextDirection)`](table.go#L1202) - Set text direction
- [`GetCellTextDirection(row, col int)`](table.go#L1223) - Get text direction
- [`ClearCellContent(row, col int)`](table.go#L1138) - Clear cell content
- [`ClearCellFormat(row, col int)`](table.go#L1156) - Clear cell format

### Table Overall Operations
- [`ClearTable()`](table.go#L575) - Clear table content
- [`CopyTable()`](table.go#L593) - Copy table
- [`ElementType()`](table.go#L66) - Get table element type

### Row Height Settings
- [`SetRowHeight(rowIndex int, config *RowHeightConfig)`](table.go#L1318) - Set row height
- [`GetRowHeight(rowIndex int)`](table.go#L1339) - Get row height
- [`SetRowHeightRange(startRow, endRow int, config *RowHeightConfig)`](table.go#L1371) - Set row height range

### Table Layout and Alignment
- [`SetTableLayout(config *TableLayoutConfig)`](table.go#L1447) - Set table layout
- [`GetTableLayout()`](table.go#L1473) - Get table layout
- [`SetTableAlignment(alignment TableAlignment)`](table.go#L1488) - Set table alignment

### Row Properties Setting
- [`SetRowKeepTogether(rowIndex int, keepTogether bool)`](table.go#L1529) - Set row keep together
- [`SetRowAsHeader(rowIndex int, isHeader bool)`](table.go#L1552) - Set row as header row
- [`SetHeaderRows(startRow, endRow int)`](table.go#L1575) - Set multiple rows as header rows
- [`IsRowHeader(rowIndex int)`](table.go#L1600) - Check if header row
- [`IsRowKeepTogether(rowIndex int)`](table.go#L1614) - Check if row keeps together
- [`SetRowKeepWithNext(rowIndex int, keepWithNext bool)`](table.go#L1645) - Set row keep with next

### Table Page Break Settings
- [`SetTablePageBreak(config *TablePageBreakConfig)`](table.go#L1636) - Set table page break
- [`GetTableBreakInfo()`](table.go#L1657) - Get page break information

### Table Styles
- [`ApplyTableStyle(config *TableStyleConfig)`](table.go#L1956) - Apply table style
- [`CreateCustomTableStyle(styleID, styleName string, borderConfig *TableBorderConfig, shadingConfig *ShadingConfig, firstRowBold bool)`](table.go#L2213) - Create custom table style

### Border Settings
- [`SetTableBorders(config *TableBorderConfig)`](table.go#L2038) - Set table borders
- [`SetCellBorders(row, col int, config *CellBorderConfig)`](table.go#L2085) - Set cell borders
- [`RemoveTableBorders()`](table.go#L2168) - Remove table borders
- [`RemoveCellBorders(row, col int)`](table.go#L2194) - Remove cell borders

### Background and Shading
- [`SetTableShading(config *ShadingConfig)`](table.go#L2069) - Set table shading
- [`SetCellShading(row, col int, config *ShadingConfig)`](table.go#L2121) - Set cell shading
- [`SetAlternatingRowColors(evenRowColor, oddRowColor string)`](table.go#L2142) - Set alternating row colors

### Cell Image Functionality ✨ **New Feature**

Support adding images to table cells:

- [`AddCellImage(table *Table, row, col int, config *CellImageConfig)`](image.go#L1106) - Add image to cell (complete configuration)
- [`AddCellImageFromFile(table *Table, row, col int, filePath string, widthMM float64)`](image.go#L1214) - Add image to cell from file
- [`AddCellImageFromData(table *Table, row, col int, data []byte, widthMM float64)`](image.go#L1236) - Add image to cell from binary data

#### CellImageConfig - Cell Image Configuration
```go
type CellImageConfig struct {
    FilePath        string      // Image file path
    Data            []byte      // Image binary data (choose one with FilePath)
    Format          ImageFormat // Image format (required when using Data)
    Width           float64     // Image width (millimeters), 0 for auto
    Height          float64     // Image height (millimeters), 0 for auto
    KeepAspectRatio bool        // Whether to keep aspect ratio
    AltText         string      // Image alternative text
    Title           string      // Image title
}
```

#### Table Cell Image Usage Example
```go
// Create table
table, err := doc.AddTable(&document.TableConfig{
    Rows:  2,
    Cols:  2,
    Width: 8000,
})

// Method 1: Add image to cell from file
imageInfo, err := doc.AddCellImageFromFile(table, 0, 0, "logo.png", 30) // 30mm width

// Method 2: Add image from binary data
imageData := []byte{...} // Image binary data
imageInfo, err := doc.AddCellImageFromData(table, 0, 1, imageData, 25) // 25mm width

// Method 3: Use complete configuration
config := &document.CellImageConfig{
    FilePath:        "product.jpg",
    Width:           50,     // 50mm width
    Height:          40,     // 40mm height
    KeepAspectRatio: false,  // Don't keep aspect ratio
    AltText:         "Product image",
    Title:           "Product showcase",
}
imageInfo, err := doc.AddCellImage(table, 1, 0, config)
```

**Notes**:
- Images are added through `Document` object methods because image resources need to be managed at document level
- Support PNG, JPEG, GIF image formats
- Width/height units are millimeters, 0 uses original size
- When setting `KeepAspectRatio` to `true`, only need to set width or height, not both

### Cell Iterator Functionality ✨ **New Feature**

Provides powerful cell traversal and search functionality:

##### CellIterator - Cell Iterator
```go
// Create iterator
iterator := table.NewCellIterator()

// Traverse all cells
for iterator.HasNext() {
    cellInfo, err := iterator.Next()
    if err != nil {
        break
    }
    fmt.Printf("Cell[%d,%d]: %s\n", cellInfo.Row, cellInfo.Col, cellInfo.Text)
}

// Get progress
progress := iterator.Progress() // 0.0 - 1.0

// Reset iterator
iterator.Reset()
```

##### ForEach Batch Processing
```go
// Traverse all cells
err := table.ForEach(func(row, col int, cell *TableCell, text string) error {
    // Process each cell
    return nil
})

// Traverse by row
err := table.ForEachInRow(rowIndex, func(col int, cell *TableCell, text string) error {
    // Process each cell in row
    return nil
})

// Traverse by column
err := table.ForEachInColumn(colIndex, func(row int, cell *TableCell, text string) error {
    // Process each cell in column
    return nil
})
```

##### Range Operations
```go
// Get cells in specified range
cells, err := table.GetCellRange(startRow, startCol, endRow, endCol)
for _, cellInfo := range cells {
    fmt.Printf("Cell[%d,%d]: %s\n", cellInfo.Row, cellInfo.Col, cellInfo.Text)
}
```

##### Search Functionality
```go
// Custom condition search
cells, err := table.FindCells(func(row, col int, cell *TableCell, text string) bool {
    return strings.Contains(text, "keyword")
})

// Search by text
exactCells, err := table.FindCellsByText("exact match", true)
fuzzyCells, err := table.FindCellsByText("fuzzy", false)
```

##### CellInfo Structure
```go
type CellInfo struct {
    Row    int        // Row index
    Col    int        // Column index
    Cell   *TableCell // Cell reference
    Text   string     // Cell text
    IsLast bool       // Whether last cell
}
```

## Utility Functions

### Logging System
- [`NewLogger(level LogLevel, output io.Writer)`](logger.go#L56) - Create new logger
- [`SetGlobalLevel(level LogLevel)`](logger.go#L129) - Set global log level
- [`SetGlobalOutput(output io.Writer)`](logger.go#L134) - Set global log output
- [`Debug(msg string)`](logger.go#L159) - Output debug message
- [`Info(msg string)`](logger.go#L164) - Output info message
- [`Warn(msg string)`](logger.go#L169) - Output warning
- [`Error(msg string)`](logger.go#L174) - Output error

### Error Handling
- [`NewDocumentError(operation string, cause error, context string)`](errors.go#L47) - Create document error
- [`WrapError(operation string, err error)`](errors.go#L56) - Wrap error
- [`WrapErrorWithContext(operation string, err error, context string)`](errors.go#L64) - Wrap error with context
- [`NewValidationError(field, value, message string)`](errors.go#L84) - Create validation error

### Field Tools ✨ New Features
- [`CreateHyperlinkField(anchor string)`](field.go) - Create hyperlink field
- [`CreatePageRefField(anchor string)`](field.go) - Create page reference field

## Common Configuration Structures

### Text Format
- `TextFormat` - Text format configuration
- `AlignmentType` - Alignment type
- `SpacingConfig` - Spacing configuration

### Table Configuration
- `TableConfig` - Table basic configuration
- `CellFormat` - Cell format
- `RowHeightConfig` - Row height configuration
- `TableLayoutConfig` - Table layout configuration
- `TableStyleConfig` - Table style configuration
- `BorderConfig` - Border configuration
- `ShadingConfig` - Shading configuration

### Page Settings Configuration ✨ New
- `PageSettings` - Page settings configuration
- `PageSize` - Page size type (A4, Letter, Legal, A3, A5, Custom)
- `PageOrientation` - Page orientation (Portrait, Landscape)
- `SectionProperties` - Section properties (contain page setting information)

### Header and Footer Configuration ✨ New
- `HeaderFooterType` - Header/footer type (Default, First, Even)
- `Header` - Header structure
- `Footer` - Footer structure
- `HeaderFooterReference` - Header/footer reference
- `PageNumber` - Page number field

### Table of Contents Configuration ✨ New
- `TOCConfig` - TOC configuration
- `TOCEntry` - TOC entry
- `Bookmark` - Bookmark structure
- `BookmarkEnd` - Bookmark end marker

### Footnote and Endnote Configuration ✨ New
- `FootnoteConfig` - Footnote configuration
- `FootnoteType` - Footnote type (Footnote, Endnote)
- `FootnoteNumberFormat` - Footnote number format
- `FootnoteRestart` - Footnote restart rule
- `FootnotePosition` - Footnote position
- `Footnote` - Footnote structure
- `Endnote` - Endnote structure

### List and Numbering Configuration ✨ New
- `ListConfig` - List configuration
- `ListType` - List type (Bullet, Number, etc.)
- `BulletType` - Bullet type
- `ListItem` - List item structure
- `Numbering` - Numbering definition
- `AbstractNum` - Abstract numbering definition
- `Level` - Numbering level

### Structured Document Tag Configuration ✨ New
- `SDT` - Structured document tag
- `SDTProperties` - SDT properties
- `SDTContent` - SDT content

### Field Configuration ✨ New
- `FieldChar` - Field character
- `InstrText` - Field instruction text
- `HyperlinkField` - Hyperlink field
- `PageRefField` - Page reference field

### Image Configuration ✨ New
- `ImageConfig` - Image configuration
- `ImageSize` - Image size configuration
- `ImageFormat` - Image format (PNG, JPEG, GIF)
- `ImagePosition` - Image position (inline, floatLeft, floatRight)
- `ImageWrapText` - Text wrapping type (none, square, tight, topAndBottom)
- `ImageInfo` - Image information structure
- `AlignmentType` - Alignment type (left, center, right, justify)

## Usage Examples

```go
// Create new document
doc := document.New()

// ✨ New: Page settings example
// Set page to A4 landscape
doc.SetPageOrientation(document.OrientationLandscape)

// Set custom margins (top, bottom, left, right: 25mm)
doc.SetPageMargins(25, 25, 25, 25)

// Set custom page size (200mm x 300mm)
doc.SetCustomPageSize(200, 300)

// Or use complete page settings
pageSettings := &document.PageSettings{
    Size:           document.PageSizeLetter,
    Orientation:    document.OrientationPortrait,
    MarginTop:      30,
    MarginRight:    20,
    MarginBottom:   30,
    MarginLeft:     20,
    HeaderDistance: 15,
    FooterDistance: 15,
    GutterWidth:    0,
}
doc.SetPageSettings(pageSettings)

// ✨ New: Header and footer example
// Add header
doc.AddHeader(document.HeaderFooterTypeDefault, "This is header")

// Add footer with page number
doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "Page ", true)

// Set different first page
doc.SetDifferentFirstPage(true)

// ✨ New: Table of contents example
// Add heading with bookmark
doc.AddHeadingWithBookmark("Chapter 1 Overview", 1, "chapter1")
doc.AddHeadingWithBookmark("1.1 Background", 2, "section1_1")

// Generate table of contents
tocConfig := document.DefaultTOCConfig()
tocConfig.Title = "Table of Contents"
tocConfig.MaxLevel = 3
doc.GenerateTOC(tocConfig)

// ✨ New: Footnote example
// Add footnote
doc.AddFootnote("This is main text", "This is footnote content")

// Add endnote
doc.AddEndnote("More explanation", "This is endnote content")

// ✨ New: List example
// Add unordered list
doc.AddBulletList("List item 1", 0, document.BulletTypeDot)
doc.AddBulletList("List item 2", 1, document.BulletTypeCircle)

// Add ordered list
doc.AddNumberedList("Numbered item 1", 0, document.ListTypeNumber)

// ✨ New: Image example
// Add image from file
imageInfo, err := doc.AddImageFromFile("path/to/image.png", &document.ImageConfig{
    Size: &document.ImageSize{
        Width:  100.0, // 100mm width
        Height: 75.0,  // 75mm height
    },
    Position: document.ImagePositionInline,
    WrapText: document.ImageWrapNone,
    AltText:  "Sample image",
    Title:    "This is a sample image",
})

// Add image from data
imageData := []byte{...} // Image binary data
imageInfo2, err := doc.AddImageFromData(
    imageData,
    "example.png",
    document.ImageFormatPNG,
    200, 150, // Original pixel size
    &document.ImageConfig{
        Size: &document.ImageSize{
            Width:           60.0, // Only set width
            KeepAspectRatio: true, // Keep aspect ratio
        },
        AltText: "Data image",
    },
)

// Resize image
err = doc.ResizeImage(imageInfo, &document.ImageSize{
    Width:  80.0,
    Height: 60.0,
})

// Set image properties
err = doc.SetImagePosition(imageInfo, document.ImagePositionFloatLeft, 5.0, 0.0)
err = doc.SetImageWrapText(imageInfo, document.ImageWrapSquare)
err = doc.SetImageAltText(imageInfo, "Updated alternative text")
err = doc.SetImageTitle(imageInfo, "Updated title")

// ✨ New: Set image alignment (only for inline images)
err = doc.SetImageAlignment(imageInfo, document.AlignCenter)  // Center alignment
err = doc.SetImageAlignment(imageInfo, document.AlignLeft)    // Left alignment
err = doc.SetImageAlignment(imageInfo, document.AlignRight)   // Right alignment
doc.AddNumberedList("First item", 0, document.ListTypeDecimal)
doc.AddNumberedList("Second item", 0, document.ListTypeDecimal)

// Add paragraph
para := doc.AddParagraph("This is a paragraph")
para.SetAlignment(document.AlignCenter)

// Create table
table := doc.CreateTable(&document.TableConfig{
    Rows:  3,
    Cols:  3,
    Width: 5000,
})

// Set cell content
table.SetCellText(0, 0, "Title")

// Save document
doc.Save("example.docx")
```

## Notes

1. All position indices start from 0
2. Width units use points (pt), 1 point = 20 twips
3. Colors use hexadecimal format, e.g., "FF0000" represents red
4. Before operating on tables, ensure row/column indices are valid, otherwise errors may be returned
5. Header/footer types include: Default, First, Even
6. TOC functionality requires first adding headings with bookmarks, then calling generate TOC method
7. Footnotes and endnotes are auto-numbered, supporting multiple numbering formats and restart rules
8. Lists support multi-level nesting, maximum 9 levels of indentation
9. Structured document tags are mainly used for special functionalities like TOC
10. Images support PNG, JPEG, GIF formats, automatically embedded in document
11. Image size can be specified in millimeters or pixels, supporting aspect ratio-preserving scaling
12. Image position supports inline, left-floating, right-floating and other layout methods
13. Image alignment functionality only applies to inline images (ImagePositionInline), use position control for floating images

## Markdown to Word Functionality

go-word supports converting Markdown documents to Word format, implementing high-quality conversion based on the goldmark parsing engine.

### Markdown Package API

#### Converter Interface
- [`NewConverter(options *ConvertOptions)`](../markdown/converter.go) - Create new Markdown converter
- [`DefaultOptions()`](../markdown/config.go) - Get default conversion options
- [`HighQualityOptions()`](../markdown/config.go) - Get high-quality conversion options

#### Conversion Methods
- [`ConvertString(content string, options *ConvertOptions)`](../markdown/converter.go) - Convert Markdown string to Word document
- [`ConvertBytes(content []byte, options *ConvertOptions)`](../markdown/converter.go) - Convert Markdown byte array to Word document
- [`ConvertFile(mdPath, docxPath string, options *ConvertOptions)`](../markdown/converter.go) - Convert Markdown file to Word file
- [`BatchConvert(inputs []string, outputDir string, options *ConvertOptions)`](../markdown/converter.go) - Batch convert Markdown files

#### Configuration Options (`ConvertOptions`)
- `EnableGFM` - Enable GitHub Flavored Markdown support
- `EnableFootnotes` - Enable footnote support
- `EnableTables` - Enable table support
- `EnableTaskList` - Enable task list support
- `StyleMapping` - Custom style mapping
- `DefaultFontFamily` - Default font family
- `DefaultFontSize` - Default font size
- `ImageBasePath` - Image base path
- `EmbedImages` - Whether to embed images
- `MaxImageWidth` - Maximum image width (inches)
- `PreserveLinkStyle` - Preserve link style
- `ConvertToBookmarks` - Convert internal links to bookmarks
- `GenerateTOC` - Generate table of contents
- `TOCMaxLevel` - Table of contents maximum level
- `PageSettings` - Page settings
- `StrictMode` - Strict mode
- `IgnoreErrors` - Ignore conversion errors
- `ErrorCallback` - Error callback function
- `ProgressCallback` - Progress callback function

### Supported Markdown Syntax

#### Basic Syntax
- **Headings** (`# ## ### #### ##### ######`) - Convert to Word heading styles 1-6
- **Paragraphs** - Convert to Word text paragraphs
- **Bold** (`**text**`) - Convert to bold format
- **Italic** (`*text*`) - Convert to italic format
- **Inline Code** (`` `code` ``) - Convert to monospace font
- **Code Block** (``` ```) - Convert to code block style

#### List Support
- **Unordered Lists** (`- * +`) - Convert to Word bullet lists
- **Ordered Lists** (`1. 2. 3.`) - Convert to Word numbered lists
- **Multi-Level Lists** - Support nested list structure

#### GitHub Flavored Markdown Extensions ✨ **New**
- **Tables** (`| col1 | col2 |`) - Convert to Word tables
  - Auto-identify table headers and set styles
  - Support alignment control (left `:---`, center `:---:`, right `---:`)
  - Auto-set table borders and cell format
- **Task Lists** (`- [x] Completed` / `- [ ] Incomplete`) - Convert to checkbox symbols
  - ☑ Indicates completed task
  - ☐ Indicates incomplete task
  - Support nested task lists
  - Support mixed formatting (bold, italic, code, etc.)

#### Other Elements
- **Block Quotes** (`> quoted text`) - Convert to italic quote style
- **Horizontal Rule** (`---`) - Convert to horizontal line
- **Links** (`[text](URL)`) - Convert to blue text (hyperlink support coming soon)
- **Images** (`![alt](src)`) - Convert to image placeholder (image embedding support coming soon)

### Usage Examples

#### Basic String Conversion
```go
import "github.com/drumkitai/go-word/pkg/markdown"

// Create converter
converter := markdown.NewConverter(markdown.DefaultOptions())

// Convert Markdown string
markdownText := `# Title

This is a paragraph with **bold** and *italic*.

## Subtitle

- List item 1
- List item 2

> Quoted text

` + "`" + `code example` + "`" + `
`

doc, err := converter.ConvertString(markdownText, nil)
if err != nil {
    log.Fatal(err)
}

// Save Word document
err = doc.Save("output.docx")
```

#### Table and Task List Example ✨ **New**
```go
// Enable table and task list features
options := markdown.DefaultOptions()
options.EnableTables = true
options.EnableTaskList = true

converter := markdown.NewConverter(options)

// Markdown with tables and task lists
markdownWithTable := `# Project Progress

## Feature Implementation Status

| Feature | Status | Owner |
|:--------|:------:|------:|
| Table Conversion | ✅ | John |
| Task Lists | ✅ | Jane |
| Image Processing | 🚧 | Bob |

## TODO Items

- [x] Implement table conversion
  - [x] Basic table support
  - [x] Alignment handling
  - [x] Header style setting
- [ ] Improve task list features
  - [x] Checkbox display
  - [ ] Interactive features
- [ ] Image embedding support
  - [ ] PNG format
  - [ ] JPEG format

## Notes

> Tables support **left alignment**, ` + "`" + `center alignment` + "`" + ` and ***right alignment*** three methods
`

doc, err := converter.ConvertString(markdownWithTable, options)
if err != nil {
    log.Fatal(err)
}

err = doc.Save("project_status.docx")
```

#### Advanced Configuration Example
```go
// Create high-quality conversion configuration
options := &markdown.ConvertOptions{
    EnableGFM:         true,
    EnableFootnotes:   true,
    EnableTables:      true,
    GenerateTOC:       true,
    TOCMaxLevel:       3,
    DefaultFontFamily: "Calibri",
    DefaultFontSize:   11.0,
    EmbedImages:       true,
    MaxImageWidth:     6.0,
    PageSettings: &document.PageSettings{
        Size:        document.PageSizeA4,
        Orientation: document.OrientationPortrait,
        MarginTop:   25,
        MarginRight: 20,
        MarginBottom: 25,
        MarginLeft:  20,
    },
    ProgressCallback: func(current, total int) {
        fmt.Printf("Conversion progress: %d/%d\n", current, total)
    },
}

converter := markdown.NewConverter(options)
```

#### File Conversion Example
```go
// Single file conversion
err := converter.ConvertFile("input.md", "output.docx", nil)

// Batch file conversion
files := []string{"doc1.md", "doc2.md", "doc3.md"}
err := converter.BatchConvert(files, "output/", options)
```

#### Custom Style Mapping
```go
options := markdown.DefaultOptions()
options.StyleMapping = map[string]string{
    "heading1": "CustomTitle",
    "heading2": "CustomSubtitle",
    "quote":    "CustomQuote",
    "code":     "CustomCode",
}

converter := markdown.NewConverter(options)
```

## Word to Markdown Functionality ✨ **New Features**

go-word now supports reverse conversion from Word documents to Markdown format, providing complete bidirectional conversion capability.

### Word Exporter API

#### Exporter Interface
- [`NewExporter(options *ExportOptions)`](../markdown/exporter.go) - Create new Word exporter
- [`DefaultExportOptions()`](../markdown/exporter.go) - Get default export options
- [`HighQualityExportOptions()`](../markdown/exporter.go) - Get high-quality export options

#### Export Methods
- [`ExportToFile(docxPath, mdPath string, options *ExportOptions)`](../markdown/exporter.go) - Export Word document to Markdown file
- [`ExportToString(doc *Document, options *ExportOptions)`](../markdown/exporter.go) - Export Word document to Markdown string
- [`ExportToBytes(doc *Document, options *ExportOptions)`](../markdown/exporter.go) - Export Word document to Markdown byte array
- [`BatchExport(inputs []string, outputDir string, options *ExportOptions)`](../markdown/exporter.go) - Batch export Word documents

#### Export Configuration Options (`ExportOptions`)
- `UseGFMTables` - Use GitHub flavored Markdown tables
- `PreserveFootnotes` - Preserve footnotes
- `PreserveLineBreaks` - Preserve line breaks
- `WrapLongLines` - Auto-wrap long lines
- `MaxLineLength` - Maximum line length
- `ExtractImages` - Export image files
- `ImageOutputDir` - Image output directory
- `ImageNamePattern` - Image naming pattern
- `ImageRelativePath` - Use relative paths
- `PreserveBookmarks` - Preserve bookmarks
- `ConvertHyperlinks` - Convert hyperlinks
- `PreserveCodeStyle` - Preserve code style
- `DefaultCodeLang` - Default code language
- `IgnoreUnknownStyles` - Ignore unknown styles
- `PreserveTOC` - Preserve table of contents
- `IncludeMetadata` - Include document metadata
- `StripComments` - Remove comments
- `UseSetext` - Use Setext style headings
- `BulletListMarker` - Bullet list marker
- `EmphasisMarker` - Emphasis marker
- `StrictMode` - Strict mode
- `IgnoreErrors` - Ignore errors
- `ErrorCallback` - Error callback function
- `ProgressCallback` - Progress callback function

### Word to Markdown Conversion Mapping

| Word Element | Markdown Syntax | Description |
|-------------|-----------------|-------------|
| Heading1-6 | `# ## ### #### ##### ######` | Heading level mapping |
| Bold | `**bold**` | Text format |
| Italic | `*italic*` | Text format |
| Strikethrough | `~~strikethrough~~` | Text format |
| Inline Code | `` `code` `` | Code format |
| Code Block | ```` code block ```` | Code block |
| Hyperlink | `[link text](URL)` | Link conversion |
| Image | `![image](path)` | Image reference |
| Table | `\| table \|` | GFM table format |
| Unordered List | `- item` | List item |
| Ordered List | `1. item` | Numbered list |
| Block Quote | `> quoted content` | Quote format |

### Word to Markdown Usage Examples

#### Basic File Export
```go
import "github.com/drumkitai/go-word/pkg/markdown"

// Create exporter
exporter := markdown.NewExporter(markdown.DefaultExportOptions())

// Export Word document to Markdown
err := exporter.ExportToFile("document.docx", "output.md", nil)
if err != nil {
    log.Fatal(err)
}
```

#### Export to String
```go
// Open Word document
doc, err := document.Open("document.docx")
if err != nil {
    log.Fatal(err)
}

// Export to Markdown string
exporter := markdown.NewExporter(markdown.DefaultExportOptions())
markdownText, err := exporter.ExportToString(doc, nil)
if err != nil {
    log.Fatal(err)
}

fmt.Println(markdownText)
```

#### High-Quality Export Configuration
```go
// High-quality export configuration
options := &markdown.ExportOptions{
    UseGFMTables:      true,              // Use GFM tables
    ExtractImages:     true,              // Export images
    ImageOutputDir:    "./images",        // Image directory
    PreserveFootnotes: true,              // Preserve footnotes
    IncludeMetadata:   true,              // Include metadata
    ConvertHyperlinks: true,              // Convert hyperlinks
    PreserveCodeStyle: true,              // Preserve code style
    UseSetext:         false,             // Use ATX headings
    BulletListMarker:  "-",              // Use hyphens
    EmphasisMarker:    "*",              // Use asterisks
    ProgressCallback: func(current, total int) {
        fmt.Printf("Export progress: %d/%d\n", current, total)
    },
}

exporter := markdown.NewExporter(options)
err := exporter.ExportToFile("complex_document.docx", "output.md", options)
```

#### Batch Export Example
```go
// Batch export Word documents
files := []string{"doc1.docx", "doc2.docx", "doc3.docx"}

options := &markdown.ExportOptions{
    ExtractImages:     true,
    ImageOutputDir:    "extracted_images/",
    UseGFMTables:      true,
    ProgressCallback: func(current, total int) {
        fmt.Printf("Batch export progress: %d/%d\n", current, total)
    },
}

exporter := markdown.NewExporter(options)
err := exporter.BatchExport(files, "markdown_output/", options)
```

## Bidirectional Converter ✨ **Unified Interface**

### Bidirectional Converter API
- [`NewBidirectionalConverter(mdOpts *ConvertOptions, exportOpts *ExportOptions)`](../markdown/exporter.go) - Create bidirectional converter
- [`AutoConvert(inputPath, outputPath string)`](../markdown/exporter.go) - Auto-detect file type and convert

### Bidirectional Conversion Usage Examples

#### Auto Conversion
```go
import "github.com/drumkitai/go-word/pkg/markdown"

// Create bidirectional converter
converter := markdown.NewBidirectionalConverter(
    markdown.HighQualityOptions(),        // Markdown→Word options
    markdown.HighQualityExportOptions(),  // Word→Markdown options
)

// Auto-detect file type and convert
err := converter.AutoConvert("input.docx", "output.md")     // Word→Markdown
err = converter.AutoConvert("input.md", "output.docx")     // Markdown→Word
```

#### Configure Independent Conversion Directions
```go
// Markdown to Word configuration
mdToWordOpts := &markdown.ConvertOptions{
    EnableGFM:         true,
    EnableTables:      true,
    GenerateTOC:       true,
    DefaultFontFamily: "Calibri",
    DefaultFontSize:   11.0,
}

// Word to Markdown configuration
wordToMdOpts := &markdown.ExportOptions{
    UseGFMTables:      true,
    ExtractImages:     true,
    ImageOutputDir:    "./images",
    PreserveFootnotes: true,
    ConvertHyperlinks: true,
}

// Create bidirectional converter
converter := markdown.NewBidirectionalConverter(mdToWordOpts, wordToMdOpts)

// Execute conversion
err := converter.AutoConvert("document.docx", "document.md")
```

### Technical Features

#### Architecture Design
- **goldmark Integration** - Uses high-performance goldmark parsing engine
- **AST Traversal** - Conversion processing based on abstract syntax tree
- **API Reuse** - Fully reuses existing WordZero document API
- **Backward Compatible** - Does not affect existing document package functionality

#### Performance Advantages
- **Stream Processing** - Support stream conversion for large documents
- **Memory Efficiency** - Optimized memory usage patterns
- **Concurrency Support** - Batch conversion supports concurrent processing
- **Error Recovery** - Intelligent error handling and recovery mechanism

#### Extensibility
- **Plugin Architecture** - Support custom renderer extensions
- **Configuration-Driven** - Rich configuration options for different needs
- **Style System** - Flexible style mapping and customization capability
- **Callback Mechanism** - Progress and error callback support

### Notes

1. **Compatibility** - Based on CommonMark 0.31.2 standard, highly compatible with GitHub Markdown
2. **Image Processing** - Current version converts images to placeholders, complete image support is planned
3. **Table Support** ✨ **Improved** - Full support for GFM table syntax, including alignment control and header styles
4. **Task Lists** ✨ **Implemented** - Support task checkboxes displayed as Unicode symbols (☑/☐)
5. **Link Processing** - Currently converts to blue text, hyperlink functionality in development
6. **Style Mapping** - Customize Markdown element to Word style mapping through StyleMapping
7. **Error Handling** - Recommended to enable error callback in production environment for quality monitoring
8. **Performance Consideration** - When batch converting many files, recommend batch processing to avoid memory pressure
9. **Encoding Support** - Full UTF-8 support including multi-byte characters like Chinese
10. **Configuration Requirements** - Table and task list features require explicit enabling in ConvertOptions
11. **Backward Compatibility** - New features do not affect existing document package API, fully compatible
