/*
Package document provides a Go library for creating, editing and manipulating Microsoft Word documents.

go-word focuses on the modern Office Open XML (OOXML) format (.docx files),
providing a simple and easy-to-use API to create and modify Word documents.

# Main Features

## Basic Features
- Create new Word documents
- Open and parse existing .docx files
- Add and format text content
- Set paragraph styles and alignment
- Configure fonts, colors and text formatting
- Set line spacing and paragraph spacing
- Error handling and logging

## Advanced Features
- **Headers and Footers**: Support for default, first page, and even page headers/footers
- **Table of Contents**: Auto-generate table of contents based on heading styles, with hyperlinks and page numbers
- **Footnotes and Endnotes**: Complete footnote and endnote functionality with multiple numbering formats
- **List and Numbering**: Unordered and ordered lists with multi-level nesting
- **Page Settings**: Complete page property settings including size, orientation, and margins
- **Tables**: Powerful table creation, formatting and styling functionality
- **Style System**: 18 predefined styles and custom style support

# Quick Start

Create a simple document:

	doc := document.New()
	doc.AddParagraph("Hello, World!")
	err := doc.Save("hello.docx")

Create a formatted document:

	doc := document.New()

	// Add formatted title
	titleFormat := &document.TextFormat{
		Bold:      true,
		FontSize:  18,
		FontColor: "FF0000", // Red
		FontName:  "Arial",
	}
	title := doc.AddFormattedParagraph("Document Title", titleFormat)
	title.SetAlignment(document.AlignCenter)

	// Add body paragraph
	para := doc.AddParagraph("This is body content...")
	para.SetSpacing(&document.SpacingConfig{
		LineSpacing:     1.5, // 1.5x line spacing
		BeforePara:      12,  // 12pt before paragraph
		AfterPara:       6,   // 6pt after paragraph
		FirstLineIndent: 24,  // First line indent 24pt
	})

	err := doc.Save("formatted.docx")

Open an existing document:

	doc, err := document.Open("existing.docx")
	if err != nil {
		log.Fatal(err)
	}

	// Read paragraph content
	for i, para := range doc.Body.Paragraphs {
		fmt.Printf("Paragraph %d: ", i+1)
		for _, run := range para.Runs {
			fmt.Print(run.Text.Content)
		}
		fmt.Println()
	}

# Advanced Features Examples

## Headers and Footers

	// Add header
	doc.AddHeader(document.HeaderFooterTypeDefault, "This is a header")

	// Add footer with page number
	doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "Page ", true)

	// Set first page different
	doc.SetDifferentFirstPage(true)

## Table of Contents

	// Add title with bookmark
	doc.AddHeadingWithBookmark("Chapter 1: Introduction", 1, "chapter1")
	doc.AddHeadingWithBookmark("1.1 Background", 2, "section1_1")

	// Generate table of contents
	tocConfig := document.DefaultTOCConfig()
	tocConfig.Title = "Table of Contents"
	tocConfig.MaxLevel = 3
	doc.GenerateTOC(tocConfig)

## Footnotes and Endnotes

	// Add footnote
	doc.AddFootnote("This is body text", "This is footnote content")

	// Add endnote
	doc.AddEndnote("Additional notes", "This is endnote content")

	// Custom footnote configuration
	footnoteConfig := &document.FootnoteConfig{
		NumberFormat: document.FootnoteFormatLowerRoman,
		StartNumber:  1,
		RestartEach:  document.FootnoteRestartEachPage,
		Position:     document.FootnotePositionPageBottom,
	}
	doc.SetFootnoteConfig(footnoteConfig)

## List Features

	// Unordered list
	doc.AddBulletList("List item 1", 0, document.BulletTypeDot)
	doc.AddBulletList("Sub-item", 1, document.BulletTypeCircle)

	// Ordered list
	doc.AddNumberedList("First item", 0, document.ListTypeDecimal)
	doc.AddNumberedList("Second item", 0, document.ListTypeDecimal)

	// Multi-level list
	items := []document.ListItem{
		{Text: "Level 1 item", Level: 0, Type: document.ListTypeDecimal},
		{Text: "Level 2 item", Level: 1, Type: document.ListTypeLowerLetter},
		{Text: "Level 3 item", Level: 2, Type: document.ListTypeLowerRoman},
	}
	doc.CreateMultiLevelList(items)

## Page Settings

	// Set page to A4 landscape
	doc.SetPageOrientation(document.OrientationLandscape)

	// Set page margins (millimeters)
	doc.SetPageMargins(25, 25, 25, 25)

	// Complete page settings
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

## Tables

	// Create table
	table := doc.CreateTable(&document.TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 5000,
	})

	// Set cell content
	table.SetCellText(0, 0, "Title")

	// Apply table style
	table.ApplyTableStyle(&document.TableStyleConfig{
		HeaderRow:    true,
		FirstColumn:  true,
		BandedRows:   true,
		BandedCols:   false,
	})

# Error Handling

The library provides unified error handling:

	doc, err := document.Open("nonexistent.docx")
	if err != nil {
		var docErr *document.DocumentError
		if errors.As(err, &docErr) {
			Errorf("Document operation failed - Operation: %s, Error: %v", docErr.Operation, docErr.Cause)
			fmt.Printf("Operation: %s, Error: %v\n", docErr.Operation, docErr.Cause)
		}
	}

# Logging

Configure log levels to control output:

	// Set to debug mode
	document.SetGlobalLevel(document.LogLevelDebug)

	// Only show errors
	document.SetGlobalLevel(document.LogLevelError)

# Text Formatting

The TextFormat struct supports various text formatting options:

	format := &document.TextFormat{
		Bold:      true,           // Bold
		Italic:    true,           // Italic
		FontSize:  14,             // Font size (points)
		FontColor: "0000FF",       // Font color (hexadecimal)
		FontName:  "Times New Roman", // Font name
	}

# Paragraph Alignment

Four alignment types are supported:

	para.SetAlignment(document.AlignLeft)     // Left align
	para.SetAlignment(document.AlignCenter)   // Center align
	para.SetAlignment(document.AlignRight)    // Right align
	para.SetAlignment(document.AlignJustify)  // Justify

# Spacing Configuration

Precise control over paragraph spacing:

	config := &document.SpacingConfig{
		LineSpacing:     1.5, // Line spacing (multiple)
		BeforePara:      12,  // Before paragraph spacing (points)
		AfterPara:       6,   // After paragraph spacing (points)
		FirstLineIndent: 24,  // First line indent (points)
	}
	para.SetSpacing(config)

# Important Notes

- Font size is in points, automatically converted to half-points internally by Word
- Color values use hexadecimal format, no # prefix needed
- Spacing values are in points, converted to TWIPs internally (1pt=20TWIPs)
- All text content uses UTF-8 encoding
- Header/footer types include: Default, First, and Even
- Footnotes and endnotes are auto-numbered with support for multiple numbering formats
- Lists support multi-level nesting, up to 9 indent levels
- Table of Contents requires adding titled bookmarks first, then calling the TOC generation method

For more details and examples, see the documentation for each type and function.
*/
package document
