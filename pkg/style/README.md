# Style Package - go-word Style Management System

go-word's style package provides a complete Word document style system with predefined styles, custom styles, and style inheritance.

## Features

### Predefined Styles
- **Heading styles**: Heading1-Heading9 with full hierarchy and navigation pane support
- **Document styles**: Title, Subtitle
- **Paragraph styles**: Normal, Quote, ListParagraph, CodeBlock
- **Character styles**: Emphasis, Strong, CodeChar

### Style Management
- **Style inheritance**: Automatic parent style property merging
- **Custom styles**: Quick creation and management
- **Style validation**: Existence checks and error handling
- **Type classification**: Filter and query by style type

### API
- **StyleManager**: Core style manager for low-level operations
- **QuickStyleAPI**: High-level API for common operations

## Installation

```go
import "github.com/drumkitai/go-word/pkg/style"
```

## Quick Start

### Create Style Manager

```go
// Create style manager (automatically loads predefined styles)
styleManager := style.NewStyleManager()

// Create quick API (recommended)
quickAPI := style.NewQuickStyleAPI(styleManager)

// Get all available styles
allStyles := quickAPI.GetAllStylesInfo()
fmt.Printf("Loaded %d styles\n", len(allStyles))
```

### Use Predefined Styles

```go
// Get specific style
heading1 := styleManager.GetStyle("Heading1")
if heading1 != nil {
    fmt.Printf("Found style: %s\n", heading1.Name.Val)
}

// Get all heading styles
headingStyles := styleManager.GetHeadingStyles()
fmt.Printf("Heading styles: %d\n", len(headingStyles))

// Get style details
styleInfo, err := quickAPI.GetStyleInfo("Heading1")
if err == nil {
    fmt.Printf("Name: %s, Type: %s\n", styleInfo.Name, styleInfo.Type)
}
```

### Apply Styles in Documents

```go
import "github.com/drumkitai/go-word/pkg/document"

doc := document.New()

// Use AddHeadingParagraph (recommended)
doc.AddHeadingParagraph("Chapter 1: Overview", 1)  // Auto-applies Heading1
doc.AddHeadingParagraph("1.1 Background", 2)       // Auto-applies Heading2

// Or manually set style
para := doc.AddParagraph("This is a quote")
para.SetStyle("Quote")

doc.Save("styled_document.docx")
```

## Predefined Styles

### Paragraph Styles

| Style ID | Name | Description |
|----------|------|-------------|
| Normal | Normal | Default paragraph style, Calibri 11pt, 1.15 line spacing |
| Heading1-9 | Heading 1-9 | Heading styles with navigation pane support (Heading1-3) |
| Title | Title | 28pt centered title |
| Subtitle | Subtitle | 15pt centered subtitle |
| Quote | Quote | Italic gray with 720TWIPs left/right indent |
| ListParagraph | List Paragraph | List style with left indent |
| CodeBlock | Code Block | Monospace font with gray background |

### Character Styles

| Style ID | Name | Description |
|----------|------|-------------|
| Emphasis | Emphasis | Italic text |
| Strong | Strong | Bold text |
| CodeChar | Code Character | Red monospace font |

## Creating Custom Styles

### Quick Style Creation

```go
// Create custom paragraph style
config := style.QuickStyleConfig{
    ID:      "MyTitle",
    Name:    "My Title Style",
    Type:    style.StyleTypeParagraph,
    BasedOn: "Normal",
    ParagraphConfig: &style.QuickParagraphConfig{
        Alignment:   "center",
        LineSpacing: 1.5,
        SpaceBefore: 15,
        SpaceAfter:  10,
    },
    RunConfig: &style.QuickRunConfig{
        FontName:  "Times New Roman",
        FontSize:  18,
        FontColor: "2F5496",
        Bold:      true,
    },
}

customStyle, err := quickAPI.CreateQuickStyle(config)
if err != nil {
    log.Printf("Failed to create style: %v", err)
}
```

### Character Style

```go
charConfig := style.QuickStyleConfig{
    ID:   "Highlight",
    Name: "Highlight Text",
    Type: style.StyleTypeCharacter,
    RunConfig: &style.QuickRunConfig{
        FontColor: "FF0000",
        Bold:      true,
        Highlight: "yellow",
    },
}

highlightStyle, err := quickAPI.CreateQuickStyle(charConfig)
```

## Style Query and Management

```go
// Get styles by type
paragraphStyles := quickAPI.GetParagraphStylesInfo()
characterStyles := quickAPI.GetCharacterStylesInfo()
headingStyles := quickAPI.GetHeadingStylesInfo()

// Check if style exists
if styleManager.StyleExists("Heading1") {
    fmt.Println("Heading1 exists")
}

// Get style info
styleInfo, err := quickAPI.GetStyleInfo("CustomStyle")
if err == nil {
    fmt.Printf("Found: %s\n", styleInfo.Name)
}

// Remove style
styleManager.RemoveStyle("MyCustomStyle")

// Reload predefined styles
styleManager.LoadPredefinedStyles()
```

## Style Inheritance

```go
// Get style with inherited properties
fullStyle := styleManager.GetStyleWithInheritance("Heading2")
// Automatically merges Normal + Heading2 properties

// Create inherited style
customHeading := style.QuickStyleConfig{
    ID:      "MyHeading",
    Name:    "My Heading",
    Type:    style.StyleTypeParagraph,
    BasedOn: "Heading1",  // Inherits all Heading1 properties
    RunConfig: &style.QuickRunConfig{
        FontColor: "8B0000",  // Override only color
    },
}

inheritedStyle, _ := quickAPI.CreateQuickStyle(customHeading)
```

## Configuration Reference

### QuickParagraphConfig

```go
type QuickParagraphConfig struct {
    Alignment       string  // "left", "center", "right", "justify"
    LineSpacing     float64 // Line spacing multiplier
    SpaceBefore     int     // Space before (points)
    SpaceAfter      int     // Space after (points)
    FirstLineIndent int     // First line indent (points)
    LeftIndent      int     // Left indent (points)
    RightIndent     int     // Right indent (points)
}
```

**Note**: All spacing values are in points (1 point = 20 TWIPs)

### QuickRunConfig

```go
type QuickRunConfig struct {
    FontName  string // Font name
    FontSize  int    // Font size (points)
    FontColor string // Hex color (e.g., "FF0000" for red, no # prefix)
    Bold      bool   // Bold
    Italic    bool   // Italic
    Underline bool   // Underline
    Strike    bool   // Strikethrough
    Highlight string // Highlight color: "yellow", "green", "cyan", etc.
}
```

## Complete Example

```go
package main

import (
    "fmt"
    "log"
    "github.com/drumkitai/go-word/pkg/document"
    "github.com/drumkitai/go-word/pkg/style"
)

func main() {
    doc := document.New()
    styleManager := doc.GetStyleManager()
    quickAPI := style.NewQuickStyleAPI(styleManager)

    // Create custom styles
    titleConfig := style.QuickStyleConfig{
        ID:      "CustomTitle",
        Name:    "Custom Title",
        Type:    style.StyleTypeParagraph,
        BasedOn: "Title",
        ParagraphConfig: &style.QuickParagraphConfig{
            Alignment:   "center",
            SpaceBefore: 24,
            SpaceAfter:  18,
        },
        RunConfig: &style.QuickRunConfig{
            FontSize:  20,
            FontColor: "1F4E79",
            Bold:      true,
        },
    }
    quickAPI.CreateQuickStyle(titleConfig)

    // Build document
    title := doc.AddParagraph("go-word Style System Guide")
    title.SetStyle("CustomTitle")

    doc.AddHeadingParagraph("1. Overview", 1)
    doc.AddParagraph("go-word provides a complete style management system.")

    quote := doc.AddParagraph("Styles are the core of document formatting.")
    quote.SetStyle("Quote")

    err := doc.Save("styled_document.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

## Testing

```bash
# Run tests
go test ./pkg/style/

# Run with coverage
go test -cover ./pkg/style/

# Run demo
go run ./examples/style_demo/
```

## Related Documentation

- [Main README](../../README.md) - Project overview
- [Document API](../document/) - Core document operations
- [Examples](../../examples/) - Usage examples

## Contributing

Contributions welcome! Please ensure:
1. New styles follow Word standard specifications
2. Include complete test cases
3. Update relevant documentation

## License

This package follows the project's MIT license.
