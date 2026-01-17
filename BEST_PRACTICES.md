# Best Practices
This chapter summarizes best practices for developing with go-word, including performance optimization, error handling, code organization, and solutions to common problems.

## 🚀 Performance Optimization

### 1. Batch Operation Optimization
```golang
// ✅ Recommended: Batch content addition
func createLargeDocument() {
    doc := document.New()
    
    // Pre-allocate slice capacity
    data := make([]string, 0, 1000)
    for i := 0; i < 1000; i++ {
        data = append(data, fmt.Sprintf("Content of paragraph %d", i))
    }
    
    // Batch add paragraphs
    for _, text := range data {
        para := doc.AddParagraph(text)
        para.SetStyle(style.StyleNormal)
    }
}

// ❌ Avoid: Frequent small operations
func inefficientCreation() {
    doc := document.New()
    
    for i := 0; i < 1000; i++ {
        // Reformatting string each time
        text := fmt.Sprintf("Content of paragraph %d", i)
        para := doc.AddParagraph(text)
        para.SetStyle(style.StyleNormal)
        // Saving individually each time (wrong approach)
        // doc.Save(fmt.Sprintf("temp_%d.docx", i))
    }
}
```

### 2. Memory Management
```golang
// ✅ Recommended: Proper document lifecycle management
func processMultipleDocuments(files []string) error {
    for _, file := range files {
        func() {
            doc := document.New()
            defer func() {
                // Ensure document resources are released
                doc = nil
            }()
            
            // Process document
            processDocument(doc, file)
            
            // Save and release immediately
            if err := doc.Save(file); err != nil {
                log.Printf("Failed to save document: %v", err)
            }
        }()
    }
    return nil
}
```

### 3. Style Reuse
```golang
// ✅ Recommended: Create style constants
var (
    HeaderFormat = &document.TextFormat{
        FontName:  "Calibri",
        FontSize:  14,
        Bold:      true,
        FontColor: "2F5496",
    }
    
    BodyFormat = &document.TextFormat{
        FontName: "Calibri",
        FontSize: 11,
    }
    
    EmphasisFormat = &document.TextFormat{
        Bold:      true,
        FontColor: "FF0000",
    }
)

func addFormattedContent(doc *document.Document) {
    // Reuse predefined formats
    title := doc.AddParagraph("")
    title.AddFormattedText("Title", HeaderFormat)
    
    content := doc.AddParagraph("")
    content.AddFormattedText("Body text", BodyFormat)
    content.AddFormattedText("Emphasis", EmphasisFormat)
}
```

## 🔧 Code Organization

### 1. Modular Design
```golang
// Document builder pattern
type DocumentBuilder struct {
    doc    *document.Document
    config *BuilderConfig
}

type BuilderConfig struct {
    Title       string
    Author      string
    PageSize    document.PageSize
    PageMargins [4]int // Top, Right, Bottom, Left
}

func NewDocumentBuilder(config *BuilderConfig) *DocumentBuilder {
    return &DocumentBuilder{
        doc:    document.New(),
        config: config,
    }
}

func (b *DocumentBuilder) SetupPage() *DocumentBuilder {
    b.doc.SetPageSize(b.config.PageSize)
    b.doc.SetPageMargins(
        b.config.PageMargins[0],
        b.config.PageMargins[1], 
        b.config.PageMargins[2],
        b.config.PageMargins[3],
    )
    return b
}

func (b *DocumentBuilder) AddTitle() *DocumentBuilder {
    title := b.doc.AddParagraph(b.config.Title)
    title.SetStyle(style.StyleTitle)
    return b
}

func (b *DocumentBuilder) AddContent(content string) *DocumentBuilder {
    para := b.doc.AddParagraph(content)
    para.SetStyle(style.StyleNormal)
    return b
}

func (b *DocumentBuilder) Build() *document.Document {
    return b.doc
}

// Usage example
func createStandardDocument() {
    config := &BuilderConfig{
        Title:    "Standard Document",
        Author:   "go-word",
        PageSize: document.PageSizeA4,
        PageMargins: [4]int{72, 72, 72, 72},
    }
    
    doc := NewDocumentBuilder(config).
        SetupPage().
        AddTitle().
        AddContent("Document content...").
        Build()
        
    doc.Save("standard_document.docx")
}
```

## 🛡️ Error Handling
### 1. Comprehensive Error Checking
```golang
// ✅ Recommended: Check all errors
func createDocumentSafely(filename string) error {
    doc := document.New()
    
    // Add content with error checking
    title := doc.AddParagraph("Document Title")
    if title == nil {
        return fmt.Errorf("failed to add title paragraph")
    }
    
    // Set style with validation
    if err := title.SetStyle(style.StyleTitle); err != nil {
        return fmt.Errorf("failed to set title style: %w", err)
    }
    
    // Save with error handling
    if err := doc.Save(filename); err != nil {
        return fmt.Errorf("failed to save document to %s: %w", filename, err)
    }
    
    return nil
}
```

### 2. Graceful Degradation
```golang
// Handle missing features gracefully
func addOptionalFeatures(doc *document.Document) {
    // Try to add advanced features, fall back to basic ones
    if err := doc.AddTableOfContents(); err != nil {
        log.Printf("Warning: Could not add TOC, skipping: %v", err)
        // Add simple heading instead
        doc.AddParagraph("Table of Contents").SetStyle(style.StyleHeading1)
    }
}
```

## 📋 Common Issues and Solutions
### 1. File Path Issues
```golang
// ✅ Recommended: Proper path handling
func saveToSafeLocation(doc *document.Document, filename string) error {
    // Ensure directory exists
    dir := filepath.Dir(filename)
    if err := os.MkdirAll(dir, 0755); err != nil {
        return fmt.Errorf("failed to create directory %s: %w", dir, err)
    }
    
    // Clean filename
    cleanName := filepath.Clean(filename)
    
    // Check if file already exists
    if _, err := os.Stat(cleanName); err == nil {
        log.Printf("Warning: File %s already exists, will be overwritten", cleanName)
    }
    
    return doc.Save(cleanName)
}
```

### 2. Resource Management
```golang
// ✅ Recommended: Proper cleanup
func processLargeDataset(data [][]string) error {
    doc := document.New()
    
    // Process in chunks to avoid memory issues
    chunkSize := 100
    for i := 0; i < len(data); i += chunkSize {
        end := i + chunkSize
        if end > len(data) {
            end = len(data)
        }
        
        chunk := data[i:end]
        if err := processChunk(doc, chunk); err != nil {
            return fmt.Errorf("failed to process chunk %d-%d: %w", i, end, err)
        }
        
        // Force garbage collection periodically
        if i%500 == 0 {
            runtime.GC()
        }
    }
    
    return doc.Save("large_dataset.docx")
}
```

## 🎯 Testing Best Practices
## 1. Unit Testing
```golang
func TestDocumentCreation(t *testing.T) {
    doc := document.New()
    
    // Test basic functionality
    para := doc.AddParagraph("Test paragraph")
    assert.NotNil(t, para, "Paragraph should not be nil")
    
    // Test style application
    err := para.SetStyle(style.StyleNormal)
    assert.NoError(t, err, "Should be able to set normal style")
    
    // Test document saving
    tempFile := filepath.Join(t.TempDir(), "test.docx")
    err = doc.Save(tempFile)
    assert.NoError(t, err, "Should be able to save document")
    
    // Verify file exists
    _, err = os.Stat(tempFile)
    assert.NoError(t, err, "Saved file should exist")
}
```

### 2. Integration Testing
```golang
func TestCompleteWorkflow(t *testing.T) {
    // Test complete document creation workflow
    tempDir := t.TempDir()
    filename := filepath.Join(tempDir, "integration_test.docx")
    
    // Create complex document
    doc := document.New()
    
    // Add various content types
    doc.AddParagraph("Title").SetStyle(style.StyleTitle)
    doc.AddParagraph("Subtitle").SetStyle(style.StyleSubtitle)
    
    // Add table
    table := doc.AddTable(3, 3)
    table.SetCellText(0, 0, "Header 1")
    table.SetCellText(0, 1, "Header 2")
    table.SetCellText(0, 2, "Header 3")
    
    // Save and verify
    err := doc.Save(filename)
    assert.NoError(t, err)
    
    // Verify file is valid DOCX
    verifyDocxFile(t, filename)
}
```

Following these best practices will help you create robust, efficient, and maintainable applications with go-word.