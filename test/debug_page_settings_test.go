package test

import (
	"fmt"
	"math"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

func TestDebugPageSettings(t *testing.T) {
	doc := document.New()

	settings := &document.PageSettings{
		Size:           document.PageSizeLetter,
		Orientation:    document.OrientationLandscape,
		MarginTop:      25,
		MarginRight:    20,
		MarginBottom:   30,
		MarginLeft:     25,
		HeaderDistance: 12,
		FooterDistance: 15,
		GutterWidth:    5,
	}

	err := doc.SetPageSettings(settings)
	if err != nil {
		t.Fatalf("Failed to set page settings: %v", err)
	}

	currentSettings := doc.GetPageSettings()
	fmt.Printf("After page settings:\n")
	fmt.Printf("  Size: %s\n", currentSettings.Size)
	fmt.Printf("  Orientation: %s\n", currentSettings.Orientation)

	doc.AddParagraph("Test page settings save and load")

	// 保存文档
	testFile := "debug_page_settings.docx"
	err = doc.Save(testFile)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	fmt.Printf("Document saved to: %s\n", testFile)

	loadedDoc, err := document.Open(testFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}

	fmt.Printf("Number of Body.Elements after loading: %d\n", len(loadedDoc.Body.Elements))
	for i, element := range loadedDoc.Body.Elements {
		switch elem := element.(type) {
		case *document.SectionProperties:
			fmt.Printf("  Element %d: SectionProperties found!\n", i)
			if elem.PageSize != nil {
				fmt.Printf("    PageSize: w=%s, h=%s, orient=%s\n", elem.PageSize.W, elem.PageSize.H, elem.PageSize.Orient)
			} else {
				fmt.Printf("    PageSize: nil\n")
			}
		case *document.Paragraph:
			fmt.Printf("  Element %d: Paragraph\n", i)
		default:
			fmt.Printf("  Element %d: Other type (%T)\n", i, element)
		}
	}

	loadedSettings := loadedDoc.GetPageSettings()
	fmt.Printf("After page settings:\n")
	fmt.Printf("  Size: %s\n", loadedSettings.Size)
	fmt.Printf("  Orientation: %s\n", loadedSettings.Orientation)

	if loadedSettings.Size != settings.Size {
		t.Errorf("After loading, page size does not match: expected %s, got %s", settings.Size, loadedSettings.Size)
	}

	if loadedSettings.Orientation != settings.Orientation {
		t.Errorf("After loading, page orientation does not match: expected %s, got %s", settings.Orientation, loadedSettings.Orientation)
	}

	parts := loadedDoc.GetParts()
	if docXML, exists := parts["word/document.xml"]; exists {
		fmt.Printf("First 500 characters of document.xml:\n%s\n", string(docXML)[:min(500, len(docXML))])

		fmt.Printf("Debug page size conversion:\n")

		// Letter dimensions: 215.9mm x 279.4mm
		// After landscape, it should be: 279.4mm x 215.9mm
		// Conversion to twips: 279.4 * 56.69 ≈ 15840, 215.9 * 56.69 ≈ 12240

		width_twips := 15840.0
		height_twips := 12240.0
		width_mm := width_twips / 56.692913385827
		height_mm := height_twips / 56.692913385827

		fmt.Printf("  Read from XML: width=%d twips, height=%d twips\n", int(width_twips), int(height_twips))
		fmt.Printf("  Convert to millimeters: width=%.1fmm, height=%.1fmm\n", width_mm, height_mm)

		fmt.Printf("  Letter portrait dimensions: 215.9mm x 279.4mm\n")
		fmt.Printf("  Letter landscape dimensions: 279.4mm x 215.9mm\n")
		fmt.Printf("  Actual parsed dimensions: %.1fmm x %.1fmm\n", width_mm, height_mm)

		tolerance := 1.0
		letter_width := 215.9
		letter_height := 279.4

		landscape_match := (math.Abs(width_mm-letter_height) < tolerance && math.Abs(height_mm-letter_width) < tolerance)
		fmt.Printf("  Landscape Letter match: %t (tolerance=%.1fmm)\n", landscape_match, tolerance)

		portrait_match := (math.Abs(width_mm-letter_width) < tolerance && math.Abs(height_mm-letter_height) < tolerance)
		fmt.Printf("  Portrait Letter match: %t (tolerance=%.1fmm)\n", portrait_match, tolerance)
	} else {
		fmt.Printf("document.xml not found\n")
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
