// Package test test TOC update functionality
package test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestTOCUpdate
func TestTOCUpdate(t *testing.T) {
	doc := document.New()

	// configure table of contents
	tocConfig := &document.TOCConfig{
		Title:       "Table of Contents",
		MaxLevel:    3,    // include headings up to which level
		ShowPageNum: true, // show page numbers
		DotLeader:   true, // use dot leader
	}

	// add cover page
	doc.AddParagraph("Cover Page Example")

	// generate table of contents
	err := doc.GenerateTOC(tocConfig)
	if err != nil {
		t.Fatalf("GenerateTOC failed: %v", err)
	}

	// add headings
	doc.AddHeadingParagraph("Chapter 1", 1)
	doc.AddHeadingParagraph("1.1", 2)
	doc.AddHeadingParagraph("Chapter 2", 1)

	// update table of contents
	err = doc.UpdateTOC()
	if err != nil {
		t.Fatalf("UpdateTOC failed: %v", err)
	}

	// save document for inspection
	outputDir := filepath.Join("test", "output")
	err = os.MkdirAll(outputDir, 0755)
	if err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	outputPath := filepath.Join(outputDir, "toc_update_test.docx")

	err = doc.Save(outputPath)
	if err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	t.Logf("document saved to: %s", outputPath)

	// check if headings have been collected
	headings := doc.ListHeadings()
	if len(headings) == 0 {
		t.Error("Expected headings to be collected, but got none")
	}

	expectedHeadings := []struct {
		text  string
		level int
	}{
		{"Chapter 1", 1},
		{"1.1", 2},
		{"Chapter 2", 1},
	}

	if len(headings) != len(expectedHeadings) {
		t.Errorf("Expected %d headings, got %d", len(expectedHeadings), len(headings))
	}

	for i, expected := range expectedHeadings {
		if i < len(headings) {
			if headings[i].Text != expected.text {
				t.Errorf("Heading %d: expected text '%s', got '%s'", i, expected.text, headings[i].Text)
			}
			if headings[i].Level != expected.level {
				t.Errorf("Heading %d: expected level %d, got %d", i, expected.level, headings[i].Level)
			}
		}
	}
}
