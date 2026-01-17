package test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

func TestTextFormatting(t *testing.T) {
	// set test log level
	document.SetGlobalLevel(document.LogLevelWarn)

	// create test document
	doc := document.New()

	// test basic formatting
	format := &document.TextFormat{
		Bold:       true,
		Italic:     true,
		FontSize:   14,
		FontColor:  "FF0000",
		FontFamily: "Arial",
	}

	p := doc.AddFormattedParagraph("test formatted text", format)
	if p == nil {
		t.Error("AddFormattedParagraph returned nil")
	}

	// check if paragraph is correctly added
	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) != 1 {
		t.Errorf("expected 1 paragraph, but got %d", len(paragraphs))
	}

	// check run properties
	if len(paragraphs[0].Runs) == 0 {
		t.Error("paragraph has no runs")
	} else {
		run := paragraphs[0].Runs[0]
		if run.Properties == nil {
			t.Error("run properties are empty")
		} else {
			if run.Properties.Bold == nil {
				t.Error("bold property is not set")
			}
			if run.Properties.Italic == nil {
				t.Error("italic property is not set")
			}
			if run.Properties.FontSize == nil {
				t.Error("font size property is not set")
			}
			if run.Properties.Color == nil {
				t.Error("color property is not set")
			}
			if run.Properties.FontFamily == nil {
				t.Error("font family property is not set")
			}
		}
	}
}

func TestParagraphAlignment(t *testing.T) {
	doc := document.New()

	// test various alignment types
	alignments := []document.AlignmentType{
		document.AlignLeft,
		document.AlignCenter,
		document.AlignRight,
		document.AlignJustify,
	}

	for _, align := range alignments {
		p := doc.AddParagraph("test alignment")
		p.SetAlignment(align)

		if p.Properties == nil {
			t.Errorf("paragraph properties are empty, alignment: %s", align)
		} else if p.Properties.Justification == nil {
			t.Errorf("alignment property is empty, alignment: %s", align)
		} else if p.Properties.Justification.Val != string(align) {
			t.Errorf("alignment does not match, expected: %s, actual: %s", align, p.Properties.Justification.Val)
		}
	}
}

func TestParagraphSpacing(t *testing.T) {
	doc := document.New()

	p := doc.AddParagraph("test spacing")

	config := &document.SpacingConfig{
		LineSpacing:     1.5,
		BeforePara:      12,
		AfterPara:       6,
		FirstLineIndent: 24,
	}

	p.SetSpacing(config)

	if p.Properties == nil {
		t.Error("paragraph properties are empty")
		return
	}

	if p.Properties.Spacing == nil {
		t.Error("spacing property is empty")
	} else {
		// 检查间距值（TWIPs单位）
		if p.Properties.Spacing.Before != "240" { // 12 * 20
			t.Errorf("before spacing is incorrect, expected: 240, actual: %s", p.Properties.Spacing.Before)
		}
		if p.Properties.Spacing.After != "120" { // 6 * 20
			t.Errorf("after spacing is incorrect, expected: 120, actual: %s", p.Properties.Spacing.After)
		}
		if p.Properties.Spacing.Line != "360" { // 1.5 * 240
			t.Errorf("line spacing is incorrect, expected: 360, actual: %s", p.Properties.Spacing.Line)
		}
	}

	if p.Properties.Indentation == nil {
		t.Error("indentation property is empty")
	} else {
		if p.Properties.Indentation.FirstLine != "480" { // 24 * 20
			t.Errorf("first line indentation is incorrect, expected: 480, actual: %s", p.Properties.Indentation.FirstLine)
		}
	}
}

func TestAddFormattedText(t *testing.T) {
	doc := document.New()

	p := doc.AddParagraph("basic text")

	// add formatted text
	format := &document.TextFormat{
		Bold:      true,
		FontColor: "0000FF",
	}

	p.AddFormattedText("additional formatted text", format)

	if len(p.Runs) != 2 {
		t.Errorf("expected 2 runs, but got %d", len(p.Runs))
	}

	// check second run's properties
	if len(p.Runs) >= 2 {
		run := p.Runs[1]
		if run.Properties == nil {
			t.Error("second run's properties are empty")
		} else {
			if run.Properties.Bold == nil {
				t.Error("second run's bold property is not set")
			}
			if run.Properties.Color == nil {
				t.Error("second run's color property is not set")
			}
		}
	}
}

func TestDocumentSaveAndOpen(t *testing.T) {
	// create temporary directory
	tempDir := filepath.Join(os.TempDir(), "go-word_test")
	os.MkdirAll(tempDir, 0755)
	defer os.RemoveAll(tempDir)

	filename := filepath.Join(tempDir, "test_formatted.docx")

	// create formatted document
	doc := document.New()

	format := &document.TextFormat{
		Bold:     true,
		FontSize: 16,
	}

	p := doc.AddFormattedParagraph("test save and load", format)
	p.SetAlignment(document.AlignCenter)

	// save document
	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("save document failed: %v", err)
	}

	// check if file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Fatal("saved file does not exist")
	}

	// reopen document
	openedDoc, err := document.Open(filename)
	if err != nil {
		t.Fatalf("open document failed: %v", err)
	}

	// check content
	paragraphs := openedDoc.Body.GetParagraphs()
	if len(paragraphs) != 1 {
		t.Errorf("expected 1 paragraph, but got %d", len(paragraphs))
	}

	if len(paragraphs) > 0 {
		para := paragraphs[0]

		// check alignment
		if para.Properties == nil || para.Properties.Justification == nil {
			t.Error("paragraph alignment property is lost")
		} else if para.Properties.Justification.Val != string(document.AlignCenter) {
			t.Errorf("alignment does not match, expected: %s, actual: %s",
				document.AlignCenter, para.Properties.Justification.Val)
		}

		// check text content
		if len(para.Runs) > 0 {
			if para.Runs[0].Text.Content != "test save and load" {
				t.Errorf("text content does not match, expected: %s, actual: %s",
					"test save and load", para.Runs[0].Text.Content)
			}

			// check format properties
			if para.Runs[0].Properties == nil {
				t.Error("run properties are lost")
			} else {
				if para.Runs[0].Properties.Bold == nil {
					t.Error("bold property is lost")
				}
				if para.Runs[0].Properties.FontSize == nil {
					t.Error("font size property is lost")
				}
			}
		}
	}
}
