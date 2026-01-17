package test

import (
	"os"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

func TestCreateDocument(t *testing.T) {
	doc := document.New()

	doc.AddParagraph("Hello, World!")
	doc.AddParagraph("This is a document created using go-word.")

	err := doc.Save("test_output/test_document.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	if _, err := os.Stat("test_output/test_document.docx"); os.IsNotExist(err) {
		t.Fatal("Document file was not created")
	}

	defer os.RemoveAll("test_output")

	t.Log("Document created successfully")
}

func TestOpenDocument(t *testing.T) {
	doc := document.New()
	doc.AddParagraph("Test paragraph")

	testFile := "test_output/test_open.docx"
	err := doc.Save(testFile)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	openedDoc, err := document.Open(testFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}

	paragraphs := openedDoc.Body.GetParagraphs()
	if len(paragraphs) != 1 {
		t.Fatalf("Expected 1 paragraph, got %d", len(paragraphs))
	}

	if paragraphs[0].Runs[0].Text.Content != "Test paragraph" {
		t.Fatalf("Paragraph content mismatch")
	}

	defer os.RemoveAll("test_output")

	t.Log("Document opened successfully")
}

func TestOpenModifySaveReopen(t *testing.T) {
	doc := document.New()
	doc.AddParagraph("Original paragraph")

	testFile := "test_output/test_modify.docx"
	err := doc.Save(testFile)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	openedDoc, err := document.Open(testFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}

	openedDoc.AddParagraph("Added paragraph")

	modifiedFile := "test_output/test_modify_saved.docx"
	err = openedDoc.Save(modifiedFile)
	if err != nil {
		t.Fatalf("Failed to save modified document: %v", err)
	}

	reopenedDoc, err := document.Open(modifiedFile)
	if err != nil {
		t.Fatalf("Failed to reopen modified document: %v", err)
	}

	paragraphs := reopenedDoc.Body.GetParagraphs()
	if len(paragraphs) != 2 {
		t.Fatalf("Expected 2 paragraphs, got %d", len(paragraphs))
	}

	if paragraphs[0].Runs[0].Text.Content != "Original paragraph" {
		t.Fatalf("First paragraph content mismatch")
	}

	if paragraphs[1].Runs[0].Text.Content != "Added paragraph" {
		t.Fatalf("Second paragraph content mismatch")
	}

	defer os.RemoveAll("test_output")

	t.Log("Document open-modify-save-reopen test passed")
}
