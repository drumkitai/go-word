package document

import (
	"os"
	"strconv"
	"strings"
	"testing"
)

// TestChineseFilename tests whether images with Chinese filenames are saved and opened correctly
func TestChineseFilename(t *testing.T) {
	doc := New()
	doc.AddParagraph("Test Chinese filename")

	imageData := createTestImage(100, 75)

	_, err := doc.AddImageFromData(
		imageData,
		"test image.png",
		ImageFormatPNG,
		100, 75,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
			AltText:   "test image",
			Title:     "test image title",
		},
	)
	if err != nil {
		t.Fatalf("failed to add image with Chinese filename: %v", err)
	}

	doc.AddParagraph("Text below image")

	testFile := "test_chinese_filename.docx"
	err = doc.Save(testFile)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile)

	foundSafeFilename := false
	foundChineseFilename := false
	for partName := range doc.parts {
		if strings.Contains(partName, "word/media/") {
			if strings.Contains(partName, "image0.png") {
				foundSafeFilename = true
			}
			if strings.Contains(partName, "test") {
				foundChineseFilename = true
			}
			t.Logf("found image: %s", partName)
		}
	}

	if !foundSafeFilename {
		t.Error("safe filename (image0.png) not found, Chinese filename conversion failed")
	}

	if foundChineseFilename {
		t.Error("Chinese filename found, should have been converted to safe ASCII filename")
	}

	foundImageRelationship := false
	for _, rel := range doc.documentRelationships.Relationships {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
			foundImageRelationship = true
			if !strings.Contains(rel.Target, "image0.png") {
				t.Errorf("image relationship does not use safe filename, Target=%s", rel.Target)
			}
			if strings.Contains(rel.Target, "test") {
				t.Errorf("image relationship contains Chinese characters, Target=%s", rel.Target)
			}
			t.Logf("image relationship: ID=%s, Target=%s", rel.ID, rel.Target)
			break
		}
	}

	if !foundImageRelationship {
		t.Error("image relationship not found")
	}

	doc2, err := Open(testFile)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	if _, exists := doc2.parts["word/media/image0.png"]; !exists {
		t.Error("image data not found in opened document")
	}

	t.Log("✓ Chinese filename test passed: automatically converted to safe ASCII filename")
}

// TestMultipleNonASCIIFilenames tests multiple non-ASCII filenames
func TestMultipleNonASCIIFilenames(t *testing.T) {
	doc := New()
	doc.AddParagraph("Test multiple non-ASCII filenames")

	imageData := createTestImage(50, 50)

	testFilenames := []string{
		"Chinese image.png",
		"Japanese.png",
		"Korean.png",
		"Russian.png",
		"Arabic.png",
	}

	for i, filename := range testFilenames {
		_, err := doc.AddImageFromData(imageData, filename, ImageFormatPNG, 50, 50, nil)
		if err != nil {
			t.Fatalf("failed to add image %s: %v", filename, err)
		}

		expectedSafeFilename := "image" + strconv.Itoa(i) + ".png"
		if _, exists := doc.parts["word/media/"+expectedSafeFilename]; !exists {
			t.Errorf("image %s does not use safe filename %s", filename, expectedSafeFilename)
		}
	}

	testFile := "test_multiple_nonascii_filenames.docx"
	err := doc.Save(testFile)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile)

	doc2, err := Open(testFile)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	for i := 0; i < len(testFilenames); i++ {
		expectedSafeFilename := "image" + strconv.Itoa(i) + ".png"
		if _, exists := doc2.parts["word/media/"+expectedSafeFilename]; !exists {
			t.Errorf("image %s not found in opened document", expectedSafeFilename)
		}
	}

	t.Logf("✓ Multi-language filename test passed: all %d images correctly converted", len(testFilenames))
}
