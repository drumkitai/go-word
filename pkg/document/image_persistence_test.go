package document

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"testing"
)

// TestImagePersistenceAfterOpenAndSave
func TestImagePersistenceAfterOpenAndSave(t *testing.T) {
	doc1 := New()
	doc1.AddParagraph("test document - image persistence test")

	imageData := createTestImageForPersistence(100, 75, color.RGBA{255, 100, 100, 255})

	imageInfo, err := doc1.AddImageFromData(
		imageData,
		"test_image.png",
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
		t.Fatalf("failed to add image: %v", err)
	}

	doc1.AddParagraph("text below the image")

	testFile1 := "test_image_persistence_1.docx"
	err = doc1.Save(testFile1)
	if err != nil {
		t.Fatalf("failed to save first document: %v", err)
	}
	defer os.Remove(testFile1)

	if _, exists := doc1.parts["word/media/image0.png"]; !exists {
		t.Fatal("first document has no image data")
	}

	doc2, err := Open(testFile1)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	if _, exists := doc2.parts["word/media/image0.png"]; !exists {
		t.Fatal("opened document has no image data")
	}

	foundImageRelationship := false
	for _, rel := range doc2.documentRelationships.Relationships {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
			foundImageRelationship = true
			t.Logf("found image relationship: ID=%s, Target=%s", rel.ID, rel.Target)
			break
		}
	}
	if !foundImageRelationship {
		t.Fatal("opened document has no image relationship")
	}

	doc2.AddParagraph("this is a new paragraph")

	testFile2 := "test_image_persistence_2.docx"
	err = doc2.Save(testFile2)
	if err != nil {
		t.Fatalf("failed to save second document: %v", err)
	}
	defer os.Remove(testFile2)

	doc3, err := Open(testFile2)
	if err != nil {
		t.Fatalf("failed to open second document: %v", err)
	}

	if _, exists := doc3.parts["word/media/image0.png"]; !exists {
		t.Fatal("second document has no image data - image lost after saving")
	}

	foundImageRelationship = false
	for _, rel := range doc3.documentRelationships.Relationships {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
			foundImageRelationship = true
			t.Logf("found image relationship in second document: ID=%s, Target=%s", rel.ID, rel.Target)
			break
		}
	}
	if !foundImageRelationship {
		t.Fatal("second document has no image relationship - image relationship lost after saving")
	}

	originalImageData := doc1.parts["word/media/image0.png"]
	finalImageData := doc3.parts["word/media/image0.png"]

	if !bytes.Equal(originalImageData, finalImageData) {
		t.Fatal("image data changed after saving and reopening")
	}

	t.Log("✓ image persistence test passed: image exists after modifying and saving document")
	t.Logf("✓ original image information: ID=%s, format=%s, size=%dx%d",
		imageInfo.ID, imageInfo.Format, imageInfo.Width, imageInfo.Height)
}

// TestAddImageToOpenedDocument
func TestAddImageToOpenedDocument(t *testing.T) {
	doc1 := New()
	doc1.AddParagraph("original document")

	imageData1 := createTestImageForPersistence(100, 75, color.RGBA{255, 0, 0, 255})
	_, err := doc1.AddImageFromData(
		imageData1,
		"image1.png",
		ImageFormatPNG,
		100, 75,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
		},
	)
	if err != nil {
		t.Fatalf("failed to add first image: %v", err)
	}

	testFile1 := "test_add_image_to_opened_1.docx"
	err = doc1.Save(testFile1)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile1)

	doc2, err := Open(testFile1)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	doc2.AddParagraph("add second image")

	imageData2 := createTestImageForPersistence(100, 75, color.RGBA{0, 0, 255, 255})
	_, err = doc2.AddImageFromData(
		imageData2,
		"image2.png",
		ImageFormatPNG,
		100, 75,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
		},
	)
	if err != nil {
		t.Fatalf("failed to add second image: %v", err)
	}

	testFile2 := "test_add_image_to_opened_2.docx"
	err = doc2.Save(testFile2)
	if err != nil {
		t.Fatalf("failed to save document with two images: %v", err)
	}
	defer os.Remove(testFile2)

	doc3, err := Open(testFile2)
	if err != nil {
		t.Fatalf("failed to open document with two images: %v", err)
	}

	if _, exists := doc3.parts["word/media/image0.png"]; !exists {
		t.Fatal("first image data lost")
	}

	if _, exists := doc3.parts["word/media/image1.png"]; !exists {
		t.Fatal("second image data lost")
	}

	imageRelCount := 0
	for _, rel := range doc3.documentRelationships.Relationships {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
			imageRelCount++
			t.Logf("found image relationship: ID=%s, Target=%s", rel.ID, rel.Target)
		}
	}

	if imageRelCount != 2 {
		t.Fatalf("expected 2 image relationships, actual %d", imageRelCount)
	}

	t.Log("✓ image added to opened document test passed: both images are correctly saved")
}

// TestImageIDCounterAfterOpen
func TestImageIDCounterAfterOpen(t *testing.T) {
	doc1 := New()
	doc1.AddParagraph("test image ID counter")

	imageData := createTestImageForPersistence(50, 50, color.RGBA{255, 0, 0, 255})

	_, err := doc1.AddImageFromData(imageData, "img1.png", ImageFormatPNG, 50, 50, nil)
	if err != nil {
		t.Fatalf("failed to add first image: %v", err)
	}

	_, err = doc1.AddImageFromData(imageData, "img2.png", ImageFormatPNG, 50, 50, nil)
	if err != nil {
		t.Fatalf("failed to add second image: %v", err)
	}

	testFile := "test_image_id_counter.docx"
	err = doc1.Save(testFile)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile)

	doc2, err := Open(testFile)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	if doc2.nextImageID < 2 {
		t.Fatalf("nextImageID not updated correctly: expected >= 2, actual = %d", doc2.nextImageID)
	}

	t.Logf("✓ nextImageID = %d (expected)", doc2.nextImageID)

	_, err = doc2.AddImageFromData(imageData, "img3.png", ImageFormatPNG, 50, 50, nil)
	if err != nil {
		t.Fatalf("failed to add third image: %v", err)
	}

	testFile2 := "test_image_id_counter_2.docx"
	err = doc2.Save(testFile2)
	if err != nil {
		t.Fatalf("failed to save document with three images: %v", err)
	}
	defer os.Remove(testFile2)

	doc3, err := Open(testFile2)
	if err != nil {
		t.Fatalf("failed to open document with three images: %v", err)
	}

	images := []string{"image0.png", "image1.png", "image2.png"}
	for _, imgName := range images {
		if _, exists := doc3.parts["word/media/"+imgName]; !exists {
			t.Fatalf("image %s lost", imgName)
		}
	}

	imageRelCount := 0
	for _, rel := range doc3.documentRelationships.Relationships {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
			imageRelCount++
		}
	}

	if imageRelCount != 3 {
		t.Fatalf("expected 3 image relationships, actual %d", imageRelCount)
	}

	t.Log("✓ image ID counter test passed: all image IDs are correct and conflict-free")
}

// createTestImageForPersistence
func createTestImageForPersistence(width, height int, bgColor color.RGBA) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, bgColor)
		}
	}

	borderColor := color.RGBA{0, 0, 0, 255}
	for x := 0; x < width; x++ {
		img.Set(x, 0, borderColor)
		img.Set(x, height-1, borderColor)
	}
	for y := 0; y < height; y++ {
		img.Set(0, y, borderColor)
		img.Set(width-1, y, borderColor)
	}

	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}
