package document

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"testing"
)

// TestFloatingImageLeftWithTightWrap
func TestFloatingImageLeftWithTightWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapTight,
		Size: &ImageSize{
			Width:  23.6,
			Height: 13,
		},
		AltText: "左浮动测试图片",
		Title:   "测试",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("添加左浮动图片失败: %v", err)
	}

	filename := "test_float_left_tight.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("✓ left floating + tight wrap image test passed")
}

// TestFloatingImageRightWithSquareWrap
func TestFloatingImageRightWithSquareWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatRight,
		WrapText: ImageWrapSquare,
		Size: &ImageSize{
			Width:  30,
			Height: 20,
		},
		AltText: "right floating test image",
		Title:   "test",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add right floating image: %v", err)
	}

	filename := "test_float_right_square.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("打开保存的文档失败: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("文档中没有元素")
	}

	t.Logf("✓ 右浮动 + 四周环绕图片测试通过")
}

// TestFloatingImageWithTopAndBottomWrap
func TestFloatingImageWithTopAndBottomWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapTopAndBottom,
		Size: &ImageSize{
			Width:  40,
			Height: 30,
		},
		AltText: "top and bottom wrap test image",
		Title:   "test",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add floating image: %v", err)
	}

	filename := "test_float_topbottom.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("✓ top and bottom wrap image test passed")
}

// TestFloatingImageWithNoWrap
func TestFloatingImageWithNoWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatRight,
		WrapText: ImageWrapNone,
		Size: &ImageSize{
			Width:  25,
			Height: 15,
		},
		AltText: "no wrap test image",
		Title:   "test",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add floating image: %v", err)
	}

	filename := "test_float_nowrap.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("✓ no wrap image test passed")
}

// TestMultipleFloatingImages
func TestMultipleFloatingImages(t *testing.T) {
	doc := New()

	doc.AddParagraph("这是一个包含多个浮动图片的文档测试。")

	imageData := createTestImageRGBA(80, 80)

	config1 := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapSquare,
		Size: &ImageSize{
			Width:  20,
			Height: 20,
		},
	}
	_, err := doc.AddImageFromData(imageData, "test1.png", ImageFormatPNG, 80, 80, config1)
	if err != nil {
		t.Fatalf("failed to add first floating image: %v", err)
	}

	doc.AddParagraph("first image added.")

	config2 := &ImageConfig{
		Position: ImagePositionFloatRight,
		WrapText: ImageWrapTight,
		Size: &ImageSize{
			Width:  20,
			Height: 20,
		},
	}
	_, err = doc.AddImageFromData(imageData, "test2.png", ImageFormatPNG, 80, 80, config2)
	if err != nil {
		t.Fatalf("failed to add second floating image: %v", err)
	}

	doc.AddParagraph("second image also added.")

	filename := "test_multiple_float.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	const expectedElements = 5 // 3 text paragraphs + 2 image paragraphs
	if len(doc2.Body.Elements) != expectedElements {
		t.Errorf("expected %d elements (3 text paragraphs + 2 image paragraphs), actual %d elements", expectedElements, len(doc2.Body.Elements))
	}

	t.Logf("✓ multiple floating images test passed")
}

// createTestImageRGBA
func createTestImageRGBA(width, height int) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			r := uint8(x * 255 / width)
			g := uint8(y * 255 / height)
			b := uint8((x + y) * 255 / (width + height))
			img.SetRGBA(x, y, color.RGBA{r, g, b, 255})
		}
	}

	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}

// TestInlineImageNotAffected
func TestInlineImageNotAffected(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position:  ImagePositionInline,
		Alignment: AlignCenter,
		Size: &ImageSize{
			Width:  30,
			Height: 30,
		},
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add inline image: %v", err)
	}

	filename := "test_inline_image.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("✓ inline image not affected test passed")
}
