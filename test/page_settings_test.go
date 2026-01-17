// Package test Page settings integration test
package test

import (
	"fmt"
	"os"
	"path/filepath"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestPageSettingsIntegration
func TestPageSettingsIntegration(t *testing.T) {
	doc := document.New()

	t.Run("Basic page settings", func(t *testing.T) {
		testBasicPageSettings(t, doc)
	})

	t.Run("Page sizes settings", func(t *testing.T) {
		testPageSizes(t, doc)
	})

	t.Run("Page orientation settings", func(t *testing.T) {
		testPageOrientation(t, doc)
	})

	t.Run("Page margins settings", func(t *testing.T) {
		testPageMargins(t, doc)
	})

	t.Run("Custom page size settings", func(t *testing.T) {
		testCustomPageSize(t, doc)
	})

	t.Run("Document save and load settings", func(t *testing.T) {
		testDocumentSaveLoad(t, doc)
	})
}

// testBasicPageSettings
func testBasicPageSettings(t *testing.T, doc *document.Document) {
	settings := doc.GetPageSettings()

	if settings.Size != document.PageSizeA4 {
		t.Errorf("默认页面尺寸应为A4，实际为: %s", settings.Size)
	}

	if settings.Orientation != document.OrientationPortrait {
		t.Errorf("默认页面方向应为纵向，实际为: %s", settings.Orientation)
	}

	doc.AddParagraph("页面设置集成测试 - 基本设置")
}

// testPageSizes
func testPageSizes(t *testing.T, doc *document.Document) {
	sizes := []document.PageSize{
		document.PageSizeLetter,
		document.PageSizeLegal,
		document.PageSizeA3,
		document.PageSizeA5,
		document.PageSizeA4,
	}

	for _, size := range sizes {
		err := doc.SetPageSize(size)
		if err != nil {
			t.Errorf("设置页面尺寸 %s 失败: %v", size, err)
			continue
		}

		settings := doc.GetPageSettings()
		if settings.Size != size {
			t.Errorf("页面尺寸设置不正确，期望: %s, 实际: %s", size, settings.Size)
		}

		doc.AddParagraph(fmt.Sprintf("页面尺寸已设置为: %s", size))
	}
}

// testPageOrientation
func testPageOrientation(t *testing.T, doc *document.Document) {
	err := doc.SetPageOrientation(document.OrientationLandscape)
	if err != nil {
		t.Errorf("设置横向页面失败: %v", err)
		return
	}

	settings := doc.GetPageSettings()
	if settings.Orientation != document.OrientationLandscape {
		t.Errorf("页面方向应为横向，实际为: %s", settings.Orientation)
	}

	doc.AddParagraph("页面方向已设置为横向")

	err = doc.SetPageOrientation(document.OrientationPortrait)
	if err != nil {
		t.Errorf("设置纵向页面失败: %v", err)
		return
	}

	settings = doc.GetPageSettings()
	if settings.Orientation != document.OrientationPortrait {
		t.Errorf("页面方向应为纵向，实际为: %s", settings.Orientation)
	}

	doc.AddParagraph("页面方向已恢复为纵向")
}

// testPageMargins
func testPageMargins(t *testing.T, doc *document.Document) {
	top, right, bottom, left := 30.0, 20.0, 25.0, 35.0
	err := doc.SetPageMargins(top, right, bottom, left)
	if err != nil {
		t.Errorf("设置页面边距失败: %v", err)
		return
	}

	settings := doc.GetPageSettings()
	if abs(settings.MarginTop-top) > 0.1 {
		t.Errorf("上边距不匹配，期望: %.1fmm, 实际: %.1fmm", top, settings.MarginTop)
	}
	if abs(settings.MarginRight-right) > 0.1 {
		t.Errorf("右边距不匹配，期望: %.1fmm, 实际: %.1fmm", right, settings.MarginRight)
	}
	if abs(settings.MarginBottom-bottom) > 0.1 {
		t.Errorf("下边距不匹配，期望: %.1fmm, 实际: %.1fmm", bottom, settings.MarginBottom)
	}
	if abs(settings.MarginLeft-left) > 0.1 {
		t.Errorf("左边距不匹配，期望: %.1fmm, 实际: %.1fmm", left, settings.MarginLeft)
	}

	doc.AddParagraph("页面边距已设置为自定义值")

	header, footer := 15.0, 20.0
	err = doc.SetHeaderFooterDistance(header, footer)
	if err != nil {
		t.Errorf("设置页眉页脚距离失败: %v", err)
		return
	}

	settings = doc.GetPageSettings()
	if abs(settings.HeaderDistance-header) > 0.1 {
		t.Errorf("页眉距离不匹配，期望: %.1fmm, 实际: %.1fmm", header, settings.HeaderDistance)
	}
	if abs(settings.FooterDistance-footer) > 0.1 {
		t.Errorf("页脚距离不匹配，期望: %.1fmm, 实际: %.1fmm", footer, settings.FooterDistance)
	}

	gutter := 8.0
	err = doc.SetGutterWidth(gutter)
	if err != nil {
		t.Errorf("设置装订线宽度失败: %v", err)
		return
	}

	settings = doc.GetPageSettings()
	if abs(settings.GutterWidth-gutter) > 0.1 {
		t.Errorf("装订线宽度不匹配，期望: %.1fmm, 实际: %.1fmm", gutter, settings.GutterWidth)
	}

	doc.AddParagraph("页眉页脚距离和装订线已设置")
}

// testCustomPageSize
func testCustomPageSize(t *testing.T, doc *document.Document) {
	width, height := 200.0, 250.0
	err := doc.SetCustomPageSize(width, height)
	if err != nil {
		t.Errorf("设置自定义页面尺寸失败: %v", err)
		return
	}

	settings := doc.GetPageSettings()
	if settings.Size != document.PageSizeCustom {
		t.Errorf("页面尺寸应为Custom，实际为: %s", settings.Size)
	}
	if abs(settings.CustomWidth-width) > 0.1 {
		t.Errorf("自定义宽度不匹配，期望: %.1fmm, 实际: %.1fmm", width, settings.CustomWidth)
	}
	if abs(settings.CustomHeight-height) > 0.1 {
		t.Errorf("自定义高度不匹配，期望: %.1fmm, 实际: %.1fmm", height, settings.CustomHeight)
	}

	doc.AddParagraph("页面已设置为自定义尺寸")

	err = doc.SetPageSize(document.PageSizeA4)
	if err != nil {
		t.Errorf("恢复A4页面尺寸失败: %v", err)
	}
}

// testDocumentSaveLoad
func testDocumentSaveLoad(t *testing.T, doc *document.Document) {
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
		t.Errorf("设置完整页面配置失败: %v", err)
		return
	}

	doc.AddParagraph("页面设置集成测试完成 - 最终配置")
	doc.AddParagraph("文档将以Letter横向格式保存")

	testFile := filepath.Join("testdata", "page_settings_integration_test.docx")

	err = os.MkdirAll(filepath.Dir(testFile), 0755)
	if err != nil {
		t.Errorf("创建测试目录失败: %v", err)
		return
	}

	err = doc.Save(testFile)
	if err != nil {
		t.Errorf("保存测试文档失败: %v", err)
		return
	}

	// verify file exists
	if _, err := os.Stat(testFile); os.IsNotExist(err) {
		t.Errorf("保存的文档文件不存在: %s", testFile)
		return
	}

	loadedDoc, err := document.Open(testFile)
	if err != nil {
		t.Errorf("重新打开文档失败: %v", err)
		return
	}

	loadedSettings := loadedDoc.GetPageSettings()

	if loadedSettings.Size != settings.Size {
		t.Errorf("加载后页面尺寸不匹配，期望: %s, 实际: %s", settings.Size, loadedSettings.Size)
	}

	if loadedSettings.Orientation != settings.Orientation {
		t.Errorf("加载后页面方向不匹配，期望: %s, 实际: %s", settings.Orientation, loadedSettings.Orientation)
	}

	os.Remove(testFile)
}

// abs
func abs(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}
