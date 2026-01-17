package style

import (
	"testing"
)

// TestNewStyleManager 测试样式管理器创建
func TestNewStyleManager(t *testing.T) {
	sm := NewStyleManager()

	if sm == nil {
		t.Fatal("StyleManager should not be nil")
	}

	// 验证预定义样式是否加载
	styles := sm.GetAllStyles()
	if len(styles) == 0 {
		t.Error("Should have predefined styles loaded")
	}

	// 验证基本样式存在
	expectedStyles := []string{"Normal", "Heading1", "Heading2", "Title", "Subtitle"}
	for _, styleID := range expectedStyles {
		if !sm.StyleExists(styleID) {
			t.Errorf("Style %s should exist", styleID)
		}
	}
}

// TestStyleExists 测试样式存在性检查
func TestStyleExists(t *testing.T) {
	sm := NewStyleManager()

	// 测试存在的样式
	if !sm.StyleExists("Normal") {
		t.Error("Normal style should exist")
	}

	if !sm.StyleExists("Heading1") {
		t.Error("Heading1 style should exist")
	}

	// 测试不存在的样式
	if sm.StyleExists("NonExistentStyle") {
		t.Error("NonExistentStyle should not exist")
	}
}

// TestGetStyle 测试获取样式
func TestGetStyle(t *testing.T) {
	sm := NewStyleManager()

	// 测试获取存在的样式
	normalStyle := sm.GetStyle("Normal")
	if normalStyle == nil {
		t.Fatal("Normal style should not be nil")
	}

	if normalStyle.StyleID != "Normal" {
		t.Errorf("Expected StyleID Normal, got %s", normalStyle.StyleID)
	}

	// 测试获取不存在的样式
	nonExistent := sm.GetStyle("NonExistentStyle")
	if nonExistent != nil {
		t.Error("NonExistentStyle should return nil")
	}
}

// TestGetHeadingStyles 测试获取标题样式
func TestGetHeadingStyles(t *testing.T) {
	sm := NewStyleManager()

	headingStyles := sm.GetHeadingStyles()

	// 应该有9个标题样式
	if len(headingStyles) != 9 {
		t.Errorf("Expected 9 heading styles, got %d", len(headingStyles))
	}

	// 验证标题样式ID
	expectedHeadings := []string{"Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Heading7", "Heading8", "Heading9"}
	styleMap := make(map[string]bool)
	for _, style := range headingStyles {
		styleMap[style.StyleID] = true
	}

	for _, expected := range expectedHeadings {
		if !styleMap[expected] {
			t.Errorf("Heading style %s should be included", expected)
		}
	}
}

// TestAddStyle 测试添加自定义样式
func TestAddStyle(t *testing.T) {
	sm := NewStyleManager()

	customStyle := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "CustomTest",
		Name:    &StyleName{Val: "测试样式"},
		RunPr: &RunProperties{
			Bold:  &Bold{},
			Color: &Color{Val: "FF0000"},
		},
	}

	sm.AddStyle(customStyle)

	// 验证样式添加
	if !sm.StyleExists("CustomTest") {
		t.Error("Custom style should exist after adding")
	}

	// 验证样式内容
	retrieved := sm.GetStyle("CustomTest")
	if retrieved == nil {
		t.Fatal("Retrieved custom style should not be nil")
	}

	if retrieved.StyleID != "CustomTest" {
		t.Errorf("Expected StyleID CustomTest, got %s", retrieved.StyleID)
	}

	if retrieved.Name.Val != "测试样式" {
		t.Errorf("Expected name 测试样式, got %s", retrieved.Name.Val)
	}
}

// TestRemoveStyle 测试移除样式
func TestRemoveStyle(t *testing.T) {
	sm := NewStyleManager()

	// 先添加一个测试样式
	testStyle := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "TestRemove",
		Name:    &StyleName{Val: "待删除样式"},
	}

	sm.AddStyle(testStyle)

	// 验证样式存在
	if !sm.StyleExists("TestRemove") {
		t.Fatal("Test style should exist before removal")
	}

	// 移除样式
	sm.RemoveStyle("TestRemove")

	// 验证样式已移除
	if sm.StyleExists("TestRemove") {
		t.Error("Test style should not exist after removal")
	}

	// 尝试移除不存在的样式（不应该报错）
	sm.RemoveStyle("NonExistentStyle")
}

// TestGetStyleWithInheritance 测试样式继承
func TestGetStyleWithInheritance(t *testing.T) {
	sm := NewStyleManager()

	// 获取带继承的Heading1样式
	heading1 := sm.GetStyleWithInheritance("Heading1")
	if heading1 == nil {
		t.Fatal("Heading1 with inheritance should not be nil")
	}

	// Heading1基于Normal，应该继承Normal的属性
	if heading1.BasedOn == nil {
		t.Error("Heading1 should have BasedOn reference")
	}

	// 验证继承的属性
	if heading1.RunPr == nil {
		t.Error("Heading1 should have run properties")
	}

	// 测试不存在的样式
	nonExistent := sm.GetStyleWithInheritance("NonExistentStyle")
	if nonExistent != nil {
		t.Error("Non-existent style with inheritance should return nil")
	}
}

// TestQuickStyleAPI 测试快速API功能
func TestQuickStyleAPI(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	if api == nil {
		t.Fatal("QuickStyleAPI should not be nil")
	}

	if api.styleManager != sm {
		t.Error("QuickStyleAPI should reference the provided StyleManager")
	}

	// 测试获取所有样式信息
	stylesInfo := api.GetAllStylesInfo()
	if len(stylesInfo) == 0 {
		t.Error("Should have style information")
	}

	// 验证返回的信息结构
	for _, info := range stylesInfo {
		if info.ID == "" {
			t.Error("Style info should have ID")
		}
		if info.Name == "" {
			t.Error("Style info should have Name")
		}
		if info.Type == "" {
			t.Error("Style info should have Type")
		}
	}
}

// TestQuickStyleAPI_GetStyleInfo 测试获取单个样式信息
func TestQuickStyleAPI_GetStyleInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	// 测试获取存在的样式信息
	info, err := api.GetStyleInfo("Normal")
	if err != nil {
		t.Fatalf("Error getting Normal style info: %v", err)
	}

	if info.ID != "Normal" {
		t.Errorf("Expected ID Normal, got %s", info.ID)
	}

	if info.Name != "Normal" {
		t.Errorf("Expected name Normal, got %s", info.Name)
	}

	// 测试获取不存在的样式信息
	_, err = api.GetStyleInfo("NonExistentStyle")
	if err == nil {
		t.Error("Should return error for non-existent style")
	}
}

// TestQuickStyleAPI_CreateStyle 测试快速创建样式
func TestQuickStyleAPI_CreateStyle(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	config := QuickStyleConfig{
		ID:      "QuickTest",
		Name:    "快速测试样式",
		Type:    StyleTypeParagraph,
		BasedOn: "Normal",
		ParagraphConfig: &QuickParagraphConfig{
			Alignment:   "center",
			LineSpacing: 1.5,
			SpaceBefore: 12,
			SpaceAfter:  6,
		},
		RunConfig: &QuickRunConfig{
			FontName:  "宋体",
			FontSize:  14,
			FontColor: "FF0000",
			Bold:      true,
			Italic:    false,
		},
	}

	style, err := api.CreateQuickStyle(config)
	if err != nil {
		t.Fatalf("Failed to create quick style: %v", err)
	}

	// 验证样式创建
	if style.StyleID != "QuickTest" {
		t.Errorf("Expected StyleID QuickTest, got %s", style.StyleID)
	}

	if style.Name.Val != "快速测试样式" {
		t.Errorf("Expected name 快速测试样式, got %s", style.Name.Val)
	}

	// 验证样式添加到管理器
	if !sm.StyleExists("QuickTest") {
		t.Error("Quick style should exist in style manager")
	}

	// 验证段落属性
	if style.ParagraphPr == nil {
		t.Fatal("Paragraph properties should not be nil")
	}

	if style.ParagraphPr.Justification == nil || style.ParagraphPr.Justification.Val != "center" {
		t.Error("Alignment should be center")
	}

	// 验证字符属性
	if style.RunPr == nil {
		t.Fatal("Run properties should not be nil")
	}

	if style.RunPr.Bold == nil {
		t.Error("Should be bold")
	}

	if style.RunPr.Color == nil || style.RunPr.Color.Val != "FF0000" {
		t.Error("Color should be FF0000")
	}

	if style.RunPr.FontSize == nil || style.RunPr.FontSize.Val != "28" {
		t.Error("Font size should be 28 (14*2)")
	}
}

// TestQuickStyleAPI_StylesByType 测试按类型获取样式
func TestQuickStyleAPI_StylesByType(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	// 测试获取段落样式
	paragraphStyles := api.GetParagraphStylesInfo()
	if len(paragraphStyles) == 0 {
		t.Error("Should have paragraph styles")
	}

	// 验证所有返回的都是段落样式
	for _, info := range paragraphStyles {
		if info.Type != "paragraph" {
			t.Errorf("Expected paragraph type, got %s", info.Type)
		}
	}

	// 测试获取字符样式
	characterStyles := api.GetCharacterStylesInfo()
	for _, info := range characterStyles {
		if info.Type != "character" {
			t.Errorf("Expected character type, got %s", info.Type)
		}
	}

	// 测试获取标题样式
	headingStyles := api.GetHeadingStylesInfo()
	if len(headingStyles) != 9 {
		t.Errorf("Expected 9 heading styles, got %d", len(headingStyles))
	}
}

// BenchmarkStyleLookup 基准测试 - 样式查找性能
func BenchmarkStyleLookup(b *testing.B) {
	sm := NewStyleManager()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		sm.GetStyle("Heading1")
	}
}

// BenchmarkStyleWithInheritance 基准测试 - 继承样式性能
func BenchmarkStyleWithInheritance(b *testing.B) {
	sm := NewStyleManager()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		sm.GetStyleWithInheritance("Heading1")
	}
}
