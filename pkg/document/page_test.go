// Package document 页面设置功能测试
package document

import (
	"testing"
)

// TestDefaultPageSettings 测试默认页面设置
func TestDefaultPageSettings(t *testing.T) {
	settings := DefaultPageSettings()

	if settings.Size != PageSizeA4 {
		t.Errorf("默认页面尺寸应为A4，实际为: %s", settings.Size)
	}

	if settings.Orientation != OrientationPortrait {
		t.Errorf("默认页面方向应为纵向，实际为: %s", settings.Orientation)
	}

	if settings.MarginTop != 25.4 {
		t.Errorf("默认上边距应为25.4mm，实际为: %.1fmm", settings.MarginTop)
	}
}

// TestSetPageSize 测试设置页面尺寸
func TestSetPageSize(t *testing.T) {
	doc := New()

	// 测试设置为Letter尺寸
	err := doc.SetPageSize(PageSizeLetter)
	if err != nil {
		t.Errorf("设置页面尺寸失败: %v", err)
	}

	settings := doc.GetPageSettings()
	if settings.Size != PageSizeLetter {
		t.Errorf("页面尺寸应为Letter，实际为: %s", settings.Size)
	}
}

// TestSetCustomPageSize 测试设置自定义页面尺寸
func TestSetCustomPageSize(t *testing.T) {
	doc := New()

	// 测试有效的自定义尺寸
	err := doc.SetCustomPageSize(200, 300)
	if err != nil {
		t.Errorf("设置自定义页面尺寸失败: %v", err)
	}

	settings := doc.GetPageSettings()
	if settings.Size != PageSizeCustom {
		t.Errorf("页面尺寸应为Custom，实际为: %s", settings.Size)
	}

	if abs(settings.CustomWidth-200) > 0.1 {
		t.Errorf("自定义宽度应为200mm，实际为: %.1fmm", settings.CustomWidth)
	}

	if abs(settings.CustomHeight-300) > 0.1 {
		t.Errorf("自定义高度应为300mm，实际为: %.1fmm", settings.CustomHeight)
	}

	// 测试无效的自定义尺寸
	err = doc.SetCustomPageSize(-100, 200)
	if err == nil {
		t.Error("设置负数尺寸应该返回错误")
	}

	err = doc.SetCustomPageSize(100, 0)
	if err == nil {
		t.Error("设置零高度应该返回错误")
	}
}

// TestSetPageOrientation 测试设置页面方向
func TestSetPageOrientation(t *testing.T) {
	doc := New()

	// 测试设置为横向
	err := doc.SetPageOrientation(OrientationLandscape)
	if err != nil {
		t.Errorf("设置页面方向失败: %v", err)
	}

	settings := doc.GetPageSettings()
	if settings.Orientation != OrientationLandscape {
		t.Errorf("页面方向应为横向，实际为: %s", settings.Orientation)
	}
}

// TestSetPageMargins 测试设置页面边距
func TestSetPageMargins(t *testing.T) {
	doc := New()

	// 测试有效的边距设置
	err := doc.SetPageMargins(20, 15, 25, 30)
	if err != nil {
		t.Errorf("设置页面边距失败: %v", err)
	}

	settings := doc.GetPageSettings()
	if abs(settings.MarginTop-20) > 0.1 {
		t.Errorf("上边距应为20mm，实际为: %.1fmm", settings.MarginTop)
	}
	if abs(settings.MarginRight-15) > 0.1 {
		t.Errorf("右边距应为15mm，实际为: %.1fmm", settings.MarginRight)
	}
	if abs(settings.MarginBottom-25) > 0.1 {
		t.Errorf("下边距应为25mm，实际为: %.1fmm", settings.MarginBottom)
	}
	if abs(settings.MarginLeft-30) > 0.1 {
		t.Errorf("左边距应为30mm，实际为: %.1fmm", settings.MarginLeft)
	}

	// 测试负数边距
	err = doc.SetPageMargins(-10, 15, 25, 30)
	if err == nil {
		t.Error("设置负数边距应该返回错误")
	}
}

// TestSetHeaderFooterDistance 测试设置页眉页脚距离
func TestSetHeaderFooterDistance(t *testing.T) {
	doc := New()

	// 测试有效的页眉页脚距离
	err := doc.SetHeaderFooterDistance(10, 15)
	if err != nil {
		t.Errorf("设置页眉页脚距离失败: %v", err)
	}

	settings := doc.GetPageSettings()
	if abs(settings.HeaderDistance-10) > 0.1 {
		t.Errorf("页眉距离应为10mm，实际为: %.1fmm", settings.HeaderDistance)
	}
	if abs(settings.FooterDistance-15) > 0.1 {
		t.Errorf("页脚距离应为15mm，实际为: %.1fmm", settings.FooterDistance)
	}

	// 测试负数距离
	err = doc.SetHeaderFooterDistance(-5, 15)
	if err == nil {
		t.Error("设置负数页眉距离应该返回错误")
	}
}

// TestSetGutterWidth 测试设置装订线宽度
func TestSetGutterWidth(t *testing.T) {
	doc := New()

	// 测试有效的装订线宽度
	err := doc.SetGutterWidth(5)
	if err != nil {
		t.Errorf("设置装订线宽度失败: %v", err)
	}

	settings := doc.GetPageSettings()
	if abs(settings.GutterWidth-5) > 0.1 {
		t.Errorf("装订线宽度应为5mm，实际为: %.1fmm", settings.GutterWidth)
	}

	// 测试负数装订线宽度
	err = doc.SetGutterWidth(-2)
	if err == nil {
		t.Error("设置负数装订线宽度应该返回错误")
	}
}

// TestPageDimensions 测试页面尺寸计算
func TestPageDimensions(t *testing.T) {
	tests := []struct {
		name      string
		settings  *PageSettings
		expWidth  float64
		expHeight float64
	}{
		{
			name: "A4纵向",
			settings: &PageSettings{
				Size:        PageSizeA4,
				Orientation: OrientationPortrait,
			},
			expWidth:  210,
			expHeight: 297,
		},
		{
			name: "A4横向",
			settings: &PageSettings{
				Size:        PageSizeA4,
				Orientation: OrientationLandscape,
			},
			expWidth:  297,
			expHeight: 210,
		},
		{
			name: "自定义尺寸",
			settings: &PageSettings{
				Size:         PageSizeCustom,
				CustomWidth:  150,
				CustomHeight: 200,
				Orientation:  OrientationPortrait,
			},
			expWidth:  150,
			expHeight: 200,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			width, height := getPageDimensions(tt.settings)

			if width != tt.expWidth {
				t.Errorf("宽度不匹配，期望: %.1fmm, 实际: %.1fmm", tt.expWidth, width)
			}

			if height != tt.expHeight {
				t.Errorf("高度不匹配，期望: %.1fmm, 实际: %.1fmm", tt.expHeight, height)
			}
		})
	}
}

// TestIdentifyPageSize 测试页面尺寸识别
func TestIdentifyPageSize(t *testing.T) {
	tests := []struct {
		name     string
		width    float64
		height   float64
		expected PageSize
	}{
		{
			name:     "A4纵向",
			width:    210,
			height:   297,
			expected: PageSizeA4,
		},
		{
			name:     "A4横向",
			width:    297,
			height:   210,
			expected: PageSizeA4,
		},
		{
			name:     "Letter",
			width:    215.9,
			height:   279.4,
			expected: PageSizeLetter,
		},
		{
			name:     "自定义尺寸",
			width:    100,
			height:   150,
			expected: PageSizeCustom,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := identifyPageSize(tt.width, tt.height)

			if result != tt.expected {
				t.Errorf("页面尺寸识别错误，期望: %s, 实际: %s", tt.expected, result)
			}
		})
	}
}

// TestValidatePageSettings tests page settings validation
func TestValidatePageSettings(t *testing.T) {
	// 测试有效设置
	validSettings := &PageSettings{
		Size:         PageSizeA4,
		Orientation:  OrientationPortrait,
		CustomWidth:  0,
		CustomHeight: 0,
	}

	err := validatePageSettings(validSettings)
	if err != nil {
		t.Errorf("有效设置应该通过验证，错误: %v", err)
	}

	// 测试无效的自定义尺寸
	invalidCustomSize := &PageSettings{
		Size:         PageSizeCustom,
		Orientation:  OrientationPortrait,
		CustomWidth:  -100,
		CustomHeight: 200,
	}

	err = validatePageSettings(invalidCustomSize)
	if err == nil {
		t.Error("负数自定义尺寸应该验证失败")
	}

	// 测试过大的自定义尺寸
	oversizeCustom := &PageSettings{
		Size:         PageSizeCustom,
		Orientation:  OrientationPortrait,
		CustomWidth:  600, // 超过最大尺寸
		CustomHeight: 200,
	}

	err = validatePageSettings(oversizeCustom)
	if err == nil {
		t.Error("过大的自定义尺寸应该验证失败")
	}

	// 测试无效方向
	invalidOrientation := &PageSettings{
		Size:        PageSizeA4,
		Orientation: PageOrientation("invalid"),
	}

	err = validatePageSettings(invalidOrientation)
	if err == nil {
		t.Error("无效方向应该验证失败")
	}
}

// TestMmToTwips 测试毫米到Twips的转换
func TestMmToTwips(t *testing.T) {
	// 测试几个已知的转换值
	tests := []struct {
		mm       float64
		expected float64
	}{
		{25.4, 1440}, // 1英寸 = 1440 twips
		{0, 0},       // 0毫米 = 0 twips
		{10, 566.93}, // 约567 twips
	}

	for _, tt := range tests {
		result := mmToTwips(tt.mm)
		// 允许小数点误差
		if abs(result-tt.expected) > 1 {
			t.Errorf("毫米转换错误，输入: %.1fmm, 期望: %.0f twips, 实际: %.0f twips",
				tt.mm, tt.expected, result)
		}
	}
}

// TestTwipsToMM 测试Twips到毫米的转换
func TestTwipsToMM(t *testing.T) {
	// 测试反向转换
	tests := []struct {
		twips    float64
		expected float64
	}{
		{1440, 25.4}, // 1440 twips = 1英寸 = 25.4mm
		{0, 0},       // 0 twips = 0mm
		{567, 10.0},  // 约10mm
	}

	for _, tt := range tests {
		result := twipsToMM(tt.twips)
		// 允许小数点误差
		if abs(result-tt.expected) > 0.1 {
			t.Errorf("Twips转换错误，输入: %.0f twips, 期望: %.1fmm, 实际: %.1fmm",
				tt.twips, tt.expected, result)
		}
	}
}

// TestCompletePageSettings 测试完整的页面设置流程
func TestCompletePageSettings(t *testing.T) {
	doc := New()

	// 创建完整的页面设置
	settings := &PageSettings{
		Size:           PageSizeLetter,
		Orientation:    OrientationLandscape,
		MarginTop:      20,
		MarginRight:    15,
		MarginBottom:   25,
		MarginLeft:     30,
		HeaderDistance: 8,
		FooterDistance: 12,
		GutterWidth:    5,
	}

	// 应用设置
	err := doc.SetPageSettings(settings)
	if err != nil {
		t.Errorf("设置页面属性失败: %v", err)
	}

	// 验证设置是否正确应用
	retrieved := doc.GetPageSettings()

	if retrieved.Size != settings.Size {
		t.Errorf("页面尺寸不匹配，期望: %s, 实际: %s", settings.Size, retrieved.Size)
	}

	if retrieved.Orientation != settings.Orientation {
		t.Errorf("页面方向不匹配，期望: %s, 实际: %s", settings.Orientation, retrieved.Orientation)
	}

	if abs(retrieved.MarginTop-settings.MarginTop) > 0.1 {
		t.Errorf("上边距不匹配，期望: %.1fmm, 实际: %.1fmm", settings.MarginTop, retrieved.MarginTop)
	}
}
