// Package document 模板功能测试
package document

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestNewTemplateEngine 测试创建模板引擎
func TestNewTemplateEngine(t *testing.T) {
	engine := NewTemplateEngine()
	if engine == nil {
		t.Fatal("Expected template engine to be created")
	}

	if engine.cache == nil {
		t.Fatal("Expected cache to be initialized")
	}
}

// TestTemplateVariableReplacement 测试变量替换功能
func TestTemplateVariableReplacement(t *testing.T) {
	engine := NewTemplateEngine()

	// 创建包含变量的模板
	templateContent := "Hello {{name}}, welcome to {{company}}!"
	template, err := engine.LoadTemplate("test_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// 验证模板变量解析
	if len(template.Variables) != 2 {
		t.Errorf("Expected 2 variables, got %d", len(template.Variables))
	}

	if _, exists := template.Variables["name"]; !exists {
		t.Error("Expected 'name' variable to be found")
	}

	if _, exists := template.Variables["company"]; !exists {
		t.Error("Expected 'company' variable to be found")
	}

	// 创建模板数据
	data := NewTemplateData()
	data.SetVariable("name", "John Doe")
	data.SetVariable("company", "go-word")

	// render template
	doc, err := engine.RenderToDocument("test_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// check document content
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// TestTemplateConditionalStatements
func TestTemplateConditionalStatements(t *testing.T) {
	engine := NewTemplateEngine()

	// create template with conditional statements
	templateContent := `{{#if showWelcome}}Welcome to go-word!{{/if}}
{{#if showDescription}}This is a powerful Word document operation library.{{/if}}`

	template, err := engine.LoadTemplate("conditional_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// verify conditional block parsing
	if len(template.Blocks) < 2 {
		t.Errorf("Expected at least 2 blocks, got %d", len(template.Blocks))
	}

	// test condition is true
	data := NewTemplateData()
	data.SetCondition("showWelcome", true)
	data.SetCondition("showDescription", false)

	doc, err := engine.RenderToDocument("conditional_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}
}

// TestTemplateLoopStatements tests loop statements functionality
func TestTemplateLoopStatements(t *testing.T) {
	engine := NewTemplateEngine()

	// create template with loop statements
	templateContent := `Product list:
{{#each products}}
- Product name: {{name}}, price: {{price}} dollars
{{/each}}`

	template, err := engine.LoadTemplate("loop_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// verify loop block parsing
	foundEachBlock := false
	for _, block := range template.Blocks {
		if block.Type == "each" && block.Variable == "products" {
			foundEachBlock = true
			break
		}
	}

	if !foundEachBlock {
		t.Error("Expected to find 'each products' block")
	}

	// create list data
	data := NewTemplateData()
	products := []interface{}{
		map[string]interface{}{
			"name":  "iPhone",
			"price": 8999,
		},
		map[string]interface{}{
			"name":  "iPad",
			"price": 5999,
		},
	}
	data.SetList("products", products)

	doc, err := engine.RenderToDocument("loop_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}
}

// TestTemplateInheritance
func TestTemplateInheritance(t *testing.T) {
	engine := NewTemplateEngine()

	// create base template
	baseTemplateContent := `Document title: {{title}}
Base content: This is the content of the base template.`

	_, err := engine.LoadTemplate("base_template", baseTemplateContent)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// create child template
	childTemplateContent := `{{extends "base_template"}}
Extended content: This is the content of the child template.`

	childTemplate, err := engine.LoadTemplate("child_template", childTemplateContent)
	if err != nil {
		t.Fatalf("Failed to load child template: %v", err)
	}

	// verify inheritance relation
	if childTemplate.Parent == nil {
		t.Error("Expected child template to have parent")
	}

	if childTemplate.Parent.Name != "base_template" {
		t.Errorf("Expected parent template name to be 'base_template', got '%s'", childTemplate.Parent.Name)
	}
}

// TestTemplateValidation tests template validation functionality
func TestTemplateValidation(t *testing.T) {
	engine := NewTemplateEngine()

	// test valid template
	validTemplate := `Hello {{name}}!
{{#if showMessage}}This is a message.{{/if}}
{{#each items}}Item: {{this}}{{/each}}`

	template, err := engine.LoadTemplate("valid_template", validTemplate)
	if err != nil {
		t.Fatalf("Failed to load valid template: %v", err)
	}

	err = engine.ValidateTemplate(template)
	if err != nil {
		t.Errorf("Expected valid template to pass validation, got error: %v", err)
	}

	// test invalid template - mismatched brackets
	invalidTemplate1 := `Hello {{name}!`
	template1, err := engine.LoadTemplate("invalid_template1", invalidTemplate1)
	if err != nil {
		t.Fatalf("Failed to load invalid template: %v", err)
	}

	err = engine.ValidateTemplate(template1)
	if err == nil {
		t.Error("Expected invalid template (mismatched brackets) to fail validation")
	}

	// test invalid template - mismatched if statements
	invalidTemplate2 := `{{#if condition}}Hello`
	template2, err := engine.LoadTemplate("invalid_template2", invalidTemplate2)
	if err != nil {
		t.Fatalf("Failed to load invalid template: %v", err)
	}

	err = engine.ValidateTemplate(template2)
	if err == nil {
		t.Error("Expected invalid template (mismatched if statements) to fail validation")
	}
}

// TestTemplateData tests template data functionality
func TestTemplateData(t *testing.T) {
	data := NewTemplateData()

	// test set and get variable
	data.SetVariable("name", "John Doe")
	value, exists := data.GetVariable("name")
	if !exists {
		t.Error("Expected variable 'name' to exist")
	}
	if value != "John Doe" {
		t.Errorf("Expected variable value to be 'John Doe', got '%v'", value)
	}

	// test set and get list
	items := []interface{}{"item1", "item2", "item3"}
	data.SetList("items", items)
	list, exists := data.GetList("items")
	if !exists {
		t.Error("Expected list 'items' to exist")
	}
	if len(list) != 3 {
		t.Errorf("Expected list length to be 3, got %d", len(list))
	}

	// test set and get condition
	data.SetCondition("enabled", true)
	condition, exists := data.GetCondition("enabled")
	if !exists {
		t.Error("Expected condition 'enabled' to exist")
	}
	if !condition {
		t.Error("Expected condition value to be true")
	}

	// test batch set variables
	variables := map[string]interface{}{
		"title":   "Test Title",
		"content": "Test Content",
	}
	data.SetVariables(variables)

	title, exists := data.GetVariable("title")
	if !exists || title != "Test Title" {
		t.Error("Expected batch set variables to work")
	}
}

// TestTemplateDataFromStruct tests template data from struct
func TestTemplateDataFromStruct(t *testing.T) {
	type TestStruct struct {
		Name    string
		Age     int
		Enabled bool
	}

	testData := TestStruct{
		Name:    "John Doe",
		Age:     30,
		Enabled: true,
	}

	templateData := NewTemplateData()
	err := templateData.FromStruct(testData)
	if err != nil {
		t.Fatalf("Failed to create template data from struct: %v", err)
	}

	// verify variables are set correctly
	name, exists := templateData.GetVariable("name")
	if !exists || name != "John Doe" {
		t.Error("Expected 'name' variable to be set correctly")
	}

	age, exists := templateData.GetVariable("age")
	if !exists || age != 30 {
		t.Error("Expected 'age' variable to be set correctly")
	}

	enabled, exists := templateData.GetVariable("enabled")
	if !exists || enabled != true {
		t.Error("Expected 'enabled' variable to be set correctly")
	}
}

// TestTemplateMerge tests template data merge
func TestTemplateMerge(t *testing.T) {
	data1 := NewTemplateData()
	data1.SetVariable("name", "John Doe")
	data1.SetCondition("enabled", true)

	data2 := NewTemplateData()
	data2.SetVariable("age", 30)
	data2.SetList("items", []interface{}{"item1", "item2"})

	// merge data
	data1.Merge(data2)

	// verify merge result
	name, exists := data1.GetVariable("name")
	if !exists || name != "John Doe" {
		t.Error("Expected original variable to remain")
	}

	age, exists := data1.GetVariable("age")
	if !exists || age != 30 {
		t.Error("Expected merged variable to be present")
	}

	enabled, exists := data1.GetCondition("enabled")
	if !exists || !enabled {
		t.Error("Expected original condition to remain")
	}

	items, exists := data1.GetList("items")
	if !exists || len(items) != 2 {
		t.Error("Expected merged list to be present")
	}
}

// TestTemplateCache tests template cache functionality
func TestTemplateCache(t *testing.T) {
	engine := NewTemplateEngine()

	// load template
	templateContent := "Hello {{name}}!"
	template1, err := engine.LoadTemplate("cached_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// get template from cache
	template2, err := engine.GetTemplate("cached_template")
	if err != nil {
		t.Fatalf("Failed to get template from cache: %v", err)
	}

	// verify same template instance
	if template1 != template2 {
		t.Error("Expected to get same template instance from cache")
	}

	// clear cache
	engine.ClearCache()

	// try to get template after cache clear
	_, err = engine.GetTemplate("cached_template")
	if err == nil {
		t.Error("Expected error when getting template after cache clear")
	}
}

// TestComplexTemplateRendering tests complex template rendering
func TestComplexTemplateRendering(t *testing.T) {
	engine := NewTemplateEngine()

	// create complex template
	complexTemplate := `Report title: {{title}}
Author: {{author}}

{{#if showSummary}}
Summary: {{summary}}
{{/if}}

Detailed content:
{{#each sections}}
Section {{@index}}: {{title}}
Content: {{content}}

{{/each}}

{{#if showFooter}}
Report finished.
{{/if}}`

	_, err := engine.LoadTemplate("complex_template", complexTemplate)
	if err != nil {
		t.Fatalf("Failed to load complex template: %v", err)
	}

	// create complex data
	data := NewTemplateData()
	data.SetVariable("title", "Test Report")
	data.SetVariable("author", "Development Team")
	data.SetVariable("summary", "This report tests the template functionality of word-zero.")

	data.SetCondition("showSummary", true)
	data.SetCondition("showFooter", true)

	sections := []interface{}{
		map[string]interface{}{
			"title":   "基础功能",
			"content": "测试了基础的文档操作功能。",
		},
		map[string]interface{}{
			"title":   "模板功能",
			"content": "测试了模板引擎的各种功能。",
		},
	}
	data.SetList("sections", sections)

	// render complex template
	doc, err := engine.RenderToDocument("complex_template", data)
	if err != nil {
		t.Fatalf("Failed to render complex template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// verify document has content
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// TestImagePlaceholder tests image placeholder functionality
func TestImagePlaceholder(t *testing.T) {
	engine := NewTemplateEngine()

	// test basic image placeholder parsing
	t.Run("解析图片占位符", func(t *testing.T) {
		templateContent := `Document title: {{title}}

这里有一个图片：
{{#image testImage}}

更多内容...`

		template, err := engine.LoadTemplate("image_test", templateContent)
		if err != nil {
			t.Fatalf("Failed to load template: %v", err)
		}

		// check if image placeholder is parsed correctly
		hasImageBlock := false
		for _, block := range template.Blocks {
			if block.Type == "image" && block.Name == "testImage" {
				hasImageBlock = true
				break
			}
		}

		if !hasImageBlock {
			t.Error("Expected template to contain image block")
		}
	})

	// test image placeholder rendering (string template)
	t.Run("render image placeholder to string", func(t *testing.T) {
		templateContent := `Product introduction: {{productName}}

Product image:
{{#image productImage}}

描述：{{description}}`

		_, err := engine.LoadTemplate("product", templateContent)
		if err != nil {
			t.Fatalf("Failed to load template: %v", err)
		}

		data := NewTemplateData()
		data.SetVariable("productName", "Test Product")
		data.SetVariable("description", "This is a test product")

		// create image config
		imageConfig := &ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
			Size: &ImageSize{
				Width:           100,
				KeepAspectRatio: true,
			},
			AltText: "Test Image",
			Title:   "Test Product Image",
		}

		// set image data (using example binary data)
		imageData := createTestImageData()
		data.SetImageFromData("productImage", imageData, imageConfig)

		// render template
		doc, err := engine.RenderToDocument("product", data)
		if err != nil {
			t.Fatalf("Failed to render template: %v", err)
		}

		if doc == nil {
			t.Error("Expected rendered result to be not nil")
		}
	})

	// test rendering image placeholder from document template
	t.Run("render image placeholder from document template", func(t *testing.T) {
		// 创建基础文档
		baseDoc := New()
		baseDoc.AddParagraph("Report title: {{title}}")
		baseDoc.AddParagraph("{{#image reportChart}}")
		baseDoc.AddParagraph("Summary: {{summary}}")

		// 从文档创建模板
		template, err := engine.LoadTemplateFromDocument("report_template", baseDoc)
		if err != nil {
			t.Fatalf("Failed to create template from document: %v", err)
		}

		if len(template.Variables) == 0 {
			t.Error("Expected template to contain variables")
		}

		// prepare data
		data := NewTemplateData()
		data.SetVariable("title", "Monthly Report")
		data.SetVariable("summary", "Data shows good growth trend")

		chartConfig := &ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
			Size: &ImageSize{
				Width: 120,
			},
		}

		imageData := createTestImageData()
		data.SetImageFromData("reportChart", imageData, chartConfig)

		// use RenderTemplateToDocument method (recommended for document templates)
		doc, err := engine.RenderTemplateToDocument("report_template", data)
		if err != nil {
			t.Fatalf("Failed to render document template: %v", err)
		}

		if doc == nil {
			t.Fatal("Expected rendered result to be not nil")
		}

		// check document contains elements
		if len(doc.Body.Elements) == 0 {
			t.Error("Expected document to contain elements")
		}
	})

	// test image data management method
	t.Run("test image data management", func(t *testing.T) {
		data := NewTemplateData()

		// test SetImage method
		config := &ImageConfig{
			Position: ImagePositionInline,
			Size:     &ImageSize{Width: 50},
		}
		data.SetImage("test1", "path/to/image.jpg", config)

		// test SetImageFromData method
		imageData := createTestImageData()
		data.SetImageFromData("test2", imageData, config)

		// test SetImageWithDetails method
		data.SetImageWithDetails("test3", "path/to/image2.jpg", imageData, config, "alt text", "title")

		// test GetImage method
		img1, exists1 := data.GetImage("test1")
		if !exists1 || img1.FilePath != "path/to/image.jpg" {
			t.Error("Expected image1 data to be correct")
		}

		img2, exists2 := data.GetImage("test2")
		if !exists2 || len(img2.Data) == 0 {
			t.Error("Expected image2 data to be correct")
		}

		img3, exists3 := data.GetImage("test3")
		if !exists3 || img3.AltText != "alt text" || img3.Title != "title" {
			t.Error("Expected image3 data to be correct")
		}

		// test non-existent image
		_, exists4 := data.GetImage("nonexistent")
		if exists4 {
			t.Error("Expected non-existent image to return false")
		}
	})

	// test image placeholder compatibility with other template syntax
	t.Run("image placeholder compatibility with other template syntax", func(t *testing.T) {
		templateContent := `{{#if showImage}}
Image title: {{imageTitle}}
{{#image dynamicImage}}
{{/if}}

{{#each items}}
项目：{{name}}
{{#image itemImage}}
描述：{{description}}
{{/each}}`

		_, err := engine.LoadTemplate("complex", templateContent)
		if err != nil {
			t.Fatalf("加载复杂模板失败: %v", err)
		}

		data := NewTemplateData()
		data.SetCondition("showImage", true)
		data.SetVariable("imageTitle", "主要图片")

		items := []interface{}{
			map[string]interface{}{
				"name":        "项目1",
				"description": "项目1描述",
			},
		}
		data.SetList("items", items)

		config := &ImageConfig{Position: ImagePositionInline}
		imageData := createTestImageData()
		data.SetImageFromData("dynamicImage", imageData, config)
		data.SetImageFromData("itemImage", imageData, config)

		// 渲染不应该出错
		doc, err := engine.RenderToDocument("complex", data)
		if err != nil {
			t.Fatalf("渲染复杂模板失败: %v", err)
		}

		if doc == nil {
			t.Error("渲染结果不应为空")
		}
	})
}

// createTestImageData 创建测试用的图片数据
func createTestImageData() []byte {
	// 创建一个最小的PNG图片数据用于测试
	return []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
		0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
		0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
		0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
		0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
		0x42, 0x60, 0x82,
	}
}

// TestNestedLoops 测试嵌套循环功能
func TestNestedLoops(t *testing.T) {
	engine := NewTemplateEngine()

	// 创建包含嵌套循环的模板
	templateContent := `会议纪要

日期：{{date}}

参会人员：
{{#each attendees}}
- {{name}} ({{role}})
  任务清单：
  {{#each tasks}}
  * {{taskName}} - 状态: {{status}}
  {{/each}}
{{/each}}

会议总结：{{summary}}`

	template, err := engine.LoadTemplate("meeting_minutes", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template with nested loops: %v", err)
	}

	if len(template.Blocks) < 1 {
		t.Error("Expected at least 1 block in template")
	}

	// 创建嵌套数据结构
	data := NewTemplateData()
	data.SetVariable("date", "2024-12-01")
	data.SetVariable("summary", "会议圆满结束")

	attendees := []interface{}{
		map[string]interface{}{
			"name": "张三",
			"role": "项目经理",
			"tasks": []interface{}{
				map[string]interface{}{
					"taskName": "制定项目计划",
					"status":   "进行中",
				},
				map[string]interface{}{
					"taskName": "分配资源",
					"status":   "已完成",
				},
			},
		},
		map[string]interface{}{
			"name": "李四",
			"role": "开发工程师",
			"tasks": []interface{}{
				map[string]interface{}{
					"taskName": "实现核心功能",
					"status":   "进行中",
				},
				map[string]interface{}{
					"taskName": "编写单元测试",
					"status":   "待开始",
				},
			},
		},
	}
	data.SetList("attendees", attendees)

	// render template
	doc, err := engine.RenderToDocument("meeting_minutes", data)
	if err != nil {
		t.Fatalf("Failed to render template with nested loops: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}

	// 检查生成的内容是否包含预期的嵌套数据
	foundNestedContent := false
	for _, element := range doc.Body.Elements {
		if para, ok := element.(*Paragraph); ok {
			fullText := ""
			for _, run := range para.Runs {
				fullText += run.Text.Content
			}

			// 检查是否包含嵌套循环生成的内容（任务名称）
			if fullText == "  * 制定项目计划 - 状态: 进行中" ||
				fullText == "  * 实现核心功能 - 状态: 进行中" {
				foundNestedContent = true
			}

			// 确保没有未处理的模板语法
			if fullText == "{{#each tasks}}" || fullText == "  * {{taskName}} - 状态: {{status}}" {
				t.Errorf("Found unprocessed template syntax in output: %s", fullText)
			}
		}
	}

	if !foundNestedContent {
		t.Error("Expected to find nested loop content in rendered document")
	}
}

// TestDeepNestedLoops 测试深度嵌套循环（三层）
func TestDeepNestedLoops(t *testing.T) {
	engine := NewTemplateEngine()

	// 创建三层嵌套循环的模板
	templateContent := `组织架构：
{{#each departments}}
部门：{{name}}
{{#each teams}}
  团队：{{teamName}}
  {{#each members}}
    成员：{{memberName}} - {{position}}
  {{/each}}
{{/each}}
{{/each}}`

	_, err := engine.LoadTemplate("org_structure", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template with deep nested loops: %v", err)
	}

	// 创建三层嵌套数据
	data := NewTemplateData()

	departments := []interface{}{
		map[string]interface{}{
			"name": "技术部",
			"teams": []interface{}{
				map[string]interface{}{
					"teamName": "前端团队",
					"members": []interface{}{
						map[string]interface{}{
							"memberName": "王五",
							"position":   "前端工程师",
						},
						map[string]interface{}{
							"memberName": "赵六",
							"position":   "UI设计师",
						},
					},
				},
				map[string]interface{}{
					"teamName": "后端团队",
					"members": []interface{}{
						map[string]interface{}{
							"memberName": "孙七",
							"position":   "后端工程师",
						},
					},
				},
			},
		},
	}
	data.SetList("departments", departments)

	// render template
	doc, err := engine.RenderToDocument("org_structure", data)
	if err != nil {
		t.Fatalf("Failed to render template with deep nested loops: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// 验证第三层嵌套内容是否正确渲染
	foundDeepContent := false
	for _, element := range doc.Body.Elements {
		if para, ok := element.(*Paragraph); ok {
			fullText := ""
			for _, run := range para.Runs {
				fullText += run.Text.Content
			}

			// 检查第三层嵌套内容
			if fullText == "    成员：王五 - 前端工程师" ||
				fullText == "    成员：孙七 - 后端工程师" {
				foundDeepContent = true
			}

			// 确保没有未处理的模板语法
			if fullText == "{{#each members}}" || fullText == "    成员：{{memberName}} - {{position}}" {
				t.Errorf("Found unprocessed template syntax in deep nested output: %s", fullText)
			}
		}
	}

	if !foundDeepContent {
		t.Error("Expected to find deep nested loop content in rendered document")
	}
}

// TestHeaderFooterTemplateVariables 测试页眉页脚中的模板变量识别和替换
func TestHeaderFooterTemplateVariables(t *testing.T) {
	// 创建包含页眉页脚的文档
	doc := New()

	// 添加主体内容
	doc.AddParagraph("{{title}}")
	doc.AddParagraph("文档内容")

	// 添加带有模板变量的页眉
	err := doc.AddHeader(HeaderFooterTypeDefault, "{{headerTitle}} - {{headerID}}")
	if err != nil {
		t.Fatalf("添加页眉失败: %v", err)
	}

	// 添加带有模板变量的页脚
	err = doc.AddFooter(HeaderFooterTypeDefault, "{{footerText}} - 第 {{pageNum}} 页")
	if err != nil {
		t.Fatalf("添加页脚失败: %v", err)
	}

	// 创建模板引擎并加载文档作为模板
	engine := NewTemplateEngine()
	template, err := engine.LoadTemplateFromDocument("header_footer_test", doc)
	if err != nil {
		t.Fatalf("从文档加载模板失败: %v", err)
	}

	// 验证模板变量被正确识别
	expectedVars := []string{"title", "headerTitle", "headerID", "footerText", "pageNum"}
	for _, varName := range expectedVars {
		if _, exists := template.Variables[varName]; !exists {
			t.Errorf("模板变量 '%s' 应该被识别但未找到", varName)
		}
	}

	// 测试使用TemplateRenderer分析包含页眉页脚的模板
	// 创建一个新的带页眉页脚的文档用于测试分析功能
	doc2 := New()
	doc2.AddParagraph("{{mainContent}}")
	err = doc2.AddHeader(HeaderFooterTypeDefault, "{{documentTitle}}")
	if err != nil {
		t.Fatalf("添加页眉失败: %v", err)
	}

	// 通过engine加载
	engine2 := NewTemplateEngine()
	_, err = engine2.LoadTemplateFromDocument("analyze_test", doc2)
	if err != nil {
		t.Fatalf("从文档加载模板失败: %v", err)
	}

	// 创建renderer并使用已加载的模板
	renderer := &TemplateRenderer{
		engine: engine2,
		logger: &TemplateLogger{enabled: false},
	}

	// 分析模板
	analysis, err := renderer.AnalyzeTemplate("analyze_test")
	if err != nil {
		t.Fatalf("分析模板失败: %v", err)
	}

	// 验证分析结果包含页眉中的变量
	if _, exists := analysis.Variables["documentTitle"]; !exists {
		t.Error("分析结果应该包含页眉变量 'documentTitle'")
	}
	if _, exists := analysis.Variables["mainContent"]; !exists {
		t.Error("分析结果应该包含主体变量 'mainContent'")
	}

	t.Logf("分析到的变量: %v", analysis.Variables)
}

// TestHeaderFooterVariableReplacement 测试页眉页脚中的变量替换功能
func TestHeaderFooterVariableReplacement(t *testing.T) {
	// 创建包含页眉页脚的文档
	doc := New()

	// 添加主体内容
	doc.AddParagraph("{{title}}")
	doc.AddParagraph("正文内容")

	// 添加带有模板变量的页眉
	err := doc.AddHeader(HeaderFooterTypeDefault, "报告编号: {{reportID}}")
	if err != nil {
		t.Fatalf("添加页眉失败: %v", err)
	}

	// 添加带有模板变量的页脚
	err = doc.AddFooter(HeaderFooterTypeDefault, "作者: {{author}}")
	if err != nil {
		t.Fatalf("添加页脚失败: %v", err)
	}

	// 创建模板引擎并加载文档作为模板
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("replacement_test", doc)
	if err != nil {
		t.Fatalf("从文档加载模板失败: %v", err)
	}

	// 准备模板数据
	data := NewTemplateData()
	data.SetVariable("title", "测试报告标题")
	data.SetVariable("reportID", "RPT-2024-001")
	data.SetVariable("author", "测试作者")

	// render template
	resultDoc, err := engine.RenderTemplateToDocument("replacement_test", data)
	if err != nil {
		t.Fatalf("渲染模板失败: %v", err)
	}

	// 验证页眉中的变量被替换
	headerReplaced := false
	footerReplaced := false

	for partName, partData := range resultDoc.parts {
		content := string(partData)

		if partName == "word/header1.xml" {
			if !strings.Contains(content, "{{reportID}}") && strings.Contains(content, "RPT-2024-001") {
				headerReplaced = true
			}
			t.Logf("页眉内容: %s", content)
		}

		if partName == "word/footer1.xml" {
			if !strings.Contains(content, "{{author}}") && strings.Contains(content, "测试作者") {
				footerReplaced = true
			}
			t.Logf("页脚内容: %s", content)
		}
	}

	if !headerReplaced {
		t.Error("页眉中的变量应该被替换")
	}

	if !footerReplaced {
		t.Error("页脚中的变量应该被替换")
	}
}

// TestTemplateFromFileWithParagraphSectionProperties 确保段落内的节属性仍能保留页眉页脚
func TestTemplateFromFileWithParagraphSectionProperties(t *testing.T) {
	doc := New()
	doc.AddParagraph("{{title}}")

	if err := doc.AddHeader(HeaderFooterTypeDefault, "报告编号: {{reportID}}"); err != nil {
		t.Fatalf("添加页眉失败: %v", err)
	}
	if err := doc.AddFooter(HeaderFooterTypeDefault, "撰写人: {{author}}"); err != nil {
		t.Fatalf("添加页脚失败: %v", err)
	}

	sectionMarker := "__SECTION_BREAK__"
	doc.AddParagraph(sectionMarker)

	tmpDir := t.TempDir()
	basePath := filepath.Join(tmpDir, "base_paragraph_section.docx")
	if err := doc.Save(basePath); err != nil {
		t.Fatalf("保存基础文档失败: %v", err)
	}

	modifiedPath := filepath.Join(tmpDir, "paragraph_section_template.docx")
	if err := moveSectPrIntoParagraph(basePath, modifiedPath, sectionMarker); err != nil {
		t.Fatalf("调整节属性位置失败: %v", err)
	}

	loadedDoc, err := Open(modifiedPath)
	if err != nil {
		t.Fatalf("打开修改后的文档失败: %v", err)
	}

	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("paragraph_section_template", loadedDoc)
	if err != nil {
		t.Fatalf("加载模板失败: %v", err)
	}

	data := NewTemplateData()
	data.SetVariable("title", "段落节属性测试")
	data.SetVariable("reportID", "RPT-2024-009")
	data.SetVariable("author", "测试作者")

	renderedDoc, err := engine.RenderTemplateToDocument("paragraph_section_template", data)
	if err != nil {
		t.Fatalf("渲染模板失败: %v", err)
	}

	headerContent := string(renderedDoc.parts["word/header1.xml"])
	if strings.Contains(headerContent, "{{reportID}}") {
		t.Error("页眉中的变量应该被替换，即使节属性位于段落内")
	}
	if !strings.Contains(headerContent, "RPT-2024-009") {
		t.Error("页眉中缺少替换后的值")
	}

	footerContent := string(renderedDoc.parts["word/footer1.xml"])
	if strings.Contains(footerContent, "{{author}}") {
		t.Error("页脚中的变量应该被替换")
	}
	if !strings.Contains(footerContent, "测试作者") {
		t.Error("页脚中缺少替换后的值")
	}
}

func moveSectPrIntoParagraph(srcPath, dstPath, marker string) error {
	reader, err := zip.OpenReader(srcPath)
	if err != nil {
		return err
	}
	defer reader.Close()

	output, err := os.Create(dstPath)
	if err != nil {
		return err
	}
	defer output.Close()

	zipWriter := zip.NewWriter(output)
	defer zipWriter.Close()

	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return err
		}

		if file.Name == "word/document.xml" {
			data, err = rewriteSectPrIntoParagraph(data, marker)
			if err != nil {
				return err
			}
		}

		writer, err := zipWriter.Create(file.Name)
		if err != nil {
			return err
		}
		if _, err := writer.Write(data); err != nil {
			return err
		}
	}

	return nil
}

func rewriteSectPrIntoParagraph(xmlData []byte, marker string) ([]byte, error) {
	content := string(xmlData)
	sectStart := strings.Index(content, "<w:sectPr")
	if sectStart == -1 {
		return nil, fmt.Errorf("未找到sectPr")
	}

	sectEndRel := strings.Index(content[sectStart:], "</w:sectPr>")
	if sectEndRel == -1 {
		return nil, fmt.Errorf("sectPr缺少结束标签")
	}
	sectEnd := sectStart + sectEndRel + len("</w:sectPr>")
	sectBlock := content[sectStart:sectEnd]

	sanitized := removeHeaderFooterReferences(sectBlock)
	content = content[:sectStart] + sanitized + content[sectEnd:]

	markerIndex := strings.Index(content, marker)
	if markerIndex == -1 {
		return nil, fmt.Errorf("未找到标记段落")
	}

	pStart := strings.LastIndex(content[:markerIndex], "<w:p")
	if pStart == -1 {
		return nil, fmt.Errorf("未找到段落起始标签")
	}
	openEnd := strings.Index(content[pStart:], ">")
	if openEnd == -1 {
		return nil, fmt.Errorf("段落标签未闭合")
	}
	insertPos := pStart + openEnd + 1

	insert := "<w:pPr>" + sectBlock + "</w:pPr>"
	modified := content[:insertPos] + insert + content[insertPos:]

	return []byte(modified), nil
}

func removeHeaderFooterReferences(block string) string {
	block = stripReferenceTag(block, "<w:headerReference")
	block = stripReferenceTag(block, "<w:footerReference")
	return block
}

func stripReferenceTag(block, tag string) string {
	for {
		start := strings.Index(block, tag)
		if start == -1 {
			break
		}
		end := strings.Index(block[start:], "/>")
		if end == -1 {
			break
		}
		block = block[:start] + block[start+end+2:]
	}
	return block
}

// TestTemplateDocumentPartsPreservation 测试模板渲染时文档部件的完整保留
func TestTemplateDocumentPartsPreservation(t *testing.T) {
	// 创建包含多种文档部件的源文档
	doc := New()

	// 添加页眉和页脚
	err := doc.AddHeader(HeaderFooterTypeDefault, "Template Header - {{headerVar}}")
	if err != nil {
		t.Fatalf("添加页眉失败: %v", err)
	}

	err = doc.AddFooter(HeaderFooterTypeDefault, "Template Footer - {{footerVar}}")
	if err != nil {
		t.Fatalf("添加页脚失败: %v", err)
	}

	// 设置页面设置
	settings := DefaultPageSettings()
	settings.Size = PageSizeA4
	settings.Orientation = OrientationPortrait
	err = doc.SetPageSettings(settings)
	if err != nil {
		t.Fatalf("设置页面设置失败: %v", err)
	}

	// 添加标题和内容
	doc.AddHeadingParagraph("{{docTitle}}", 1)
	doc.AddParagraph("Content with {{variable1}} and more text.")

	// save original document
	originalPath := "test_parts_preservation_original.docx"
	err = doc.Save(originalPath)
	if err != nil {
		t.Fatalf("保存原文档失败: %v", err)
	}
	defer func() {
		if err := os.Remove(originalPath); err != nil {
			t.Logf("清理原文档失败: %v", err)
		}
	}()

	// open original document as template
	templateDoc, err := Open(originalPath)
	if err != nil {
		t.Fatalf("打开模板文档失败: %v", err)
	}

	// 记录原文档的parts
	originalParts := make(map[string]bool)
	for partName := range templateDoc.parts {
		originalParts[partName] = true
	}

	// create template engine and load template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("parts_test", templateDoc)
	if err != nil {
		t.Fatalf("加载模板失败: %v", err)
	}

	// render template
	data := NewTemplateData()
	data.SetVariable("headerVar", "Header Value")
	data.SetVariable("footerVar", "Footer Value")
	data.SetVariable("docTitle", "Document Title")
	data.SetVariable("variable1", "Variable 1 Value")

	renderedDoc, err := engine.RenderTemplateToDocument("parts_test", data)
	if err != nil {
		t.Fatalf("渲染模板失败: %v", err)
	}

	// 保存渲染后的文档
	renderedPath := "test_parts_preservation_rendered.docx"
	err = renderedDoc.Save(renderedPath)
	if err != nil {
		t.Fatalf("保存渲染后的文档失败: %v", err)
	}
	defer func() {
		if err := os.Remove(renderedPath); err != nil {
			t.Logf("清理渲染文档失败: %v", err)
		}
	}()

	// 检查渲染后文档的parts
	renderedParts := make(map[string]bool)
	for partName := range renderedDoc.parts {
		renderedParts[partName] = true
	}

	// 验证关键部件被保留
	criticalParts := []string{
		"word/styles.xml",
		"word/header1.xml",
		"word/footer1.xml",
	}

	for _, part := range criticalParts {
		if originalParts[part] && !renderedParts[part] {
			t.Errorf("关键部件 %s 在原文档中存在但在渲染后的文档中丢失", part)
		}
	}

	// 验证页眉页脚变量被替换
	headerContent := string(renderedDoc.parts["word/header1.xml"])
	if strings.Contains(headerContent, "{{headerVar}}") {
		t.Error("页眉中的变量应该被替换")
	}
	if !strings.Contains(headerContent, "Header Value") {
		t.Error("页眉中应该包含替换后的值")
	}

	footerContent := string(renderedDoc.parts["word/footer1.xml"])
	if strings.Contains(footerContent, "{{footerVar}}") {
		t.Error("页脚中的变量应该被替换")
	}
	if !strings.Contains(footerContent, "Footer Value") {
		t.Error("页脚中应该包含替换后的值")
	}

	t.Log("文档部件保留测试通过")
}

// TestTemplateSectionPropertiesPreservation 测试节属性在模板渲染时的保留
func TestTemplateSectionPropertiesPreservation(t *testing.T) {
	// 创建包含节属性的源文档
	doc := New()

	// 设置页面设置（这会创建SectionProperties）
	settings := DefaultPageSettings()
	settings.Size = PageSizeA4
	settings.MarginTop = 30.0
	settings.MarginBottom = 25.0
	settings.MarginLeft = 20.0
	settings.MarginRight = 20.0
	err := doc.SetPageSettings(settings)
	if err != nil {
		t.Fatalf("设置页面设置失败: %v", err)
	}

	// 添加内容
	doc.AddParagraph("Content with {{variable}}")

	// save original document
	originalPath := "test_section_props_original.docx"
	err = doc.Save(originalPath)
	if err != nil {
		t.Fatalf("保存原文档失败: %v", err)
	}
	defer func() {
		if err := os.Remove(originalPath); err != nil {
			t.Logf("清理原文档失败: %v", err)
		}
	}()

	// open original document as template
	templateDoc, err := Open(originalPath)
	if err != nil {
		t.Fatalf("打开模板文档失败: %v", err)
	}

	// 获取原文档的页面设置
	originalSettings := templateDoc.GetPageSettings()

	// create template engine and load template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("section_test", templateDoc)
	if err != nil {
		t.Fatalf("加载模板失败: %v", err)
	}

	// render template
	data := NewTemplateData()
	data.SetVariable("variable", "Value")

	renderedDoc, err := engine.RenderTemplateToDocument("section_test", data)
	if err != nil {
		t.Fatalf("渲染模板失败: %v", err)
	}

	// 获取渲染后文档的页面设置
	renderedSettings := renderedDoc.GetPageSettings()

	// 验证页面设置被保留
	if renderedSettings.Size != originalSettings.Size {
		t.Errorf("页面大小不匹配: 期望 %v, 实际 %v", originalSettings.Size, renderedSettings.Size)
	}

	// 允许1mm的误差
	tolerance := 1.0
	if abs(renderedSettings.MarginTop-originalSettings.MarginTop) > tolerance {
		t.Errorf("上边距不匹配: 期望 %.1f, 实际 %.1f", originalSettings.MarginTop, renderedSettings.MarginTop)
	}
	if abs(renderedSettings.MarginBottom-originalSettings.MarginBottom) > tolerance {
		t.Errorf("下边距不匹配: 期望 %.1f, 实际 %.1f", originalSettings.MarginBottom, renderedSettings.MarginBottom)
	}
	if abs(renderedSettings.MarginLeft-originalSettings.MarginLeft) > tolerance {
		t.Errorf("左边距不匹配: 期望 %.1f, 实际 %.1f", originalSettings.MarginLeft, renderedSettings.MarginLeft)
	}
	if abs(renderedSettings.MarginRight-originalSettings.MarginRight) > tolerance {
		t.Errorf("右边距不匹配: 期望 %.1f, 实际 %.1f", originalSettings.MarginRight, renderedSettings.MarginRight)
	}

	t.Log("节属性保留测试通过")
}

// TestTemplateNumberingPropertiesPreservation 测试模板渲染时编号属性的保留
func TestTemplateNumberingPropertiesPreservation(t *testing.T) {
	// 创建包含编号段落的文档
	doc := New()

	// 添加带有编号的列表项
	config := &ListConfig{
		Type:        ListTypeNumber,
		IndentLevel: 0,
		StartNumber: 1,
	}
	doc.AddListItem("第一条 {{itemTitle}}", config)
	doc.AddListItem("第二条 {{itemContent}}", config)

	// save original document
	originalPath := "test_numbering_preservation_original.docx"
	err := doc.Save(originalPath)
	if err != nil {
		t.Fatalf("保存原文档失败: %v", err)
	}
	defer func() {
		if err := os.Remove(originalPath); err != nil {
			t.Logf("清理原文档失败: %v", err)
		}
	}()

	// open original document as template
	templateDoc, err := Open(originalPath)
	if err != nil {
		t.Fatalf("打开模板文档失败: %v", err)
	}

	// 验证原文档的编号属性被正确解析
	paragraphs := templateDoc.Body.GetParagraphs()
	if len(paragraphs) < 2 {
		t.Fatalf("期望至少2个段落，实际 %d 个", len(paragraphs))
	}

	// 检查第一个段落的编号属性
	if paragraphs[0].Properties == nil || paragraphs[0].Properties.NumberingProperties == nil {
		t.Error("第一个段落的编号属性应该被解析")
	}

	// create template engine and load template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("numbering_test", templateDoc)
	if err != nil {
		t.Fatalf("加载模板失败: %v", err)
	}

	// render template
	data := NewTemplateData()
	data.SetVariable("itemTitle", "合作项目情况")
	data.SetVariable("itemContent", "合作项目背景")

	renderedDoc, err := engine.RenderTemplateToDocument("numbering_test", data)
	if err != nil {
		t.Fatalf("渲染模板失败: %v", err)
	}

	// 保存渲染后的文档
	renderedPath := "test_numbering_preservation_rendered.docx"
	err = renderedDoc.Save(renderedPath)
	if err != nil {
		t.Fatalf("保存渲染后的文档失败: %v", err)
	}
	defer func() {
		if err := os.Remove(renderedPath); err != nil {
			t.Logf("清理渲染文档失败: %v", err)
		}
	}()

	// 验证渲染后文档的编号属性被保留
	renderedParagraphs := renderedDoc.Body.GetParagraphs()
	if len(renderedParagraphs) < 2 {
		t.Fatalf("渲染后期望至少2个段落，实际 %d 个", len(renderedParagraphs))
	}

	// 检查渲染后段落的编号属性是否被保留
	for i, para := range renderedParagraphs[:2] {
		if para.Properties == nil {
			t.Errorf("段落 %d 的属性不应为空", i+1)
			continue
		}
		if para.Properties.NumberingProperties == nil {
			t.Errorf("段落 %d 的编号属性不应为空", i+1)
			continue
		}
		if para.Properties.NumberingProperties.NumID == nil {
			t.Errorf("段落 %d 的编号ID不应为空", i+1)
		}
		if para.Properties.NumberingProperties.ILevel == nil {
			t.Errorf("段落 %d 的编号级别不应为空", i+1)
		}
	}

	// 验证变量已被替换
	firstParaText := ""
	for _, run := range renderedParagraphs[0].Runs {
		firstParaText += run.Text.Content
	}
	if !strings.Contains(firstParaText, "合作项目情况") {
		t.Errorf("第一个段落应该包含替换后的变量值，实际内容: %s", firstParaText)
	}

	t.Log("编号属性保留测试通过")
}
