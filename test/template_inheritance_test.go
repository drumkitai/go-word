// Package test provides template inheritance functionality tests
package test

import (
	"os"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestTemplateInheritanceComplete tests template inheritance
func TestTemplateInheritanceComplete(t *testing.T) {
	outputDir := "output"
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		err = os.Mkdir(outputDir, 0755)
		if err != nil {
			t.Fatalf("Failed to create output directory: %v", err)
		}
	}

	engine := document.NewTemplateEngine()

	t.Run("basic template inheritance", func(t *testing.T) {
		testBasicInheritance(t, engine)
	})

	t.Run("block override functionality", func(t *testing.T) {
		testBlockOverride(t, engine)
	})

	t.Cleanup(func() {
		cleanupInheritanceTestFiles()
	})
}

func testBasicInheritance(t *testing.T, engine *document.TemplateEngine) {
	doc := document.New()

	doc.AddParagraph("{{companyName}} Official Documentation")
	doc.AddParagraph("")
	doc.AddParagraph("Version: {{version}}")
	doc.AddParagraph("Created: {{createTime}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#block \"content\"}}")
	doc.AddParagraph("默认内容区域")
	doc.AddParagraph("{{/block}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#block \"footer\"}}")
	doc.AddParagraph("版权所有 © {{year}} {{companyName}}")
	doc.AddParagraph("{{/block}}")

	_, err := engine.LoadTemplateFromDocument("base", doc)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// prepare data
	data := document.NewTemplateData()
	data.SetVariable("companyName", "Drumkit")
	data.SetVariable("version", "v1.0.0")
	data.SetVariable("createTime", "2024-12-01")
	data.SetVariable("year", "2024")

	// render template
	resultDoc, err := engine.RenderTemplateToDocument("base", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// save document
	filename := "output/test_basic_inheritance.docx"
	err = resultDoc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// verify document content
	if len(resultDoc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}

	t.Logf("Basic inheritance test completed: %s", filename)
}

// testBlockOverride
func testBlockOverride(t *testing.T, engine *document.TemplateEngine) {
	// create base document
	doc := document.New()

	// add enterprise report template
	doc.AddParagraph("enterprise report template")
	doc.AddParagraph("")
	doc.AddParagraph("{{#block \"header\"}}")
	doc.AddParagraph("standard report header")
	doc.AddParagraph("{{/block}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#block \"main_content\"}}")
	doc.AddParagraph("standard content area")
	doc.AddParagraph("{{/block}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#block \"footer\"}}")
	doc.AddParagraph("standard footer")
	doc.AddParagraph("{{/block}}")

	_, err := engine.LoadTemplateFromDocument("report_base", doc)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// prepare data
	data := document.NewTemplateData()
	data.SetVariable("reportPeriod", "2024年11月")
	data.SetVariable("totalSales", "1,250,000")
	data.SetVariable("growthRate", "15.8")

	// render template
	resultDoc, err := engine.RenderTemplateToDocument("report_base", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// save document
	filename := "output/test_block_override.docx"
	err = resultDoc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// verify document content
	if len(resultDoc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}

	t.Logf("Block override test completed: %s", filename)
}

// TestTemplateInheritanceValidation
func TestTemplateInheritanceValidation(t *testing.T) {
	engine := document.NewTemplateEngine()

	// test 1: block syntax validation
	t.Run("block syntax validation", func(t *testing.T) {
		// create document with correct block syntax
		doc := document.New()
		doc.AddParagraph("{{#block \"content\"}}")
		doc.AddParagraph("this is a valid block")
		doc.AddParagraph("{{/block}}")

		_, err := engine.LoadTemplateFromDocument("valid_blocks", doc)
		if err != nil {
			t.Fatalf("Valid block syntax should not cause error: %v", err)
		}

		t.Log("Block syntax validation passed")
	})

	// test 2: inheritance chain validation
	t.Run("inheritance chain validation", func(t *testing.T) {
		// create base template
		baseDoc := document.New()
		baseDoc.AddParagraph("base template")
		baseDoc.AddParagraph("{{#block \"content\"}}")
		baseDoc.AddParagraph("default content")
		baseDoc.AddParagraph("{{/block}}")

		_, err := engine.LoadTemplateFromDocument("inheritance_base", baseDoc)
		if err != nil {
			t.Fatalf("Failed to load inheritance base template: %v", err)
		}

		// verify that the template is correctly loaded
		template, err := engine.GetTemplate("inheritance_base")
		if err != nil {
			t.Fatalf("Failed to get template: %v", err)
		}

		if template.Name != "inheritance_base" {
			t.Errorf("Expected template name 'inheritance_base', got '%s'", template.Name)
		}

		t.Log("Inheritance chain validation passed")
	})
}

// cleanupInheritanceTestFiles
func cleanupInheritanceTestFiles() {
	files := []string{
		"output/test_basic_inheritance.docx",
		"output/test_block_override.docx",
	}

	for _, file := range files {
		if _, err := os.Stat(file); err == nil {
			os.Remove(file)
		}
	}
}
