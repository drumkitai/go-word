package test

import (
	"fmt"
	"testing"
	"time"

	"github.com/drumkitai/go-word/pkg/document"
	"github.com/drumkitai/go-word/pkg/style"
)

// TestTemplateStylePreservation tests style preservation during template rendering
func TestTemplateStylePreservation(t *testing.T) {
	fmt.Println("=== template style preservation test ===")

	templateDoc := document.New()

	titlePara := templateDoc.AddParagraph("Project Report: {{projectName}}")
	titlePara.SetStyle(style.StyleHeading1)

	subtitlePara := templateDoc.AddParagraph("Reporter: {{author}}")
	subtitlePara.SetStyle(style.StyleHeading2)

	bodyPara := templateDoc.AddParagraph("")

	normalRun := &document.Run{
		Text: document.Text{Content: "project status: "},
		Properties: &document.RunProperties{
			FontFamily: &document.FontFamily{
				ASCII:    "Arial",
				HAnsi:    "Arial",
				EastAsia: "Microsoft YaHei",
			},
			FontSize: &document.FontSize{Val: "24"},
		},
	}

	boldRun := &document.Run{
		Properties: &document.RunProperties{
			Bold: &document.Bold{},
			FontFamily: &document.FontFamily{
				ASCII:    "Arial",
				HAnsi:    "Arial",
				EastAsia: "Microsoft YaHei",
			},
			FontSize: &document.FontSize{Val: "24"},
			Color:    &document.Color{Val: "FF0000"},
		},
		Text: document.Text{Content: "{{status}}"},
	}

	endRun := &document.Run{
		Text: document.Text{Content: ", progress: "},
		Properties: &document.RunProperties{
			FontFamily: &document.FontFamily{
				ASCII:    "Arial",
				HAnsi:    "Arial",
				EastAsia: "Microsoft YaHei",
			},
			FontSize: &document.FontSize{Val: "24"},
		},
	}

	progressRun := &document.Run{
		Properties: &document.RunProperties{
			Bold:   &document.Bold{},
			Italic: &document.Italic{},
			FontFamily: &document.FontFamily{
				ASCII:    "Arial",
				HAnsi:    "Arial",
				EastAsia: "Microsoft YaHei",
			},
			FontSize: &document.FontSize{Val: "28"},
			Color:    &document.Color{Val: "008000"},
		},
		Text: document.Text{Content: "{{progress}}%"},
	}

	bodyPara.Runs = []document.Run{*normalRun, *boldRun, *endRun, *progressRun}

	// create table with styles
	tableConfig := &document.TableConfig{
		Rows: 2,
		Cols: 3,
	}
	table, err := templateDoc.AddTable(tableConfig)
	if err != nil {
		t.Fatalf("create table failed: %v", err)
	}

	headerRow := table.Rows[0]
	for i, cell := range headerRow.Cells {
		headers := []string{"task", "assignee", "status"}
		para := &document.Paragraph{
			Properties: &document.ParagraphProperties{
				ParagraphStyle: &document.ParagraphStyle{Val: style.StyleHeading3},
			},
			Runs: []document.Run{{
				Text: document.Text{Content: headers[i]},
				Properties: &document.RunProperties{
					Bold: &document.Bold{},
					FontFamily: &document.FontFamily{
						ASCII:    "Calibri",
						HAnsi:    "Calibri",
						EastAsia: "宋体",
					},
					FontSize: &document.FontSize{Val: "22"},  // 11pt
					Color:    &document.Color{Val: "FFFFFF"}, // white
				},
			}},
		}
		cell.Paragraphs = []document.Paragraph{*para}
	}

	// set data row template
	dataRow := table.Rows[1]
	templates := []string{"{{#each tasks}}", "{{name}}", "{{assignee}}", "{{status}}", "{{/each}}"}
	for i, cell := range dataRow.Cells {
		if i < len(templates) {
			para := &document.Paragraph{
				Properties: &document.ParagraphProperties{
					ParagraphStyle: &document.ParagraphStyle{Val: style.StyleNormal},
				},
				Runs: []document.Run{{
					Text: document.Text{Content: templates[i]},
					Properties: &document.RunProperties{
						FontFamily: &document.FontFamily{
							ASCII:    "Times New Roman",
							HAnsi:    "Times New Roman",
							EastAsia: "宋体",
						},
						FontSize: &document.FontSize{Val: "20"},  // 10pt
						Color:    &document.Color{Val: "000080"}, // dark blue
					},
				}},
			}
			cell.Paragraphs = []document.Paragraph{*para}
		}
	}

	// save template document
	templateFile := "test/output/style_template.docx"
	err = templateDoc.Save(templateFile)
	if err != nil {
		t.Fatalf("save template document failed: %v", err)
	}
	fmt.Printf("✓ create style template document: %s\n", templateFile)

	// 2. load template from template document
	engine := document.NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("style_template", templateDoc)
	if err != nil {
		t.Fatalf("load template failed: %v", err)
	}

	// 3. prepare test data
	data := document.NewTemplateData()
	data.SetVariable("projectName", "go-word development project")
	data.SetVariable("author", "Winston")
	data.SetVariable("status", "in progress")
	data.SetVariable("progress", "55")

	// set table data
	tasks := []interface{}{
		map[string]interface{}{
			"name":     "document parsing",
			"assignee": "John Doe",
			"status":   "completed",
		},
		map[string]interface{}{
			"name":     "style system",
			"assignee": "Jane Smith",
			"status":   "in progress",
		},
		map[string]interface{}{
			"name":     "test case",
			"assignee": "Jim Beam",
			"status":   "pending",
		},
	}
	data.SetList("tasks", tasks)

	// 4. render template
	resultDoc, err := engine.RenderTemplateToDocument("style_template", data)
	if err != nil {
		t.Fatalf("render template failed: %v", err)
	}

	// 5. save result document
	outputFile := "test/output/style_result_" + time.Now().Format("20060102_150405") + ".docx"
	err = resultDoc.Save(outputFile)
	if err != nil {
		t.Fatalf("save result document failed: %v", err)
	}

	fmt.Printf("✓ generate result document: %s\n", outputFile)

	// 6. verify if styles are preserved
	verifyDocumentStyles(t, resultDoc)

	fmt.Println("✓ template style preservation test completed")
}

// verifyDocumentStyles
func verifyDocumentStyles(t *testing.T, doc *document.Document) {
	fmt.Println("\n=== verify if document styles are correctly preserved ===")

	// check document elements
	if len(doc.Body.Elements) == 0 {
		t.Error("document has no elements")
		return
	}

	elementCount := 0
	styledElements := 0

	for i, element := range doc.Body.Elements {
		elementCount++

		switch elem := element.(type) {
		case *document.Paragraph:
			fmt.Printf("paragraph %d: ", i+1)

			// check paragraph style
			if elem.Properties != nil && elem.Properties.ParagraphStyle != nil {
				fmt.Printf("paragraph style=%s, ", elem.Properties.ParagraphStyle.Val)
				styledElements++
			} else {
				fmt.Printf("paragraph style=none, ")
			}

			// check run style
			runStyleCount := 0
			for j, run := range elem.Runs {
				if run.Properties != nil {
					runStyleCount++

					// check key style attributes
					hasFont := run.Properties.FontFamily != nil
					hasBold := run.Properties.Bold != nil
					hasColor := run.Properties.Color != nil
					hasSize := run.Properties.FontSize != nil

					fmt.Printf("run%d(font:%t,bold:%t,color:%t,size:%t) ",
						j+1, hasFont, hasBold, hasColor, hasSize)
				}
			}

			fmt.Printf("(total %d runs with styles)\n", runStyleCount)

		case *document.Table:
			fmt.Printf("table %d: %d rows %d columns\n", i+1, len(elem.Rows), len(elem.Rows[0].Cells))

			// check table style
			tableStyledCells := 0
			for rowIdx, row := range elem.Rows {
				for cellIdx, cell := range row.Cells {
					for paraIdx, para := range cell.Paragraphs {
						if para.Properties != nil && para.Properties.ParagraphStyle != nil {
							tableStyledCells++
							fmt.Printf("  row%d column%d paragraph%d: style=%s\n",
								rowIdx+1, cellIdx+1, paraIdx+1, para.Properties.ParagraphStyle.Val)
						}

						for runIdx, run := range para.Runs {
							if run.Properties != nil {
								hasFont := run.Properties.FontFamily != nil
								hasBold := run.Properties.Bold != nil
								hasColor := run.Properties.Color != nil
								hasSize := run.Properties.FontSize != nil

								if hasFont || hasBold || hasColor || hasSize {
									fmt.Printf("    run%d: font:%t,bold:%t,color:%t,size:%t\n",
										runIdx+1, hasFont, hasBold, hasColor, hasSize)
								}
							}
						}
					}
				}
			}

			if tableStyledCells > 0 {
				styledElements++
				fmt.Printf("  table has %d styled cells\n", tableStyledCells)
			}
		}
	}

	fmt.Printf("\nsummary: total %d elements, %d elements with styles\n", elementCount, styledElements)

	// basic verification
	if elementCount == 0 {
		t.Error("document has no elements")
	}

	if styledElements == 0 {
		t.Error("❌ serious problem: all styles are lost!")
	} else {
		fmt.Printf("✓ detected %d elements with styles\n", styledElements)
	}
}

// TestTemplateStyleIssues
func TestTemplateStyleIssues(t *testing.T) {
	fmt.Println("\n=== template style issues diagnosis test ===")

	// create simple template document
	templateDoc := document.New()

	// add a styled paragraph
	para := templateDoc.AddParagraph("test variable: {{testVar}}")
	para.SetStyle(style.StyleHeading1)

	// check template document styles
	fmt.Println("模板文档检查:")
	if para.Properties != nil && para.Properties.ParagraphStyle != nil {
		fmt.Printf("✓ 段落样式: %s\n", para.Properties.ParagraphStyle.Val)
	} else {
		fmt.Println("❌ 段落没有样式")
	}

	// check style manager
	styleManager := templateDoc.GetStyleManager()
	heading1Style := styleManager.GetStyle(style.StyleHeading1)
	if heading1Style != nil {
		fmt.Printf("✓ StyleManager中存在Heading1样式: %s\n", heading1Style.StyleID)
	} else {
		fmt.Println("❌ StyleManager中缺少Heading1样式")
	}

	// load as template
	engine := document.NewTemplateEngine()
	template, err := engine.LoadTemplateFromDocument("test_template", templateDoc)
	if err != nil {
		t.Fatalf("load template failed: %v", err)
	}

	// check template's BaseDoc
	if template.BaseDoc != nil {
		fmt.Println("✓ template has BaseDoc")

		// check BaseDoc's style manager
		if template.BaseDoc.GetStyleManager() != nil {
			fmt.Println("✓ BaseDoc has style manager")

			baseHeading1 := template.BaseDoc.GetStyleManager().GetStyle(style.StyleHeading1)
			if baseHeading1 != nil {
				fmt.Printf("✓ BaseDoc style manager has Heading1: %s\n", baseHeading1.StyleID)
			} else {
				fmt.Println("❌ BaseDoc style manager is missing Heading1")
			}
		} else {
			fmt.Println("❌ BaseDoc has no style manager")
		}
	} else {
		fmt.Println("❌ template has no BaseDoc")
	}

	// render template
	data := document.NewTemplateData()
	data.SetVariable("testVar", "test value")

	resultDoc, err := engine.RenderTemplateToDocument("test_template", data)
	if err != nil {
		t.Fatalf("render template failed: %v", err)
	}

	// check result document
	fmt.Println("\nresult document check:")
	if len(resultDoc.Body.Elements) > 0 {
		if para, ok := resultDoc.Body.Elements[0].(*document.Paragraph); ok {
			if para.Properties != nil && para.Properties.ParagraphStyle != nil {
				fmt.Printf("✓ result paragraph style: %s\n", para.Properties.ParagraphStyle.Val)
			} else {
				fmt.Println("❌ result paragraph has no style")
			}

			// check text content
			fullText := ""
			for _, run := range para.Runs {
				fullText += run.Text.Content
			}
			fmt.Printf("text content: %s\n", fullText)
		}
	}

	// check result document's style manager
	resultStyleManager := resultDoc.GetStyleManager()
	if resultStyleManager != nil {
		fmt.Println("✓ result document has style manager")

		resultHeading1 := resultStyleManager.GetStyle(style.StyleHeading1)
		if resultHeading1 != nil {
			fmt.Printf("✓ result style manager has Heading1: %s\n", resultHeading1.StyleID)
		} else {
			fmt.Println("❌ result style manager is missing Heading1")
		}
	} else {
		fmt.Println("❌ result document has no style manager")
	}

	// save result for manual inspection
	outputFile := "test/output/style_diagnosis_" + time.Now().Format("20060102_150405") + ".docx"
	err = resultDoc.Save(outputFile)
	if err != nil {
		t.Fatalf("save result document failed: %v", err)
	}
	fmt.Printf("✓ save diagnosis result to: %s\n", outputFile)
}
