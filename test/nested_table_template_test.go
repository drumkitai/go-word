// Package test Nested table template test
package test

import (
	"os"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestNestedTableTemplate Test nested table template functionality
func TestNestedTableTemplate(t *testing.T) {
	outputDir := "output"
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		err = os.Mkdir(outputDir, 0755)
		if err != nil {
			t.Fatalf("创建输出目录失败: %v", err)
		}
	}

	doc := document.New()
	doc.AddParagraph("嵌套表格测试文档")
	doc.AddParagraph("")

	outerTable, err := doc.CreateTable(&document.TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 9000,
	})
	if err != nil {
		t.Fatalf("Failed to create outer table: %v", err)
	}
	doc.Body.Elements = append(doc.Body.Elements, outerTable)

	outerTable.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text.Content = "Resume"
	outerTable.Rows[0].Cells[1].Paragraphs[0].Runs[0].Text.Content = "{{resume}}"

	outerTable.Rows[1].Cells[0].Paragraphs[0].Runs[0].Text.Content = "Family members and important social relationships"

	innerTable := &document.Table{
		Properties: &document.TableProperties{
			TableW: &document.TableWidth{
				W:    "4000",
				Type: "dxa",
			},
		},
		Rows: []document.TableRow{
			{
				Cells: []document.TableCell{
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "Name"},
									},
								},
							},
						},
					},
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "Age"},
									},
								},
							},
						},
					},
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "Gender"},
									},
								},
							},
						},
					},
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "Relationship"},
									},
								},
							},
						},
					},
				},
			},
			{
				Cells: []document.TableCell{
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "{{#each family_members}}{{name}}"},
									},
								},
							},
						},
					},
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "{{age}}{{/each}}"},
									},
								},
							},
						},
					},
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "{{gender}}"},
									},
								},
							},
						},
					},
					{
						Paragraphs: []document.Paragraph{
							{
								Runs: []document.Run{
									{
										Text: document.Text{Content: "{{relationship}}"},
									},
								},
							},
						},
					},
				},
			},
		},
	}

	outerTable.Rows[1].Cells[1].Tables = []document.Table{*innerTable}
	if len(outerTable.Rows[1].Cells[1].Paragraphs) == 0 {
		outerTable.Rows[1].Cells[1].Paragraphs = []document.Paragraph{
			{
				Runs: []document.Run{
					{
						Text: document.Text{Content: ""},
					},
				},
			},
		}
	}

	templatePath := "output/nested_table_template.docx"
	err = doc.Save(templatePath)
	if err != nil {
		t.Fatalf("Failed to save template document: %v", err)
	}

	engine := document.NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("nested_template", doc)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	data := document.NewTemplateData()
	data.SetVariable("resume", "Resume")
	data.SetList("family_members", []interface{}{
		map[string]interface{}{
			"name":         "John Doe",
			"age":          "45",
			"gender":       "Male",
			"relationship": "Father",
		},
		map[string]interface{}{
			"name":         "Jane Doe",
			"age":          "43",
			"gender":       "Female",
			"relationship": "Mother",
		},
		map[string]interface{}{
			"name":         "Jack Doe",
			"age":          "20",
			"gender":       "Male",
			"relationship": "Myself",
		},
	})

	renderedDoc, err := engine.RenderTemplateToDocument("nested_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	outputPath := "output/nested_table_rendered.docx"
	err = renderedDoc.Save(outputPath)
	if err != nil {
		t.Fatalf("Failed to save rendered document: %v", err)
	}

	tables := renderedDoc.Body.GetTables()
	if len(tables) < 1 {
		t.Fatalf("Rendered document should至少包含1个表格，实际包含 %d 个", len(tables))
	}

	outerTableRendered := tables[0]
	if len(outerTableRendered.Rows) < 2 {
		t.Fatalf("Outer table should have at least 2 rows, got %d", len(outerTableRendered.Rows))
	}

	nestedTables := outerTableRendered.Rows[1].Cells[1].Tables
	if len(nestedTables) == 0 {
		t.Fatalf("Nested table disappeared after rendering! Should exist 1 nested table")
	}

	nestedTable := nestedTables[0]
	expectedRows := 4 // 1 header + 3 data rows
	if len(nestedTable.Rows) != expectedRows {
		t.Errorf("Nested table should have %d rows, got %d", expectedRows, len(nestedTable.Rows))
	}

	if len(nestedTable.Rows) >= 2 {
		firstDataRow := nestedTable.Rows[1]
		if len(firstDataRow.Cells) >= 1 {
			nameCell := firstDataRow.Cells[0]
			if len(nameCell.Paragraphs) > 0 && len(nameCell.Paragraphs[0].Runs) > 0 {
				name := nameCell.Paragraphs[0].Runs[0].Text.Content
				if name != "John Doe" {
					t.Errorf("First row data name should be 'John Doe', got '%s'", name)
				}
			}
		}
	}

	t.Logf("Nested table template test passed!")
	t.Logf("Template document: %s", templatePath)
	t.Logf("Rendered document: %s", outputPath)
}
