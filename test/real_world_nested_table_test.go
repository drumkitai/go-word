// Package test real world nested table test
package test

import (
	"os"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestRealWorldNestedTableScenario real world nested table test
// this test simulates the scenario described in the issue
func TestRealWorldNestedTableScenario(t *testing.T) {
	// ensure output directory exists
	outputDir := "output"
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		err = os.Mkdir(outputDir, 0755)
		if err != nil {
			t.Fatalf("创建输出目录失败: %v", err)
		}
	}

	// create a document simulating a real resume
	doc := document.New()
	doc.AddParagraph("个人简历模板")
	doc.AddParagraph("")

	// create the main table (resume table)
	mainTable, err := doc.CreateTable(&document.TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 9000,
	})
	if err != nil {
		t.Fatalf("create main table failed: %v", err)
	}
	doc.Body.Elements = append(doc.Body.Elements, mainTable)

	// first row: resume title
	mainTable.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text.Content = "简历"
	mainTable.Rows[0].Cells[1].Paragraphs[0].Runs[0].Text.Content = "{{resume}}"

	// second row: family members information
	mainTable.Rows[1].Cells[0].Paragraphs[0].Runs[0].Text.Content = "家庭主要成员及重要社会关系"

	// create the nested family members table (this is the critical part)
	familyTable := &document.Table{
		Properties: &document.TableProperties{
			TableW: &document.TableWidth{
				W:    "4000",
				Type: "dxa",
			},
			TableBorders: &document.TableBorders{
				Top: &document.TableBorder{
					Val:   "single",
					Sz:    "4",
					Color: "000000",
				},
				Left: &document.TableBorder{
					Val:   "single",
					Sz:    "4",
					Color: "000000",
				},
				Bottom: &document.TableBorder{
					Val:   "single",
					Sz:    "4",
					Color: "000000",
				},
				Right: &document.TableBorder{
					Val:   "single",
					Sz:    "4",
					Color: "000000",
				},
			},
		},
		Rows: []document.TableRow{
			// title row
			{
				Cells: []document.TableCell{
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "姓名"}}}}}},
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "年龄"}}}}}},
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "性别"}}}}}},
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "关系"}}}}}},
				},
			},
			// data row (using template syntax)
			{
				Cells: []document.TableCell{
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "{{#each family_members}}{{name}}"}}}}}},
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "{{age}}"}}}}}},
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "{{gender}}"}}}}}},
					{Paragraphs: []document.Paragraph{{Runs: []document.Run{{Text: document.Text{Content: "{{relationship}}{{/each}}"}}}}}},
				},
			},
		},
	}

	// nest the family members table into the main table's cell
	mainTable.Rows[1].Cells[1].Tables = []document.Table{*familyTable}

	// save the template document
	templatePath := "output/resume_template.docx"
	err = doc.Save(templatePath)
	if err != nil {
		t.Fatalf("save template document failed: %v", err)
	}
	t.Logf("template document saved: %s", templatePath)

	// create the template engine and render
	engine := document.NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("resume", doc)
	if err != nil {
		t.Fatalf("load template failed: %v", err)
	}

	// prepare data
	data := document.NewTemplateData()
	data.SetVariable("resume", "张小明的个人简历")
	data.SetList("family_members", []interface{}{
		map[string]interface{}{
			"name":         "张大明",
			"age":          "50",
			"gender":       "男",
			"relationship": "父亲",
		},
		map[string]interface{}{
			"name":         "李红",
			"age":          "48",
			"gender":       "女",
			"relationship": "母亲",
		},
	})

	// render the document
	renderedDoc, err := engine.RenderTemplateToDocument("resume", data)
	if err != nil {
		t.Fatalf("render template failed: %v", err)
	}

	// save the rendered document
	outputPath := "output/resume_rendered.docx"
	err = renderedDoc.Save(outputPath)
	if err != nil {
		t.Fatalf("save rendered document failed: %v", err)
	}
	t.Logf("rendered document saved: %s", outputPath)

	// verify the nested table exists
	tables := renderedDoc.Body.GetTables()
	if len(tables) < 1 {
		t.Fatalf("should have at least 1 main table, actually has %d", len(tables))
	}

	mainTableResult := tables[0]
	if len(mainTableResult.Rows) < 2 {
		t.Fatalf("main table should have 2 rows, actually has %d rows", len(mainTableResult.Rows))
	}

	// critical verification: the nested table must exist
	nestedTables := mainTableResult.Rows[1].Cells[1].Tables
	if len(nestedTables) == 0 {
		t.Fatal("❌ BUG REPRODUCED: the nested table disappeared after rendering! this is the problem described in the issue.")
	}

	t.Log("✅ BUG FIXED: the nested table is successfully preserved after rendering!")

	// verify the nested table's data
	nestedTable := nestedTables[0]
	// should have 3 rows: 1 title row + 2 data rows
	expectedRows := 3
	if len(nestedTable.Rows) != expectedRows {
		t.Errorf("nested table should have %d rows, actually has %d rows", expectedRows, len(nestedTable.Rows))
	}

	// verify the first row data
	if len(nestedTable.Rows) >= 2 {
		firstRow := nestedTable.Rows[1]
		if len(firstRow.Cells) >= 4 {
			name := firstRow.Cells[0].Paragraphs[0].Runs[0].Text.Content
			age := firstRow.Cells[1].Paragraphs[0].Runs[0].Text.Content
			gender := firstRow.Cells[2].Paragraphs[0].Runs[0].Text.Content
			relation := firstRow.Cells[3].Paragraphs[0].Runs[0].Text.Content

			t.Logf("家庭成员数据 - 姓名: %s, 年龄: %s, 性别: %s, 关系: %s",
				name, age, gender, relation)

			if name != "张大明" || age != "50" || gender != "男" || relation != "父亲" {
				t.Errorf("数据渲染不正确")
			}
		}
	}

	t.Log("✅ 真实场景测试通过：简历中的家庭成员嵌套表格正确渲染")
}
