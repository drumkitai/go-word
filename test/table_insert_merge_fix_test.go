package test

import (
	"fmt"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestTableInsertAndMergeFix
func TestTableInsertAndMergeFix(t *testing.T) {
	document.SetGlobalLevel(document.LogLevelInfo)

	t.Run("verify that the properties are deep copied", func(t *testing.T) {
		doc := document.New()

		// create the initial table
		config := &document.TableConfig{
			Rows:  2,
			Cols:  3,
			Width: 6000,
		}

		table, err := doc.AddTable(config)
		if err != nil {
			t.Fatalf("创建表格失败: %v", err)
		}

		// set the text of the first row's cells
		table.SetCellText(0, 0, "Header1")
		table.SetCellText(0, 1, "Header2")
		table.SetCellText(0, 2, "Header3")

		// add a new row
		err = table.AppendRow([]string{"Row2-1", "Row2-2", "Row2-3"})
		if err != nil {
			t.Fatalf("add row failed: %v", err)
		}

		// get the properties of the first row and the new added row's cells
		cell1, _ := table.GetCell(0, 0)
		cell2, _ := table.GetCell(2, 0)

		// verify that the properties are independent (not the same pointer)
		if cell1.Properties == cell2.Properties {
			t.Error("the cell properties should be independent copies, not shared pointers")
		}

		// modify the properties of the new row, it should not affect the first row
		if cell2.Properties != nil && cell2.Properties.TableCellW != nil {
			cell2.Properties.TableCellW.W = "3000"
		}

		// verify that the properties of the first row have not been changed
		if cell1.Properties != nil && cell1.Properties.TableCellW != nil {
			if cell1.Properties.TableCellW.W == "3000" {
				t.Error("modifying the properties of the new row should not affect the first row")
			}
		}
	})

	t.Run("large table dynamic addition and merge", func(t *testing.T) {
		doc := document.New()

		// create a 28 row table
		config := &document.TableConfig{
			Rows:  28,
			Cols:  5,
			Width: 10000,
		}

		table, err := doc.AddTable(config)
		if err != nil {
			t.Fatalf("create table failed: %v", err)
		}

		// fill the data
		for i := 0; i < 28; i++ {
			for j := 0; j < 5; j++ {
				table.SetCellText(i, j, fmt.Sprintf("Cell-%d-%d", i+1, j+1))
			}
		}

		// dynamically add multiple rows
		for i := 29; i <= 35; i++ {
			rowData := make([]string, 5)
			for j := 0; j < 5; j++ {
				rowData[j] = fmt.Sprintf("Cell-%d-%d", i, j+1)
			}
			err := table.AppendRow(rowData)
			if err != nil {
				t.Fatalf("添加第%d行失败: %v", i, err)
			}
		}

		// 在不同位置进行合并
		testCases := []struct {
			name     string
			row      int
			startCol int
			endCol   int
		}{
			{"合并第1行", 0, 1, 3},
			{"合并第15行", 14, 0, 2},
			{"合并第28行", 27, 2, 4},
			{"合并第30行（动态添加的）", 29, 1, 3},
			{"合并最后一行", 34, 0, 1},
		}

		for _, tc := range testCases {
			t.Run(tc.name, func(t *testing.T) {
				err := table.MergeCellsHorizontal(tc.row, tc.startCol, tc.endCol)
				if err != nil {
					t.Errorf("%s失败: %v", tc.name, err)
				}

				// 验证合并后的单元格数
				row := table.Rows[tc.row]
				expectedCells := 5 - (tc.endCol - tc.startCol)
				if len(row.Cells) != expectedCells {
					t.Errorf("%s后单元格数不正确: 期望%d，实际%d",
						tc.name, expectedCells, len(row.Cells))
				}

				// 验证GridSpan设置
				cell, _ := table.GetCell(tc.row, tc.startCol)
				if cell.Properties == nil || cell.Properties.GridSpan == nil {
					t.Errorf("%s后GridSpan未设置", tc.name)
				} else {
					expectedSpan := fmt.Sprintf("%d", tc.endCol-tc.startCol+1)
					if cell.Properties.GridSpan.Val != expectedSpan {
						t.Errorf("%s后GridSpan值不正确: 期望%s，实际%s",
							tc.name, expectedSpan, cell.Properties.GridSpan.Val)
					}
				}
			})
		}

		// 保存文档
		err = doc.Save("test/output/large_table_merge_test.docx")
		if err != nil {
			t.Errorf("保存文档失败: %v", err)
		}
	})

	t.Run("混合操作测试", func(t *testing.T) {
		doc := document.New()

		// 创建初始表格
		config := &document.TableConfig{
			Rows:  5,
			Cols:  4,
			Width: 8000,
		}

		table, err := doc.AddTable(config)
		if err != nil {
			t.Fatalf("创建表格失败: %v", err)
		}

		// 先合并一些单元格
		err = table.MergeCellsHorizontal(1, 1, 2)
		if err != nil {
			t.Fatalf("初始合并失败: %v", err)
		}

		// 添加新行
		for i := 0; i < 3; i++ {
			err := table.AppendRow([]string{"New1", "New2", "New3", "New4"})
			if err != nil {
				t.Fatalf("添加行失败: %v", err)
			}
		}

		// 在新添加的行上进行合并
		err = table.MergeCellsHorizontal(6, 0, 1)
		if err != nil {
			t.Fatalf("合并新行失败: %v", err)
		}

		// 验证表格结构完整性
		if table.GetRowCount() != 8 {
			t.Errorf("表格行数不正确: 期望8，实际%d", table.GetRowCount())
		}

		// 验证每行的单元格数是否正确
		expectedCellCounts := []int{4, 3, 4, 4, 4, 4, 3, 4} // 第1行和第6行有合并
		for i, row := range table.Rows {
			if len(row.Cells) != expectedCellCounts[i] {
				t.Errorf("第%d行单元格数不正确: 期望%d，实际%d",
					i, expectedCellCounts[i], len(row.Cells))
			}
		}

		// 保存文档
		err = doc.Save("test/output/mixed_operations_test.docx")
		if err != nil {
			t.Errorf("保存文档失败: %v", err)
		}
	})
}

// TestTableGridConsistencyAfterFix 测试修复后的表格网格一致性
func TestTableGridConsistencyAfterFix(t *testing.T) {
	doc := document.New()

	// 创建带有自定义列宽的表格
	config := &document.TableConfig{
		Rows:      3,
		Cols:      4,
		Width:     8000,
		ColWidths: []int{1500, 2000, 2500, 2000},
	}

	table, err := doc.AddTable(config)
	if err != nil {
		t.Fatalf("创建表格失败: %v", err)
	}

	// 记录原始列宽
	originalWidths := make([]string, len(table.Grid.Cols))
	for i, col := range table.Grid.Cols {
		originalWidths[i] = col.W
	}

	// 动态添加10行
	for i := 0; i < 10; i++ {
		err := table.AppendRow([]string{
			fmt.Sprintf("A%d", i+4),
			fmt.Sprintf("B%d", i+4),
			fmt.Sprintf("C%d", i+4),
			fmt.Sprintf("D%d", i+4),
		})
		if err != nil {
			t.Fatalf("添加第%d行失败: %v", i+4, err)
		}
	}

	// 验证所有行的单元格宽度与网格定义一致
	for i, row := range table.Rows {
		for j, cell := range row.Cells {
			if cell.Properties == nil || cell.Properties.TableCellW == nil {
				t.Errorf("行%d列%d缺少宽度属性", i, j)
				continue
			}

			expectedWidth := originalWidths[j]
			actualWidth := cell.Properties.TableCellW.W
			if actualWidth != expectedWidth {
				t.Errorf("行%d列%d宽度不一致: 期望%s，实际%s",
					i, j, expectedWidth, actualWidth)
			}
		}
	}

	// 进行一些合并操作
	err = table.MergeCellsHorizontal(5, 1, 2)
	if err != nil {
		t.Fatalf("合并失败: %v", err)
	}

	err = table.MergeCellsVertical(8, 10, 0)
	if err != nil {
		t.Fatalf("垂直合并失败: %v", err)
	}

	// 再次验证未合并单元格的宽度保持一致
	for i, row := range table.Rows {
		cellIndex := 0
		for j := 0; j < len(originalWidths); j++ {
			if cellIndex >= len(row.Cells) {
				break
			}

			cell := row.Cells[cellIndex]

			// 跳过被合并掉的单元格
			if i == 5 && (j == 2 || j == 3) {
				// 这些单元格在第5行被水平合并了
				continue
			}

			if cell.Properties != nil && cell.Properties.TableCellW != nil {
				expectedWidth := originalWidths[j]
				actualWidth := cell.Properties.TableCellW.W
				if actualWidth != expectedWidth {
					// 合并的单元格可能有不同的宽度
					if cell.Properties.GridSpan == nil && cell.Properties.VMerge == nil {
						t.Errorf("行%d单元格%d宽度不一致: 期望%s，实际%s",
							i, cellIndex, expectedWidth, actualWidth)
					}
				}
			}

			cellIndex++
		}
	}

	// 保存文档
	err = doc.Save("test/output/grid_consistency_after_fix.docx")
	if err != nil {
		t.Errorf("保存文档失败: %v", err)
	}
}
