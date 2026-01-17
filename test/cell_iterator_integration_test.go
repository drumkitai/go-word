package test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

// TestCellIteratorIntegration
func TestCellIteratorIntegration(t *testing.T) {
	// 创建测试文档
	doc := document.New()

	// 创建测试表格
	config := &document.TableConfig{
		Rows:  4,
		Cols:  3,
		Width: 6000,
		Data: [][]string{
			{"Name", "Age", "City"},
			{"Zhang San", "25", "Beijing"},
			{"Li Si", "30", "Shanghai"},
			{"Wang Wu", "28", "Guangzhou"},
		},
	}

	table, err := doc.AddTable(config)
	if err != nil {
		t.Fatalf("Failed to create table: %v", err)
	}

	// 测试1: 基本迭代器功能
	t.Run("Basic iterator", func(t *testing.T) {
		iterator := table.NewCellIterator()

		// 验证总数
		expectedTotal := 12
		if iterator.Total() != expectedTotal {
			t.Errorf("Total cell count is incorrect: expected %d, got %d", expectedTotal, iterator.Total())
		}

		// 验证完整遍历
		count := 0
		for iterator.HasNext() {
			cellInfo, err := iterator.Next()
			if err != nil {
				t.Errorf("迭代器错误: %v", err)
				break
			}

			if cellInfo == nil {
				t.Error("Cell information is empty")
				continue
			}

			if cellInfo.Cell == nil {
				t.Error("Cell reference is empty")
			}

			count++
		}

		if count != expectedTotal {
			t.Errorf("Actual traversal count is incorrect: expected %d, got %d", expectedTotal, count)
		}
	})

	// 测试2: 重置功能
	t.Run("Iterator reset", func(t *testing.T) {
		iterator := table.NewCellIterator()

		// 迭代几个单元格
		iterator.Next()
		iterator.Next()

		// 重置并验证
		iterator.Reset()
		row, col := iterator.Current()
		if row != 0 || col != 0 {
			t.Errorf("After reset, position is incorrect: expected (0,0), got (%d,%d)", row, col)
		}
	})

	// 测试3: ForEach功能
	t.Run("ForEach traversal", func(t *testing.T) {
		visitedCount := 0
		err := table.ForEach(func(row, col int, cell *document.TableCell, text string) error {
			visitedCount++
			if cell == nil {
				t.Errorf("Cell at position (%d,%d) is empty", row, col)
			}
			return nil
		})

		if err != nil {
			t.Errorf("ForEach failed: %v", err)
		}

		if visitedCount != 12 {
			t.Errorf("ForEach visited count is incorrect: expected 12, got %d", visitedCount)
		}
	})

	// 测试4: 按行遍历
	t.Run("ForEach in row", func(t *testing.T) {
		for row := 0; row < table.GetRowCount(); row++ {
			visitedCount := 0
			err := table.ForEachInRow(row, func(col int, cell *document.TableCell, text string) error {
				visitedCount++
				return nil
			})

			if err != nil {
				t.Errorf("ForEach in row %d failed: %v", row, err)
			}

			if visitedCount != 3 {
				t.Errorf("ForEach in row %d cell count is incorrect: expected 3, got %d", row, visitedCount)
			}
		}
	})

	// 测试5: 按列遍历
	t.Run("ForEach in column", func(t *testing.T) {
		for col := 0; col < table.GetColumnCount(); col++ {
			visitedCount := 0
			err := table.ForEachInColumn(col, func(row int, cell *document.TableCell, text string) error {
				visitedCount++
				return nil
			})

			if err != nil {
				t.Errorf("ForEach in column %d failed: %v", col, err)
			}

			if visitedCount != 4 {
				t.Errorf("ForEach in column %d cell count is incorrect: expected 4, got %d", col, visitedCount)
			}
		}
	})

	// 测试6: 范围获取
	t.Run("Cell range", func(t *testing.T) {
		// 获取数据区域 (1,0) 到 (3,2)
		cells, err := table.GetCellRange(1, 0, 3, 2)
		if err != nil {
			t.Errorf("Get cell range failed: %v", err)
		}

		expectedCount := 9 // 3行x3列
		if len(cells) != expectedCount {
			t.Errorf("Cell range count is incorrect: expected %d, got %d", expectedCount, len(cells))
		}

		// 验证范围内容
		if cells[0].Row != 1 || cells[0].Col != 0 {
			t.Errorf("Cell range start position is incorrect: expected (1,0), got (%d,%d)", cells[0].Row, cells[0].Col)
		}

		lastIndex := len(cells) - 1
		if cells[lastIndex].Row != 3 || cells[lastIndex].Col != 2 {
			t.Errorf("Cell range end position is incorrect: expected (3,2), got (%d,%d)",
				cells[lastIndex].Row, cells[lastIndex].Col)
		}
	})

	t.Run("	Cell find", func(t *testing.T) {
		cells, err := table.FindCellsByText("Zhang", false)
		if err != nil {
			t.Errorf("Cell find failed: %v", err)
		}

		if len(cells) != 1 {
			t.Errorf("Cell find result count is incorrect: expected 1, got %d", len(cells))
		}

		if len(cells) > 0 && cells[0].Text != "Zhang San" {
			t.Errorf("Cell find content is incorrect: expected 'Zhang San', got '%s'", cells[0].Text)
		}

		// 精确查找
		exactCells, err := table.FindCellsByText("25", true)
		if err != nil {
			t.Errorf("Exact cell find failed: %v", err)
		}

		if len(exactCells) != 1 {
			t.Errorf("Exact cell find result count is incorrect: expected 1, got %d", len(exactCells))
		}
	})

	t.Run("Custom cell find", func(t *testing.T) {
		ageCells, err := table.FindCells(func(row, col int, cell *document.TableCell, text string) bool {
			if col == 1 && row > 0 {
				return text == "30" || text == "28"
			}
			return false
		})

		if err != nil {
			t.Errorf("Custom cell find failed: %v", err)
		}

		if len(ageCells) != 2 {
			t.Errorf("Custom cell find result count is incorrect: expected 2, got %d", len(ageCells))
		}
	})

	outputDir := "../examples/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		t.Logf("Failed to create output directory: %v", err)
	}

	filename := filepath.Join(outputDir, "cell_iterator_integration_test.docx")
	if err := doc.Save(filename); err != nil {
		t.Errorf("Failed to save test document: %v", err)
	} else {
		t.Logf("Test document saved to: %s", filename)
	}
}
