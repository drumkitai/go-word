// Package document provides table operation functionality for Word documents
package document

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// Table represents a table
type Table struct {
	XMLName    xml.Name         `xml:"w:tbl"`
	Properties *TableProperties `xml:"w:tblPr,omitempty"`
	Grid       *TableGrid       `xml:"w:tblGrid,omitempty"`
	Rows       []TableRow       `xml:"w:tr"`
}

// TableProperties represents table properties
type TableProperties struct {
	XMLName      xml.Name          `xml:"w:tblPr"`
	TableW       *TableWidth       `xml:"w:tblW,omitempty"`
	TableJc      *TableJc          `xml:"w:jc,omitempty"`
	TableLook    *TableLook        `xml:"w:tblLook,omitempty"`
	TableStyle   *TableStyle       `xml:"w:tblStyle,omitempty"`
	TableBorders *TableBorders     `xml:"w:tblBorders,omitempty"`
	Shd          *TableShading     `xml:"w:shd,omitempty"`
	TableCellMar *TableCellMargins `xml:"w:tblCellMar,omitempty"`
	TableLayout  *TableLayoutType  `xml:"w:tblLayout,omitempty"`
	TableInd     *TableIndentation `xml:"w:tblInd,omitempty"`
}

// TableWidth represents table width
type TableWidth struct {
	XMLName xml.Name `xml:"w:tblW"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// TableJc represents table alignment
type TableJc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

// TableLook represents table appearance
type TableLook struct {
	XMLName  xml.Name `xml:"w:tblLook"`
	Val      string   `xml:"w:val,attr"`
	FirstRow string   `xml:"w:firstRow,attr,omitempty"`
	LastRow  string   `xml:"w:lastRow,attr,omitempty"`
	FirstCol string   `xml:"w:firstColumn,attr,omitempty"`
	LastCol  string   `xml:"w:lastColumn,attr,omitempty"`
	NoHBand  string   `xml:"w:noHBand,attr,omitempty"`
	NoVBand  string   `xml:"w:noVBand,attr,omitempty"`
}

// TableGrid represents table grid
type TableGrid struct {
	XMLName xml.Name       `xml:"w:tblGrid"`
	Cols    []TableGridCol `xml:"w:gridCol"`
}

// TableGridCol represents a table grid column
type TableGridCol struct {
	XMLName xml.Name `xml:"w:gridCol"`
	W       string   `xml:"w:w,attr,omitempty"`
}

// TableRow represents a table row
type TableRow struct {
	XMLName    xml.Name            `xml:"w:tr"`
	Properties *TableRowProperties `xml:"w:trPr,omitempty"`
	Cells      []TableCell         `xml:"w:tc"`
}

// TableRowProperties represents table row properties
type TableRowProperties struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	TableRowH *TableRowH `xml:"w:trHeight,omitempty"`
	CantSplit *CantSplit `xml:"w:cantSplit,omitempty"` // prevent page break within row
	TblHeader *TblHeader `xml:"w:tblHeader,omitempty"` // repeat header row
}

// TableRowH represents table row height
type TableRowH struct {
	XMLName xml.Name `xml:"w:trHeight"`
	Val     string   `xml:"w:val,attr,omitempty"`
	HRule   string   `xml:"w:hRule,attr,omitempty"`
}

// TableCell represents a table cell
type TableCell struct {
	XMLName    xml.Name             `xml:"w:tc"`
	Properties *TableCellProperties `xml:"w:tcPr,omitempty"`
	Paragraphs []Paragraph          `xml:"w:p"`
	Tables     []Table              `xml:"w:tbl"` // supports nested tables
}

// MarshalXML custom XML serialization to ensure nested tables are properly serialized
// OOXML requirement: cell content should output paragraphs and tables in their original document order
func (tc *TableCell) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// start element <w:tc>
	start.Name = xml.Name{Local: "w:tc"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// serialize properties <w:tcPr>
	if tc.Properties != nil {
		if err := e.Encode(tc.Properties); err != nil {
			return err
		}
	}

	// serialize paragraphs <w:p>
	for i := range tc.Paragraphs {
		if err := e.Encode(&tc.Paragraphs[i]); err != nil {
			return err
		}
	}

	// serialize nested tables <w:tbl>
	for i := range tc.Tables {
		if err := e.Encode(&tc.Tables[i]); err != nil {
			return err
		}
	}

	// end element </w:tc>
	return e.EncodeToken(start.End())
}

// TableCellProperties represents table cell properties
type TableCellProperties struct {
	XMLName       xml.Name              `xml:"w:tcPr"`
	TableCellW    *TableCellW           `xml:"w:tcW,omitempty"`
	VAlign        *VAlign               `xml:"w:vAlign,omitempty"`
	GridSpan      *GridSpan             `xml:"w:gridSpan,omitempty"`
	VMerge        *VMerge               `xml:"w:vMerge,omitempty"`
	TextDirection *TextDirection        `xml:"w:textDirection,omitempty"`
	Shd           *TableCellShading     `xml:"w:shd,omitempty"`
	TcBorders     *TableCellBorders     `xml:"w:tcBorders,omitempty"`
	TcMar         *TableCellMarginsCell `xml:"w:tcMar,omitempty"`
	NoWrap        *NoWrap               `xml:"w:noWrap,omitempty"`
	HideMark      *HideMark             `xml:"w:hideMark,omitempty"`
}

// TableCellMarginsCell represents cell margins (different XML structure from table margins)
type TableCellMarginsCell struct {
	XMLName xml.Name            `xml:"w:tcMar"`
	Top     *TableCellSpaceCell `xml:"w:top,omitempty"`
	Left    *TableCellSpaceCell `xml:"w:left,omitempty"`
	Bottom  *TableCellSpaceCell `xml:"w:bottom,omitempty"`
	Right   *TableCellSpaceCell `xml:"w:right,omitempty"`
}

// TableCellSpaceCell represents cell spacing settings
type TableCellSpaceCell struct {
	W    string `xml:"w:w,attr"`
	Type string `xml:"w:type,attr"`
}

// TableCellW represents cell width
type TableCellW struct {
	XMLName xml.Name `xml:"w:tcW"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// VAlign represents vertical alignment
type VAlign struct {
	XMLName xml.Name `xml:"w:vAlign"`
	Val     string   `xml:"w:val,attr"`
}

// GridSpan represents grid span (column merge)
type GridSpan struct {
	XMLName xml.Name `xml:"w:gridSpan"`
	Val     string   `xml:"w:val,attr"`
}

// VMerge represents vertical merge (row merge)
type VMerge struct {
	XMLName xml.Name `xml:"w:vMerge"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// TableConfig represents table configuration
type TableConfig struct {
	Rows      int        // number of rows
	Cols      int        // number of columns
	Width     int        // table total width (points)
	ColWidths []int      // each column width (points), if empty then distribute evenly
	Data      [][]string // initial data
	Emphases  [][]int    // cell style 1=italic 2=bold
}

// CreateTable creates a new table
// Parameters:
//   - config: table configuration
//
// Returns:
//   - *Table: the created table object
//   - error: returns error if configuration is invalid
func (d *Document) CreateTable(config *TableConfig) (*Table, error) {
	if config.Rows <= 0 || config.Cols <= 0 {
		Error("Table rows and columns must be greater than 0")
		return nil, NewValidationError("TableConfig", "", "Table rows and columns must be greater than 0")
	}

	table := &Table{
		Properties: &TableProperties{
			TableW: &TableWidth{
				W:    fmt.Sprintf("%d", config.Width),
				Type: "dxa", // unit in points
			},
			TableJc: &TableJc{
				Val: "center", // default center alignment
			},
			TableLook: &TableLook{
				Val:      "04A0",
				FirstRow: "1",
				LastRow:  "0",
				FirstCol: "1",
				LastCol:  "0",
				NoHBand:  "0",
				NoVBand:  "1",
			},
			// add default table borders, using single line border style same as reference table
			TableBorders: &TableBorders{
				Top: &TableBorder{
					Val:   "single", // single line border style
					Sz:    "4",      // border thickness (1/8 points)
					Space: "0",      // border spacing
					Color: "auto",   // automatic color
				},
				Left: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Bottom: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Right: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				InsideH: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				InsideV: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
			},
			// add table layout and cell margin settings, consistent with reference table
			TableLayout: &TableLayoutType{
				Type: "autofit",
			},
			TableCellMar: &TableCellMargins{
				Left: &TableCellSpace{
					W:    "108",
					Type: "dxa",
				},
				Right: &TableCellSpace{
					W:    "108",
					Type: "dxa",
				},
			},
		},
		Grid: &TableGrid{},
		Rows: make([]TableRow, 0, config.Rows),
	}

	colWidths := config.ColWidths
	if len(colWidths) == 0 {
		avgWidth := config.Width / config.Cols
		colWidths = make([]int, config.Cols)
		for i := range colWidths {
			colWidths[i] = avgWidth
		}
	} else if len(colWidths) != config.Cols {
		Error("column width count does not match column count")
		return nil, NewValidationError("TableConfig.ColWidths", "", "column width count does not match column count")
	}

	for _, width := range colWidths {
		table.Grid.Cols = append(table.Grid.Cols, TableGridCol{
			W: fmt.Sprintf("%d", width),
		})
	}


	for i := 0; i < config.Rows; i++ {
		row := TableRow{
			Cells: make([]TableCell, 0, config.Cols),
		}

		for j := 0; j < config.Cols; j++ {
			cell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: &TableCellW{
						W:    fmt.Sprintf("%d", colWidths[j]),
						Type: "dxa",
					},
					VAlign: &VAlign{
						Val: "center",
					},
				},
				Paragraphs: []Paragraph{
					{
						Runs: []Run{
							{
								Text: Text{
									Content: "",
								},
							},
						},
					},
				},
			}

			// if initial data exists, set cell content
			if config.Data != nil && i < len(config.Data) && j < len(config.Data[i]) {
				cell.Paragraphs[0].Runs[0].Text.Content = config.Data[i][j]
			}

			if config.Emphases != nil && i < len(config.Emphases) && j < len(config.Emphases[i]) {
				switch config.Emphases[i][j] {
				case 1:
					cell.Paragraphs[0].Runs[0].Properties = &RunProperties{Italic: &Italic{}}
				case 2:
					cell.Paragraphs[0].Runs[0].Properties = &RunProperties{Bold: &Bold{}}
				}
			}

			row.Cells = append(row.Cells, cell)
		}

		table.Rows = append(table.Rows, row)
	}

	Info(fmt.Sprintf("table created successfully: %d rows x %d columns", config.Rows, config.Cols))
	return table, nil
}

// AddTable adds a table to the document
// Parameters:
//   - config: table configuration
//
// Returns:
//   - *Table: the added table object
//   - error: error if configuration is invalid
func (d *Document) AddTable(config *TableConfig) (*Table, error) {
	table, err := d.CreateTable(config)
	if err != nil {
		return nil, err
	}

	d.Body.Elements = append(d.Body.Elements, table)

	Info(fmt.Sprintf("table added to document, document now contains %d tables", len(d.Body.GetTables())))
	return table, nil
}

// InsertRow inserts a row at the specified position
func (t *Table) InsertRow(position int, data []string) error {
	if position < 0 || position > len(t.Rows) {
		return fmt.Errorf("invalid insertion position: %d, table has %d rows", position, len(t.Rows))
	}

	if len(t.Rows) == 0 {
		return fmt.Errorf("table has no column definitions, cannot insert row")
	}

	colCount := len(t.Rows[0].Cells)
	if len(data) > colCount {
		return fmt.Errorf("data column count (%d) exceeds table column count (%d)", len(data), colCount)
	}

	newRow := TableRow{
		Cells: make([]TableCell, colCount),
	}

	templateRow := t.Rows[0]
	for i := 0; i < colCount; i++ {
		var cellProps *TableCellProperties
		if templateRow.Cells[i].Properties != nil {
			cellProps = &TableCellProperties{}
			// 复制宽度
			if templateRow.Cells[i].Properties.TableCellW != nil {
				cellProps.TableCellW = &TableCellW{
					W:    templateRow.Cells[i].Properties.TableCellW.W,
					Type: templateRow.Cells[i].Properties.TableCellW.Type,
				}
			}
			// 复制垂直对齐
			if templateRow.Cells[i].Properties.VAlign != nil {
				cellProps.VAlign = &VAlign{
					Val: templateRow.Cells[i].Properties.VAlign.Val,
				}
			}
			// 复制其他必要的属性
			// 注意：不要复制GridSpan和VMerge，因为这些是合并相关的属性
		}

		newRow.Cells[i] = TableCell{
			Properties: cellProps,
			Paragraphs: []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{
								Content: "",
							},
						},
					},
				},
			},
		}

		// 设置数据
		if i < len(data) {
			newRow.Cells[i].Paragraphs[0].Runs[0].Text.Content = data[i]
		}
	}

	// 插入行
	if position == len(t.Rows) {
		// 在末尾添加
		t.Rows = append(t.Rows, newRow)
	} else {
		// 在中间插入
		t.Rows = append(t.Rows[:position+1], t.Rows[position:]...)
		t.Rows[position] = newRow
	}

	Info(fmt.Sprintf("在位置%d插入行成功", position))
	return nil
}

// AppendRow 在表格末尾添加行
func (t *Table) AppendRow(data []string) error {
	return t.InsertRow(len(t.Rows), data)
}

// DeleteRow 删除指定行
func (t *Table) DeleteRow(rowIndex int) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	if len(t.Rows) <= 1 {
		return fmt.Errorf("表格至少需要保留一行")
	}

	// 删除行
	t.Rows = append(t.Rows[:rowIndex], t.Rows[rowIndex+1:]...)

	Info(fmt.Sprintf("删除第%d行成功", rowIndex))
	return nil
}

// DeleteRows 删除指定范围的行
func (t *Table) DeleteRows(startIndex, endIndex int) error {
	if startIndex < 0 || endIndex >= len(t.Rows) || startIndex > endIndex {
		return fmt.Errorf("行索引范围无效：[%d, %d]，表格共有%d行", startIndex, endIndex, len(t.Rows))
	}

	deleteCount := endIndex - startIndex + 1
	if len(t.Rows)-deleteCount < 1 {
		return fmt.Errorf("删除后表格至少需要保留一行")
	}

	// 删除行范围
	t.Rows = append(t.Rows[:startIndex], t.Rows[endIndex+1:]...)

	Info(fmt.Sprintf("删除第%d到%d行成功", startIndex, endIndex))
	return nil
}

// InsertColumn 在指定位置插入列
func (t *Table) InsertColumn(position int, data []string, width int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有行，无法插入列")
	}

	colCount := len(t.Rows[0].Cells)
	if position < 0 || position > colCount {
		return fmt.Errorf("插入位置无效：%d，表格共有%d列", position, colCount)
	}

	if len(data) > len(t.Rows) {
		return fmt.Errorf("数据行数(%d)超过表格行数(%d)", len(data), len(t.Rows))
	}

	// 更新表格网格
	newGridCol := TableGridCol{
		W: fmt.Sprintf("%d", width),
	}
	if position == len(t.Grid.Cols) {
		t.Grid.Cols = append(t.Grid.Cols, newGridCol)
	} else {
		t.Grid.Cols = append(t.Grid.Cols[:position+1], t.Grid.Cols[position:]...)
		t.Grid.Cols[position] = newGridCol
	}

	// 为每一行插入新单元格
	for i := range t.Rows {
		newCell := TableCell{
			Properties: &TableCellProperties{
				TableCellW: &TableCellW{
					W:    fmt.Sprintf("%d", width),
					Type: "dxa",
				},
				VAlign: &VAlign{
					Val: "center",
				},
			},
			Paragraphs: []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{
								Content: "",
							},
						},
					},
				},
			},
		}

		// 设置数据
		if i < len(data) {
			newCell.Paragraphs[0].Runs[0].Text.Content = data[i]
		}

		// 插入单元格
		if position == len(t.Rows[i].Cells) {
			t.Rows[i].Cells = append(t.Rows[i].Cells, newCell)
		} else {
			t.Rows[i].Cells = append(t.Rows[i].Cells[:position+1], t.Rows[i].Cells[position:]...)
			t.Rows[i].Cells[position] = newCell
		}
	}

	Info(fmt.Sprintf("在位置%d插入列成功", position))
	return nil
}

// AppendColumn 在表格末尾添加列
func (t *Table) AppendColumn(data []string, width int) error {
	colCount := 0
	if len(t.Rows) > 0 {
		colCount = len(t.Rows[0].Cells)
	}
	return t.InsertColumn(colCount, data, width)
}

// DeleteColumn 删除指定列
func (t *Table) DeleteColumn(colIndex int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有行")
	}

	colCount := len(t.Rows[0].Cells)
	if colIndex < 0 || colIndex >= colCount {
		return fmt.Errorf("列索引无效：%d，表格共有%d列", colIndex, colCount)
	}

	if colCount <= 1 {
		return fmt.Errorf("表格至少需要保留一列")
	}

	// 删除网格列
	t.Grid.Cols = append(t.Grid.Cols[:colIndex], t.Grid.Cols[colIndex+1:]...)

	// 删除每行的对应单元格
	for i := range t.Rows {
		t.Rows[i].Cells = append(t.Rows[i].Cells[:colIndex], t.Rows[i].Cells[colIndex+1:]...)
	}

	Info(fmt.Sprintf("删除第%d列成功", colIndex))
	return nil
}

// DeleteColumns 删除指定范围的列
func (t *Table) DeleteColumns(startIndex, endIndex int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有行")
	}

	colCount := len(t.Rows[0].Cells)
	if startIndex < 0 || endIndex >= colCount || startIndex > endIndex {
		return fmt.Errorf("列索引范围无效：[%d, %d]，表格共有%d列", startIndex, endIndex, colCount)
	}

	deleteCount := endIndex - startIndex + 1
	if colCount-deleteCount < 1 {
		return fmt.Errorf("删除后表格至少需要保留一列")
	}

	// 删除网格列范围
	t.Grid.Cols = append(t.Grid.Cols[:startIndex], t.Grid.Cols[endIndex+1:]...)

	// 删除每行的对应单元格范围
	for i := range t.Rows {
		t.Rows[i].Cells = append(t.Rows[i].Cells[:startIndex], t.Rows[i].Cells[endIndex+1:]...)
	}

	Info(fmt.Sprintf("删除第%d到%d列成功", startIndex, endIndex))
	return nil
}

// GetCell 获取指定位置的单元格
func (t *Table) GetCell(row, col int) (*TableCell, error) {
	if row < 0 || row >= len(t.Rows) {
		return nil, fmt.Errorf("行索引无效：%d，表格共有%d行", row, len(t.Rows))
	}

	if col < 0 || col >= len(t.Rows[row].Cells) {
		return nil, fmt.Errorf("列索引无效：%d，第%d行共有%d列", col, row, len(t.Rows[row].Cells))
	}

	return &t.Rows[row].Cells[col], nil
}

// SetCellText sets cell text
func (t *Table) SetCellText(row, col int, text string) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有段落和运行
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: text},
					},
				},
			},
		}
	} else {
		if len(cell.Paragraphs[0].Runs) == 0 {
			cell.Paragraphs[0].Runs = []Run{
				{
					Text: Text{Content: text},
				},
			}
		} else {
			cell.Paragraphs[0].Runs[0].Text.Content = text
		}
	}

	return nil
}

// GetCellText 获取单元格文本
func (t *Table) GetCellText(row, col int) (string, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return "", err
	}

	if len(cell.Paragraphs) == 0 {
		return "", nil
	}

	var result string
	for idx, para := range cell.Paragraphs {
		for _, run := range para.Runs {
			result += run.Text.Content
		}
		// 段落之间添加软换行符（除最后一个段落）
		if idx < len(cell.Paragraphs)-1 {
			result += "\n"
		}
	}
	return result, nil
}

// GetRowCount 获取表格行数
func (t *Table) GetRowCount() int {
	return len(t.Rows)
}

// GetColumnCount 获取表格列数
func (t *Table) GetColumnCount() int {
	if len(t.Rows) == 0 {
		return 0
	}
	return len(t.Rows[0].Cells)
}

// ClearTable 清空表格内容（保留结构）
func (t *Table) ClearTable() {
	for i := range t.Rows {
		for j := range t.Rows[i].Cells {
			t.Rows[i].Cells[j].Paragraphs = []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{Content: ""},
						},
					},
				},
			}
		}
	}
	Info("表格内容已清空")
}

// CopyTable 复制表格
func (t *Table) CopyTable() *Table {
	// 深拷贝表格结构
	newTable := &Table{
		Properties: t.Properties,
		Grid:       t.Grid,
		Rows:       make([]TableRow, len(t.Rows)),
	}

	// 复制所有行和单元格
	for i, row := range t.Rows {
		newTable.Rows[i] = TableRow{
			Properties: row.Properties,
			Cells:      make([]TableCell, len(row.Cells)),
		}

		for j, cell := range row.Cells {
			newTable.Rows[i].Cells[j] = TableCell{
				Properties: cell.Properties,
				Paragraphs: make([]Paragraph, len(cell.Paragraphs)),
			}

			// 复制段落内容
			for k, para := range cell.Paragraphs {
				newTable.Rows[i].Cells[j].Paragraphs[k] = Paragraph{
					Properties: para.Properties,
					Runs:       make([]Run, len(para.Runs)),
				}

				for l, run := range para.Runs {
					newTable.Rows[i].Cells[j].Paragraphs[k].Runs[l] = Run{
						Properties: run.Properties,
						Text:       Text{Content: run.Text.Content},
					}
				}
			}
		}
	}

	Info("表格复制成功")
	return newTable
}

// CellAlignment 单元格对齐方式
type CellAlignment string

const (
	// CellAlignLeft 左对齐
	CellAlignLeft CellAlignment = "left"
	// CellAlignCenter 居中对齐
	CellAlignCenter CellAlignment = "center"
	// CellAlignRight 右对齐
	CellAlignRight CellAlignment = "right"
	// CellAlignJustify 两端对齐
	CellAlignJustify CellAlignment = "both"
)

// CellVerticalAlignment 单元格垂直对齐方式
type CellVerticalAlignment string

const (
	// CellVAlignTop 顶部对齐
	CellVAlignTop CellVerticalAlignment = "top"
	// CellVAlignCenter 居中对齐
	CellVAlignCenter CellVerticalAlignment = "center"
	// CellVAlignBottom 底部对齐
	CellVAlignBottom CellVerticalAlignment = "bottom"
)

// CellTextDirection 单元格文字方向
type CellTextDirection string

const (
	// TextDirectionLR 从左到右（默认）
	TextDirectionLR CellTextDirection = "lrTb"
	// TextDirectionTB 从上到下
	TextDirectionTB CellTextDirection = "tbRl"
	// TextDirectionBT 从下到上
	TextDirectionBT CellTextDirection = "btLr"
	// TextDirectionRL 从右到左
	TextDirectionRL CellTextDirection = "rlTb"
	// TextDirectionTBV 从上到下，垂直显示
	TextDirectionTBV CellTextDirection = "tbLrV"
	// TextDirectionBTV 从下到上，垂直显示
	TextDirectionBTV CellTextDirection = "btLrV"
)

// CellFormat 单元格格式配置
type CellFormat struct {
	TextFormat      *TextFormat           // 文字格式
	HorizontalAlign CellAlignment         // 水平对齐
	VerticalAlign   CellVerticalAlignment // 垂直对齐
	TextDirection   CellTextDirection     // 文字方向
	BackgroundColor string                // 背景颜色
	BorderStyle     string                // border style
	Padding         int                   // 内边距（points）
}

// SetCellFormat sets cell format
func (t *Table) SetCellFormat(row, col int, format *CellFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有属性
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	// 设置垂直对齐
	if format.VerticalAlign != "" {
		cell.Properties.VAlign = &VAlign{
			Val: string(format.VerticalAlign),
		}
	}

	// set text direction
	if format.TextDirection != "" {
		cell.Properties.TextDirection = &TextDirection{
			Val: string(format.TextDirection),
		}
	}

	// 确保单元格有段落
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{{}}
	}

	// 设置水平对齐
	if format.HorizontalAlign != "" {
		if cell.Paragraphs[0].Properties == nil {
			cell.Paragraphs[0].Properties = &ParagraphProperties{}
		}
		cell.Paragraphs[0].Properties.Justification = &Justification{
			Val: string(format.HorizontalAlign),
		}
	}

	// 设置文字格式
	if format.TextFormat != nil {
		// 确保有运行
		if len(cell.Paragraphs[0].Runs) == 0 {
			cell.Paragraphs[0].Runs = []Run{{}}
		}

		run := &cell.Paragraphs[0].Runs[0]
		if run.Properties == nil {
			run.Properties = &RunProperties{}
		}

		// 设置粗体
		if format.TextFormat.Bold {
			run.Properties.Bold = &Bold{}
		}

		// 设置斜体
		if format.TextFormat.Italic {
			run.Properties.Italic = &Italic{}
		}

		// 设置字体大小
		if format.TextFormat.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.TextFormat.FontSize*2), // Word使用半points为单位
			}
		}

		// 设置字体颜色
		if format.TextFormat.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.TextFormat.FontColor,
			}
		}

		// 设置字体名称
		if format.TextFormat.FontFamily != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.TextFormat.FontFamily,
			}
		}
	}

	Info(fmt.Sprintf("set cell (%d,%d)format successfully", row, col))
	return nil
}

// SetCellFormattedText sets cell formatted text content
func (t *Table) SetCellFormattedText(row, col int, text string, format *TextFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 创建格式化的运行
	run := Run{
		Text: Text{Content: text},
	}

	if format != nil {
		run.Properties = &RunProperties{}

		if format.FontFamily != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.FontFamily,
			}
		}

		if format.Bold {
			run.Properties.Bold = &Bold{}
		}

		if format.Italic {
			run.Properties.Italic = &Italic{}
		}

		if format.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.FontColor,
			}
		}

		if format.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.FontSize*2),
			}
		}
	}

	// set cell content
	cell.Paragraphs = []Paragraph{
		{
			Runs: []Run{run},
		},
	}

	Info(fmt.Sprintf("set cell (%d,%d)rich text content successfully", row, col))
	return nil
}

// AddCellFormattedText 添加格式化文本到单元格（追加模式）
func (t *Table) AddCellFormattedText(row, col int, text string, format *TextFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有段落
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{{}}
	}

	// 创建格式化的运行
	run := Run{
		Text: Text{Content: text},
	}

	if format != nil {
		run.Properties = &RunProperties{}

		if format.FontFamily != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.FontFamily,
			}
		}

		if format.Bold {
			run.Properties.Bold = &Bold{}
		}

		if format.Italic {
			run.Properties.Italic = &Italic{}
		}

		if format.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.FontColor,
			}
		}

		if format.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.FontSize*2),
			}
		}
	}

	// 添加运行到第一个段落
	cell.Paragraphs[0].Runs = append(cell.Paragraphs[0].Runs, run)

	Info(fmt.Sprintf("添加格式化文本到单元格(%d,%d)成功", row, col))
	return nil
}

// MergeCellsHorizontal 水平合并单元格（合并列）
func (t *Table) MergeCellsHorizontal(row, startCol, endCol int) error {
	if row < 0 || row >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d", row)
	}

	if startCol < 0 || endCol >= len(t.Rows[row].Cells) || startCol > endCol {
		return fmt.Errorf("列索引范围无效：[%d, %d]", startCol, endCol)
	}

	if startCol == endCol {
		return fmt.Errorf("起始列和结束列不能相同")
	}

	// 设置起始单元格的网格跨度
	startCell := &t.Rows[row].Cells[startCol]
	if startCell.Properties == nil {
		startCell.Properties = &TableCellProperties{}
	}

	spanCount := endCol - startCol + 1
	startCell.Properties.GridSpan = &GridSpan{
		Val: fmt.Sprintf("%d", spanCount),
	}

	// 删除被合并的单元格
	newCells := make([]TableCell, 0, len(t.Rows[row].Cells)-(endCol-startCol))
	newCells = append(newCells, t.Rows[row].Cells[:startCol+1]...)
	if endCol+1 < len(t.Rows[row].Cells) {
		newCells = append(newCells, t.Rows[row].Cells[endCol+1:]...)
	}
	t.Rows[row].Cells = newCells

	Info(fmt.Sprintf("水平合并单元格：行%d，列%d到%d", row, startCol, endCol))
	return nil
}

// MergeCellsVertical 垂直合并单元格（合并行）
func (t *Table) MergeCellsVertical(startRow, endRow, col int) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("行索引范围无效：[%d, %d]", startRow, endRow)
	}

	if col < 0 {
		return fmt.Errorf("列索引无效：%d", col)
	}

	if startRow == endRow {
		return fmt.Errorf("起始行和结束行不能相同")
	}

	// 检查所有行的列数
	for i := startRow; i <= endRow; i++ {
		if col >= len(t.Rows[i].Cells) {
			return fmt.Errorf("第%d行没有第%d列", i, col)
		}
	}

	// 设置起始单元格为合并起始
	startCell := &t.Rows[startRow].Cells[col]
	if startCell.Properties == nil {
		startCell.Properties = &TableCellProperties{}
	}
	startCell.Properties.VMerge = &VMerge{
		Val: "restart",
	}

	// 设置后续单元格为合并继续
	for i := startRow + 1; i <= endRow; i++ {
		cell := &t.Rows[i].Cells[col]
		if cell.Properties == nil {
			cell.Properties = &TableCellProperties{}
		}
		cell.Properties.VMerge = &VMerge{
			Val: "continue",
		}
		// 清空被合并单元格的内容
		cell.Paragraphs = []Paragraph{{}}
	}

	Info(fmt.Sprintf("垂直合并单元格：行%d到%d，列%d", startRow, endRow, col))
	return nil
}

// MergeCellsRange 合并单元格区域（多行多列）
func (t *Table) MergeCellsRange(startRow, endRow, startCol, endCol int) error {
	// 验证范围
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("行索引范围无效：[%d, %d]", startRow, endRow)
	}

	// 先水平合并每一行
	for i := startRow; i <= endRow; i++ {
		if startCol >= len(t.Rows[i].Cells) || endCol >= len(t.Rows[i].Cells) {
			return fmt.Errorf("第%d行列索引范围无效：[%d, %d]", i, startCol, endCol)
		}

		if startCol != endCol {
			err := t.MergeCellsHorizontal(i, startCol, endCol)
			if err != nil {
				return fmt.Errorf("水平合并第%d行失败：%v", i, err)
			}
		}
	}

	// 然后垂直合并第一列
	if startRow != endRow {
		err := t.MergeCellsVertical(startRow, endRow, startCol)
		if err != nil {
			return fmt.Errorf("垂直合并失败：%v", err)
		}
	}

	Info(fmt.Sprintf("合并单元格区域：行%d到%d，列%d到%d", startRow, endRow, startCol, endCol))
	return nil
}

// UnmergeCells 取消单元格合并
func (t *Table) UnmergeCells(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		return fmt.Errorf("单元格没有合并")
	}

	// 检查是否有水平合并
	if cell.Properties.GridSpan != nil {
		// 获取合并的列数
		spanCount := 1
		if cell.Properties.GridSpan.Val != "" {
			fmt.Sscanf(cell.Properties.GridSpan.Val, "%d", &spanCount)
		}

		// 插入被合并的单元格
		for i := 1; i < spanCount; i++ {
			newCell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: cell.Properties.TableCellW,
					VAlign:     cell.Properties.VAlign,
				},
				Paragraphs: []Paragraph{{}},
			}

			// 在指定位置插入新单元格
			insertPos := col + i
			if insertPos <= len(t.Rows[row].Cells) {
				t.Rows[row].Cells = append(t.Rows[row].Cells[:insertPos], append([]TableCell{newCell}, t.Rows[row].Cells[insertPos:]...)...)
			}
		}

		// 移除水平合并属性
		cell.Properties.GridSpan = nil
	}

	// 检查是否有垂直合并
	if cell.Properties.VMerge != nil {
		// 移除垂直合并属性
		cell.Properties.VMerge = nil

		// 查找并恢复被合并的单元格
		for i := row + 1; i < len(t.Rows); i++ {
			if col < len(t.Rows[i].Cells) {
				otherCell := &t.Rows[i].Cells[col]
				if otherCell.Properties != nil && otherCell.Properties.VMerge != nil {
					if otherCell.Properties.VMerge.Val == "continue" {
						// 恢复单元格内容
						otherCell.Properties.VMerge = nil
						if len(otherCell.Paragraphs) == 0 {
							otherCell.Paragraphs = []Paragraph{{}}
						}
					} else {
						break
					}
				} else {
					break
				}
			}
		}
	}

	Info(fmt.Sprintf("取消单元格(%d,%d)合并成功", row, col))
	return nil
}

// IsCellMerged 检查单元格是否被合并
func (t *Table) IsCellMerged(row, col int) (bool, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return false, err
	}

	if cell.Properties == nil {
		return false, nil
	}

	// 检查水平合并
	if cell.Properties.GridSpan != nil && cell.Properties.GridSpan.Val != "" && cell.Properties.GridSpan.Val != "1" {
		return true, nil
	}

	// 检查垂直合并
	if cell.Properties.VMerge != nil {
		return true, nil
	}

	return false, nil
}

// GetMergedCellInfo 获取合并单元格信息
func (t *Table) GetMergedCellInfo(row, col int) (map[string]interface{}, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	info := make(map[string]interface{})
	info["is_merged"] = false

	if cell.Properties == nil {
		return info, nil
	}

	// 检查水平合并
	if cell.Properties.GridSpan != nil && cell.Properties.GridSpan.Val != "" {
		spanCount := 1
		fmt.Sscanf(cell.Properties.GridSpan.Val, "%d", &spanCount)
		if spanCount > 1 {
			info["is_merged"] = true
			info["horizontal_span"] = spanCount
			info["merge_type"] = "horizontal"
		}
	}

	// 检查垂直合并
	if cell.Properties.VMerge != nil {
		info["is_merged"] = true
		if cell.Properties.VMerge.Val == "restart" {
			info["vertical_merge_start"] = true
			info["merge_type"] = "vertical"
		} else if cell.Properties.VMerge.Val == "continue" {
			info["vertical_merge_continue"] = true
			info["merge_type"] = "vertical"
		}
	}

	return info, nil
}

// ClearCellContent 清空单元格内容但保留格式
func (t *Table) ClearCellContent(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 保留格式，只清空文本内容
	for i := range cell.Paragraphs {
		for j := range cell.Paragraphs[i].Runs {
			cell.Paragraphs[i].Runs[j].Text.Content = ""
		}
	}

	Info(fmt.Sprintf("清空单元格(%d,%d)内容成功", row, col))
	return nil
}

// ClearCellFormat 清空单元格格式但保留内容
func (t *Table) ClearCellFormat(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 清除单元格属性中的格式
	if cell.Properties != nil {
		// 保留合并信息和基本宽度，清除其他格式
		oldGridSpan := cell.Properties.GridSpan
		oldVMerge := cell.Properties.VMerge
		oldWidth := cell.Properties.TableCellW

		cell.Properties = &TableCellProperties{
			TableCellW: oldWidth,
			GridSpan:   oldGridSpan,
			VMerge:     oldVMerge,
		}
	}

	// 清除段落和运行的格式
	for i := range cell.Paragraphs {
		cell.Paragraphs[i].Properties = nil
		for j := range cell.Paragraphs[i].Runs {
			cell.Paragraphs[i].Runs[j].Properties = nil
		}
	}

	Info(fmt.Sprintf("清空单元格(%d,%d)format successfully", row, col))
	return nil
}

// SetCellPadding sets cell padding
func (t *Table) SetCellPadding(row, col int, padding int) error {
	_, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 单元格内边距通过表格属性设置，这里先预留接口
	// 实际实现需要在表格级别设置默认内边距
	Info(fmt.Sprintf("set cell (%d,%d)padding is %dpoints", row, col, padding))
	return nil
}

// SetCellTextDirection sets cell text direction
func (t *Table) SetCellTextDirection(row, col int, direction CellTextDirection) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有属性
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	// set text direction
	cell.Properties.TextDirection = &TextDirection{
		Val: string(direction),
	}

	Info(fmt.Sprintf("set cell (%d,%d)text direction is %s", row, col, direction))
	return nil
}

// GetCellTextDirection 获取单元格文字方向
func (t *Table) GetCellTextDirection(row, col int) (CellTextDirection, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return TextDirectionLR, err
	}

	if cell.Properties != nil && cell.Properties.TextDirection != nil {
		return CellTextDirection(cell.Properties.TextDirection.Val), nil
	}

	// 默认返回从左到右
	return TextDirectionLR, nil
}

// GetCellFormat 获取单元格格式信息
func (t *Table) GetCellFormat(row, col int) (*CellFormat, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	format := &CellFormat{}

	// 获取垂直对齐
	if cell.Properties != nil && cell.Properties.VAlign != nil {
		format.VerticalAlign = CellVerticalAlignment(cell.Properties.VAlign.Val)
	}

	// 获取文字方向
	if cell.Properties != nil && cell.Properties.TextDirection != nil {
		format.TextDirection = CellTextDirection(cell.Properties.TextDirection.Val)
	}

	// 获取水平对齐
	if len(cell.Paragraphs) > 0 && cell.Paragraphs[0].Properties != nil && cell.Paragraphs[0].Properties.Justification != nil {
		format.HorizontalAlign = CellAlignment(cell.Paragraphs[0].Properties.Justification.Val)
	}

	// 获取文字格式
	if len(cell.Paragraphs) > 0 && len(cell.Paragraphs[0].Runs) > 0 {
		run := &cell.Paragraphs[0].Runs[0]
		if run.Properties != nil {
			format.TextFormat = &TextFormat{}

			if run.Properties.Bold != nil {
				format.TextFormat.Bold = true
			}

			if run.Properties.Italic != nil {
				format.TextFormat.Italic = true
			}

			if run.Properties.FontSize != nil {
				fmt.Sscanf(run.Properties.FontSize.Val, "%d", &format.TextFormat.FontSize)
				format.TextFormat.FontSize /= 2 // 转换为points
			}

			if run.Properties.Color != nil {
				format.TextFormat.FontColor = run.Properties.Color.Val
			}

			if run.Properties.FontFamily != nil {
				format.TextFormat.FontFamily = run.Properties.FontFamily.ASCII
			}
		}
	}

	return format, nil
}

// TextDirection 文字方向
type TextDirection struct {
	XMLName xml.Name `xml:"w:textDirection"`
	Val     string   `xml:"w:val,attr"`
}

// RowHeightRule 行高规则
type RowHeightRule string

const (
	// RowHeightAuto 自动调整行高
	RowHeightAuto RowHeightRule = "auto"
	// RowHeightMinimum 最小行高
	RowHeightMinimum RowHeightRule = "atLeast"
	// RowHeightExact 固定行高
	RowHeightExact RowHeightRule = "exact"
)

// RowHeightConfig 行高配置
type RowHeightConfig struct {
	Height int           // 行高值（points，1points=20twips）
	Rule   RowHeightRule // 行高规则
}

// SetRowHeight 设置行高
func (t *Table) SetRowHeight(rowIndex int, config *RowHeightConfig) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil {
		row.Properties = &TableRowProperties{}
	}

	// 设置行高属性
	row.Properties.TableRowH = &TableRowH{
		Val:   fmt.Sprintf("%d", config.Height*20), // 转换为twips (1points=20twips)
		HRule: string(config.Rule),
	}

	Info(fmt.Sprintf("设置第%d行高度为%dpoints，规则为%s", rowIndex, config.Height, config.Rule))
	return nil
}

// GetRowHeight 获取行高配置
func (t *Table) GetRowHeight(rowIndex int) (*RowHeightConfig, error) {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return nil, fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil || row.Properties.TableRowH == nil {
		// 返回默认自动高度
		return &RowHeightConfig{
			Height: 0,
			Rule:   RowHeightAuto,
		}, nil
	}

	height := 0
	if row.Properties.TableRowH.Val != "" {
		fmt.Sscanf(row.Properties.TableRowH.Val, "%d", &height)
		height /= 20 // 转换为points
	}

	rule := RowHeightAuto
	if row.Properties.TableRowH.HRule != "" {
		rule = RowHeightRule(row.Properties.TableRowH.HRule)
	}

	return &RowHeightConfig{
		Height: height,
		Rule:   rule,
	}, nil
}

// SetRowHeightRange 批量设置行高
func (t *Table) SetRowHeightRange(startRow, endRow int, config *RowHeightConfig) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("行索引范围无效：[%d, %d]，表格共有%d行", startRow, endRow, len(t.Rows))
	}

	for i := startRow; i <= endRow; i++ {
		err := t.SetRowHeight(i, config)
		if err != nil {
			return fmt.Errorf("设置第%d行高度失败：%v", i, err)
		}
	}

	Info(fmt.Sprintf("批量设置第%d到%d行高度成功", startRow, endRow))
	return nil
}

// TableTextWrap 表格文字环绕类型
type TableTextWrap string

const (
	// TextWrapNone 无环绕（默认）
	TextWrapNone TableTextWrap = "none"
	// TextWrapAround 环绕表格
	TextWrapAround TableTextWrap = "around"
)

// TablePosition 表格定位类型
type TablePosition string

const (
	// PositionInline 行内定位（默认）
	PositionInline TablePosition = "inline"
	// PositionFloating 浮动定位
	PositionFloating TablePosition = "floating"
)

// TableAlignment 表格对齐类型
type TableAlignment string

const (
	// TableAlignLeft 左对齐
	TableAlignLeft TableAlignment = "left"
	// TableAlignCenter 居中对齐
	TableAlignCenter TableAlignment = "center"
	// TableAlignRight 右对齐
	TableAlignRight TableAlignment = "right"
	// TableAlignInside 内侧对齐
	TableAlignInside TableAlignment = "inside"
	// TableAlignOutside 外侧对齐
	TableAlignOutside TableAlignment = "outside"
)

// TablePositioning 表格定位配置
type TablePositioning struct {
	XMLName        xml.Name `xml:"w:tblpPr"`
	LeftFromText   string   `xml:"w:leftFromText,attr,omitempty"`   // 距离左侧文字的距离
	RightFromText  string   `xml:"w:rightFromText,attr,omitempty"`  // 距离右侧文字的距离
	TopFromText    string   `xml:"w:topFromText,attr,omitempty"`    // 距离上方文字的距离
	BottomFromText string   `xml:"w:bottomFromText,attr,omitempty"` // 距离下方文字的距离
	VertAnchor     string   `xml:"w:vertAnchor,attr,omitempty"`     // 垂直锚点
	HorzAnchor     string   `xml:"w:horzAnchor,attr,omitempty"`     // 水平锚点
	TblpXSpec      string   `xml:"w:tblpXSpec,attr,omitempty"`      // 水平对齐规格
	TblpYSpec      string   `xml:"w:tblpYSpec,attr,omitempty"`      // 垂直对齐规格
	TblpX          string   `xml:"w:tblpX,attr,omitempty"`          // X坐标
	TblpY          string   `xml:"w:tblpY,attr,omitempty"`          // Y坐标
}

// TableLayoutConfig 表格布局配置
type TableLayoutConfig struct {
	Alignment   TableAlignment    // 表格对齐方式
	TextWrap    TableTextWrap     // 文字环绕类型
	Position    TablePosition     // 定位类型
	Positioning *TablePositioning // 定位详细配置（仅在Position为Floating时有效）
}

// SetTableLayout 设置表格布局和定位
func (t *Table) SetTableLayout(config *TableLayoutConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	// 设置表格对齐
	if config.Alignment != "" {
		t.Properties.TableJc = &TableJc{
			Val: string(config.Alignment),
		}
	}

	// 设置定位属性（仅在浮动定位时生效）
	if config.Position == PositionFloating && config.Positioning != nil {
		// 在OOXML中，浮动表格定位需要特殊的TablePositioning属性
		// 这里将配置信息存储到表格属性中
		Info("设置表格为浮动定位模式")
		// 注意：完整的浮动定位实现需要更复杂的XML结构支持
	}

	Info(fmt.Sprintf("设置表格布局：对齐=%s，环绕=%s，定位=%s",
		config.Alignment, config.TextWrap, config.Position))
	return nil
}

// GetTableLayout 获取表格布局配置
func (t *Table) GetTableLayout() *TableLayoutConfig {
	config := &TableLayoutConfig{
		Alignment: TableAlignLeft, // 默认值
		TextWrap:  TextWrapNone,
		Position:  PositionInline,
	}

	if t.Properties != nil && t.Properties.TableJc != nil {
		config.Alignment = TableAlignment(t.Properties.TableJc.Val)
	}

	return config
}

// SetTableAlignment 设置表格对齐方式（快捷方法）
func (t *Table) SetTableAlignment(alignment TableAlignment) error {
	return t.SetTableLayout(&TableLayoutConfig{
		Alignment: alignment,
		TextWrap:  TextWrapNone,
		Position:  PositionInline,
	})
}

// TableBreakRule 表格分页规则
type TableBreakRule string

const (
	// BreakAuto 自动分页（默认）
	BreakAuto TableBreakRule = "auto"
	// BreakPage 强制分页
	BreakPage TableBreakRule = "page"
	// BreakColumn 强制分栏
	BreakColumn TableBreakRule = "column"
)

// RowBreakConfig 行分页配置
type RowBreakConfig struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	CantSplit *CantSplit `xml:"w:cantSplit,omitempty"` // prevent page break
	TrHeight  *TableRowH `xml:"w:trHeight,omitempty"`  // 行高
	TblHeader *TblHeader `xml:"w:tblHeader,omitempty"` // 标题行重复
}

// CantSplit 禁止分割
type CantSplit struct {
	XMLName xml.Name `xml:"w:cantSplit"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// TblHeader 表格标题行
type TblHeader struct {
	XMLName xml.Name `xml:"w:tblHeader"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// SetRowKeepTogether sets row to prevent page break
func (t *Table) SetRowKeepTogether(rowIndex int, keepTogether bool) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil {
		row.Properties = &TableRowProperties{}
	}

	if keepTogether {
		row.Properties.CantSplit = &CantSplit{
			Val: "1",
		}
	} else {
		row.Properties.CantSplit = nil
	}

	Info(fmt.Sprintf("设置第%d行跨页分割为：%t", rowIndex, !keepTogether))
	return nil
}

// SetRowAsHeader 设置行为重复的标题行
func (t *Table) SetRowAsHeader(rowIndex int, isHeader bool) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil {
		row.Properties = &TableRowProperties{}
	}

	if isHeader {
		row.Properties.TblHeader = &TblHeader{
			Val: "1",
		}
	} else {
		row.Properties.TblHeader = nil
	}

	Info(fmt.Sprintf("设置第%d行为标题行：%t", rowIndex, isHeader))
	return nil
}

// SetHeaderRows 设置表格标题行范围
func (t *Table) SetHeaderRows(startRow, endRow int) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("行索引范围无效：[%d, %d]，表格共有%d行", startRow, endRow, len(t.Rows))
	}

	// 清除所有现有的标题行设置
	for i := range t.Rows {
		if t.Rows[i].Properties != nil {
			t.Rows[i].Properties.TblHeader = nil
		}
	}

	// 设置指定范围为标题行
	for i := startRow; i <= endRow; i++ {
		err := t.SetRowAsHeader(i, true)
		if err != nil {
			return fmt.Errorf("设置第%d行为标题行失败：%v", i, err)
		}
	}

	Info(fmt.Sprintf("设置第%d到%d行为标题行", startRow, endRow))
	return nil
}

// IsRowHeader 检查行是否为标题行
func (t *Table) IsRowHeader(rowIndex int) (bool, error) {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return false, fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties != nil && row.Properties.TblHeader != nil {
		return row.Properties.TblHeader.Val == "1", nil
	}

	return false, nil
}

// IsRowKeepTogether checks if row prevents page break
func (t *Table) IsRowKeepTogether(rowIndex int) (bool, error) {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return false, fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties != nil && row.Properties.CantSplit != nil {
		return row.Properties.CantSplit.Val == "1", nil
	}

	return false, nil
}

// TablePageBreakConfig 表格分页配置
type TablePageBreakConfig struct {
	KeepWithNext    bool // 与下一段落保持在一起
	KeepLines       bool // 保持行在一起
	PageBreakBefore bool // 段落前分页
	WidowControl    bool // 孤行控制
}

// SetTablePageBreak 设置表格分页控制
func (t *Table) SetTablePageBreak(config *TablePageBreakConfig) error {
	// 表格级别的分页控制通常在表格属性中设置
	// 这里先记录配置，实际XML输出时需要相应的实现
	Info(fmt.Sprintf("设置表格分页控制：保持与下一段落=%t，保持行=%t，段前分页=%t，孤行控制=%t",
		config.KeepWithNext, config.KeepLines, config.PageBreakBefore, config.WidowControl))
	return nil
}

// SetRowKeepWithNext 设置行与下一行保持在同一页
func (t *Table) SetRowKeepWithNext(rowIndex int, keepWithNext bool) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	// 这个功能需要在行属性中设置特定的分页属性
	// 实际实现时需要扩展TableRowProperties结构
	Info(fmt.Sprintf("设置第%d行与下一行保持在同一页：%t", rowIndex, keepWithNext))
	return nil
}

// GetTableBreakInfo 获取表格分页信息
func (t *Table) GetTableBreakInfo() map[string]interface{} {
	info := make(map[string]interface{})

	headerRowCount := 0
	keepTogetherCount := 0

	for i := range t.Rows {
		isHeader, _ := t.IsRowHeader(i)
		if isHeader {
			headerRowCount++
		}

		keepTogether, _ := t.IsRowKeepTogether(i)
		if keepTogether {
			keepTogetherCount++
		}
	}

	info["total_rows"] = len(t.Rows)
	info["header_rows"] = headerRowCount
	info["keep_together_rows"] = keepTogetherCount

	return info
}

// 扩展TableRowProperties以支持分页控制
type TableRowPropertiesExtended struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	TableRowH *TableRowH `xml:"w:trHeight,omitempty"`
	CantSplit *CantSplit `xml:"w:cantSplit,omitempty"`
	TblHeader *TblHeader `xml:"w:tblHeader,omitempty"`
	KeepNext  *KeepNext  `xml:"w:keepNext,omitempty"`
	KeepLines *KeepLines `xml:"w:keepLines,omitempty"`
}

// 扩展现有的TableRowProperties结构
func (trp *TableRowProperties) SetCantSplit(cantSplit bool) {
	if cantSplit {
		trp.CantSplit = &CantSplit{Val: "1"}
	} else {
		trp.CantSplit = nil
	}
}

func (trp *TableRowProperties) SetTblHeader(isHeader bool) {
	if isHeader {
		trp.TblHeader = &TblHeader{Val: "1"}
	} else {
		trp.TblHeader = nil
	}
}

// TableStyle 表格样式引用
type TableStyle struct {
	XMLName xml.Name `xml:"w:tblStyle"`
	Val     string   `xml:"w:val,attr"`
}

// TableBorders 表格边框
type TableBorders struct {
	XMLName xml.Name     `xml:"w:tblBorders"`
	Top     *TableBorder `xml:"w:top,omitempty"`     // 上边框
	Left    *TableBorder `xml:"w:left,omitempty"`    // 左边框
	Bottom  *TableBorder `xml:"w:bottom,omitempty"`  // 下边框
	Right   *TableBorder `xml:"w:right,omitempty"`   // 右边框
	InsideH *TableBorder `xml:"w:insideH,omitempty"` // 内部水平边框
	InsideV *TableBorder `xml:"w:insideV,omitempty"` // 内部垂直边框
}

// TableBorder 边框定义
type TableBorder struct {
	Val        string `xml:"w:val,attr"`                  // border style
	Sz         string `xml:"w:sz,attr"`                   // 边框粗细（1/8points）
	Space      string `xml:"w:space,attr"`                // 边框间距
	Color      string `xml:"w:color,attr"`                // 边框颜色
	ThemeColor string `xml:"w:themeColor,attr,omitempty"` // theme color
}

// TableShading 表格底纹/背景
type TableShading struct {
	XMLName   xml.Name `xml:"w:shd"`
	Val       string   `xml:"w:val,attr"`                 // shading style
	Color     string   `xml:"w:color,attr,omitempty"`     // 前景色
	Fill      string   `xml:"w:fill,attr,omitempty"`      // 背景色
	ThemeFill string   `xml:"w:themeFill,attr,omitempty"` // theme fill color
}

// TableCellMargins 表格单元格边距
type TableCellMargins struct {
	XMLName xml.Name        `xml:"w:tblCellMar"`
	Top     *TableCellSpace `xml:"w:top,omitempty"`
	Left    *TableCellSpace `xml:"w:left,omitempty"`
	Bottom  *TableCellSpace `xml:"w:bottom,omitempty"`
	Right   *TableCellSpace `xml:"w:right,omitempty"`
}

// TableCellSpace 表格单元格空间
type TableCellSpace struct {
	W    string `xml:"w:w,attr"`
	Type string `xml:"w:type,attr"`
}

// TableLayoutType 表格布局类型
type TableLayoutType struct {
	XMLName xml.Name `xml:"w:tblLayout"`
	Type    string   `xml:"w:type,attr"` // fixed, autofit
}

// TableIndentation 表格缩进
type TableIndentation struct {
	XMLName xml.Name `xml:"w:tblInd"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// TableCellShading 单元格背景
type TableCellShading struct {
	XMLName   xml.Name `xml:"w:shd"`
	Val       string   `xml:"w:val,attr"`                 // shading style
	Color     string   `xml:"w:color,attr,omitempty"`     // 前景色
	Fill      string   `xml:"w:fill,attr,omitempty"`      // 背景色
	ThemeFill string   `xml:"w:themeFill,attr,omitempty"` // theme fill color
}

// TableCellBorders 单元格边框
type TableCellBorders struct {
	XMLName xml.Name         `xml:"w:tcBorders"`
	Top     *TableCellBorder `xml:"w:top,omitempty"`     // 上边框
	Left    *TableCellBorder `xml:"w:left,omitempty"`    // 左边框
	Bottom  *TableCellBorder `xml:"w:bottom,omitempty"`  // 下边框
	Right   *TableCellBorder `xml:"w:right,omitempty"`   // 右边框
	InsideH *TableCellBorder `xml:"w:insideH,omitempty"` // 内部水平边框
	InsideV *TableCellBorder `xml:"w:insideV,omitempty"` // 内部垂直边框
	TL2BR   *TableCellBorder `xml:"w:tl2br,omitempty"`   // 左上到右下对角线
	TR2BL   *TableCellBorder `xml:"w:tr2bl,omitempty"`   // 右上到左下对角线
}

// TableCellBorder 单元格边框定义
type TableCellBorder struct {
	Val        string `xml:"w:val,attr"`                  // border style
	Sz         string `xml:"w:sz,attr"`                   // 边框粗细（1/8points）
	Space      string `xml:"w:space,attr"`                // 边框间距
	Color      string `xml:"w:color,attr"`                // 边框颜色
	ThemeColor string `xml:"w:themeColor,attr,omitempty"` // theme color
}

// NoWrap 禁止换行
type NoWrap struct {
	XMLName xml.Name `xml:"w:noWrap"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// HideMark 隐藏标记
type HideMark struct {
	XMLName xml.Name `xml:"w:hideMark"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// ============== 表格样式和外观功能 ==============

// BorderStyle 边框样式常量
type BorderStyle string

const (
	BorderStyleNone                   BorderStyle = "none"                   // 无边框
	BorderStyleSingle                 BorderStyle = "single"                 // 单线
	BorderStyleThick                  BorderStyle = "thick"                  // 粗线
	BorderStyleDouble                 BorderStyle = "double"                 // 双线
	BorderStyleDotted                 BorderStyle = "dotted"                 // 点线
	BorderStyleDashed                 BorderStyle = "dashed"                 // 虚线
	BorderStyleDotDash                BorderStyle = "dotDash"                // 点划线
	BorderStyleDotDotDash             BorderStyle = "dotDotDash"             // 双点划线
	BorderStyleTriple                 BorderStyle = "triple"                 // 三线
	BorderStyleThinThickSmallGap      BorderStyle = "thinThickSmallGap"      // 细粗细线（小间距）
	BorderStyleThickThinSmallGap      BorderStyle = "thickThinSmallGap"      // 粗细粗线（小间距）
	BorderStyleThinThickThinSmallGap  BorderStyle = "thinThickThinSmallGap"  // 细粗细线（小间距）
	BorderStyleThinThickMediumGap     BorderStyle = "thinThickMediumGap"     // 细粗细线（中间距）
	BorderStyleThickThinMediumGap     BorderStyle = "thickThinMediumGap"     // 粗细粗线（中间距）
	BorderStyleThinThickThinMediumGap BorderStyle = "thinThickThinMediumGap" // 细粗细线（中间距）
	BorderStyleThinThickLargeGap      BorderStyle = "thinThickLargeGap"      // 细粗细线（大间距）
	BorderStyleThickThinLargeGap      BorderStyle = "thickThinLargeGap"      // 粗细粗线（大间距）
	BorderStyleThinThickThinLargeGap  BorderStyle = "thinThickThinLargeGap"  // 细粗细线（大间距）
	BorderStyleWave                   BorderStyle = "wave"                   // 波浪线
	BorderStyleDoubleWave             BorderStyle = "doubleWave"             // 双波浪线
	BorderStyleDashSmallGap           BorderStyle = "dashSmallGap"           // 虚线（小间距）
	BorderStyleDashDotStroked         BorderStyle = "dashDotStroked"         // 划点线
	BorderStyleThreeDEmboss           BorderStyle = "threeDEmboss"           // 3D浮雕
	BorderStyleThreeDEngrave          BorderStyle = "threeDEngrave"          // 3D雕刻
	BorderStyleOutset                 BorderStyle = "outset"                 // 外凸
	BorderStyleInset                  BorderStyle = "inset"                  // 内凹
)

// ShadingPattern 底纹图案类型
type ShadingPattern string

const (
	ShadingPatternClear             ShadingPattern = "clear"             // 透明/实色
	ShadingPatternSolid             ShadingPattern = "clear"             // 实色（使用clear实现）
	ShadingPatternPct5              ShadingPattern = "pct5"              // 5%
	ShadingPatternPct10             ShadingPattern = "pct10"             // 10%
	ShadingPatternPct20             ShadingPattern = "pct20"             // 20%
	ShadingPatternPct25             ShadingPattern = "pct25"             // 25%
	ShadingPatternPct30             ShadingPattern = "pct30"             // 30%
	ShadingPatternPct40             ShadingPattern = "pct40"             // 40%
	ShadingPatternPct50             ShadingPattern = "pct50"             // 50%
	ShadingPatternPct60             ShadingPattern = "pct60"             // 60%
	ShadingPatternPct70             ShadingPattern = "pct70"             // 70%
	ShadingPatternPct75             ShadingPattern = "pct75"             // 75%
	ShadingPatternPct80             ShadingPattern = "pct80"             // 80%
	ShadingPatternPct90             ShadingPattern = "pct90"             // 90%
	ShadingPatternHorzStripe        ShadingPattern = "horzStripe"        // 水平条纹
	ShadingPatternVertStripe        ShadingPattern = "vertStripe"        // 垂直条纹
	ShadingPatternReverseDiagStripe ShadingPattern = "reverseDiagStripe" // 反对角条纹
	ShadingPatternDiagStripe        ShadingPattern = "diagStripe"        // 对角条纹
	ShadingPatternHorzCross         ShadingPattern = "horzCross"         // 水平十字
	ShadingPatternDiagCross         ShadingPattern = "diagCross"         // 对角十字
)

// TableStyleTemplate 表格样式模板
type TableStyleTemplate string

const (
	TableStyleTemplateNormal    TableStyleTemplate = "TableNormal"    // 普通表格
	TableStyleTemplateGrid      TableStyleTemplate = "TableGrid"      // 网格表格
	TableStyleTemplateList      TableStyleTemplate = "TableList"      // 列表表格
	TableStyleTemplateColorful1 TableStyleTemplate = "TableColorful1" // 彩色表格1
	TableStyleTemplateColorful2 TableStyleTemplate = "TableColorful2" // 彩色表格2
	TableStyleTemplateColorful3 TableStyleTemplate = "TableColorful3" // 彩色表格3
	TableStyleTemplateColumns1  TableStyleTemplate = "TableColumns1"  // 列样式1
	TableStyleTemplateColumns2  TableStyleTemplate = "TableColumns2"  // 列样式2
	TableStyleTemplateColumns3  TableStyleTemplate = "TableColumns3"  // 列样式3
	TableStyleTemplateRows1     TableStyleTemplate = "TableRows1"     // 行样式1
	TableStyleTemplateRows2     TableStyleTemplate = "TableRows2"     // 行样式2
	TableStyleTemplateRows3     TableStyleTemplate = "TableRows3"     // 行样式3
	TableStyleTemplatePlain1    TableStyleTemplate = "TablePlain1"    // 简洁表格1
	TableStyleTemplatePlain2    TableStyleTemplate = "TablePlain2"    // 简洁表格2
	TableStyleTemplatePlain3    TableStyleTemplate = "TablePlain3"    // 简洁表格3
)

// TableStyleConfig 表格样式配置
type TableStyleConfig struct {
	Template          TableStyleTemplate // 样式模板
	StyleID           string             // 自定义样式ID
	FirstRowHeader    bool               // 首行作为标题
	LastRowTotal      bool               // 最后一行作为总计
	FirstColumnHeader bool               // 首列作为标题
	LastColumnTotal   bool               // 最后一列作为总计
	BandedRows        bool               // 交替行颜色
	BandedColumns     bool               // 交替列颜色
}

// BorderConfig 边框配置
type BorderConfig struct {
	Style BorderStyle // border style
	Width int         // 边框宽度（1/8points）
	Color string      // 边框颜色（十六进制，如 "FF0000"）
	Space int         // 边框间距
}

// ShadingConfig 底纹配置
type ShadingConfig struct {
	Pattern         ShadingPattern // 底纹图案
	ForegroundColor string         // 前景色（十六进制）
	BackgroundColor string         // 背景色（十六进制）
}

// TableBorderConfig 表格边框配置
type TableBorderConfig struct {
	Top     *BorderConfig // 上边框
	Left    *BorderConfig // 左边框
	Bottom  *BorderConfig // 下边框
	Right   *BorderConfig // 右边框
	InsideH *BorderConfig // 内部水平边框
	InsideV *BorderConfig // 内部垂直边框
}

// CellBorderConfig 单元格边框配置
type CellBorderConfig struct {
	Top      *BorderConfig // 上边框
	Left     *BorderConfig // 左边框
	Bottom   *BorderConfig // 下边框
	Right    *BorderConfig // 右边框
	DiagDown *BorderConfig // 左上到右下对角线
	DiagUp   *BorderConfig // 右上到左下对角线
}

// ApplyTableStyle 应用表格样式
func (t *Table) ApplyTableStyle(config *TableStyleConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	// 设置样式模板
	if config.Template != "" {
		t.Properties.TableStyle = &TableStyle{
			Val: string(config.Template),
		}
	} else if config.StyleID != "" {
		t.Properties.TableStyle = &TableStyle{
			Val: config.StyleID,
		}
	}

	// 设置表格外观选项
	if t.Properties.TableLook == nil {
		t.Properties.TableLook = &TableLook{}
	}

	// 构建TableLook值
	lookVal := "0000"
	if config.FirstRowHeader {
		t.Properties.TableLook.FirstRow = "1"
		lookVal = "0400"
	} else {
		t.Properties.TableLook.FirstRow = "0"
	}

	if config.LastRowTotal {
		t.Properties.TableLook.LastRow = "1"
		if lookVal == "0400" {
			lookVal = "0440"
		} else {
			lookVal = "0040"
		}
	} else {
		t.Properties.TableLook.LastRow = "0"
	}

	if config.FirstColumnHeader {
		t.Properties.TableLook.FirstCol = "1"
		switch lookVal {
		case "0400":
			lookVal = "0500"
		case "0040":
			lookVal = "0140"
		case "0440":
			lookVal = "0540"
		default:
			lookVal = "0100"
		}
	} else {
		t.Properties.TableLook.FirstCol = "0"
	}

	if config.LastColumnTotal {
		t.Properties.TableLook.LastCol = "1"
	} else {
		t.Properties.TableLook.LastCol = "0"
	}

	if config.BandedRows {
		t.Properties.TableLook.NoHBand = "0"
	} else {
		t.Properties.TableLook.NoHBand = "1"
	}

	if config.BandedColumns {
		t.Properties.TableLook.NoVBand = "0"
	} else {
		t.Properties.TableLook.NoVBand = "1"
	}

	t.Properties.TableLook.Val = lookVal

	Info(fmt.Sprintf("应用表格样式成功：%s", config.Template))
	return nil
}

// SetTableBorders 设置表格边框
func (t *Table) SetTableBorders(config *TableBorderConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	t.Properties.TableBorders = &TableBorders{}

	if config.Top != nil {
		t.Properties.TableBorders.Top = createTableBorder(config.Top)
	}
	if config.Left != nil {
		t.Properties.TableBorders.Left = createTableBorder(config.Left)
	}
	if config.Bottom != nil {
		t.Properties.TableBorders.Bottom = createTableBorder(config.Bottom)
	}
	if config.Right != nil {
		t.Properties.TableBorders.Right = createTableBorder(config.Right)
	}
	if config.InsideH != nil {
		t.Properties.TableBorders.InsideH = createTableBorder(config.InsideH)
	}
	if config.InsideV != nil {
		t.Properties.TableBorders.InsideV = createTableBorder(config.InsideV)
	}

	Info("设置表格边框成功")
	return nil
}

// SetTableShading 设置表格背景
func (t *Table) SetTableShading(config *ShadingConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	t.Properties.Shd = &TableShading{
		Val:   string(config.Pattern),
		Color: config.ForegroundColor,
		Fill:  config.BackgroundColor,
	}

	Info("设置表格背景成功")
	return nil
}

// SetCellBorders sets cell borders
func (t *Table) SetCellBorders(row, col int, config *CellBorderConfig) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	cell.Properties.TcBorders = &TableCellBorders{}

	if config.Top != nil {
		cell.Properties.TcBorders.Top = createTableCellBorder(config.Top)
	}
	if config.Left != nil {
		cell.Properties.TcBorders.Left = createTableCellBorder(config.Left)
	}
	if config.Bottom != nil {
		cell.Properties.TcBorders.Bottom = createTableCellBorder(config.Bottom)
	}
	if config.Right != nil {
		cell.Properties.TcBorders.Right = createTableCellBorder(config.Right)
	}
	if config.DiagDown != nil {
		cell.Properties.TcBorders.TL2BR = createTableCellBorder(config.DiagDown)
	}
	if config.DiagUp != nil {
		cell.Properties.TcBorders.TR2BL = createTableCellBorder(config.DiagUp)
	}

	Info(fmt.Sprintf("set cell (%d,%d)borders successfully", row, col))
	return nil
}

// SetCellShading sets cell background
func (t *Table) SetCellShading(row, col int, config *ShadingConfig) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	cell.Properties.Shd = &TableCellShading{
		Val:   string(config.Pattern),
		Color: config.ForegroundColor,
		Fill:  config.BackgroundColor,
	}

	Info(fmt.Sprintf("set cell (%d,%d)background successfully", row, col))
	return nil
}

// SetAlternatingRowColors 设置奇偶行颜色交替
func (t *Table) SetAlternatingRowColors(evenRowColor, oddRowColor string) error {
	for i := range t.Rows {
		var bgColor string
		if i%2 == 0 {
			bgColor = evenRowColor
		} else {
			bgColor = oddRowColor
		}

		// 为该行的所有单元格设置背景色
		for j := range t.Rows[i].Cells {
			err := t.SetCellShading(i, j, &ShadingConfig{
				Pattern:         ShadingPatternSolid,
				BackgroundColor: bgColor,
			})
			if err != nil {
				return fmt.Errorf("设置第%d行第%d列背景色失败: %v", i, j, err)
			}
		}
	}

	Info("设置奇偶行颜色交替成功")
	return nil
}

// RemoveTableBorders 移除表格边框
func (t *Table) RemoveTableBorders() error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	// 设置所有边框为无
	noBorderConfig := &BorderConfig{
		Style: BorderStyleNone,
		Width: 0,
		Color: "auto",
		Space: 0,
	}

	borderConfig := &TableBorderConfig{
		Top:     noBorderConfig,
		Left:    noBorderConfig,
		Bottom:  noBorderConfig,
		Right:   noBorderConfig,
		InsideH: noBorderConfig,
		InsideV: noBorderConfig,
	}

	return t.SetTableBorders(borderConfig)
}

// RemoveCellBorders 移除单元格边框
func (t *Table) RemoveCellBorders(row, col int) error {
	noBorderConfig := &BorderConfig{
		Style: BorderStyleNone,
		Width: 0,
		Color: "auto",
		Space: 0,
	}

	cellBorderConfig := &CellBorderConfig{
		Top:    noBorderConfig,
		Left:   noBorderConfig,
		Bottom: noBorderConfig,
		Right:  noBorderConfig,
	}

	return t.SetCellBorders(row, col, cellBorderConfig)
}

// CreateCustomTableStyle 创建自定义表格样式
func (t *Table) CreateCustomTableStyle(styleID, styleName string,
	borderConfig *TableBorderConfig,
	shadingConfig *ShadingConfig,
	firstRowBold bool) error {

	// 应用样式到表格
	config := &TableStyleConfig{
		StyleID:        styleID,
		FirstRowHeader: firstRowBold,
		BandedRows:     shadingConfig != nil,
	}

	err := t.ApplyTableStyle(config)
	if err != nil {
		return err
	}

	// 设置边框
	if borderConfig != nil {
		err = t.SetTableBorders(borderConfig)
		if err != nil {
			return err
		}
	}

	// 设置背景
	if shadingConfig != nil {
		err = t.SetTableShading(shadingConfig)
		if err != nil {
			return err
		}
	}

	Info(fmt.Sprintf("创建自定义表格样式成功：%s", styleID))
	return nil
}

// 辅助函数：创建表格边框
func createTableBorder(config *BorderConfig) *TableBorder {
	return &TableBorder{
		Val:   string(config.Style),
		Sz:    fmt.Sprintf("%d", config.Width),
		Space: fmt.Sprintf("%d", config.Space),
		Color: config.Color,
	}
}

// 辅助函数：创建单元格边框
func createTableCellBorder(config *BorderConfig) *TableCellBorder {
	return &TableCellBorder{
		Val:   string(config.Style),
		Sz:    fmt.Sprintf("%d", config.Width),
		Space: fmt.Sprintf("%d", config.Space),
		Color: config.Color,
	}
}

// CellIterator 单元格迭代器
type CellIterator struct {
	table      *Table
	currentRow int
	currentCol int
	totalRows  int
	totalCols  int
}

// CellInfo 单元格信息
type CellInfo struct {
	Row    int        // 行索引
	Col    int        // 列索引
	Cell   *TableCell // 单元格引用
	Text   string     // 单元格文本
	IsLast bool       // 是否为最后一个单元格
}

// NewCellIterator 创建新的单元格迭代器
func (t *Table) NewCellIterator() *CellIterator {
	totalRows := t.GetRowCount()
	totalCols := 0
	if totalRows > 0 {
		totalCols = t.GetColumnCount()
	}

	return &CellIterator{
		table:      t,
		currentRow: 0,
		currentCol: 0,
		totalRows:  totalRows,
		totalCols:  totalCols,
	}
}

// HasNext 检查是否还有下一个单元格
func (iter *CellIterator) HasNext() bool {
	if iter.totalRows == 0 || iter.totalCols == 0 {
		return false
	}

	// 检查当前位置是否超出范围
	return iter.currentRow < iter.totalRows &&
		(iter.currentRow < iter.totalRows-1 || iter.currentCol < iter.totalCols)
}

// Next 获取下一个单元格信息
func (iter *CellIterator) Next() (*CellInfo, error) {
	if !iter.HasNext() {
		return nil, fmt.Errorf("没有更多单元格")
	}

	// 获取当前单元格
	cell, err := iter.table.GetCell(iter.currentRow, iter.currentCol)
	if err != nil {
		return nil, fmt.Errorf("获取单元格失败: %v", err)
	}

	// 获取单元格文本
	text, _ := iter.table.GetCellText(iter.currentRow, iter.currentCol)

	// 创建单元格信息
	cellInfo := &CellInfo{
		Row:  iter.currentRow,
		Col:  iter.currentCol,
		Cell: cell,
		Text: text,
	}

	// 更新位置并检查是否为最后一个
	iter.currentCol++
	if iter.currentCol >= iter.totalCols {
		iter.currentCol = 0
		iter.currentRow++
	}

	// 检查是否为最后一个单元格
	cellInfo.IsLast = !iter.HasNext()

	return cellInfo, nil
}

// Reset 重置迭代器到开始位置
func (iter *CellIterator) Reset() {
	iter.currentRow = 0
	iter.currentCol = 0
}

// Current 获取当前位置信息（不移动迭代器）
func (iter *CellIterator) Current() (int, int) {
	return iter.currentRow, iter.currentCol
}

// Total 获取总单元格数量
func (iter *CellIterator) Total() int {
	return iter.totalRows * iter.totalCols
}

// Progress 获取迭代进度（0.0-1.0）
func (iter *CellIterator) Progress() float64 {
	if iter.totalRows == 0 || iter.totalCols == 0 {
		return 1.0
	}

	processed := iter.currentRow*iter.totalCols + iter.currentCol
	total := iter.totalRows * iter.totalCols

	return float64(processed) / float64(total)
}

// ForEach 遍历所有单元格，对每个单元格执行指定函数
func (t *Table) ForEach(fn func(row, col int, cell *TableCell, text string) error) error {
	iterator := t.NewCellIterator()

	for iterator.HasNext() {
		cellInfo, err := iterator.Next()
		if err != nil {
			return fmt.Errorf("迭代失败: %v", err)
		}

		if err := fn(cellInfo.Row, cellInfo.Col, cellInfo.Cell, cellInfo.Text); err != nil {
			return fmt.Errorf("回调函数执行失败 (行:%d, 列:%d): %v", cellInfo.Row, cellInfo.Col, err)
		}
	}

	return nil
}

// ForEachInRow 遍历指定行的所有单元格
func (t *Table) ForEachInRow(rowIndex int, fn func(col int, cell *TableCell, text string) error) error {
	if rowIndex < 0 || rowIndex >= t.GetRowCount() {
		return fmt.Errorf("行索引无效: %d", rowIndex)
	}

	colCount := t.GetColumnCount()
	for col := 0; col < colCount; col++ {
		cell, err := t.GetCell(rowIndex, col)
		if err != nil {
			return fmt.Errorf("获取单元格失败 (行:%d, 列:%d): %v", rowIndex, col, err)
		}

		text, _ := t.GetCellText(rowIndex, col)

		if err := fn(col, cell, text); err != nil {
			return fmt.Errorf("回调函数执行失败 (行:%d, 列:%d): %v", rowIndex, col, err)
		}
	}

	return nil
}

// ForEachInColumn 遍历指定列的所有单元格
func (t *Table) ForEachInColumn(colIndex int, fn func(row int, cell *TableCell, text string) error) error {
	if colIndex < 0 || colIndex >= t.GetColumnCount() {
		return fmt.Errorf("列索引无效: %d", colIndex)
	}

	rowCount := t.GetRowCount()
	for row := 0; row < rowCount; row++ {
		cell, err := t.GetCell(row, colIndex)
		if err != nil {
			return fmt.Errorf("获取单元格失败 (行:%d, 列:%d): %v", row, colIndex, err)
		}

		text, _ := t.GetCellText(row, colIndex)

		if err := fn(row, cell, text); err != nil {
			return fmt.Errorf("回调函数执行失败 (行:%d, 列:%d): %v", row, colIndex, err)
		}
	}

	return nil
}

// GetCellRange 获取指定范围内的所有单元格
func (t *Table) GetCellRange(startRow, startCol, endRow, endCol int) ([]*CellInfo, error) {
	// 参数验证
	if startRow < 0 || startCol < 0 || endRow >= t.GetRowCount() || endCol >= t.GetColumnCount() {
		return nil, fmt.Errorf("范围索引无效: (%d,%d) 到 (%d,%d)", startRow, startCol, endRow, endCol)
	}

	if startRow > endRow || startCol > endCol {
		return nil, fmt.Errorf("开始位置不能大于结束位置")
	}

	var cells []*CellInfo

	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cell, err := t.GetCell(row, col)
			if err != nil {
				return nil, fmt.Errorf("获取单元格失败 (行:%d, 列:%d): %v", row, col, err)
			}

			text, _ := t.GetCellText(row, col)

			cellInfo := &CellInfo{
				Row:    row,
				Col:    col,
				Cell:   cell,
				Text:   text,
				IsLast: row == endRow && col == endCol,
			}

			cells = append(cells, cellInfo)
		}
	}

	return cells, nil
}

// FindCells 查找满足条件的单元格
func (t *Table) FindCells(predicate func(row, col int, cell *TableCell, text string) bool) ([]*CellInfo, error) {
	var matchedCells []*CellInfo

	err := t.ForEach(func(row, col int, cell *TableCell, text string) error {
		if predicate(row, col, cell, text) {
			cellInfo := &CellInfo{
				Row:  row,
				Col:  col,
				Cell: cell,
				Text: text,
			}
			matchedCells = append(matchedCells, cellInfo)
		}
		return nil
	})

	if err != nil {
		return nil, err
	}

	return matchedCells, nil
}

// FindCellsByText 根据文本内容查找单元格
func (t *Table) FindCellsByText(searchText string, exactMatch bool) ([]*CellInfo, error) {
	return t.FindCells(func(row, col int, cell *TableCell, text string) bool {
		if exactMatch {
			return text == searchText
		}
		// 使用strings.Contains进行模糊匹配
		return strings.Contains(text, searchText)
	})
}

// ============== 单元格复杂内容功能 ==============
// 以下方法支持向表格单元格中添加段落、图片、列表、嵌套表格等复杂内容

// AddCellParagraph 向单元格添加段落
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//   - text: 段落文本内容
//
// Returns:
//   - *Paragraph: 新添加的段落对象
//   - error: returns error if index is invalid
func (t *Table) AddCellParagraph(row, col int, text string) (*Paragraph, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	// 创建新段落
	para := &Paragraph{
		Runs: []Run{
			{
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	// add to cell
	cell.Paragraphs = append(cell.Paragraphs, *para)

	Info(fmt.Sprintf("向单元格(%d,%d)添加段落成功", row, col))
	return &cell.Paragraphs[len(cell.Paragraphs)-1], nil
}

// AddCellFormattedParagraph 向单元格添加格式化段落
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//   - text: 段落文本内容
//   - format: 文本格式配置
//
// Returns:
//   - *Paragraph: 新添加的段落对象
//   - error: returns error if index is invalid
func (t *Table) AddCellFormattedParagraph(row, col int, text string, format *TextFormat) (*Paragraph, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	// create run properties
	runProps := &RunProperties{}

	if format != nil {
		if format.FontFamily != "" {
			runProps.FontFamily = &FontFamily{
				ASCII:    format.FontFamily,
				HAnsi:    format.FontFamily,
				EastAsia: format.FontFamily,
				CS:       format.FontFamily,
			}
		}

		if format.Bold {
			runProps.Bold = &Bold{}
		}

		if format.Italic {
			runProps.Italic = &Italic{}
		}

		if format.FontColor != "" {
			color := strings.TrimPrefix(format.FontColor, "#")
			runProps.Color = &Color{Val: color}
		}

		if format.FontSize > 0 {
			runProps.FontSize = &FontSize{Val: fmt.Sprintf("%d", format.FontSize*2)}
		}

		if format.Underline {
			runProps.Underline = &Underline{Val: "single"}
		}

		if format.Strike {
			runProps.Strike = &Strike{}
		}

		if format.Highlight != "" {
			runProps.Highlight = &Highlight{Val: format.Highlight}
		}
	}

	// 创建新段落
	para := &Paragraph{
		Runs: []Run{
			{
				Properties: runProps,
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	// add to cell
	cell.Paragraphs = append(cell.Paragraphs, *para)

	Info(fmt.Sprintf("向单元格(%d,%d)添加格式化段落成功", row, col))
	return &cell.Paragraphs[len(cell.Paragraphs)-1], nil
}

// ClearCellParagraphs 清空单元格中的所有段落，只保留一个空段落
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//
// Returns:
//   - error: returns error if index is invalid
func (t *Table) ClearCellParagraphs(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 清空段落，只保留一个空段落（OOXML规范要求单元格至少有一个段落）
	cell.Paragraphs = []Paragraph{
		{
			Runs: []Run{
				{
					Text: Text{Content: ""},
				},
			},
		},
	}

	Info(fmt.Sprintf("清空单元格(%d,%d)段落成功", row, col))
	return nil
}

// GetCellParagraphs 获取单元格中的所有段落
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//
// Returns:
//   - []Paragraph: 单元格中的所有段落
//   - error: returns error if index is invalid
func (t *Table) GetCellParagraphs(row, col int) ([]Paragraph, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	return cell.Paragraphs, nil
}

// AddNestedTable 向单元格添加嵌套表格
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//   - config: 嵌套表格的配置
//
// Returns:
//   - *Table: 新创建的嵌套表格对象
//   - error: 如果索引无效或配置无效则返回错误
func (t *Table) AddNestedTable(row, col int, config *TableConfig) (*Table, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	if config.Rows <= 0 || config.Cols <= 0 {
		Error("嵌套表格行数和列数必须大于0")
		return nil, NewValidationError("TableConfig", "", "嵌套表格行数和列数必须大于0")
	}

	// 创建嵌套表格
	nestedTable := &Table{
		Properties: &TableProperties{
			TableW: &TableWidth{
				W:    fmt.Sprintf("%d", config.Width),
				Type: "dxa",
			},
			TableJc: &TableJc{
				Val: "center",
			},
			TableLook: &TableLook{
				Val:      "04A0",
				FirstRow: "1",
				LastRow:  "0",
				FirstCol: "1",
				LastCol:  "0",
				NoHBand:  "0",
				NoVBand:  "1",
			},
			TableBorders: &TableBorders{
				Top:     &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Left:    &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Bottom:  &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Right:   &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				InsideH: &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				InsideV: &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
			},
			TableLayout: &TableLayoutType{
				Type: "autofit",
			},
			TableCellMar: &TableCellMargins{
				Left:  &TableCellSpace{W: "108", Type: "dxa"},
				Right: &TableCellSpace{W: "108", Type: "dxa"},
			},
		},
		Grid: &TableGrid{},
		Rows: make([]TableRow, 0, config.Rows),
	}

	// 设置列宽
	colWidths := config.ColWidths
	if len(colWidths) == 0 {
		avgWidth := config.Width / config.Cols
		colWidths = make([]int, config.Cols)
		for i := range colWidths {
			colWidths[i] = avgWidth
		}
	} else if len(colWidths) != config.Cols {
		Error("嵌套表格列宽数量与列数不匹配")
		return nil, NewValidationError("TableConfig.ColWidths", "", "列宽数量与列数不匹配")
	}

	// 创建表格网格
	for _, width := range colWidths {
		nestedTable.Grid.Cols = append(nestedTable.Grid.Cols, TableGridCol{
			W: fmt.Sprintf("%d", width),
		})
	}

	// 创建表格行和单元格
	for i := 0; i < config.Rows; i++ {
		tableRow := TableRow{
			Cells: make([]TableCell, 0, config.Cols),
		}

		for j := 0; j < config.Cols; j++ {
			tableCell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: &TableCellW{
						W:    fmt.Sprintf("%d", colWidths[j]),
						Type: "dxa",
					},
					VAlign: &VAlign{
						Val: "center",
					},
				},
				Paragraphs: []Paragraph{
					{
						Runs: []Run{
							{
								Text: Text{Content: ""},
							},
						},
					},
				},
			}

			// if initial data exists, set cell content
			if config.Data != nil && i < len(config.Data) && j < len(config.Data[i]) {
				tableCell.Paragraphs[0].Runs[0].Text.Content = config.Data[i][j]
			}

			tableRow.Cells = append(tableRow.Cells, tableCell)
		}

		nestedTable.Rows = append(nestedTable.Rows, tableRow)
	}

	// add to cell的嵌套表格列表
	cell.Tables = append(cell.Tables, *nestedTable)

	Info(fmt.Sprintf("向单元格(%d,%d)添加嵌套表格成功：%d行 x %d列", row, col, config.Rows, config.Cols))
	return &cell.Tables[len(cell.Tables)-1], nil
}

// GetNestedTables 获取单元格中的所有嵌套表格
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//
// Returns:
//   - []Table: 单元格中的所有嵌套表格
//   - error: returns error if index is invalid
func (t *Table) GetNestedTables(row, col int) ([]Table, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	return cell.Tables, nil
}

// CellListConfig 单元格列表配置
type CellListConfig struct {
	Type         ListType   // list type
	BulletSymbol BulletType // 项目符号（仅用于无序列表）
	Items        []string   // 列表项内容
}

// AddCellList 向单元格添加列表
// Parameters:
//   - row: row index (starting from 0)
//   - col: column index (starting from 0)
//   - config: 列表配置
//
// Returns:
//   - error: returns error if index is invalid
func (t *Table) AddCellList(row, col int, config *CellListConfig) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if config == nil || len(config.Items) == 0 {
		return NewValidationError("CellListConfig", "", "列表配置不能为空且必须包含列表项")
	}

	// 根据列表类型确定前缀
	for i, item := range config.Items {
		var prefix string
		switch config.Type {
		case ListTypeBullet:
			// 使用项目符号
			bulletSymbol := config.BulletSymbol
			if bulletSymbol == "" {
				bulletSymbol = BulletTypeDot
			}
			prefix = string(bulletSymbol) + " "
		case ListTypeNumber, ListTypeDecimal:
			// 使用数字编号
			prefix = fmt.Sprintf("%d. ", i+1)
		case ListTypeLowerLetter:
			// 使用小写字母
			prefix = fmt.Sprintf("%c. ", 'a'+i)
		case ListTypeUpperLetter:
			// 使用大写字母
			prefix = fmt.Sprintf("%c. ", 'A'+i)
		case ListTypeLowerRoman:
			// 使用小写罗马数字
			prefix = fmt.Sprintf("%s. ", toRomanLower(i+1))
		case ListTypeUpperRoman:
			// 使用大写罗马数字
			prefix = fmt.Sprintf("%s. ", toRomanUpper(i+1))
		default:
			// 默认使用项目符号
			prefix = string(BulletTypeDot) + " "
		}

		// 创建列表项段落
		para := Paragraph{
			Runs: []Run{
				{
					Text: Text{
						Content: prefix + item,
						Space:   "preserve",
					},
				},
			},
		}

		// add to cell
		cell.Paragraphs = append(cell.Paragraphs, para)
	}

	Info(fmt.Sprintf("向单元格(%d,%d)添加列表成功：%d个列表项", row, col, len(config.Items)))
	return nil
}

// toRomanLower 将数字转换为小写罗马数字
func toRomanLower(num int) string {
	return strings.ToLower(toRomanUpper(num))
}

// toRomanUpper 将数字转换为大写罗马数字
func toRomanUpper(num int) string {
	values := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	symbols := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}

	if num <= 0 || num > 3999 {
		return fmt.Sprintf("%d", num)
	}

	result := ""
	for i, value := range values {
		for num >= value {
			result += symbols[i]
			num -= value
		}
	}
	return result
}

// CellImageConfig 单元格图片配置
type CellImageConfig struct {
	// 图片来源 - 文件路径
	FilePath string
	// 图片来源 - 二进制数据
	Data []byte
	// 图片格式（当使用Data时需要指定）
	Format ImageFormat
	// 图片宽度（毫米），0表示自动
	Width float64
	// 图片高度（毫米），0表示自动
	Height float64
	// 是否保持宽高比
	KeepAspectRatio bool
	// 图片替代文字
	AltText string
	// 图片标题
	Title string
}
