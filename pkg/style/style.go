// Package style provides Word document style management functionality
package style

import (
	"encoding/xml"
	"fmt"
)

// StyleType represents the type of style
type StyleType string

const (
	StyleTypeParagraph StyleType = "paragraph"
	StyleTypeCharacter StyleType = "character"
	StyleTypeTable     StyleType = "table"
	StyleTypeNumbering StyleType = "numbering"
)

// Style defines a Word document style
type Style struct {
	XMLName     xml.Name             `xml:"w:style"`
	Type        string               `xml:"w:type,attr"`
	StyleID     string               `xml:"w:styleId,attr"`
	Name        *StyleName           `xml:"w:name,omitempty"`
	BasedOn     *BasedOn             `xml:"w:basedOn,omitempty"`
	Next        *Next                `xml:"w:next,omitempty"`
	Default     bool                 `xml:"w:default,attr,omitempty"`
	CustomStyle bool                 `xml:"w:customStyle,attr,omitempty"`
	ParagraphPr *ParagraphProperties `xml:"w:pPr,omitempty"`
	RunPr       *RunProperties       `xml:"w:rPr,omitempty"`
	TablePr     *TableProperties     `xml:"w:tblPr,omitempty"`
	TableRowPr  *TableRowProperties  `xml:"w:trPr,omitempty"`
	TableCellPr *TableCellProperties `xml:"w:tcPr,omitempty"`
}

// StyleName specifies the name of a style
type StyleName struct {
	XMLName xml.Name `xml:"w:name"`
	Val     string   `xml:"w:val,attr"`
}

// BasedOn specifies the parent style
type BasedOn struct {
	XMLName xml.Name `xml:"w:basedOn"`
	Val     string   `xml:"w:val,attr"`
}

// Next specifies the style that follows
type Next struct {
	XMLName xml.Name `xml:"w:next"`
	Val     string   `xml:"w:val,attr"`
}

// ParagraphProperties defines paragraph properties
// Note: field order must comply with OpenXML specification
type ParagraphProperties struct {
	XMLName         xml.Name         `xml:"w:pPr"`
	KeepNext        *KeepNext        `xml:"w:keepNext,omitempty"`
	KeepLines       *KeepLines       `xml:"w:keepLines,omitempty"`
	PageBreak       *PageBreak       `xml:"w:pageBreakBefore,omitempty"`
	ParagraphBorder *ParagraphBorder `xml:"w:pBdr,omitempty"`
	Shading         *Shading         `xml:"w:shd,omitempty"`
	SnapToGrid      *SnapToGrid      `xml:"w:snapToGrid,omitempty"`
	Spacing         *Spacing         `xml:"w:spacing,omitempty"`
	Indentation     *Indentation     `xml:"w:ind,omitempty"`
	Justification   *Justification   `xml:"w:jc,omitempty"`
	OutlineLevel    *OutlineLevel    `xml:"w:outlineLvl,omitempty"`
}

// ParagraphBorder defines paragraph borders
type ParagraphBorder struct {
	XMLName xml.Name             `xml:"w:pBdr"`
	Top     *ParagraphBorderLine `xml:"w:top,omitempty"`
	Left    *ParagraphBorderLine `xml:"w:left,omitempty"`
	Bottom  *ParagraphBorderLine `xml:"w:bottom,omitempty"`
	Right   *ParagraphBorderLine `xml:"w:right,omitempty"`
}

// ParagraphBorderLine defines paragraph border line properties
type ParagraphBorderLine struct {
	XMLName xml.Name `xml:""`
	Val     string   `xml:"w:val,attr"`
	Color   string   `xml:"w:color,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
}

// Shading defines shading and fill color
type Shading struct {
	XMLName xml.Name `xml:"w:shd"`
	Fill    string   `xml:"w:fill,attr"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// RunProperties defines character properties
// Note: field order must comply with OpenXML specification, w:rFonts must come before w:color
type RunProperties struct {
	XMLName    xml.Name    `xml:"w:rPr"`
	FontFamily *FontFamily `xml:"w:rFonts,omitempty"`
	Bold       *Bold       `xml:"w:b,omitempty"`
	Italic     *Italic     `xml:"w:i,omitempty"`
	Underline  *Underline  `xml:"w:u,omitempty"`
	Strike     *Strike     `xml:"w:strike,omitempty"`
	Color      *Color      `xml:"w:color,omitempty"`
	FontSize   *FontSize   `xml:"w:sz,omitempty"`
	Highlight  *Highlight  `xml:"w:highlight,omitempty"`
}

// TableProperties defines table properties
type TableProperties struct {
	XMLName    xml.Name       `xml:"w:tblPr"`
	TblInd     *TblIndent     `xml:"w:tblInd,omitempty"`
	TblBorders *TblBorders    `xml:"w:tblBorders,omitempty"`
	TblCellMar *TblCellMargin `xml:"w:tblCellMar,omitempty"`
}

// TblIndent specifies table indentation
type TblIndent struct {
	XMLName xml.Name `xml:"w:tblInd"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// TblBorders defines table borders
type TblBorders struct {
	XMLName xml.Name   `xml:"w:tblBorders"`
	Top     *TblBorder `xml:"w:top,omitempty"`
	Left    *TblBorder `xml:"w:left,omitempty"`
	Bottom  *TblBorder `xml:"w:bottom,omitempty"`
	Right   *TblBorder `xml:"w:right,omitempty"`
	InsideH *TblBorder `xml:"w:insideH,omitempty"`
	InsideV *TblBorder `xml:"w:insideV,omitempty"`
}

// TblBorder defines a table border
type TblBorder struct {
	Val   string `xml:"w:val,attr"`
	Sz    string `xml:"w:sz,attr"`
	Space string `xml:"w:space,attr"`
	Color string `xml:"w:color,attr"`
}

// TblCellMargin defines table cell margins
type TblCellMargin struct {
	XMLName xml.Name      `xml:"w:tblCellMar"`
	Top     *TblCellSpace `xml:"w:top,omitempty"`
	Left    *TblCellSpace `xml:"w:left,omitempty"`
	Bottom  *TblCellSpace `xml:"w:bottom,omitempty"`
	Right   *TblCellSpace `xml:"w:right,omitempty"`
}

// TblCellSpace specifies table cell spacing
type TblCellSpace struct {
	W    string `xml:"w:w,attr"`
	Type string `xml:"w:type,attr"`
}

// TableRowProperties defines table row properties
// Table row style properties to be implemented in future versions
type TableRowProperties struct {
	XMLName xml.Name `xml:"w:trPr"`
}

// TableCellProperties defines table cell properties
// Table cell style properties to be implemented in future versions
type TableCellProperties struct {
	XMLName xml.Name `xml:"w:tcPr"`
}

// Spacing defines paragraph spacing
type Spacing struct {
	XMLName  xml.Name `xml:"w:spacing"`
	Before   string   `xml:"w:before,attr,omitempty"`
	After    string   `xml:"w:after,attr,omitempty"`
	Line     string   `xml:"w:line,attr,omitempty"`
	LineRule string   `xml:"w:lineRule,attr,omitempty"`
}

type Justification struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

type Indentation struct {
	XMLName   xml.Name `xml:"w:ind"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
	Left      string   `xml:"w:left,attr,omitempty"`
	Right     string   `xml:"w:right,attr,omitempty"`
}

type KeepNext struct {
	XMLName xml.Name `xml:"w:keepNext"`
}

type KeepLines struct {
	XMLName xml.Name `xml:"w:keepLines"`
}

type PageBreak struct {
	XMLName xml.Name `xml:"w:pageBreakBefore"`
}

type OutlineLevel struct {
	XMLName xml.Name `xml:"w:outlineLvl"`
	Val     string   `xml:"w:val,attr"`
}

// SnapToGrid controls grid alignment setting
// Set to "0" to disable grid alignment, "1" to enable (complies with OOXML spec, supports only "0" or "1")
// Note: This type is intentionally duplicated in the document package to allow independent package usage
type SnapToGrid struct {
	XMLName xml.Name `xml:"w:snapToGrid"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

type Bold struct {
	XMLName xml.Name `xml:"w:b"`
}

type Italic struct {
	XMLName xml.Name `xml:"w:i"`
}

type Underline struct {
	XMLName xml.Name `xml:"w:u"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

type Strike struct {
	XMLName xml.Name `xml:"w:strike"`
}

type FontSize struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

type FontFamily struct {
	XMLName  xml.Name `xml:"w:rFonts"`
	ASCII    string   `xml:"w:ascii,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	HAnsi    string   `xml:"w:hAnsi,attr,omitempty"`
	CS       string   `xml:"w:cs,attr,omitempty"`
}

type Highlight struct {
	XMLName xml.Name `xml:"w:highlight"`
	Val     string   `xml:"w:val,attr"`
}

// Styles represents a collection of styles
type Styles struct {
	XMLName xml.Name `xml:"w:styles"`
	Xmlns   string   `xml:"xmlns:w,attr"`
	Styles  []Style  `xml:"w:style"`
}

// StyleManager manages all document styles
type StyleManager struct {
	styles map[string]*Style
}

// NewStyleManager creates a new style manager with predefined styles
func NewStyleManager() *StyleManager {
	sm := &StyleManager{
		styles: make(map[string]*Style),
	}
	sm.initializePredefinedStyles()
	return sm
}

// GetStyle retrieves a style by its ID
func (sm *StyleManager) GetStyle(styleID string) *Style {
	return sm.styles[styleID]
}

// AddStyle adds a style to the manager
func (sm *StyleManager) AddStyle(style *Style) {
	sm.styles[style.StyleID] = style
}

// GetAllStyles returns all styles in the manager
func (sm *StyleManager) GetAllStyles() []*Style {
	styles := make([]*Style, 0, len(sm.styles))
	for _, style := range sm.styles {
		styles = append(styles, style)
	}
	return styles
}

// initializePredefinedStyles initializes all predefined styles
func (sm *StyleManager) initializePredefinedStyles() {
	sm.addNormalStyle()
	sm.addHeadingStyles()
	sm.addSpecialStyles()
}

// addNormalStyle adds the Normal paragraph style
func (sm *StyleManager) addNormalStyle() {
	normalStyle := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Normal",
		Default: true,
		Name: &StyleName{
			Val: "Normal",
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "21", // 10.5pt (half-point units)
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri",
				EastAsia: "宋体",
				HAnsi:    "Calibri",
				CS:       "Times New Roman",
			},
		},
	}
	sm.AddStyle(normalStyle)
}

// addHeadingStyles adds heading styles for levels 1-9
func (sm *StyleManager) addHeadingStyles() {
	heading1 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading1",
		Name: &StyleName{
			Val: "heading 1",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "240", // 12pt spacing before
				After:  "0",   // 0pt spacing after
			},
			OutlineLevel: &OutlineLevel{
				Val: "0",
			},
		},
		RunPr: &RunProperties{
			Bold: &Bold{},
			FontSize: &FontSize{
				Val: "32", // 16pt
			},
			Color: &Color{
				Val: "2F5496", // dark blue
			},
		},
	}
	sm.AddStyle(heading1)

	heading2 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading2",
		Name: &StyleName{
			Val: "heading 2",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "1",
			},
		},
		RunPr: &RunProperties{
			Bold: &Bold{},
			FontSize: &FontSize{
				Val: "26", // 13pt
			},
			Color: &Color{
				Val: "2F5496", // dark blue
			},
		},
	}
	sm.AddStyle(heading2)

	heading3 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading3",
		Name: &StyleName{
			Val: "heading 3",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "2",
			},
		},
		RunPr: &RunProperties{
			Bold: &Bold{},
			FontSize: &FontSize{
				Val: "24", // 12pt
			},
			Color: &Color{
				Val: "1F3763", // dark blue
			},
		},
	}
	sm.AddStyle(heading3)

	heading4 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading4",
		Name: &StyleName{
			Val: "heading 4",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "3",
			},
		},
		RunPr: &RunProperties{
			Bold:   &Bold{},
			Italic: &Italic{},
			FontSize: &FontSize{
				Val: "22", // 11pt
			},
			Color: &Color{
				Val: "2F5496", // dark blue
			},
		},
	}
	sm.AddStyle(heading4)

	heading5 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading5",
		Name: &StyleName{
			Val: "heading 5",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "4",
			},
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "22", // 11pt
			},
			Color: &Color{
				Val: "2F5496", // dark blue
			},
		},
	}
	sm.AddStyle(heading5)

	heading6 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading6",
		Name: &StyleName{
			Val: "heading 6",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "5",
			},
		},
		RunPr: &RunProperties{
			Italic: &Italic{},
			FontSize: &FontSize{
				Val: "22", // 11pt
			},
			Color: &Color{
				Val: "1F3763", // dark blue
			},
		},
	}
	sm.AddStyle(heading6)

	heading7 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading7",
		Name: &StyleName{
			Val: "heading 7",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "6",
			},
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "20", // 10pt
			},
			Color: &Color{
				Val: "1F3763", // dark blue
			},
		},
	}
	sm.AddStyle(heading7)

	heading8 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading8",
		Name: &StyleName{
			Val: "heading 8",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "7",
			},
		},
		RunPr: &RunProperties{
			Italic: &Italic{},
			FontSize: &FontSize{
				Val: "20", // 10pt
			},
			Color: &Color{
				Val: "272727", // dark gray
			},
		},
	}
	sm.AddStyle(heading8)

	heading9 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Heading9",
		Name: &StyleName{
			Val: "heading 9",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			KeepNext:  &KeepNext{},
			KeepLines: &KeepLines{},
			Spacing: &Spacing{
				Before: "120", // 6 points before spacing
				After:  "0",   // 0 points after spacing
			},
			OutlineLevel: &OutlineLevel{
				Val: "8",
			},
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "18", // 9pt
			},
			Color: &Color{
				Val: "272727", // dark gray
			},
		},
	}
	sm.AddStyle(heading9)
}

// addSpecialStyles adds special styles like title, subtitle, quotes, and code blocks
func (sm *StyleManager) addSpecialStyles() {
	title := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Title",
		Name: &StyleName{
			Val: "标题",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Justification: &Justification{
				Val: "center",
			},
			Spacing: &Spacing{
				Before: "240", // 12pt spacing before
				After:  "60",  // 3pt spacing after
			},
		},
		RunPr: &RunProperties{
			Bold: &Bold{},
			FontSize: &FontSize{
				Val: "56", // 28pt
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri Light",
				EastAsia: "微软雅黑 Light",
				HAnsi:    "Calibri Light",
				CS:       "Calibri Light",
			},
			Color: &Color{
				Val: "2F5496", // dark blue
			},
		},
	}
	sm.AddStyle(title)

	subtitle := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Subtitle",
		Name: &StyleName{
			Val: "副标题",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Justification: &Justification{
				Val: "center",
			},
			Spacing: &Spacing{
				Before: "0",   // 0pt spacing before
				After:  "160", // 8pt spacing after
			},
		},
		RunPr: &RunProperties{
			Italic: &Italic{},
			FontSize: &FontSize{
				Val: "30", // 15pt
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri",
				EastAsia: "微软雅黑",
				HAnsi:    "Calibri",
				CS:       "Calibri",
			},
			Color: &Color{
				Val: "7030A0", // purple
			},
		},
	}
	sm.AddStyle(subtitle)

	sm.addTOCStyles()

	listParagraph := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "ListParagraph",
		Name: &StyleName{
			Val: "列表段落",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Indentation: &Indentation{
				Left: "720", // left indent 0.5in
			},
			Spacing: &Spacing{
				After:    "120", // 6pt spacing after
				Line:     "276", // 1.15x line spacing
				LineRule: "auto",
			},
		},
	}
	sm.AddStyle(listParagraph)

	emphasis := &Style{
		Type:    string(StyleTypeCharacter),
		StyleID: "Emphasis",
		Name: &StyleName{
			Val: "强调",
		},
		RunPr: &RunProperties{
			Italic: &Italic{},
		},
	}
	sm.AddStyle(emphasis)

	strong := &Style{
		Type:    string(StyleTypeCharacter),
		StyleID: "Strong",
		Name: &StyleName{
			Val: "加粗",
		},
		RunPr: &RunProperties{
			Bold: &Bold{},
		},
	}
	sm.AddStyle(strong)

	quote := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "Quote",
		Name: &StyleName{
			Val: "引用",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Indentation: &Indentation{
				Left:  "720", // left indent 0.5in
				Right: "720", // right indent 0.5in
			},
			Spacing: &Spacing{
				Before: "120", // 6pt spacing before
				After:  "120", // 6pt spacing after
			},
		},
		RunPr: &RunProperties{
			Italic: &Italic{},
			Color: &Color{
				Val: "666666", // medium gray
			},
		},
	}
	sm.AddStyle(quote)

	codeBlock := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "CodeBlock",
		Name: &StyleName{
			Val: "代码块",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Indentation: &Indentation{
				Left: "360", // left indent 0.25in
			},
			Spacing: &Spacing{
				Before: "60", // 3pt spacing before
				After:  "60", // 3pt spacing after
			},
			ParagraphBorder: &ParagraphBorder{
				Top: &ParagraphBorderLine{
					Val:   "thick",
					Color: "E9E7E7",
					Sz:    "8",
					Space: "8",
				},
				Left: &ParagraphBorderLine{
					Val:   "thick",
					Color: "E9E7E7",
					Sz:    "8",
					Space: "8",
				},
				Bottom: &ParagraphBorderLine{
					Val:   "thick",
					Color: "E9E7E7",
					Sz:    "8",
					Space: "8",
				},
				Right: &ParagraphBorderLine{
					Val:   "thick",
					Color: "E9E7E7",
					Sz:    "8",
					Space: "8",
				},
			},
			Shading: &Shading{
				Fill: "F6F5F5",
				Val:  "clear",
			},
		},
		RunPr: &RunProperties{
			FontFamily: &FontFamily{
				ASCII:    "Consolas",
				EastAsia: "Consolas",
				HAnsi:    "Consolas",
				CS:       "Consolas",
			},
			FontSize: &FontSize{
				Val: "18", // 9pt
			},
		},
	}
	sm.AddStyle(codeBlock)

	codeChar := &Style{
		Type:    string(StyleTypeCharacter),
		StyleID: "CodeChar",
		Name: &StyleName{
			Val: "代码字符",
		},
		RunPr: &RunProperties{
			FontFamily: &FontFamily{
				ASCII:    "Consolas",
				EastAsia: "Consolas",
				HAnsi:    "Consolas",
				CS:       "Consolas",
			},
			FontSize: &FontSize{
				Val: "18", // 9pt
			},
		},
	}
	sm.AddStyle(codeChar)

	sm.addTableStyles()
}

// addTOCStyles adds table of contents styles for levels 1-9
func (sm *StyleManager) addTOCStyles() {
	toc1 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "13",
		Name: &StyleName{
			Val: "toc 1",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Spacing: &Spacing{
				After: "100", // 5pt spacing after
			},
			Indentation: &Indentation{
				Left: "0",
			},
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "22", // 11pt
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri",
				EastAsia: "宋体",
				HAnsi:    "Calibri",
				CS:       "Times New Roman",
			},
		},
	}
	sm.AddStyle(toc1)

	toc2 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "14",
		Name: &StyleName{
			Val: "toc 2",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Spacing: &Spacing{
				After: "100", // 5pt spacing after
			},
			Indentation: &Indentation{
				Left: "240", // left indent 240 twips (12pt)
			},
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "22", // 11pt
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri",
				EastAsia: "宋体",
				HAnsi:    "Calibri",
				CS:       "Times New Roman",
			},
		},
	}
	sm.AddStyle(toc2)

	toc3 := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "15",
		Name: &StyleName{
			Val: "toc 3",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Spacing: &Spacing{
				After: "100", // 5pt spacing after
			},
			Indentation: &Indentation{
				Left: "480", // left indent 480 twips (24pt)
			},
		},
		RunPr: &RunProperties{
			FontSize: &FontSize{
				Val: "22", // 11pt
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri",
				EastAsia: "宋体",
				HAnsi:    "Calibri",
				CS:       "Times New Roman",
			},
		},
	}
	sm.AddStyle(toc3)

	// TOC 4-9 levels
	for level := 4; level <= 9; level++ {
		styleID := fmt.Sprintf("%d", 12+level) // 16, 17, 18, 19, 20, 21
		tocStyle := &Style{
			Type:    string(StyleTypeParagraph),
			StyleID: styleID,
			Name: &StyleName{
				Val: fmt.Sprintf("toc %d", level),
			},
			BasedOn: &BasedOn{
				Val: "Normal",
			},
			Next: &Next{
				Val: "Normal",
			},
			ParagraphPr: &ParagraphProperties{
				Spacing: &Spacing{
					After: "100", // 5pt spacing after
				},
				Indentation: &Indentation{
					Left: fmt.Sprintf("%d", level*240), // each level adds 240 twips (12pt)
				},
			},
			RunPr: &RunProperties{
				FontSize: &FontSize{
					Val: "22", // 11pt
				},
				FontFamily: &FontFamily{
					ASCII:    "Calibri",
					EastAsia: "宋体",
					HAnsi:    "Calibri",
					CS:       "Times New Roman",
				},
			},
		}
		sm.AddStyle(tocStyle)
	}

	tocBase := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "12",
		Name: &StyleName{
			Val: "TOCHeading",
		},
		BasedOn: &BasedOn{
			Val: "Normal",
		},
		Next: &Next{
			Val: "Normal",
		},
		ParagraphPr: &ParagraphProperties{
			Spacing: &Spacing{
				Before: "240", // 12pt spacing before
				After:  "120", // 6pt spacing after
			},
			Justification: &Justification{
				Val: "center",
			},
		},
		RunPr: &RunProperties{
			Bold: &Bold{},
			FontSize: &FontSize{
				Val: "26", // 13pt
			},
			FontFamily: &FontFamily{
				ASCII:    "Calibri",
				EastAsia: "宋体",
				HAnsi:    "Calibri",
				CS:       "Times New Roman",
			},
		},
	}
	sm.AddStyle(tocBase)
}

// GetStyleWithInheritance retrieves a style with inherited properties
// Merges parent style properties if style is based on another style
func (sm *StyleManager) GetStyleWithInheritance(styleID string) *Style {
	style := sm.GetStyle(styleID)
	if style == nil {
		return nil
	}

	if style.BasedOn == nil {
		return style
	}

	baseStyle := sm.GetStyleWithInheritance(style.BasedOn.Val)
	if baseStyle == nil {
		return style
	}

	mergedStyle := &Style{
		Type:        style.Type,
		StyleID:     style.StyleID,
		Name:        style.Name,
		BasedOn:     style.BasedOn,
		Next:        style.Next,
		Default:     style.Default,
		CustomStyle: style.CustomStyle,
	}

	mergedStyle.ParagraphPr = mergeParagraphProperties(baseStyle.ParagraphPr, style.ParagraphPr)
	mergedStyle.RunPr = mergeRunProperties(baseStyle.RunPr, style.RunPr)
	if style.TablePr != nil {
		mergedStyle.TablePr = style.TablePr
	} else if baseStyle.TablePr != nil {
		mergedStyle.TablePr = baseStyle.TablePr
	}

	return mergedStyle
}

// mergeParagraphProperties merges paragraph properties with priority: override > base
func mergeParagraphProperties(base, override *ParagraphProperties) *ParagraphProperties {
	if base == nil {
		return override
	}
	if override == nil {
		return base
	}

	merged := &ParagraphProperties{}

	if override.Spacing != nil {
		merged.Spacing = override.Spacing
	} else if base.Spacing != nil {
		merged.Spacing = base.Spacing
	}

	if override.Justification != nil {
		merged.Justification = override.Justification
	} else if base.Justification != nil {
		merged.Justification = base.Justification
	}

	if override.Indentation != nil {
		merged.Indentation = override.Indentation
	} else if base.Indentation != nil {
		merged.Indentation = base.Indentation
	}

	if override.ParagraphBorder != nil {
		merged.ParagraphBorder = override.ParagraphBorder
	} else if base.ParagraphBorder != nil {
		merged.ParagraphBorder = base.ParagraphBorder
	}

	if override.Shading != nil {
		merged.Shading = override.Shading
	} else if base.Shading != nil {
		merged.Shading = base.Shading
	}
	if override.KeepNext != nil {
		merged.KeepNext = override.KeepNext
	} else if base.KeepNext != nil {
		merged.KeepNext = base.KeepNext
	}

	if override.KeepLines != nil {
		merged.KeepLines = override.KeepLines
	} else if base.KeepLines != nil {
		merged.KeepLines = base.KeepLines
	}

	if override.PageBreak != nil {
		merged.PageBreak = override.PageBreak
	} else if base.PageBreak != nil {
		merged.PageBreak = base.PageBreak
	}

	if override.OutlineLevel != nil {
		merged.OutlineLevel = override.OutlineLevel
	} else if base.OutlineLevel != nil {
		merged.OutlineLevel = base.OutlineLevel
	}

	return merged
}

// mergeRunProperties merges run properties with priority: override > base
func mergeRunProperties(base, override *RunProperties) *RunProperties {
	if base == nil {
		return override
	}
	if override == nil {
		return base
	}

	merged := &RunProperties{}

	if override.Bold != nil {
		merged.Bold = override.Bold
	} else if base.Bold != nil {
		merged.Bold = base.Bold
	}

	if override.Italic != nil {
		merged.Italic = override.Italic
	} else if base.Italic != nil {
		merged.Italic = base.Italic
	}

	if override.Underline != nil {
		merged.Underline = override.Underline
	} else if base.Underline != nil {
		merged.Underline = base.Underline
	}

	if override.Strike != nil {
		merged.Strike = override.Strike
	} else if base.Strike != nil {
		merged.Strike = base.Strike
	}
	if override.FontSize != nil {
		merged.FontSize = override.FontSize
	} else if base.FontSize != nil {
		merged.FontSize = base.FontSize
	}

	if override.Color != nil {
		merged.Color = override.Color
	} else if base.Color != nil {
		merged.Color = base.Color
	}

	if override.FontFamily != nil {
		merged.FontFamily = override.FontFamily
	} else if base.FontFamily != nil {
		merged.FontFamily = base.FontFamily
	}

	if override.Highlight != nil {
		merged.Highlight = override.Highlight
	} else if base.Highlight != nil {
		merged.Highlight = base.Highlight
	}

	return merged
}

// CreateCustomStyle creates a new custom style with optional parent style
func (sm *StyleManager) CreateCustomStyle(styleID, name string, styleType StyleType, basedOn string) *Style {
	style := &Style{
		Type:        string(styleType),
		StyleID:     styleID,
		CustomStyle: true,
		Name: &StyleName{
			Val: name,
		},
	}

	if basedOn != "" {
		style.BasedOn = &BasedOn{
			Val: basedOn,
		}
	}

	sm.AddStyle(style)
	return style
}

// RemoveStyle removes a style by its ID
func (sm *StyleManager) RemoveStyle(styleID string) {
	delete(sm.styles, styleID)
}

// StyleExists checks if a style with the given ID exists
func (sm *StyleManager) StyleExists(styleID string) bool {
	_, exists := sm.styles[styleID]
	return exists
}

// Clone creates a deep copy of the style manager to avoid conflicts during template rendering
func (sm *StyleManager) Clone() *StyleManager {
	clonedSM := &StyleManager{
		styles: make(map[string]*Style),
	}

	for styleID, style := range sm.styles {
		clonedSM.styles[styleID] = sm.cloneStyle(style)
	}

	return clonedSM
}

// cloneStyle creates a deep copy of a single style
func (sm *StyleManager) cloneStyle(source *Style) *Style {
	if source == nil {
		return nil
	}

	cloned := &Style{
		Type:        source.Type,
		StyleID:     source.StyleID,
		Default:     source.Default,
		CustomStyle: source.CustomStyle,
	}

	if source.Name != nil {
		cloned.Name = &StyleName{Val: source.Name.Val}
	}

	if source.BasedOn != nil {
		cloned.BasedOn = &BasedOn{Val: source.BasedOn.Val}
	}

	if source.Next != nil {
		cloned.Next = &Next{Val: source.Next.Val}
	}

	if source.ParagraphPr != nil {
		cloned.ParagraphPr = sm.cloneParagraphProperties(source.ParagraphPr)
	}

	if source.RunPr != nil {
		cloned.RunPr = sm.cloneRunProperties(source.RunPr)
	}

	if source.TablePr != nil {
		cloned.TablePr = sm.cloneTableProperties(source.TablePr)
	}

	if source.TableRowPr != nil {
		cloned.TableRowPr = sm.cloneTableRowProperties(source.TableRowPr)
	}

	if source.TableCellPr != nil {
		cloned.TableCellPr = sm.cloneTableCellProperties(source.TableCellPr)
	}

	return cloned
}

// cloneParagraphProperties creates a deep copy of paragraph properties
func (sm *StyleManager) cloneParagraphProperties(source *ParagraphProperties) *ParagraphProperties {
	if source == nil {
		return nil
	}

	cloned := &ParagraphProperties{}

	if source.Spacing != nil {
		cloned.Spacing = &Spacing{
			Before:   source.Spacing.Before,
			After:    source.Spacing.After,
			Line:     source.Spacing.Line,
			LineRule: source.Spacing.LineRule,
		}
	}

	if source.Justification != nil {
		cloned.Justification = &Justification{
			Val: source.Justification.Val,
		}
	}

	if source.Indentation != nil {
		cloned.Indentation = &Indentation{
			FirstLine: source.Indentation.FirstLine,
			Left:      source.Indentation.Left,
			Right:     source.Indentation.Right,
		}
	}
	if source.ParagraphBorder != nil {
		cloned.ParagraphBorder = &ParagraphBorder{}
		if source.ParagraphBorder.Top != nil {
			cloned.ParagraphBorder.Top = &ParagraphBorderLine{
				Val:   source.ParagraphBorder.Top.Val,
				Color: source.ParagraphBorder.Top.Color,
				Sz:    source.ParagraphBorder.Top.Sz,
				Space: source.ParagraphBorder.Top.Space,
			}
		}
		if source.ParagraphBorder.Left != nil {
			cloned.ParagraphBorder.Left = &ParagraphBorderLine{
				Val:   source.ParagraphBorder.Left.Val,
				Color: source.ParagraphBorder.Left.Color,
				Sz:    source.ParagraphBorder.Left.Sz,
				Space: source.ParagraphBorder.Left.Space,
			}
		}
		if source.ParagraphBorder.Bottom != nil {
			cloned.ParagraphBorder.Bottom = &ParagraphBorderLine{
				Val:   source.ParagraphBorder.Bottom.Val,
				Color: source.ParagraphBorder.Bottom.Color,
				Sz:    source.ParagraphBorder.Bottom.Sz,
				Space: source.ParagraphBorder.Bottom.Space,
			}
		}
		if source.ParagraphBorder.Right != nil {
			cloned.ParagraphBorder.Right = &ParagraphBorderLine{
				Val:   source.ParagraphBorder.Right.Val,
				Color: source.ParagraphBorder.Right.Color,
				Sz:    source.ParagraphBorder.Right.Sz,
				Space: source.ParagraphBorder.Right.Space,
			}
		}
	}

	if source.Shading != nil {
		cloned.Shading = &Shading{
			Fill: source.Shading.Fill,
			Val:  source.Shading.Val,
		}
	}
	if source.KeepNext != nil {
		cloned.KeepNext = &KeepNext{}
	}

	if source.KeepLines != nil {
		cloned.KeepLines = &KeepLines{}
	}

	if source.PageBreak != nil {
		cloned.PageBreak = &PageBreak{}
	}

	if source.OutlineLevel != nil {
		cloned.OutlineLevel = &OutlineLevel{
			Val: source.OutlineLevel.Val,
		}
	}

	if source.SnapToGrid != nil {
		cloned.SnapToGrid = &SnapToGrid{
			Val: source.SnapToGrid.Val,
		}
	}

	return cloned
}

// cloneRunProperties creates a deep copy of run properties
func (sm *StyleManager) cloneRunProperties(source *RunProperties) *RunProperties {
	if source == nil {
		return nil
	}

	cloned := &RunProperties{}

	if source.Bold != nil {
		cloned.Bold = &Bold{}
	}

	if source.Italic != nil {
		cloned.Italic = &Italic{}
	}

	if source.Underline != nil {
		cloned.Underline = &Underline{Val: source.Underline.Val}
	}

	if source.Strike != nil {
		cloned.Strike = &Strike{}
	}

	if source.FontSize != nil {
		cloned.FontSize = &FontSize{Val: source.FontSize.Val}
	}

	if source.Color != nil {
		cloned.Color = &Color{Val: source.Color.Val}
	}

	if source.FontFamily != nil {
		cloned.FontFamily = &FontFamily{
			ASCII:    source.FontFamily.ASCII,
			EastAsia: source.FontFamily.EastAsia,
			HAnsi:    source.FontFamily.HAnsi,
			CS:       source.FontFamily.CS,
		}
	}

	if source.Highlight != nil {
		cloned.Highlight = &Highlight{Val: source.Highlight.Val}
	}

	return cloned
}

// cloneTableProperties creates a deep copy of table properties
func (sm *StyleManager) cloneTableProperties(source *TableProperties) *TableProperties {
	if source == nil {
		return nil
	}

	cloned := &TableProperties{}

	if source.TblInd != nil {
		cloned.TblInd = &TblIndent{
			W:    source.TblInd.W,
			Type: source.TblInd.Type,
		}
	}
	if source.TblBorders != nil {
		cloned.TblBorders = &TblBorders{}

		if source.TblBorders.Top != nil {
			cloned.TblBorders.Top = &TblBorder{
				Val:   source.TblBorders.Top.Val,
				Sz:    source.TblBorders.Top.Sz,
				Space: source.TblBorders.Top.Space,
				Color: source.TblBorders.Top.Color,
			}
		}

		if source.TblBorders.Left != nil {
			cloned.TblBorders.Left = &TblBorder{
				Val:   source.TblBorders.Left.Val,
				Sz:    source.TblBorders.Left.Sz,
				Space: source.TblBorders.Left.Space,
				Color: source.TblBorders.Left.Color,
			}
		}

		if source.TblBorders.Bottom != nil {
			cloned.TblBorders.Bottom = &TblBorder{
				Val:   source.TblBorders.Bottom.Val,
				Sz:    source.TblBorders.Bottom.Sz,
				Space: source.TblBorders.Bottom.Space,
				Color: source.TblBorders.Bottom.Color,
			}
		}

		if source.TblBorders.Right != nil {
			cloned.TblBorders.Right = &TblBorder{
				Val:   source.TblBorders.Right.Val,
				Sz:    source.TblBorders.Right.Sz,
				Space: source.TblBorders.Right.Space,
				Color: source.TblBorders.Right.Color,
			}
		}

		if source.TblBorders.InsideH != nil {
			cloned.TblBorders.InsideH = &TblBorder{
				Val:   source.TblBorders.InsideH.Val,
				Sz:    source.TblBorders.InsideH.Sz,
				Space: source.TblBorders.InsideH.Space,
				Color: source.TblBorders.InsideH.Color,
			}
		}

		if source.TblBorders.InsideV != nil {
			cloned.TblBorders.InsideV = &TblBorder{
				Val:   source.TblBorders.InsideV.Val,
				Sz:    source.TblBorders.InsideV.Sz,
				Space: source.TblBorders.InsideV.Space,
				Color: source.TblBorders.InsideV.Color,
			}
		}
	}

	if source.TblCellMar != nil {
		cloned.TblCellMar = &TblCellMargin{}

		if source.TblCellMar.Top != nil {
			cloned.TblCellMar.Top = &TblCellSpace{
				W:    source.TblCellMar.Top.W,
				Type: source.TblCellMar.Top.Type,
			}
		}

		if source.TblCellMar.Left != nil {
			cloned.TblCellMar.Left = &TblCellSpace{
				W:    source.TblCellMar.Left.W,
				Type: source.TblCellMar.Left.Type,
			}
		}

		if source.TblCellMar.Bottom != nil {
			cloned.TblCellMar.Bottom = &TblCellSpace{
				W:    source.TblCellMar.Bottom.W,
				Type: source.TblCellMar.Bottom.Type,
			}
		}

		if source.TblCellMar.Right != nil {
			cloned.TblCellMar.Right = &TblCellSpace{
				W:    source.TblCellMar.Right.W,
				Type: source.TblCellMar.Right.Type,
			}
		}
	}

	return cloned
}

// cloneTableRowProperties creates a deep copy of table row properties
func (sm *StyleManager) cloneTableRowProperties(source *TableRowProperties) *TableRowProperties {
	if source == nil {
		return nil
	}

	cloned := &TableRowProperties{}
	return cloned
}

// cloneTableCellProperties creates a deep copy of table cell properties
func (sm *StyleManager) cloneTableCellProperties(source *TableCellProperties) *TableCellProperties {
	if source == nil {
		return nil
	}

	cloned := &TableCellProperties{}
	return cloned
}

// GetStylesByType retrieves all styles of a given type
func (sm *StyleManager) GetStylesByType(styleType StyleType) []*Style {
	var styles []*Style
	for _, style := range sm.styles {
		if StyleType(style.Type) == styleType {
			styles = append(styles, style)
		}
	}
	return styles
}

// GetHeadingStyles retrieves all heading styles (levels 1-9)
func (sm *StyleManager) GetHeadingStyles() []*Style {
	var headingStyles []*Style
	for i := 1; i <= 9; i++ {
		styleID := fmt.Sprintf("Heading%d", i)
		if style := sm.GetStyle(styleID); style != nil {
			headingStyles = append(headingStyles, style)
		}
	}
	return headingStyles
}

// ApplyStyleToXML converts a style to an XML structure (for document integration)
func (sm *StyleManager) ApplyStyleToXML(styleID string) (map[string]interface{}, error) {
	style := sm.GetStyleWithInheritance(styleID)
	if style == nil {
		return nil, fmt.Errorf("style %s not found", styleID)
	}

	result := make(map[string]interface{})
	result["styleId"] = style.StyleID
	result["type"] = style.Type

	if style.ParagraphPr != nil {
		result["paragraphProperties"] = convertParagraphPropertiesToMap(style.ParagraphPr)
	}

	if style.RunPr != nil {
		result["runProperties"] = convertRunPropertiesToMap(style.RunPr)
	}

	return result, nil
}

// convertParagraphPropertiesToMap converts paragraph properties to a map
func convertParagraphPropertiesToMap(props *ParagraphProperties) map[string]interface{} {
	result := make(map[string]interface{})

	if props.Spacing != nil {
		spacing := make(map[string]string)
		if props.Spacing.Before != "" {
			spacing["before"] = props.Spacing.Before
		}
		if props.Spacing.After != "" {
			spacing["after"] = props.Spacing.After
		}
		if props.Spacing.Line != "" {
			spacing["line"] = props.Spacing.Line
		}
		if props.Spacing.LineRule != "" {
			spacing["lineRule"] = props.Spacing.LineRule
		}
		result["spacing"] = spacing
	}

	if props.Justification != nil {
		result["justification"] = props.Justification.Val
	}

	if props.Indentation != nil {
		indentation := make(map[string]string)
		if props.Indentation.FirstLine != "" {
			indentation["firstLine"] = props.Indentation.FirstLine
		}
		if props.Indentation.Left != "" {
			indentation["left"] = props.Indentation.Left
		}
		if props.Indentation.Right != "" {
			indentation["right"] = props.Indentation.Right
		}
		result["indentation"] = indentation
	}

	if props.OutlineLevel != nil {
		result["outlineLevel"] = props.OutlineLevel.Val
	}

	return result
}

// convertRunPropertiesToMap converts run properties to a map
func convertRunPropertiesToMap(props *RunProperties) map[string]interface{} {
	result := make(map[string]interface{})

	if props.Bold != nil {
		result["bold"] = true
	}

	if props.Italic != nil {
		result["italic"] = true
	}

	if props.Underline != nil {
		result["underline"] = props.Underline.Val
	}

	if props.Strike != nil {
		result["strike"] = true
	}

	if props.FontSize != nil {
		result["fontSize"] = props.FontSize.Val
	}

	if props.Color != nil {
		result["color"] = props.Color.Val
	}

	if props.FontFamily != nil {
		fontFamily := make(map[string]string)
		if props.FontFamily.ASCII != "" {
			fontFamily["ascii"] = props.FontFamily.ASCII
		}
		if props.FontFamily.EastAsia != "" {
			fontFamily["eastAsia"] = props.FontFamily.EastAsia
		}
		if props.FontFamily.HAnsi != "" {
			fontFamily["hAnsi"] = props.FontFamily.HAnsi
		}
		if props.FontFamily.CS != "" {
			fontFamily["cs"] = props.FontFamily.CS
		}
		result["fontFamily"] = fontFamily
	}

	if props.Highlight != nil {
		result["highlight"] = props.Highlight.Val
	}

	return result
}

// ParseStylesFromXML parses styles from XML data, replacing all existing styles
func (sm *StyleManager) ParseStylesFromXML(xmlData []byte) error {
	type stylesXML struct {
		XMLName xml.Name `xml:"w:styles"`
		Styles  []Style  `xml:"w:style"`
	}

	var styles stylesXML
	if err := xml.Unmarshal(xmlData, &styles); err != nil {
		return fmt.Errorf("failed to parse styles XML: %v", err)
	}

	sm.styles = make(map[string]*Style)

	for i := range styles.Styles {
		sm.AddStyle(&styles.Styles[i])
	}

	return nil
}

// MergeStylesFromXML merges styles from XML data, keeping existing styles
func (sm *StyleManager) MergeStylesFromXML(xmlData []byte) error {
	type stylesXML struct {
		XMLName xml.Name `xml:"w:styles"`
		Styles  []Style  `xml:"w:style"`
	}

	var styles stylesXML
	if err := xml.Unmarshal(xmlData, &styles); err != nil {
		return fmt.Errorf("failed to parse styles XML: %v", err)
	}

	for i := range styles.Styles {
		if !sm.StyleExists(styles.Styles[i].StyleID) {
			sm.AddStyle(&styles.Styles[i])
		}
	}

	return nil
}

// LoadStylesFromDocument loads styles from existing document XML, preserving original styles
func (sm *StyleManager) LoadStylesFromDocument(xmlData []byte) error {
	if len(xmlData) == 0 {
		sm.initializePredefinedStyles()
		return nil
	}

	if err := sm.ParseStylesFromXML(xmlData); err != nil {
		sm.initializePredefinedStyles()
		return fmt.Errorf("failed to parse existing styles, using default styles: %v", err)
	}

	if !sm.StyleExists("Normal") {
		sm.addNormalStyle()
	}

	headingStyles := []string{"Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Heading7", "Heading8", "Heading9"}
	for _, styleID := range headingStyles {
		if !sm.StyleExists(styleID) {
			sm.addHeadingStyles()
			break
		}
	}

	return nil
}

// addTableStyles adds predefined table styles
func (sm *StyleManager) addTableStyles() {
	normalTable := &Style{
		Type:    string(StyleTypeTable),
		StyleID: "a1",
		Default: true,
		Name: &StyleName{
			Val: "Normal Table",
		},
		TablePr: &TableProperties{
			TblInd: &TblIndent{
				W:    "0",
				Type: "dxa",
			},
			TblCellMar: &TblCellMargin{
				Top: &TblCellSpace{
					W:    "0",
					Type: "dxa",
				},
				Left: &TblCellSpace{
					W:    "108",
					Type: "dxa",
				},
				Bottom: &TblCellSpace{
					W:    "0",
					Type: "dxa",
				},
				Right: &TblCellSpace{
					W:    "108",
					Type: "dxa",
				},
			},
		},
	}
	sm.AddStyle(normalTable)

	tableGrid := &Style{
		Type:    string(StyleTypeTable),
		StyleID: "ab",
		Name: &StyleName{
			Val: "Table Grid",
		},
		BasedOn: &BasedOn{
			Val: "a1",
		},
		TablePr: &TableProperties{
			TblBorders: &TblBorders{
				Top: &TblBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Left: &TblBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Bottom: &TblBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Right: &TblBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				InsideH: &TblBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				InsideV: &TblBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
			},
		},
	}
	sm.AddStyle(tableGrid)
}
