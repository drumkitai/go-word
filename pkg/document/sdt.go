// Package document provides Word document SDT (Structured Document Tag) functionality
package document

import (
	"encoding/xml"
	"fmt"
)

// SDT structured document tag for special functionality like table of contents
type SDT struct {
	XMLName    xml.Name       `xml:"w:sdt"`
	Properties *SDTProperties `xml:"w:sdtPr"`
	EndPr      *SDTEndPr      `xml:"w:sdtEndPr,omitempty"`
	Content    *SDTContent    `xml:"w:sdtContent"`
}

// ElementType returns the SDT element type
func (s *SDT) ElementType() string {
	return "sdt"
}

// SDTProperties SDT properties
type SDTProperties struct {
	XMLName     xml.Name        `xml:"w:sdtPr"`
	RunPr       *RunProperties  `xml:"w:rPr,omitempty"`
	ID          *SDTID          `xml:"w:id,omitempty"`
	Color       *SDTColor       `xml:"w15:color,omitempty"`
	DocPartObj  *DocPartObj     `xml:"w:docPartObj,omitempty"`
	Placeholder *SDTPlaceholder `xml:"w:placeholder,omitempty"`
}

// SDTEndPr SDT end properties
type SDTEndPr struct {
	XMLName xml.Name       `xml:"w:sdtEndPr"`
	RunPr   *RunProperties `xml:"w:rPr,omitempty"`
}

// SDTContent SDT content
type SDTContent struct {
	XMLName  xml.Name      `xml:"w:sdtContent"`
	Elements []interface{} `xml:"-"`
}

// MarshalXML custom XML serialization
func (s *SDTContent) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, element := range s.Elements {
		if err := e.Encode(element); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// SDTID SDT identifier
type SDTID struct {
	XMLName xml.Name `xml:"w:id"`
	Val     string   `xml:"w:val,attr"`
}

// SDTColor SDT color
type SDTColor struct {
	XMLName xml.Name `xml:"w15:color"`
	Val     string   `xml:"w:val,attr"`
}

// DocPartObj document part object
type DocPartObj struct {
	XMLName        xml.Name        `xml:"w:docPartObj"`
	DocPartGallery *DocPartGallery `xml:"w:docPartGallery,omitempty"`
	DocPartUnique  *DocPartUnique  `xml:"w:docPartUnique,omitempty"`
}

// DocPartGallery 文档部件库
type DocPartGallery struct {
	XMLName xml.Name `xml:"w:docPartGallery"`
	Val     string   `xml:"w:val,attr"`
}

// DocPartUnique 文档部件唯一标识
type DocPartUnique struct {
	XMLName xml.Name `xml:"w:docPartUnique"`
}

// SDTPlaceholder SDT占位符
type SDTPlaceholder struct {
	XMLName xml.Name `xml:"w:placeholder"`
	DocPart *DocPart `xml:"w:docPart,omitempty"`
}

// DocPart 文档部件
type DocPart struct {
	XMLName xml.Name `xml:"w:docPart"`
	Val     string   `xml:"w:val,attr"`
}

// Tab 制表符
type Tab struct {
	XMLName xml.Name `xml:"w:tab"`
}

// CreateTOCSDT 创建目录SDT结构
func (d *Document) CreateTOCSDT(title string, maxLevel int) *SDT {
	sdt := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "宋体"},
				FontSize:   &FontSize{Val: "21"},
			},
			ID:    &SDTID{Val: "147476628"},
			Color: &SDTColor{Val: "DBDBDB"},
			DocPartObj: &DocPartObj{
				DocPartGallery: &DocPartGallery{Val: "Table of Contents"},
				DocPartUnique:  &DocPartUnique{},
			},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontSize: &FontSize{Val: "20"},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{},
		},
	}

	// add TOC title paragraph
	titlePara := &Paragraph{
		Properties: &ParagraphProperties{
			Spacing: &Spacing{
				Before: "0",
				After:  "0",
				Line:   "240",
			},
			Indentation: &Indentation{
				Left:      "0",
				Right:     "0",
				FirstLine: "0",
			},
			Justification: &Justification{Val: "center"},
		},
		Runs: []Run{
			{
				Text: Text{Content: title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: "宋体"},
					FontSize:   &FontSize{Val: "21"},
				},
			},
		},
	}

	// 添加书签开始 - 使用已有的BookmarkStart类型
	bookmarkStart := &BookmarkStart{
		ID:   "0",
		Name: "_Toc11693_WPSOffice_Type3",
	}

	sdt.Content.Elements = append(sdt.Content.Elements, bookmarkStart, titlePara)

	return sdt
}

// AddTOCEntry 向目录SDT添加条目
func (sdt *SDT) AddTOCEntry(text string, level int, pageNum int, entryID string) {
	// 确定目录样式ID (13=toc 1, 14=toc 2, 15=toc 3等)
	styleVal := fmt.Sprintf("%d", 12+level)

	// 创建目录条目段落
	entryPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: styleVal},
			Tabs: &Tabs{
				Tabs: []TabDef{
					{
						Val:    "right",
						Leader: "dot",
						Pos:    "8640",
					},
				},
			},
		},
		Runs: []Run{},
	}

	// 创建内嵌的SDT用于占位符文本
	placeholderSDT := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "Calibri"},
				FontSize:   &FontSize{Val: "22"},
			},
			ID: &SDTID{Val: entryID},
			Placeholder: &SDTPlaceholder{
				DocPart: &DocPart{Val: generatePlaceholderGUID(level)},
			},
			Color: &SDTColor{Val: "509DF3"},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "Calibri"},
				FontSize:   &FontSize{Val: "22"},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{
				Run{
					Text: Text{Content: text},
				},
			},
		},
	}

	// 将占位符SDT添加到段落中
	sdt.Content.Elements = append(sdt.Content.Elements, placeholderSDT)

	// 创建包含制表符和页码的文本Run
	tabRun := Run{
		Text: Text{Content: "\t"},
	}

	pageRun := Run{
		Text: Text{Content: fmt.Sprintf("%d", pageNum)},
	}

	entryPara.Runs = append(entryPara.Runs, tabRun, pageRun)

	// 添加段落到SDT内容中
	sdt.Content.Elements = append(sdt.Content.Elements, entryPara)
}

// generatePlaceholderGUID 生成占位符GUID
func generatePlaceholderGUID(level int) string {
	guids := map[int]string{
		1: "{b5fdec38-8301-4b26-9716-d8b31c00c718}",
		2: "{a500490c-aaae-4252-8340-aa59729b9870}",
		3: "{d7310822-77d9-4e43-95e1-4649f1e215b3}",
	}

	if guid, exists := guids[level]; exists {
		return guid
	}
	return "{b5fdec38-8301-4b26-9716-d8b31c00c718}" // 默认使用1级
}

// FinalizeTOCSDT 完成目录SDT构建
func (sdt *SDT) FinalizeTOCSDT() {
	// 添加书签结束 - 使用已有的BookmarkEnd类型
	bookmarkEnd := &BookmarkEnd{
		ID: "0",
	}
	sdt.Content.Elements = append(sdt.Content.Elements, bookmarkEnd)
}
