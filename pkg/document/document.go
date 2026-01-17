// Package document provides core document manipulation functionality
package document

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/drumkitai/go-word/pkg/style"
)

// Document represents a Word document
type Document struct {
	// main document content
	Body *Body
	// document relationships
	relationships *Relationships
	// document-level relationships (for headers/footers, etc)
	documentRelationships *Relationships
	// content types
	contentTypes *ContentTypes
	// style manager
	styleManager *style.StyleManager
	// temporary storage for document parts
	parts map[string][]byte
	// image ID counter, ensure each image has unique ID
	nextImageID int
}

// Body represents the document body
type Body struct {
	XMLName  xml.Name      `xml:"w:body"`
	Elements []interface{} `xml:"-"` // not serialized; use custom method
}

// MarshalXML custom XML serialization, output elements in order
func (b *Body) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// start element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// separate SectionProperties and other elements
	var sectPr *SectionProperties
	var otherElements []interface{}

	for _, element := range b.Elements {
		if sp, ok := element.(*SectionProperties); ok {
			sectPr = sp // save last SectionProperties
		} else {
			otherElements = append(otherElements, element)
		}
	}

	// serialize other elements first (paragraphs, tables, etc)
	for _, element := range otherElements {
		if err := e.Encode(element); err != nil {
			return err
		}
	}

	// serialize SectionProperties last (if exists)
	if sectPr != nil {
		if err := e.Encode(sectPr); err != nil {
			return err
		}
	}

	// end element
	return e.EncodeToken(start.End())
}

// BodyElement document body element interface
type BodyElement interface {
	ElementType() string
}

// ElementType returns the paragraph element type
func (p *Paragraph) ElementType() string {
	return "paragraph"
}

// ElementType returns the table element type
func (t *Table) ElementType() string {
	return "table"
}

// Paragraph represents a paragraph
type Paragraph struct {
	XMLName    xml.Name             `xml:"w:p"`
	Properties *ParagraphProperties `xml:"w:pPr,omitempty"`
	Runs       []Run                `xml:"w:r"`
}

// ParagraphProperties paragraph properties
type ParagraphProperties struct {
	XMLName             xml.Name             `xml:"w:pPr"`
	ParagraphStyle      *ParagraphStyle      `xml:"w:pStyle,omitempty"`
	NumberingProperties *NumberingProperties `xml:"w:numPr,omitempty"`
	ParagraphBorder     *ParagraphBorder     `xml:"w:pBdr,omitempty"`
	Tabs                *Tabs                `xml:"w:tabs,omitempty"`
	SnapToGrid          *SnapToGrid          `xml:"w:snapToGrid,omitempty"` // 网格对齐设置
	Spacing             *Spacing             `xml:"w:spacing,omitempty"`
	Indentation         *Indentation         `xml:"w:ind,omitempty"`
	Justification       *Justification       `xml:"w:jc,omitempty"`
	KeepNext            *KeepNext            `xml:"w:keepNext,omitempty"`        // 与下一段落保持在一起
	KeepLines           *KeepLines           `xml:"w:keepLines,omitempty"`       // 段落中的行保持在一起
	PageBreakBefore     *PageBreakBefore     `xml:"w:pageBreakBefore,omitempty"` // 段前分页
	WidowControl        *WidowControl        `xml:"w:widowControl,omitempty"`    // 孤行控制
	OutlineLevel        *OutlineLevel        `xml:"w:outlineLvl,omitempty"`      // 大纲级别
}

// SnapToGrid grid alignment setting
// When set to "0" or "false", grid alignment is disabled, allowing custom line spacing to take effect
// Note: This type is intentionally duplicated in the style package to allow independent package usage
type SnapToGrid struct {
	XMLName xml.Name `xml:"w:snapToGrid"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// ParagraphBorder paragraph border
type ParagraphBorder struct {
	XMLName xml.Name             `xml:"w:pBdr"`
	Top     *ParagraphBorderLine `xml:"w:top,omitempty"`
	Left    *ParagraphBorderLine `xml:"w:left,omitempty"`
	Bottom  *ParagraphBorderLine `xml:"w:bottom,omitempty"`
	Right   *ParagraphBorderLine `xml:"w:right,omitempty"`
}

// ParagraphBorderLine paragraph border line
type ParagraphBorderLine struct {
	Val   string `xml:"w:val,attr"`
	Color string `xml:"w:color,attr"`
	Sz    string `xml:"w:sz,attr"`
	Space string `xml:"w:space,attr"`
}

// Spacing spacing setting
type Spacing struct {
	XMLName  xml.Name `xml:"w:spacing"`
	Before   string   `xml:"w:before,attr,omitempty"`
	After    string   `xml:"w:after,attr,omitempty"`
	Line     string   `xml:"w:line,attr,omitempty"`
	LineRule string   `xml:"w:lineRule,attr,omitempty"`
}

// Justification alignment setting
type Justification struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// KeepNext keep next paragraph together
type KeepNext struct {
	XMLName xml.Name `xml:"w:keepNext"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// KeepLines keep lines in paragraph together
type KeepLines struct {
	XMLName xml.Name `xml:"w:keepLines"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// PageBreakBefore page break before
type PageBreakBefore struct {
	XMLName xml.Name `xml:"w:pageBreakBefore"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// WidowControl widow control
type WidowControl struct {
	XMLName xml.Name `xml:"w:widowControl"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// OutlineLevel outline level
type OutlineLevel struct {
	XMLName xml.Name `xml:"w:outlineLvl"`
	Val     string   `xml:"w:val,attr"`
}

// Run represents a text run
type Run struct {
	XMLName    xml.Name        `xml:"w:r"`
	Properties *RunProperties  `xml:"w:rPr,omitempty"`
	Text       Text            `xml:"w:t,omitempty"`
	Break      *Break          `xml:"w:br,omitempty"` // 分页符 / Page break
	Drawing    *DrawingElement `xml:"w:drawing,omitempty"`
	FieldChar  *FieldChar      `xml:"w:fldChar,omitempty"`
	InstrText  *InstrText      `xml:"w:instrText,omitempty"`
}

// MarshalXML custom Run XML serialization
// This method ensures that only non-empty elements are serialized, especially for Drawing elements
func (r *Run) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// start Run element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// serializeRunProperties (if exists)
	if r.Properties != nil {
		if err := e.EncodeElement(r.Properties, xml.StartElement{Name: xml.Name{Local: "w:rPr"}}); err != nil {
			return err
		}
	}

	// serializeText (only when there is content)
	// This is a critical fix: avoid serializing empty Text elements
	if r.Text.Content != "" {
		if err := e.EncodeElement(r.Text, xml.StartElement{Name: xml.Name{Local: "w:t"}}); err != nil {
			return err
		}
	}

	// serializeBreak (if exists)
	if r.Break != nil {
		if err := e.EncodeElement(r.Break, xml.StartElement{Name: xml.Name{Local: "w:br"}}); err != nil {
			return err
		}
	}

	// serializeDrawing (if exists)
	if r.Drawing != nil {
		if err := e.EncodeElement(r.Drawing, xml.StartElement{Name: xml.Name{Local: "w:drawing"}}); err != nil {
			return err
		}
	}

	// serializeFieldChar (if exists)
	if r.FieldChar != nil {
		if err := e.EncodeElement(r.FieldChar, xml.StartElement{Name: xml.Name{Local: "w:fldChar"}}); err != nil {
			return err
		}
	}

	// serializeInstrText (if exists)
	if r.InstrText != nil {
		if err := e.EncodeElement(r.InstrText, xml.StartElement{Name: xml.Name{Local: "w:instrText"}}); err != nil {
			return err
		}
	}

	// end Run element
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// RunProperties text properties
// Note: The field order must conform to the OpenXML standard, w:rFonts must be before w:color
type RunProperties struct {
	XMLName    xml.Name    `xml:"w:rPr"`
	FontFamily *FontFamily `xml:"w:rFonts,omitempty"`
	Bold       *Bold       `xml:"w:b,omitempty"`
	BoldCs     *BoldCs     `xml:"w:bCs,omitempty"`
	Italic     *Italic     `xml:"w:i,omitempty"`
	ItalicCs   *ItalicCs   `xml:"w:iCs,omitempty"`
	Underline  *Underline  `xml:"w:u,omitempty"`
	Strike     *Strike     `xml:"w:strike,omitempty"`
	Color      *Color      `xml:"w:color,omitempty"`
	FontSize   *FontSize   `xml:"w:sz,omitempty"`
	FontSizeCs *FontSizeCs `xml:"w:szCs,omitempty"`
	Highlight  *Highlight  `xml:"w:highlight,omitempty"`
}

// Bold bold
type Bold struct {
	XMLName xml.Name `xml:"w:b"`
}

// BoldCs complex script bold
type BoldCs struct {
	XMLName xml.Name `xml:"w:bCs"`
}

// Italic italic
type Italic struct {
	XMLName xml.Name `xml:"w:i"`
}

// ItalicCs complex script italic
type ItalicCs struct {
	XMLName xml.Name `xml:"w:iCs"`
}

// Underline underline
type Underline struct {
	XMLName xml.Name `xml:"w:u"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// Strike strike
type Strike struct {
	XMLName xml.Name `xml:"w:strike"`
}

// FontSize font size
type FontSize struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

// FontSizeCs complex script font size
type FontSizeCs struct {
	XMLName xml.Name `xml:"w:szCs"`
	Val     string   `xml:"w:val,attr"`
}

// Color color
type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

// Highlight highlight color
type Highlight struct {
	XMLName xml.Name `xml:"w:highlight"`
	Val     string   `xml:"w:val,attr"`
}

// Text text content
type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Content string   `xml:",chardata"`
}

// Break page break
// Break represents page breaks in Word documents
type Break struct {
	XMLName xml.Name `xml:"w:br"`
	Type    string   `xml:"w:type,attr,omitempty"` // "page" indicates a page break
}

// Relationships
type Relationships struct {
	XMLName       xml.Name       `xml:"Relationships"`
	Xmlns         string         `xml:"xmlns,attr"`
	Relationships []Relationship `xml:"Relationship"`
}

// Relationship
type Relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// ContentTypes
type ContentTypes struct {
	XMLName   xml.Name   `xml:"Types"`
	Xmlns     string     `xml:"xmlns,attr"`
	Defaults  []Default  `xml:"Default"`
	Overrides []Override `xml:"Override"`
}

// Default
type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// Override
type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// FontFamily
type FontFamily struct {
	XMLName  xml.Name `xml:"w:rFonts"`
	ASCII    string   `xml:"w:ascii,attr,omitempty"`
	HAnsi    string   `xml:"w:hAnsi,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	CS       string   `xml:"w:cs,attr,omitempty"`
	Hint     string   `xml:"w:hint,attr,omitempty"`
}

// TextFormat
type TextFormat struct {
	Bold       bool
	Italic     bool
	FontSize   int
	FontColor  string
	FontFamily string
	FontName   string
	Underline  bool
	Strike     bool
	Highlight  string
}

// AlignmentType
type AlignmentType string

const (
	// AlignLeft
	AlignLeft AlignmentType = "left"
	// AlignCenter
	AlignCenter AlignmentType = "center"
	// AlignRight
	AlignRight AlignmentType = "right"
	// AlignJustify
	AlignJustify AlignmentType = "both"
)

// SpacingConfig
type SpacingConfig struct {
	LineSpacing     float64
	BeforePara      int
	AfterPara       int
	FirstLineIndent int
}

// Indentation
type Indentation struct {
	XMLName   xml.Name `xml:"w:ind"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
	Left      string   `xml:"w:left,attr,omitempty"`
	Right     string   `xml:"w:right,attr,omitempty"`
}

// Tabs
type Tabs struct {
	XMLName xml.Name `xml:"w:tabs"`
	Tabs    []TabDef `xml:"w:tab"`
}

// TabDef
type TabDef struct {
	XMLName xml.Name `xml:"w:tab"`
	Val     string   `xml:"w:val,attr"`
	Leader  string   `xml:"w:leader,attr,omitempty"`
	Pos     string   `xml:"w:pos,attr"`
}

// ParagraphStyle
type ParagraphStyle struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

// NumberingProperties
type NumberingProperties struct {
	XMLName xml.Name `xml:"w:numPr"`
	ILevel  *ILevel  `xml:"w:ilvl,omitempty"`
	NumID   *NumID   `xml:"w:numId,omitempty"`
}

// ILevel
type ILevel struct {
	XMLName xml.Name `xml:"w:ilvl"`
	Val     string   `xml:"w:val,attr"`
}

// NumID
type NumID struct {
	XMLName xml.Name `xml:"w:numId"`
	Val     string   `xml:"w:val,attr"`
}

// New
func New() *Document {
	Debugf("Creating new document")

	doc := &Document{
		Body: &Body{
			Elements: make([]interface{}, 0),
		},
		styleManager: style.NewStyleManager(),
		parts:        make(map[string][]byte),
		nextImageID:  0, // initialize image ID counter, start from 0
		documentRelationships: &Relationships{
			Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{},
		},
	}

	// initialize document structure
	doc.initializeStructure()

	return doc
}

// Open
//
// filename is the path to the .docx file to open.
// This function will parse the entire document structure, including text content, formatting, and properties.
//
// If the file does not exist, is formatted incorrectly, or parsing fails, the appropriate error will be returned.
//
// Example:
//
//	doc, err := document.Open("existing.docx")
//	if err != nil {
//		log.Fatal(err)
//	}
//
//	// print all paragraph content
//	for i, para := range doc.Body.Paragraphs {
//		fmt.Printf("Paragraph %d: ", i+1)
//		for _, run := range para.Runs {
//			fmt.Print(run.Text.Content)
//		}
//		fmt.Println()
//	}
func Open(filename string) (*Document, error) {
	Infof("Opening document: %s", filename)

	reader, err := zip.OpenReader(filename)
	if err != nil {
		Errorf("Cannot open file: %s", filename)
		return nil, WrapErrorWithContext("open_file", err, filename)
	}
	defer reader.Close()

	doc, err := openFromZipReader(&reader.Reader, filename)
	if err != nil {
		return nil, err
	}

	Infof("Successfully opened document: %s", filename)
	return doc, nil
}

func OpenFromMemory(readCloser io.ReadCloser) (*Document, error) {
	defer readCloser.Close()
	Infof("Opening document")

	fileData, err := io.ReadAll(readCloser)
	if err != nil {
		return nil, fmt.Errorf("failed to read file content: %w", err)
	}

	readerAt := bytes.NewReader(fileData)
	size := int64(len(fileData))
	reader, err := zip.NewReader(readerAt, size)
	if err != nil {
		Errorf("cannot open file")
		return nil, WrapErrorWithContext("open_file", err, "")
	}

	doc, err := openFromZipReader(reader, "memory")
	if err != nil {
		return nil, err
	}

	Infof("Successfully opened document")
	return doc, nil
}

func openFromZipReader(zipReader *zip.Reader, filename string) (*Document, error) {
	doc := &Document{
		parts: make(map[string][]byte),
		documentRelationships: &Relationships{
			Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{},
		},
		nextImageID: 0, // initialize image ID counter, start from 0
	}

	// read all file parts
	for _, file := range zipReader.File {
		rc, err := file.Open()
		if err != nil {
			Errorf("cannot open file part: %s", file.Name)
			return nil, WrapErrorWithContext("open_part", err, file.Name)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			Errorf("cannot read file part: %s", file.Name)
			return nil, WrapErrorWithContext("read_part", err, file.Name)
		}

		doc.parts[file.Name] = data
		Debugf("read file part: %s (%d bytes)", file.Name, len(data))
	}

	// initialize style manager
	doc.styleManager = style.NewStyleManager()

	// parse content types
	if err := doc.parseContentTypes(); err != nil {
		Debugf("failed to parse content types, using default values: %v", err)
		// 如果解析失败，使用默认值
		doc.contentTypes = &ContentTypes{
			Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
			Defaults: []Default{
				{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
				{Extension: "xml", ContentType: "application/xml"},
			},
			Overrides: []Override{
				{PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
				{PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
			},
		}
	}

	// parse relationships
	if err := doc.parseRelationships(); err != nil {
		Debugf("failed to parse relationships, using default values: %v", err)
		// if parsing fails, use default values
		doc.relationships = &Relationships{
			Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{
				{
					ID:     "rId1",
					Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
					Target: "word/document.xml",
				},
			},
		}
	}

	// parse main document
	if err := doc.parseDocument(); err != nil {
		Errorf("failed to parse document: %s", filename)
		return nil, WrapErrorWithContext("parse_document", err, filename)
	}

	// parse styles file
	if err := doc.parseStyles(); err != nil {
		Debugf("failed to parse styles, using default styles: %v", err)
		// if styles parsing fails, reinitialize to default styles
		doc.styleManager = style.NewStyleManager()
	}

	// parse document relationships (including relationships for images, etc.)
	if err := doc.parseDocumentRelationships(); err != nil {
		Debugf("failed to parse document relationships, using default values: %v", err)
		// if parsing fails, keep the initialized empty relationship list
	}

	// update nextImageID counter based on existing image relationships
	doc.updateNextImageID()

	return doc, nil

}

// Save
//
// filename is the path to the file to save, including the file name and extension.
// If the directory does not exist, the necessary directory structure will be automatically created.
//
// The saving process includes serializing all document content, compressing to ZIP format,
// and writing to the file system.
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("example content")
//
//	// save to current directory
//	err := doc.Save("example.docx")
//
//	// save to subdirectory (directory will be automatically created)
//	err = doc.Save("output/documents/example.docx")
//
//	if err != nil {
//		log.Fatal(err)
//	}
func (d *Document) Save(filename string) error {
	Infof("Saving document: %s", filename)

	// ensure directory exists
	dir := filepath.Dir(filename)
	if err := os.MkdirAll(dir, 0755); err != nil {
		Errorf("cannot create directory: %s", dir)
		return WrapErrorWithContext("create_dir", err, dir)
	}

	// create file
	file, err := os.Create(filename)
	if err != nil {
		Errorf("cannot create file: %s", filename)
		return WrapErrorWithContext("create_file", err, filename)
	}
	defer file.Close()

	// create ZIP writer
	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// serialize main document
	if err := d.serializeDocument(); err != nil {
		Errorf("failed to serialize document")
		return WrapError("serialize_document", err)
	}

	// serialize styles
	if err := d.serializeStyles(); err != nil {
		Errorf("failed to serialize styles")
		return WrapError("serialize_styles", err)
	}

	// serialize content types
	d.serializeContentTypes()

	// serialize relationships
	d.serializeRelationships()

	// serialize document relationships
	d.serializeDocumentRelationships()

	// write all parts
	for name, data := range d.parts {
		writer, err := zipWriter.Create(name)
		if err != nil {
			Errorf("cannot create ZIP entry: %s", name)
			return WrapErrorWithContext("create_zip_entry", err, name)
		}

		if _, err := writer.Write(data); err != nil {
			Errorf("cannot write ZIP entry: %s", name)
			return WrapErrorWithContext("write_zip_entry", err, name)
		}

		Debugf("wrote ZIP entry: %s (%d bytes)", name, len(data))
	}

	Infof("successfully saved document: %s", filename)
	return nil
}

// AddParagraph
//
// text is the text content of the paragraph. The paragraph will use the default format,
// and the format and properties can be set later using the returned Paragraph pointer.
//
// Returns a pointer to the newly created paragraph, which can be used to further format.
//
// Example:
//
//	doc := document.New()
//
//	// add a normal paragraph
//	para := doc.AddParagraph("this is a paragraph")
//
//	// set paragraph properties
//	para.SetAlignment(document.AlignCenter)
//	para.SetSpacing(&document.SpacingConfig{
//		LineSpacing: 1.5,
//		BeforePara:  12,
//	})
func (d *Document) AddParagraph(text string) *Paragraph {
	Debugf("add paragraph: %s", text)
	p := &Paragraph{
		Runs: []Run{
			{
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// AddFormattedParagraph
//
// text is the text content of the paragraph.
// format specifies the text format, if nil the default format will be used.
//
// Example:
//
//	doc := document.New()
//
//	// create format configuration
//	titleFormat := &document.TextFormat{
//		Bold:      true,
//		FontSize:  18,
//		FontColor: "FF0000", // red
//		FontName:  "微软雅黑",
//	}
//
//	// add formatted title
//	title := doc.AddFormattedParagraph("document title", titleFormat)
//	title.SetAlignment(document.AlignCenter)
func (d *Document) AddFormattedParagraph(text string, format *TextFormat) *Paragraph {
	Debugf("add formatted paragraph: %s", text)

	// create run properties
	runProps := &RunProperties{}

	if format != nil {
		// compatible with FontFamily and FontName fields
		fontName := ""
		if format.FontFamily != "" {
			fontName = format.FontFamily
		} else if format.FontName != "" { // backward compatible with example code
			fontName = format.FontName
		}
		if fontName != "" {
			runProps.FontFamily = &FontFamily{ // set all related fields, ensure test and render consistency
				ASCII:    fontName,
				HAnsi:    fontName,
				EastAsia: fontName,
				CS:       fontName,
			}
		}

		if format.Bold {
			runProps.Bold = &Bold{}
		}

		if format.Italic {
			runProps.Italic = &Italic{}
		}

		if format.FontColor != "" {
			// ensure color format is correct (remove # prefix)
			color := strings.TrimPrefix(format.FontColor, "#")
			runProps.Color = &Color{Val: color}
		}

		if format.FontSize > 0 {
			// font size in Word is in half points, so it needs to be multiplied by 2
			runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
		}
		if format.Underline {
			runProps.Underline = &Underline{Val: "single"} // default single underline for underline
		}

		if format.Strike {
			runProps.Strike = &Strike{} // add strikethrough for strike
		}

		if format.Highlight != "" {
			runProps.Highlight = &Highlight{Val: format.Highlight}
		}
	}

	p := &Paragraph{
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

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// SetAlignment
//
// alignment specifies the alignment type, supported values:
//   - AlignLeft: left alignment (default)
//   - AlignCenter: center alignment
//   - AlignRight: right alignment
//   - AlignJustify: justify alignment
//
// Example:
//
//	para := doc.AddParagraph("centered title")
//	para.SetAlignment(document.AlignCenter)
//
//	para2 := doc.AddParagraph("right aligned text")
//	para2.SetAlignment(document.AlignRight)
func (p *Paragraph) SetAlignment(alignment AlignmentType) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	p.Properties.Justification = &Justification{Val: string(alignment)}
	Debugf("set paragraph alignment: %s", alignment)
}

// SetSpacing
//
// config contains various spacing settings, if nil no settings will be applied.
// supported settings:
//   - LineSpacing: line spacing (e.g. 1.5 for 1.5x line spacing)
//   - BeforePara: before paragraph spacing (in points)
//   - AfterPara: after paragraph spacing (in points)
//   - FirstLineIndent: first line indentation (in points)
//
// note: spacing values will be automatically converted to TWIPs (1 point = 20 TWIPs)
//
// Example:
//
//	para := doc.AddParagraph("paragraph with spacing")
//
//	// set complex spacing
//	para.SetSpacing(&document.SpacingConfig{
//		LineSpacing:     1.5, // 1.5x line spacing
//		BeforePara:      12,  // before paragraph 12 points
//		AfterPara:       6,   // after paragraph 6 points
//		FirstLineIndent: 24,  // first line indentation 24 points
//	})
//
//	// set only line spacing
//	para2 := doc.AddParagraph("double line spacing")
//	para2.SetSpacing(&document.SpacingConfig{
//		LineSpacing: 2.0,
//	})
func (p *Paragraph) SetSpacing(config *SpacingConfig) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if config != nil {
		spacing := &Spacing{}

		if config.BeforePara > 0 {
			// convert to TWIPs (1/20 points)
			spacing.Before = strconv.Itoa(config.BeforePara * 20)
		}

		if config.AfterPara > 0 {
			// convert to TWIPs (1/20 points)
			spacing.After = strconv.Itoa(config.AfterPara * 20)
		}

		if config.LineSpacing > 0 {
			// line spacing, 240 represents single line spacing
			spacing.Line = strconv.Itoa(int(config.LineSpacing * 240))
		}

		p.Properties.Spacing = spacing

		if config.FirstLineIndent > 0 {
			if p.Properties.Indentation == nil {
				p.Properties.Indentation = &Indentation{}
			}
			// convert to TWIPs (1/20 points)
			p.Properties.Indentation.FirstLine = strconv.Itoa(config.FirstLineIndent * 20)
		}

		Debugf("set paragraph spacing: before=%d, after=%d, line spacing=%.1f, first line indentation=%d",
			config.BeforePara, config.AfterPara, config.LineSpacing, config.FirstLineIndent)
	}
}

// AddFormattedText
//
// this method allows mixing different text formats in a paragraph.
// new text will be added as a new Run to the paragraph.
//
// text is the text content to add.
// format specifies the text format, if nil the default format will be used.
//
// Example:
//
//	para := doc.AddParagraph("this paragraph contains")
//
//	// add bold red text
//	para.AddFormattedText("bold red", &document.TextFormat{
//		Bold: true,
//		FontColor: "FF0000",
//	})
//
//	// add normal text
//	para.AddFormattedText("normal text", nil)
//
//	// add italic blue text
//	para.AddFormattedText("italic blue", &document.TextFormat{
//		Italic: true,
//		FontColor: "0000FF",
//		FontSize: 14,
//	})
func (p *Paragraph) AddFormattedText(text string, format *TextFormat) {
	// create run properties
	runProps := &RunProperties{}

	if format != nil {
		fontName := ""
		if format.FontFamily != "" {
			fontName = format.FontFamily
		} else if format.FontName != "" { // 兼容旧示例
			fontName = format.FontName
		}
		if fontName != "" {
			runProps.FontFamily = &FontFamily{
				ASCII:    fontName,
				HAnsi:    fontName,
				EastAsia: fontName,
				CS:       fontName,
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
			runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
		}
		if format.Underline {
			runProps.Underline = &Underline{Val: "single"} // default single underline
		}

		if format.Strike {
			runProps.Strike = &Strike{} // add strikethrough
		}

		if format.Highlight != "" {
			runProps.Highlight = &Highlight{Val: format.Highlight}
		}
	}

	run := Run{
		Properties: runProps,
		Text: Text{
			Content: text,
			Space:   "preserve",
		},
	}

	p.Runs = append(p.Runs, run)
	Debugf("add formatted text: %s", text)
}

// AddPageBreak
//
// this method adds a page break to the current paragraph.
// the content after the page break will be displayed on a new page.
// different from Document.AddPageBreak(), this method does not create a new paragraph, but adds a page break to the current paragraph.
//
// Example:
//
//	para := doc.AddParagraph("content of the first page")
//	para.AddPageBreak()
//	para.AddFormattedText("content of the second page", nil)
func (p *Paragraph) AddPageBreak() {
	run := Run{
		Break: &Break{
			Type: "page",
		},
	}
	p.Runs = append(p.Runs, run)
	Debugf("add page break")
}

// AddHeadingParagraph
//
// text is the text content of the heading.
// level is the heading level (1-9), corresponding to Heading1 to Heading9.
//
// Returns pointer to newly created paragraph that can be used to further set paragraph properties.
// this method automatically sets the correct style reference, ensuring the heading can be recognized by the Word navigation pane.
//
// Example:
//
//	doc := document.New()
//
//	// add heading level 1
//	h1 := doc.AddHeadingParagraph("chapter 1: overview", 1)
//
//	// add heading level 2
//	h2 := doc.AddHeadingParagraph("1.1 background", 2)
//
//	// add heading level 3
//	h3 := doc.AddHeadingParagraph("1.1.1 research target", 3)
func (d *Document) AddHeadingParagraph(text string, level int) *Paragraph {
	return d.AddHeadingParagraphWithBookmark(text, level, "")
}

// AddHeadingParagraphWithBookmark
//
// text is the text content of the heading.
// level is the heading level (1-9), corresponding to Heading1 to Heading9.
// bookmarkName is the bookmark name, if empty string no bookmark will be added.
//
// Returns pointer to newly created paragraph that can be used to further set paragraph properties.
// this method automatically sets the correct style reference, ensuring the heading can be recognized by the Word navigation pane,
// and adds a bookmark when needed to support directory navigation and hyperlinks.
//
// Example:
//
//	doc := document.New()
//
//	// add heading level 1 with bookmark
//	h1 := doc.AddHeadingParagraphWithBookmark("chapter 1: overview", 1, "chapter1")
//
//	// add heading level 2 without bookmark
//	h2 := doc.AddHeadingParagraphWithBookmark("1.1 background", 2, "")
//
//	// add heading level 3 with auto generated bookmark name
//	h3 := doc.AddHeadingParagraphWithBookmark("1.1.1 research target", 3, "auto_bookmark")
func (d *Document) AddHeadingParagraphWithBookmark(text string, level int, bookmarkName string) *Paragraph {
	if level < 1 || level > 9 {
		Debugf("heading level %d out of range, using default level 1", level)
		level = 1
	}

	styleID := fmt.Sprintf("Heading%d", level)
	Debugf("add heading paragraph: %s (level: %d, style: %s, bookmark: %s)", text, level, styleID, bookmarkName)

	// get style from style manager
	headingStyle := d.styleManager.GetStyle(styleID)
	if headingStyle == nil {
		Debugf("warning: style %s not found, using default style", styleID)
		return d.AddParagraph(text)
	}

	// create run properties, apply character format from style
	runProps := &RunProperties{}
	if headingStyle.RunPr != nil {
		if headingStyle.RunPr.Bold != nil {
			runProps.Bold = &Bold{}
		}
		if headingStyle.RunPr.Italic != nil {
			runProps.Italic = &Italic{}
		}
		if headingStyle.RunPr.FontSize != nil {
			runProps.FontSize = &FontSize{Val: headingStyle.RunPr.FontSize.Val}
		}
		if headingStyle.RunPr.Color != nil {
			runProps.Color = &Color{Val: headingStyle.RunPr.Color.Val}
		}
		if headingStyle.RunPr.FontFamily != nil {
			runProps.FontFamily = &FontFamily{ASCII: headingStyle.RunPr.FontFamily.ASCII}
		}
	}

	// create paragraph properties, apply paragraph format from style
	paraProps := &ParagraphProperties{
		ParagraphStyle: &ParagraphStyle{Val: styleID},
	}

	// apply paragraph format from style
	if headingStyle.ParagraphPr != nil {
		if headingStyle.ParagraphPr.Spacing != nil {
			paraProps.Spacing = &Spacing{
				Before: headingStyle.ParagraphPr.Spacing.Before,
				After:  headingStyle.ParagraphPr.Spacing.After,
				Line:   headingStyle.ParagraphPr.Spacing.Line,
			}
		}
		if headingStyle.ParagraphPr.Justification != nil {
			paraProps.Justification = &Justification{
				Val: headingStyle.ParagraphPr.Justification.Val,
			}
		}
		if headingStyle.ParagraphPr.Indentation != nil {
			paraProps.Indentation = &Indentation{
				FirstLine: headingStyle.ParagraphPr.Indentation.FirstLine,
				Left:      headingStyle.ParagraphPr.Indentation.Left,
				Right:     headingStyle.ParagraphPr.Indentation.Right,
			}
		}
	}

	// create runs for the paragraph
	runs := make([]Run, 0)

	// if bookmark is needed, add bookmark start marker at the beginning of the paragraph
	if bookmarkName != "" {
		// generate unique bookmark ID
		bookmarkID := fmt.Sprintf("bookmark_%d_%s", len(d.Body.Elements), bookmarkName)

		// add bookmark start marker as a separate element to the document body
		d.Body.Elements = append(d.Body.Elements, &BookmarkStart{
			ID:   bookmarkID,
			Name: bookmarkName,
		})

		Debugf("add bookmark start: ID=%s, Name=%s", bookmarkID, bookmarkName)
	}

	// add text content
	runs = append(runs, Run{
		Properties: runProps,
		Text: Text{
			Content: text,
			Space:   "preserve",
		},
	})

	// create paragraph
	p := &Paragraph{
		Properties: paraProps,
		Runs:       runs,
	}

	d.Body.Elements = append(d.Body.Elements, p)

	// if bookmark is needed, add bookmark end marker at the end of the paragraph
	if bookmarkName != "" {
		bookmarkID := fmt.Sprintf("bookmark_%d_%s", len(d.Body.Elements)-2, bookmarkName) // -2 because the paragraph has already been added

		// add bookmark end marker
		d.Body.Elements = append(d.Body.Elements, &BookmarkEnd{
			ID: bookmarkID,
		})

		Debugf("add bookmark end: ID=%s", bookmarkID)
	}

	return p
}

// AddPageBreak
//
// this method adds a page break to the current position, forcing a new page to start at the current position.
// this method creates a paragraph containing the page break.
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("content of the first page")
//	doc.AddPageBreak()
//	doc.AddParagraph("content of the second page")
func (d *Document) AddPageBreak() {
	Debugf("add page break")

	// create a paragraph containing the page break
	p := &Paragraph{
		Runs: []Run{
			{
				Break: &Break{
					Type: "page",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
}

// SetStyle
//
// styleID is the ID of the style to apply, such as "Heading1", "Normal" etc.
// this method sets the paragraph style reference, ensuring the paragraph uses the specified style.
//
// Example:
//
//	para := doc.AddParagraph("this is a paragraph")
//	para.SetStyle("Heading2")  // set to heading level 2 style
func (p *Paragraph) SetStyle(styleID string) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	p.Properties.ParagraphStyle = &ParagraphStyle{Val: styleID}
	Debugf("set paragraph style: %s", styleID)
}

// SetIndentation
//
// Parameters:
//   - firstLineCm: first line indentation, in centimeters (can be negative for hanging indentation)
//   - leftCm: left indentation, in centimeters
//   - rightCm: right indentation, in centimeters
//
// Example:
//
//	para := doc.AddParagraph("this is a paragraph with indentation")
//	para.SetIndentation(0.5, 0, 0)    // first line indentation 0.5 centimeters
//	para.SetIndentation(-0.5, 1, 0)  // hanging indentation 0.5 centimeters, left indentation 1 centimeter
func (p *Paragraph) SetIndentation(firstLineCm, leftCm, rightCm float64) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if p.Properties.Indentation == nil {
		p.Properties.Indentation = &Indentation{}
	}

	// convert centimeters to TWIPs (1 centimeter = 567 TWIPs)
	if firstLineCm != 0 {
		p.Properties.Indentation.FirstLine = strconv.Itoa(int(firstLineCm * 567))
	}

	if leftCm != 0 {
		p.Properties.Indentation.Left = strconv.Itoa(int(leftCm * 567))
	}

	if rightCm != 0 {
		p.Properties.Indentation.Right = strconv.Itoa(int(rightCm * 567))
	}

	Debugf("set paragraph indentation: first line=%.2fcm, left=%.2fcm, right=%.2fcm", firstLineCm, leftCm, rightCm)
}

// SetKeepWithNext
//
// this method ensures the current paragraph and the next paragraph are not separated by a page break,
// commonly used for title and content combinations, or content that needs to be kept continuous.
//
// Parameters:
//   - keep: true to enable, false to disable
//
// Example:
//
//	// title and next paragraph are kept together
//	title := doc.AddParagraph("chapter 1: overview")
//	title.SetKeepWithNext(true)
//	doc.AddParagraph("chapter 1: introduction")  // this content will be kept on the same page as the title
func (p *Paragraph) SetKeepWithNext(keep bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if keep {
		p.Properties.KeepNext = &KeepNext{Val: "1"}
		Debugf("set paragraph to keep with next paragraph")
	} else {
		p.Properties.KeepNext = nil
		Debugf("cancel paragraph to keep with next paragraph")
	}
}

// SetKeepLines
//
// this method ensures the current paragraph and the next paragraph are not separated by a page break,
// ensuring all lines of the paragraph are displayed on the same page.
//
// Parameters:
//   - keep: true to enable, false to disable
//
// Example:
//
//	// ensure the entire paragraph is not split across pages
//	para := doc.AddParagraph("this is an important paragraph, needs to be kept intact.")
//	para.SetKeepLines(true)
func (p *Paragraph) SetKeepLines(keep bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if keep {
		p.Properties.KeepLines = &KeepLines{Val: "1"}
		Debugf("set paragraph lines to keep together")
	} else {
		p.Properties.KeepLines = nil
		Debugf("cancel paragraph lines to keep together")
	}
}

// SetPageBreakBefore
//
// this method forces a page break before the current paragraph, ensuring the paragraph starts on a new page.
// commonly used for section titles or content that needs to be displayed on a separate page.
//
// Parameters:
//   - pageBreak: true to enable, false to disable
//
// Example:
//
//	// section title starts on a new page
//	chapter := doc.AddParagraph("section 2: detailed explanation")
//	chapter.SetPageBreakBefore(true)
func (p *Paragraph) SetPageBreakBefore(pageBreak bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if pageBreak {
		p.Properties.PageBreakBefore = &PageBreakBefore{Val: "1"}
		Debugf("set page break before paragraph")
	} else {
		p.Properties.PageBreakBefore = nil
		Debugf("cancel page break before paragraph")
	}
}

// SetWidowControl
//
// widow control is used to prevent the first or last line of a paragraph from appearing at the bottom or top of a page,
// improving the document's typography.
//
// Parameters:
//   - control: true to enable, false to disable
//
// Example:
//
//	para := doc.AddParagraph("this is a long paragraph...")
//	para.SetWidowControl(true)  // enable widow control
func (p *Paragraph) SetWidowControl(control bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if control {
		p.Properties.WidowControl = &WidowControl{Val: "1"}
		Debugf("enable paragraph widow control")
	} else {
		p.Properties.WidowControl = &WidowControl{Val: "0"}
		Debugf("disable paragraph widow control")
	}
}

// SetOutlineLevel
//
// outline level is used to display the document structure in the document navigation pane, the level range is 0-8.
// commonly used for title paragraphs, used with the directory function.
//
// Parameters:
//   - level: outline level, an integer between 0-8 (0 for body, 1-8 for heading 1-8)
//
// Example:
//
//	// set to outline level 0 (corresponds to Heading1)
//	title := doc.AddParagraph("chapter 1")
//	title.SetOutlineLevel(0)  // corresponds to Heading1
//
//	// set to outline level 1 (corresponds to Heading2)
//	subtitle := doc.AddParagraph("1.1 overview")
//	subtitle.SetOutlineLevel(1)  // corresponds to Heading2
func (p *Paragraph) SetOutlineLevel(level int) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if level < 0 || level > 8 {
		Warnf("outline level should be between 0-8, adjusted to valid range")
		if level < 0 {
			level = 0
		} else {
			level = 8
		}
	}

	p.Properties.OutlineLevel = &OutlineLevel{Val: strconv.Itoa(level)}
	Debugf("set paragraph outline level: %d", level)
}

// SetSnapToGrid
//
// snap to grid is used to control whether the lines of the paragraph are aligned to the grid of the document.
// when the document has enabled grid settings (such as the "if defined, align to grid" option in Chinese documents),
// the custom line spacing may not take effect accurately because the lines will be automatically aligned to the grid lines.
//
// by setting snapToGrid to false, the grid alignment of the paragraph can be disabled,
// allowing the custom line spacing to take effect accurately.
//
// Parameters:
//   - snapToGrid: true to enable, false to disable
//
// Example:
//
//	// disable grid alignment, allowing the custom line spacing to take effect accurately
//	para := doc.AddParagraph("this paragraph uses precise line spacing")
//	para.SetSpacing(&document.SpacingConfig{LineSpacing: 1.5})
//	para.SetSnapToGrid(false)  // disable grid alignment
func (p *Paragraph) SetSnapToGrid(snapToGrid bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if !snapToGrid {
		p.Properties.SnapToGrid = &SnapToGrid{Val: "0"}
		Debugf("disable paragraph grid alignment")
	} else {
		// 启用网格对齐时移除设置（使用默认行为）
		p.Properties.SnapToGrid = nil
		Debugf("enable paragraph grid alignment (default)")
	}
}

// ParagraphFormatConfig
//
// this structure provides a unified configuration interface for all paragraph format attributes,
// allowing multiple paragraph attributes to be set at once, improving code readability and ease of use.
type ParagraphFormatConfig struct {
	// 基础格式
	Alignment AlignmentType // alignment type (AlignLeft, AlignCenter, AlignRight, AlignJustify)
	Style     string        // 段落样式ID（如"Heading1", "Normal"等）

	// 间距设置
	LineSpacing     float64 // 行间距（倍数，如1.5表示1.5倍行距）
	BeforePara      int     // 段前间距（磅）
	AfterPara       int     // 段后间距（磅）
	FirstLineIndent int     // 首行缩进（磅）

	// 缩进设置
	FirstLineCm float64 // 首行缩进（厘米，可以为负数表示悬挂缩进）
	LeftCm      float64 // 左缩进（厘米）
	RightCm     float64 // 右缩进（厘米）

	// 分页与控制
	KeepWithNext    bool  // 与下一段落保持在同一页
	KeepLines       bool  // 段落中的所有行保持在同一页
	PageBreakBefore bool  // 段前分页
	WidowControl    bool  // 孤行控制
	SnapToGrid      *bool // 是否对齐网格（设置为false可禁用网格对齐，使自定义行间距精确生效）

	// 大纲级别
	OutlineLevel int // 大纲级别（0-8，0表示正文，1-8对应标题1-8）
}

// SetParagraphFormat 使用配置一次性设置段落的所有格式属性。
//
// 此方法提供了一种便捷的方式来设置段落的所有格式属性，
// 而不需要调用多个单独的设置方法。只有非零值的属性会被应用。
//
// Parameters:
//   - config: 段落格式配置，包含所有格式属性
//
// Example:
//
//	// 创建一个带完整格式的段落
//	para := doc.AddParagraph("重要章节标题")
//	para.SetParagraphFormat(&document.ParagraphFormatConfig{
//		Alignment:       document.AlignCenter,
//		Style:           "Heading1",
//		LineSpacing:     1.5,
//		BeforePara:      24,
//		AfterPara:       12,
//		KeepWithNext:    true,
//		PageBreakBefore: true,
//		OutlineLevel:    0,
//	})
//
//	// 设置带缩进的正文段落
//	para2 := doc.AddParagraph("正文内容...")
//	para2.SetParagraphFormat(&document.ParagraphFormatConfig{
//		Alignment:       document.AlignJustify,
//		FirstLineCm:     0.5,
//		LineSpacing:     1.5,
//		BeforePara:      6,
//		AfterPara:       6,
//		WidowControl:    true,
//	})
func (p *Paragraph) SetParagraphFormat(config *ParagraphFormatConfig) {
	if config == nil {
		return
	}

	// 设置对齐方式
	if config.Alignment != "" {
		p.SetAlignment(config.Alignment)
	}

	// 设置样式
	if config.Style != "" {
		p.SetStyle(config.Style)
	}

	// 设置间距（如果有任何间距设置）
	if config.LineSpacing > 0 || config.BeforePara > 0 || config.AfterPara > 0 || config.FirstLineIndent > 0 {
		p.SetSpacing(&SpacingConfig{
			LineSpacing:     config.LineSpacing,
			BeforePara:      config.BeforePara,
			AfterPara:       config.AfterPara,
			FirstLineIndent: config.FirstLineIndent,
		})
	}

	// 设置缩进（如果有任何缩进设置）
	if config.FirstLineCm != 0 || config.LeftCm != 0 || config.RightCm != 0 {
		p.SetIndentation(config.FirstLineCm, config.LeftCm, config.RightCm)
	}

	// 设置分页和控制属性
	p.SetKeepWithNext(config.KeepWithNext)
	p.SetKeepLines(config.KeepLines)
	p.SetPageBreakBefore(config.PageBreakBefore)
	p.SetWidowControl(config.WidowControl)

	// 设置网格对齐
	if config.SnapToGrid != nil {
		p.SetSnapToGrid(*config.SnapToGrid)
	}

	// 设置大纲级别
	if config.OutlineLevel >= 0 && config.OutlineLevel <= 8 {
		p.SetOutlineLevel(config.OutlineLevel)
	}

	Debugf("应用段落格式配置: 对齐=%s, 样式=%s, 行距=%.1f, 段前=%d, 段后=%d",
		config.Alignment, config.Style, config.LineSpacing, config.BeforePara, config.AfterPara)
}

// ParagraphBorderConfig 段落边框配置（区别于表格边框配置）
type ParagraphBorderConfig struct {
	Style BorderStyle // border style
	Size  int         // 边框粗细（1/8磅为单位，默认值建议12，即1.5磅）
	Color string      // 边框颜色（十六进制，如"000000"表示黑色）
	Space int         // 边框与文本的间距（磅，默认值建议1）
}

// SetBorder 设置段落的边框。
//
// 此方法用于为段落添加边框装饰，特别适用于实现Markdown分割线(---)的转换。
//
// Parameters:
//   - top: 上边框配置，传入nil表示不设置上边框
//   - left: 左边框配置，传入nil表示不设置左边框
//   - bottom: 下边框配置，传入nil表示不设置下边框
//   - right: 右边框配置，传入nil表示不设置右边框
//
// 边框配置包含样式、粗细、颜色和间距等属性。
//
// Example:
//
//	// 设置分割线效果（仅底边框）
//	para := doc.AddParagraph("")
//	para.SetBorder(nil, nil, &document.ParagraphBorderConfig{
//		Style: document.BorderStyleSingle,
//		Size:  12,   // 1.5磅粗细
//		Color: "000000", // 黑色
//		Space: 1,    // 1磅间距
//	}, nil)
//
//	// 设置完整边框
//	para := doc.AddParagraph("带边框的段落")
//	borderConfig := &document.ParagraphBorderConfig{
//		Style: document.BorderStyleDouble,
//		Size:  8,
//		Color: "0000FF", // 蓝色
//		Space: 2,
//	}
//	para.SetBorder(borderConfig, borderConfig, borderConfig, borderConfig)
func (p *Paragraph) SetBorder(top, left, bottom, right *ParagraphBorderConfig) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// 如果没有任何边框配置，清除边框
	if top == nil && left == nil && bottom == nil && right == nil {
		p.Properties.ParagraphBorder = nil
		return
	}

	// 创建段落边框
	if p.Properties.ParagraphBorder == nil {
		p.Properties.ParagraphBorder = &ParagraphBorder{}
	}

	// 设置上边框
	if top != nil {
		p.Properties.ParagraphBorder.Top = &ParagraphBorderLine{
			Val:   string(top.Style),
			Sz:    strconv.Itoa(top.Size),
			Color: top.Color,
			Space: strconv.Itoa(top.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Top = nil
	}

	// 设置左边框
	if left != nil {
		p.Properties.ParagraphBorder.Left = &ParagraphBorderLine{
			Val:   string(left.Style),
			Sz:    strconv.Itoa(left.Size),
			Color: left.Color,
			Space: strconv.Itoa(left.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Left = nil
	}

	// 设置下边框
	if bottom != nil {
		p.Properties.ParagraphBorder.Bottom = &ParagraphBorderLine{
			Val:   string(bottom.Style),
			Sz:    strconv.Itoa(bottom.Size),
			Color: bottom.Color,
			Space: strconv.Itoa(bottom.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Bottom = nil
	}

	// 设置右边框
	if right != nil {
		p.Properties.ParagraphBorder.Right = &ParagraphBorderLine{
			Val:   string(right.Style),
			Sz:    strconv.Itoa(right.Size),
			Color: right.Color,
			Space: strconv.Itoa(right.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Right = nil
	}

	Debugf("设置段落边框: 上=%v, 左=%v, 下=%v, 右=%v", top != nil, left != nil, bottom != nil, right != nil)
}

// SetHorizontalRule 设置水平分割线。
//
// 此方法是SetBorder的简化版本，专门用于快速创建Markdown风格的分割线效果。
// 只在段落底部添加一条水平线，适用于Markdown中的 --- 或 *** 语法。
//
// Parameters:
//   - style: 边框样式，如BorderStyleSingle、BorderStyleDouble等
//   - size: 边框粗细（1/8磅为单位，建议值12-18）
//   - color: 边框颜色（十六进制，如"000000"）
//
// Example:
//
//	// 创建简单分割线
//	para := doc.AddParagraph("")
//	para.SetHorizontalRule(document.BorderStyleSingle, 12, "000000")
//
//	// 创建粗双线分割线
//	para := doc.AddParagraph("")
//	para.SetHorizontalRule(document.BorderStyleDouble, 18, "808080")
func (p *Paragraph) SetHorizontalRule(style BorderStyle, size int, color string) {
	borderConfig := &ParagraphBorderConfig{
		Style: style,
		Size:  size,
		Color: color,
		Space: 1, // 默认1磅间距
	}

	p.SetBorder(nil, nil, borderConfig, nil)

	Debugf("设置水平分割线: 样式=%s, 粗细=%d, 颜色=%s", style, size, color)
}

// SetUnderline 设置段落中所有文本的下划线效果。
//
// 参数 underline 表示是否启用下划线。
// 当设置为 true 时，将对段落中所有运行应用单线下划线效果。
// 当设置为 false 时，将移除所有运行的下划线效果。
//
// Example:
//
//	para := doc.AddParagraph("这是下划线文本")
//	para.SetUnderline(true)
func (p *Paragraph) SetUnderline(underline bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if underline {
			p.Runs[i].Properties.Underline = &Underline{Val: "single"}
		} else {
			p.Runs[i].Properties.Underline = nil
		}
	}
	Debugf("设置段落下划线: %v", underline)
}

// SetBold 设置段落中所有文本的粗体效果。
//
// 参数 bold 表示是否启用粗体。
// 当设置为 true 时，将对段落中所有运行应用粗体效果。
// 当设置为 false 时，将移除所有运行的粗体效果。
//
// Example:
//
//	para := doc.AddParagraph("这是粗体文本")
//	para.SetBold(true)
func (p *Paragraph) SetBold(bold bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if bold {
			p.Runs[i].Properties.Bold = &Bold{}
			p.Runs[i].Properties.BoldCs = &BoldCs{}
		} else {
			p.Runs[i].Properties.Bold = nil
			p.Runs[i].Properties.BoldCs = nil
		}
	}
	Debugf("设置段落粗体: %v", bold)
}

// SetItalic 设置段落中所有文本的斜体效果。
//
// 参数 italic 表示是否启用斜体。
// 当设置为 true 时，将对段落中所有运行应用斜体效果。
// 当设置为 false 时，将移除所有运行的斜体效果。
//
// Example:
//
//	para := doc.AddParagraph("这是斜体文本")
//	para.SetItalic(true)
func (p *Paragraph) SetItalic(italic bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if italic {
			p.Runs[i].Properties.Italic = &Italic{}
			p.Runs[i].Properties.ItalicCs = &ItalicCs{}
		} else {
			p.Runs[i].Properties.Italic = nil
			p.Runs[i].Properties.ItalicCs = nil
		}
	}
	Debugf("设置段落斜体: %v", italic)
}

// SetStrike 设置段落中所有文本的删除线效果。
//
// 参数 strike 表示是否启用删除线。
// 当设置为 true 时，将对段落中所有运行应用删除线效果。
// 当设置为 false 时，将移除所有运行的删除线效果。
//
// Example:
//
//	para := doc.AddParagraph("这是删除线文本")
//	para.SetStrike(true)
func (p *Paragraph) SetStrike(strike bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if strike {
			p.Runs[i].Properties.Strike = &Strike{}
		} else {
			p.Runs[i].Properties.Strike = nil
		}
	}
	Debugf("设置段落删除线: %v", strike)
}

// SetHighlight 设置段落中所有文本的高亮颜色。
//
// 参数 color 是高亮颜色名称，支持的颜色包括：
// "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue",
// "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
// "darkGray", "lightGray", "black" 等。
// 传入空字符串将移除高亮效果。
//
// Example:
//
//	para := doc.AddParagraph("这是高亮文本")
//	para.SetHighlight("yellow")
func (p *Paragraph) SetHighlight(color string) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if color != "" {
			p.Runs[i].Properties.Highlight = &Highlight{Val: color}
		} else {
			p.Runs[i].Properties.Highlight = nil
		}
	}
	Debugf("设置段落高亮: %s", color)
}

// SetFontFamily 设置段落中所有文本的字体。
//
// 参数 name 是字体名称，如 "Arial"、"Times New Roman"、"微软雅黑" 等。
//
// Example:
//
//	para := doc.AddParagraph("这是自定义字体文本")
//	para.SetFontFamily("微软雅黑")
func (p *Paragraph) SetFontFamily(name string) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if name != "" {
			p.Runs[i].Properties.FontFamily = &FontFamily{
				ASCII:    name,
				HAnsi:    name,
				EastAsia: name,
				CS:       name,
			}
		} else {
			p.Runs[i].Properties.FontFamily = nil
		}
	}
	Debugf("设置段落字体: %s", name)
}

// SetFontSize 设置段落中所有文本的字体大小。
//
// 参数 size 是字体大小（磅），如 12、14、16 等。
// 传入 0 或负数将移除字体大小设置。
//
// Example:
//
//	para := doc.AddParagraph("这是大号文本")
//	para.SetFontSize(16)
func (p *Paragraph) SetFontSize(size int) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if size > 0 {
			// Word 使用半磅为单位，所以乘以2
			sizeStr := strconv.Itoa(size * 2)
			p.Runs[i].Properties.FontSize = &FontSize{Val: sizeStr}
			p.Runs[i].Properties.FontSizeCs = &FontSizeCs{Val: sizeStr}
		} else {
			p.Runs[i].Properties.FontSize = nil
			p.Runs[i].Properties.FontSizeCs = nil
		}
	}
	Debugf("设置段落字体大小: %d", size)
}

// SetColor 设置段落中所有文本的颜色。
//
// 参数 color 是十六进制颜色值，如 "FF0000"（红色）、"0000FF"（蓝色）等。
// 颜色值不需要 "#" 前缀，如果包含会自动移除。
// 传入空字符串将移除颜色设置。
//
// Example:
//
//	para := doc.AddParagraph("这是红色文本")
//	para.SetColor("FF0000")
func (p *Paragraph) SetColor(color string) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if color != "" {
			// 移除可能存在的 # 前缀
			colorVal := strings.TrimPrefix(color, "#")
			p.Runs[i].Properties.Color = &Color{Val: colorVal}
		} else {
			p.Runs[i].Properties.Color = nil
		}
	}
	Debugf("设置段落颜色: %s", color)
}

// GetStyleManager 获取文档的样式管理器。
//
// 返回文档的样式管理器，可用于访问和管理样式。
//
// Example:
//
//	doc := document.New()
//	styleManager := doc.GetStyleManager()
//	headingStyle := styleManager.GetStyle("Heading1")
func (d *Document) GetStyleManager() *style.StyleManager {
	return d.styleManager
}

// GetParts 获取文档部件映射
//
// 返回包含文档所有部件的映射，主要用于测试和调试。
// 键是部件名称，值是部件内容的字节数组。
//
// Example:
//
//	parts := doc.GetParts()
//	settingsXML := parts["word/settings.xml"]
func (d *Document) GetParts() map[string][]byte {
	return d.parts
}

// initializeStructure 初始化文档基础结构
func (d *Document) initializeStructure() {
	// 初始化 content types
	d.contentTypes = &ContentTypes{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
		Defaults: []Default{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
		},
		Overrides: []Override{
			{PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
			{PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
		},
	}

	// 初始化主关系
	d.relationships = &Relationships{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: []Relationship{
			{
				ID:     "rId1",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
				Target: "word/document.xml",
			},
		},
	}

	// 添加基础部件
	d.serializeContentTypes()
	d.serializeRelationships()
	d.serializeDocumentRelationships()
}

// parseDocument 解析文档内容
func (d *Document) parseDocument() error {
	Debugf("开始解析文档内容")

	// 解析主文档
	docData, ok := d.parts["word/document.xml"]
	if !ok {
		return WrapError("parse_document", ErrDocumentNotFound)
	}

	// 首先解析基本结构
	decoder := xml.NewDecoder(bytes.NewReader(docData))
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return WrapError("parse_document", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Local == "document" && t.Name.Space == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" {
				// 开始解析文档
				if err := d.parseDocumentElement(decoder); err != nil {
					return err
				}
				goto done
			}
		}
	}

done:
	Infof("Parsing completed, %d elements", len(d.Body.Elements))
	return nil
}

// parseDocumentElement
func (d *Document) parseDocumentElement(decoder *xml.Decoder) error {
	// initialize Body
	d.Body = &Body{
		Elements: make([]interface{}, 0),
	}

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return WrapError("parse_document_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch {
			case t.Name.Local == "body":
				if err := d.parseBodyElement(decoder); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "document" {
				return nil
			}
		}
	}

	return nil
}

// parseBodyElement
func (d *Document) parseBodyElement(decoder *xml.Decoder) error {
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return WrapError("parse_body_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			element, err := d.parseBodySubElement(decoder, t)
			if err != nil {
				return err
			}
			if element != nil {
				d.Body.Elements = append(d.Body.Elements, element)
			}
		case xml.EndElement:
			if t.Name.Local == "body" {
				return nil
			}
		}
	}

	return nil
}

// parseBodySubElement
func (d *Document) parseBodySubElement(decoder *xml.Decoder, startElement xml.StartElement) (interface{}, error) {
	switch startElement.Name.Local {
	case "p":
		return d.parseParagraph(decoder, startElement)
	case "tbl":
		return d.parseTable(decoder, startElement)
	case "sectPr":
		return d.parseSectionProperties(decoder, startElement)
	default:
		Debugf("跳过未知元素: %s", startElement.Name.Local)
		return nil, d.skipElement(decoder, startElement.Name.Local)
	}
}

// parseParagraph
func (d *Document) parseParagraph(decoder *xml.Decoder, startElement xml.StartElement) (*Paragraph, error) {
	paragraph := &Paragraph{
		Runs: make([]Run, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_paragraph", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pPr":
				if err := d.parseParagraphProperties(decoder, paragraph); err != nil {
					return nil, err
				}
			case "r":
				run, err := d.parseRun(decoder, t)
				if err != nil {
					return nil, err
				}
				if run != nil {
					paragraph.Runs = append(paragraph.Runs, *run)
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "p" {
				return paragraph, nil
			}
		}
	}
}

// parseParagraphProperties
func (d *Document) parseParagraphProperties(decoder *xml.Decoder, paragraph *Paragraph) error {
	paragraph.Properties = &ParagraphProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_paragraph_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pStyle":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					paragraph.Properties.ParagraphStyle = &ParagraphStyle{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "spacing":
				spacing := &Spacing{}
				spacing.Before = getAttributeValue(t.Attr, "before")
				spacing.After = getAttributeValue(t.Attr, "after")
				spacing.Line = getAttributeValue(t.Attr, "line")
				spacing.LineRule = getAttributeValue(t.Attr, "lineRule")
				paragraph.Properties.Spacing = spacing
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "jc":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					paragraph.Properties.Justification = &Justification{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "ind":
				indentation := &Indentation{}
				indentation.FirstLine = getAttributeValue(t.Attr, "firstLine")
				indentation.Left = getAttributeValue(t.Attr, "left")
				indentation.Right = getAttributeValue(t.Attr, "right")
				paragraph.Properties.Indentation = indentation
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "numPr":
				numPr, err := d.parseNumberingProperties(decoder)
				if err != nil {
					return err
				}
				paragraph.Properties.NumberingProperties = numPr
			case "sectPr":
				sectPr, err := d.parseSectionProperties(decoder, t)
				if err != nil {
					return err
				}
				d.setSectionProperties(sectPr)
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "pPr" {
				return nil
			}
		}
	}
}

// parseNumberingProperties
func (d *Document) parseNumberingProperties(decoder *xml.Decoder) (*NumberingProperties, error) {
	numPr := &NumberingProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_numbering_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "ilvl":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					numPr.ILevel = &ILevel{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "numId":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					numPr.NumID = &NumID{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "numPr" {
				return numPr, nil
			}
		}
	}
}

// parseRun
func (d *Document) parseRun(decoder *xml.Decoder, startElement xml.StartElement) (*Run, error) {
	run := &Run{
		Text: Text{},
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_run", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPr":
				if err := d.parseRunProperties(decoder, run); err != nil {
					return nil, err
				}
			case "t":
				space := getAttributeValue(t.Attr, "space")
				run.Text.Space = space

				content, err := d.readElementText(decoder, "t")
				if err != nil {
					return nil, err
				}
				run.Text.Content = content
			case "drawing":
				drawing, err := d.parseDrawingElement(decoder, t)
				if err != nil {
					return nil, err
				}
				run.Drawing = drawing
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "r" {
				return run, nil
			}
		}
	}
}

// parseRunProperties
func (d *Document) parseRunProperties(decoder *xml.Decoder, run *Run) error {
	run.Properties = &RunProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_run_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "b":
				run.Properties.Bold = &Bold{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "bCs":
				run.Properties.BoldCs = &BoldCs{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "i":
				run.Properties.Italic = &Italic{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "iCs":
				run.Properties.ItalicCs = &ItalicCs{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "u":
				val := getAttributeValue(t.Attr, "val")
				run.Properties.Underline = &Underline{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "strike":
				run.Properties.Strike = &Strike{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "sz":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.FontSize = &FontSize{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "szCs":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.FontSizeCs = &FontSizeCs{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "color":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.Color = &Color{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "highlight":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.Highlight = &Highlight{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "rFonts":
				ascii := getAttributeValue(t.Attr, "ascii")
				hAnsi := getAttributeValue(t.Attr, "hAnsi")
				eastAsia := getAttributeValue(t.Attr, "eastAsia")
				cs := getAttributeValue(t.Attr, "cs")
				hint := getAttributeValue(t.Attr, "hint")

				run.Properties.FontFamily = &FontFamily{
					ASCII:    ascii,
					HAnsi:    hAnsi,
					EastAsia: eastAsia,
					CS:       cs,
					Hint:     hint,
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "rPr" {
				return nil
			}
		}
	}
}

// parseTable
func (d *Document) parseTable(decoder *xml.Decoder, startElement xml.StartElement) (*Table, error) {
	table := &Table{
		Rows: make([]TableRow, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblPr":
				if err := d.parseTableProperties(decoder, table); err != nil {
					return nil, err
				}
			case "tblGrid":
				if err := d.parseTableGrid(decoder, table); err != nil {
					return nil, err
				}
			case "tr":
				row, err := d.parseTableRow(decoder, t)
				if err != nil {
					return nil, err
				}
				if row != nil {
					table.Rows = append(table.Rows, *row)
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tbl" {
				return table, nil
			}
		}
	}
}

// parseTableProperties
func (d *Document) parseTableProperties(decoder *xml.Decoder, table *Table) error {
	table.Properties = &TableProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_table_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblW":
				w := getAttributeValue(t.Attr, "w")
				wType := getAttributeValue(t.Attr, "type")
				if w != "" || wType != "" {
					table.Properties.TableW = &TableWidth{W: w, Type: wType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "jc":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					table.Properties.TableJc = &TableJc{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblLook":
				tableLook := &TableLook{
					Val:      getAttributeValue(t.Attr, "val"),
					FirstRow: getAttributeValue(t.Attr, "firstRow"),
					LastRow:  getAttributeValue(t.Attr, "lastRow"),
					FirstCol: getAttributeValue(t.Attr, "firstColumn"),
					LastCol:  getAttributeValue(t.Attr, "lastColumn"),
					NoHBand:  getAttributeValue(t.Attr, "noHBand"),
					NoVBand:  getAttributeValue(t.Attr, "noVBand"),
				}
				table.Properties.TableLook = tableLook
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblStyle":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					table.Properties.TableStyle = &TableStyle{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblBorders":
				borders, err := d.parseTableBorders(decoder)
				if err != nil {
					return err
				}
				table.Properties.TableBorders = borders
			case "shd":
				shd := &TableShading{
					Val:       getAttributeValue(t.Attr, "val"),
					Color:     getAttributeValue(t.Attr, "color"),
					Fill:      getAttributeValue(t.Attr, "fill"),
					ThemeFill: getAttributeValue(t.Attr, "themeFill"),
				}
				table.Properties.Shd = shd
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblCellMar":
				margins, err := d.parseTableCellMargins(decoder)
				if err != nil {
					return err
				}
				table.Properties.TableCellMar = margins
			case "tblLayout":
				layoutType := getAttributeValue(t.Attr, "type")
				if layoutType != "" {
					table.Properties.TableLayout = &TableLayoutType{Type: layoutType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblInd":
				w := getAttributeValue(t.Attr, "w")
				indType := getAttributeValue(t.Attr, "type")
				if w != "" || indType != "" {
					table.Properties.TableInd = &TableIndentation{W: w, Type: indType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tblPr" {
				return nil
			}
		}
	}
}

// parseTableGrid
func (d *Document) parseTableGrid(decoder *xml.Decoder, table *Table) error {
	table.Grid = &TableGrid{
		Cols: make([]TableGridCol, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_table_grid", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "gridCol":
				w := getAttributeValue(t.Attr, "w")
				col := TableGridCol{W: w}
				table.Grid.Cols = append(table.Grid.Cols, col)
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tblGrid" {
				return nil
			}
		}
	}
}

// parseTableRow
func (d *Document) parseTableRow(decoder *xml.Decoder, startElement xml.StartElement) (*TableRow, error) {
	row := &TableRow{
		Cells: make([]TableCell, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_row", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "trPr":
				props, err := d.parseTableRowProperties(decoder)
				if err != nil {
					return nil, err
				}
				row.Properties = props
			case "tc":
				cell, err := d.parseTableCell(decoder, t)
				if err != nil {
					return nil, err
				}
				if cell != nil {
					row.Cells = append(row.Cells, *cell)
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tr" {
				return row, nil
			}
		}
	}
}

// parseTableCell 解析表格单元格
func (d *Document) parseTableCell(decoder *xml.Decoder, startElement xml.StartElement) (*TableCell, error) {
	cell := &TableCell{
		Paragraphs: make([]Paragraph, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcPr":
				// 解析单元格属性
				props, err := d.parseTableCellProperties(decoder)
				if err != nil {
					return nil, err
				}
				cell.Properties = props
			case "p":
				// 解析段落
				para, err := d.parseParagraph(decoder, t)
				if err != nil {
					return nil, err
				}
				if para != nil {
					cell.Paragraphs = append(cell.Paragraphs, *para)
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tc" {
				return cell, nil
			}
		}
	}
}

// parseSectionProperties 解析节属性
func (d *Document) parseSectionProperties(decoder *xml.Decoder, startElement xml.StartElement) (*SectionProperties, error) {
	sectPr := &SectionProperties{
		XmlnsR: getAttributeValue(startElement.Attr, "xmlns:r"),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_section_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pgSz":
				// 解析页面尺寸
				w := getAttributeValue(t.Attr, "w")
				h := getAttributeValue(t.Attr, "h")
				orient := getAttributeValue(t.Attr, "orient")
				if w != "" || h != "" {
					sectPr.PageSize = &PageSizeXML{W: w, H: h, Orient: orient}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "pgMar":
				// 解析页面边距
				margin := &PageMargin{}
				margin.Top = getAttributeValue(t.Attr, "top")
				margin.Right = getAttributeValue(t.Attr, "right")
				margin.Bottom = getAttributeValue(t.Attr, "bottom")
				margin.Left = getAttributeValue(t.Attr, "left")
				margin.Header = getAttributeValue(t.Attr, "header")
				margin.Footer = getAttributeValue(t.Attr, "footer")
				margin.Gutter = getAttributeValue(t.Attr, "gutter")
				sectPr.PageMargins = margin
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "cols":
				// 解析分栏
				space := getAttributeValue(t.Attr, "space")
				num := getAttributeValue(t.Attr, "num")
				if space != "" || num != "" {
					sectPr.Columns = &Columns{Space: space, Num: num}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "docGrid":
				// 解析文档网格
				docGridType := getAttributeValue(t.Attr, "type")
				linePitch := getAttributeValue(t.Attr, "linePitch")
				charSpace := getAttributeValue(t.Attr, "charSpace")
				if docGridType != "" || linePitch != "" || charSpace != "" {
					sectPr.DocGrid = &DocGrid{
						Type:      docGridType,
						LinePitch: linePitch,
						CharSpace: charSpace,
					}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "headerReference":
				ref := &HeaderFooterReference{
					Type: getAttributeValue(t.Attr, "type"),
					ID:   getAttributeValue(t.Attr, "id"),
				}
				if ref.Type == "" {
					ref.Type = getAttributeValue(t.Attr, "w:type")
				}
				if ref.ID == "" {
					ref.ID = getAttributeValue(t.Attr, "r:id")
				}
				if ref.ID != "" || ref.Type != "" {
					sectPr.HeaderReferences = append(sectPr.HeaderReferences, ref)
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "footerReference":
				ref := &FooterReference{
					Type: getAttributeValue(t.Attr, "type"),
					ID:   getAttributeValue(t.Attr, "id"),
				}
				if ref.Type == "" {
					ref.Type = getAttributeValue(t.Attr, "w:type")
				}
				if ref.ID == "" {
					ref.ID = getAttributeValue(t.Attr, "r:id")
				}
				if ref.ID != "" || ref.Type != "" {
					sectPr.FooterReferences = append(sectPr.FooterReferences, ref)
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// 跳过其他节属性
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "sectPr" {
				return sectPr, nil
			}
		}
	}
}

// skipElement 跳过元素及其子元素
func (d *Document) skipElement(decoder *xml.Decoder, elementName string) error {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("skip_element", err)
		}

		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

// readElementText 读取元素的文本内容
func (d *Document) readElementText(decoder *xml.Decoder, elementName string) (string, error) {
	var content string
	for {
		token, err := decoder.Token()
		if err != nil {
			return "", WrapError("read_element_text", err)
		}

		switch t := token.(type) {
		case xml.CharData:
			content += string(t)
		case xml.EndElement:
			if t.Name.Local == elementName {
				return content, nil
			}
		}
	}
}

// getAttributeValue 获取属性值
func getAttributeValue(attrs []xml.Attr, name string) string {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

// serializeDocument 序列化文档内容
func (d *Document) serializeDocument() error {
	Debugf("开始序列化文档")

	// 创建文档结构
	type documentXML struct {
		XMLName  xml.Name `xml:"w:document"`
		Xmlns    string   `xml:"xmlns:w,attr"`
		XmlnsW15 string   `xml:"xmlns:w15,attr"`
		XmlnsWP  string   `xml:"xmlns:wp,attr"`
		XmlnsA   string   `xml:"xmlns:a,attr"`
		XmlnsPic string   `xml:"xmlns:pic,attr"`
		XmlnsR   string   `xml:"xmlns:r,attr"`
		Body     *Body    `xml:"w:body"`
	}

	doc := documentXML{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW15: "http://schemas.microsoft.com/office/word/2012/wordml",
		XmlnsWP:  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsA:   "http://schemas.openxmlformats.org/drawingml/2006/main",
		XmlnsPic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
		XmlnsR:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Body:     d.Body,
	}

	// serialize为XML
	data, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		Errorf("XML序列化失败: %v", err)
		return WrapError("marshal_xml", err)
	}

	// add XML declaration
	d.parts["word/document.xml"] = append([]byte(xml.Header), data...)

	Debugf("文档序列化完成")
	return nil
}

// serializeContentTypes 序列化内容类型
func (d *Document) serializeContentTypes() {
	data, _ := xml.MarshalIndent(d.contentTypes, "", "  ")
	d.parts["[Content_Types].xml"] = append([]byte(xml.Header), data...)
}

// serializeRelationships 序列化关系
func (d *Document) serializeRelationships() {
	data, _ := xml.MarshalIndent(d.relationships, "", "  ")
	d.parts["_rels/.rels"] = append([]byte(xml.Header), data...)
}

// serializeDocumentRelationships 序列化文档关系
func (d *Document) serializeDocumentRelationships() {
	// 获取已存在的关系，从索引1开始（保留给styles.xml）
	relationships := []Relationship{
		{
			ID:     "rId1",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
			Target: "styles.xml",
		},
	}

	// 添加动态创建的文档级关系（如页眉、页脚等）
	relationships = append(relationships, d.documentRelationships.Relationships...)

	// 创建文档关系
	docRels := &Relationships{
		Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: relationships,
	}

	data, _ := xml.MarshalIndent(docRels, "", "  ")
	d.parts["word/_rels/document.xml.rels"] = append([]byte(xml.Header), data...)
}

// serializeStyles 序列化样式
func (d *Document) serializeStyles() error {
	Debugf("开始序列化样式")

	// 如果在克隆文档时已经保留了完整的 styles.xml（含 docDefaults 等信息），
	// 这里直接跳过重新生成，避免丢失模板原有的默认段落/字符设置。
	if existing, ok := d.parts["word/styles.xml"]; ok && len(existing) > 0 {
		Debugf("检测到已有 styles.xml，跳过样式重建以保留模板默认样式")
		return nil
	}

	// 创建样式结构，包含完整的命名空间
	type stylesXML struct {
		XMLName     xml.Name       `xml:"w:styles"`
		XmlnsW      string         `xml:"xmlns:w,attr"`
		XmlnsMC     string         `xml:"xmlns:mc,attr"`
		XmlnsO      string         `xml:"xmlns:o,attr"`
		XmlnsR      string         `xml:"xmlns:r,attr"`
		XmlnsM      string         `xml:"xmlns:m,attr"`
		XmlnsV      string         `xml:"xmlns:v,attr"`
		XmlnsW14    string         `xml:"xmlns:w14,attr"`
		XmlnsW10    string         `xml:"xmlns:w10,attr"`
		XmlnsSL     string         `xml:"xmlns:sl,attr"`
		XmlnsWPS    string         `xml:"xmlns:wpsCustomData,attr"`
		MCIgnorable string         `xml:"mc:Ignorable,attr"`
		Styles      []*style.Style `xml:"w:style"`
	}

	doc := stylesXML{
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsSL:     "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
		XmlnsWPS:    "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14",
		Styles:      d.styleManager.GetAllStyles(),
	}

	// serialize为XML
	data, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		Errorf("XML序列化失败: %v", err)
		return WrapError("marshal_xml", err)
	}

	// add XML declaration
	d.parts["word/styles.xml"] = append([]byte(xml.Header), data...)

	Debugf("样式序列化完成")
	return nil
}

// parseContentTypes 解析内容类型文件
func (d *Document) parseContentTypes() error {
	Debugf("开始解析内容类型文件")

	// 查找内容类型文件
	contentTypesData, ok := d.parts["[Content_Types].xml"]
	if !ok {
		return WrapError("parse_content_types", fmt.Errorf("内容类型文件不存在"))
	}

	// parse XML
	var contentTypes ContentTypes
	if err := xml.Unmarshal(contentTypesData, &contentTypes); err != nil {
		return WrapError("parse_content_types", err)
	}

	d.contentTypes = &contentTypes
	Debugf("内容类型解析完成")
	return nil
}

// parseRelationships 解析关系文件
func (d *Document) parseRelationships() error {
	Debugf("开始解析关系文件")

	// 查找关系文件
	relsData, ok := d.parts["_rels/.rels"]
	if !ok {
		return WrapError("parse_relationships", fmt.Errorf("关系文件不存在"))
	}

	// parse XML
	var relationships Relationships
	if err := xml.Unmarshal(relsData, &relationships); err != nil {
		return WrapError("parse_relationships", err)
	}

	d.relationships = &relationships
	Debugf("关系解析完成")
	return nil
}

// parseStyles 解析样式文件
func (d *Document) parseStyles() error {
	Debugf("开始解析样式文件")

	// 查找样式文件
	stylesData, ok := d.parts["word/styles.xml"]
	if !ok {
		return WrapError("parse_styles", fmt.Errorf("样式文件不存在"))
	}

	// 使用样式管理器解析样式
	if err := d.styleManager.LoadStylesFromDocument(stylesData); err != nil {
		return WrapError("parse_styles", err)
	}

	Debugf("样式解析完成")
	return nil
}

// parseDocumentRelationships 解析文档关系文件（word/_rels/document.xml.rels）
// 该文件包含文档中图片、页眉、页脚等资源的关系
func (d *Document) parseDocumentRelationships() error {
	Debugf("开始解析文档关系文件")

	// 查找文档关系文件
	docRelsData, ok := d.parts["word/_rels/document.xml.rels"]
	if !ok {
		// 文档可能没有关系文件（没有图片等资源），这不是错误
		Debugf("文档关系文件不存在，文档可能不包含图片等资源")
		return nil
	}

	// parse XML
	var relationships Relationships
	if err := xml.Unmarshal(docRelsData, &relationships); err != nil {
		return WrapError("parse_document_relationships", err)
	}

	// 保存解析的关系（不包括styles.xml，因为它在serializeDocumentRelationships中会自动添加）
	// 过滤掉styles.xml的关系，因为它总是rId1并在保存时自动添加
	filteredRels := make([]Relationship, 0)
	for _, rel := range relationships.Relationships {
		if rel.Type != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" {
			filteredRels = append(filteredRels, rel)
		}
	}

	d.documentRelationships.Relationships = filteredRels
	Debugf("文档关系解析完成，共 %d 个关系", len(filteredRels))
	return nil
}

// updateNextImageID 根据已有的图片关系更新nextImageID计数器
// 确保新添加的图片ID不会与现有图片冲突
func (d *Document) updateNextImageID() {
	maxImageID := -1

	// 遍历所有parts，查找已存在的图片文件的最大ID
	for partName := range d.parts {
		// 检查是否是图片文件（word/media/imageN.xxx）
		if len(partName) > 11 && partName[:11] == "word/media/" {
			// 从文件名中提取图片ID（image0.png -> 0, image1.png -> 1等）
			filename := partName[11:] // 去掉"word/media/"前缀
			var id int
			if _, err := fmt.Sscanf(filename, "image%d.", &id); err == nil {
				if id > maxImageID {
					maxImageID = id
				}
			}
		}
	}

	// 设置nextImageID为最大图片ID + 1
	// 如果没有现有图片，maxImageID为-1，nextImageID应该为0
	d.nextImageID = maxImageID + 1

	Debugf("更新图片ID计数器: nextImageID = %d", d.nextImageID)
}

// ToBytes 将文档转换为字节数组
func (d *Document) ToBytes() ([]byte, error) {
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)

	// serialize文档
	if err := d.serializeDocument(); err != nil {
		return nil, err
	}

	// serialize样式
	if err := d.serializeStyles(); err != nil {
		return nil, err
	}

	// serialize内容类型
	d.serializeContentTypes()

	// serialize关系
	d.serializeRelationships()

	// serialize文档关系
	d.serializeDocumentRelationships()

	// 写入所有部件
	for name, data := range d.parts {
		writer, err := zipWriter.Create(name)
		if err != nil {
			return nil, err
		}
		if _, err := writer.Write(data); err != nil {
			return nil, err
		}
	}

	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// GetParagraphs 获取所有段落
func (b *Body) GetParagraphs() []*Paragraph {
	paragraphs := make([]*Paragraph, 0)
	for _, element := range b.Elements {
		if p, ok := element.(*Paragraph); ok {
			paragraphs = append(paragraphs, p)
		}
	}
	return paragraphs
}

// GetTables 获取所有表格
func (b *Body) GetTables() []*Table {
	tables := make([]*Table, 0)
	for _, element := range b.Elements {
		if t, ok := element.(*Table); ok {
			tables = append(tables, t)
		}
	}
	return tables
}

// AddElement 添加元素到文档主体
func (b *Body) AddElement(element interface{}) {
	b.Elements = append(b.Elements, element)
}

// RemoveParagraph 从文档中删除指定的段落。
//
// 参数 paragraph 是要删除的段落对象。
// 如果段落不存在于文档中，此方法不会产生任何效果。
//
// 返回值表示是否成功删除段落。
//
// Example:
//
//	doc := document.New()
//	para := doc.AddParagraph("要删除的段落")
//	doc.RemoveParagraph(para)
func (d *Document) RemoveParagraph(paragraph *Paragraph) bool {
	for i, element := range d.Body.Elements {
		if p, ok := element.(*Paragraph); ok && p == paragraph {
			// 删除元素
			d.Body.Elements = append(d.Body.Elements[:i], d.Body.Elements[i+1:]...)
			Debugf("删除段落: 索引 %d", i)
			return true
		}
	}
	Debugf("警告：未找到要删除的段落")
	return false
}

// RemoveParagraphAt 根据索引删除段落。
//
// 参数 index 是要删除的段落在所有段落中的索引（从0开始）。
// 如果索引超出范围，此方法会返回错误。
//
// 返回值表示是否成功删除段落。
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("第一段")
//	doc.AddParagraph("第二段")
//	doc.RemoveParagraphAt(0)  // 删除第一段
func (d *Document) RemoveParagraphAt(index int) bool {
	// 提前验证负数索引
	if index < 0 {
		Debugf("错误：段落索引不能为负数: %d", index)
		return false
	}

	// 优化：单次遍历找到目标段落及其元素索引
	paragraphCount := 0
	for i, element := range d.Body.Elements {
		if _, ok := element.(*Paragraph); ok {
			if paragraphCount == index {
				// 找到目标段落，删除它
				d.Body.Elements = append(d.Body.Elements[:i], d.Body.Elements[i+1:]...)
				Debugf("删除段落: 段落索引 %d, 元素索引 %d", index, i)
				return true
			}
			paragraphCount++
		}
	}

	Debugf("错误：段落索引 %d 超出范围 [0, %d)", index, paragraphCount)
	return false
}

// RemoveElementAt 根据元素索引删除元素（包括段落、表格等）。
//
// 参数 index 是要删除的元素在文档主体中的索引（从0开始）。
// 如果索引超出范围，此方法会返回错误。
//
// 返回值表示是否成功删除元素。
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("段落")
//	doc.AddTable(&document.TableConfig{Rows: 2, Cols: 2})
//	doc.RemoveElementAt(0)  // 删除第一个元素（段落）
func (d *Document) RemoveElementAt(index int) bool {
	if index < 0 || index >= len(d.Body.Elements) {
		Debugf("错误：元素索引 %d 超出范围 [0, %d)", index, len(d.Body.Elements))
		return false
	}

	// 删除元素
	d.Body.Elements = append(d.Body.Elements[:index], d.Body.Elements[index+1:]...)
	Debugf("删除元素: 索引 %d", index)
	return true
}

// parseTableBorders 解析表格边框
func (d *Document) parseTableBorders(decoder *xml.Decoder) (*TableBorders, error) {
	borders := &TableBorders{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_borders", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			border := &TableBorder{
				Val:        getAttributeValue(t.Attr, "val"),
				Sz:         getAttributeValue(t.Attr, "sz"),
				Space:      getAttributeValue(t.Attr, "space"),
				Color:      getAttributeValue(t.Attr, "color"),
				ThemeColor: getAttributeValue(t.Attr, "themeColor"),
			}

			switch t.Name.Local {
			case "top":
				borders.Top = border
			case "left":
				borders.Left = border
			case "bottom":
				borders.Bottom = border
			case "right":
				borders.Right = border
			case "insideH":
				borders.InsideH = border
			case "insideV":
				borders.InsideV = border
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tblBorders" {
				return borders, nil
			}
		}
	}
}

// parseTableCellMargins 解析表格单元格边距
func (d *Document) parseTableCellMargins(decoder *xml.Decoder) (*TableCellMargins, error) {
	margins := &TableCellMargins{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_margins", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			space := &TableCellSpace{
				W:    getAttributeValue(t.Attr, "w"),
				Type: getAttributeValue(t.Attr, "type"),
			}

			switch t.Name.Local {
			case "top":
				margins.Top = space
			case "left":
				margins.Left = space
			case "bottom":
				margins.Bottom = space
			case "right":
				margins.Right = space
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tblCellMar" {
				return margins, nil
			}
		}
	}
}

// parseTableCellProperties 解析表格单元格属性
func (d *Document) parseTableCellProperties(decoder *xml.Decoder) (*TableCellProperties, error) {
	props := &TableCellProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcW":
				// 解析单元格宽度
				w := getAttributeValue(t.Attr, "w")
				wType := getAttributeValue(t.Attr, "type")
				if w != "" || wType != "" {
					props.TableCellW = &TableCellW{W: w, Type: wType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "vAlign":
				// 解析垂直对齐
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					props.VAlign = &VAlign{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "gridSpan":
				// 解析网格跨度
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					props.GridSpan = &GridSpan{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "vMerge":
				// 解析垂直合并
				val := getAttributeValue(t.Attr, "val")
				props.VMerge = &VMerge{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "textDirection":
				// 解析文字方向
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					props.TextDirection = &TextDirection{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "shd":
				// 解析单元格底纹
				shd := &TableCellShading{
					Val:       getAttributeValue(t.Attr, "val"),
					Color:     getAttributeValue(t.Attr, "color"),
					Fill:      getAttributeValue(t.Attr, "fill"),
					ThemeFill: getAttributeValue(t.Attr, "themeFill"),
				}
				props.Shd = shd
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "tcBorders":
				// 解析单元格边框
				borders, err := d.parseTableCellBorders(decoder)
				if err != nil {
					return nil, err
				}
				props.TcBorders = borders
			case "tcMar":
				// 解析单元格边距
				margins, err := d.parseTableCellMarginsCell(decoder)
				if err != nil {
					return nil, err
				}
				props.TcMar = margins
			case "noWrap":
				// 解析禁止换行
				val := getAttributeValue(t.Attr, "val")
				props.NoWrap = &NoWrap{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "hideMark":
				// 解析隐藏标记
				val := getAttributeValue(t.Attr, "val")
				props.HideMark = &HideMark{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// 跳过其他未处理的单元格属性
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tcPr" {
				return props, nil
			}
		}
	}
}

// parseTableCellBorders 解析表格单元格边框
func (d *Document) parseTableCellBorders(decoder *xml.Decoder) (*TableCellBorders, error) {
	borders := &TableCellBorders{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_borders", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			border := &TableCellBorder{
				Val:        getAttributeValue(t.Attr, "val"),
				Sz:         getAttributeValue(t.Attr, "sz"),
				Space:      getAttributeValue(t.Attr, "space"),
				Color:      getAttributeValue(t.Attr, "color"),
				ThemeColor: getAttributeValue(t.Attr, "themeColor"),
			}

			switch t.Name.Local {
			case "top":
				borders.Top = border
			case "left":
				borders.Left = border
			case "bottom":
				borders.Bottom = border
			case "right":
				borders.Right = border
			case "insideH":
				borders.InsideH = border
			case "insideV":
				borders.InsideV = border
			case "tl2br":
				borders.TL2BR = border
			case "tr2bl":
				borders.TR2BL = border
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tcBorders" {
				return borders, nil
			}
		}
	}
}

// parseTableCellMarginsCell 解析表格单元格边距（单元格级别）
func (d *Document) parseTableCellMarginsCell(decoder *xml.Decoder) (*TableCellMarginsCell, error) {
	margins := &TableCellMarginsCell{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_margins_cell", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			space := &TableCellSpaceCell{
				W:    getAttributeValue(t.Attr, "w"),
				Type: getAttributeValue(t.Attr, "type"),
			}

			switch t.Name.Local {
			case "top":
				margins.Top = space
			case "left":
				margins.Left = space
			case "bottom":
				margins.Bottom = space
			case "right":
				margins.Right = space
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tcMar" {
				return margins, nil
			}
		}
	}
}

// parseTableRowProperties 解析表格行属性
func (d *Document) parseTableRowProperties(decoder *xml.Decoder) (*TableRowProperties, error) {
	props := &TableRowProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_row_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "trHeight":
				// 解析行高
				val := getAttributeValue(t.Attr, "val")
				hRule := getAttributeValue(t.Attr, "hRule")
				if val != "" || hRule != "" {
					props.TableRowH = &TableRowH{Val: val, HRule: hRule}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "cantSplit":
				// 解析禁止跨页分割
				val := getAttributeValue(t.Attr, "val")
				props.CantSplit = &CantSplit{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "tblHeader":
				// 解析标题行重复
				val := getAttributeValue(t.Attr, "val")
				props.TblHeader = &TblHeader{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// 跳过其他行属性
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "trPr" {
				return props, nil
			}
		}
	}
}

// parseDrawingElement 解析绘图元素（图片等）
// 此方法用于从XML中解析完整的绘图元素结构
func (d *Document) parseDrawingElement(decoder *xml.Decoder, startElement xml.StartElement) (*DrawingElement, error) {
	drawing := &DrawingElement{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_drawing_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "inline":
				// 解析嵌入式绘图
				inline, err := d.parseInlineDrawing(decoder, t)
				if err != nil {
					return nil, err
				}
				drawing.Inline = inline
			case "anchor":
				// 解析浮动绘图
				anchor, err := d.parseAnchorDrawing(decoder, t)
				if err != nil {
					return nil, err
				}
				drawing.Anchor = anchor
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "drawing" {
				return drawing, nil
			}
		}
	}
}

// parseInlineDrawing 解析嵌入式绘图
func (d *Document) parseInlineDrawing(decoder *xml.Decoder, startElement xml.StartElement) (*InlineDrawing, error) {
	inline := &InlineDrawing{}

	// parse properties
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case "distT":
			inline.DistT = attr.Value
		case "distB":
			inline.DistB = attr.Value
		case "distL":
			inline.DistL = attr.Value
		case "distR":
			inline.DistR = attr.Value
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_inline_drawing", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "extent":
				extent := &DrawingExtent{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						extent.Cx = attr.Value
					case "cy":
						extent.Cy = attr.Value
					}
				}
				inline.Extent = extent
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "docPr":
				docPr := &DrawingDocPr{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						docPr.ID = attr.Value
					case "name":
						docPr.Name = attr.Value
					case "descr":
						docPr.Descr = attr.Value
					case "title":
						docPr.Title = attr.Value
					}
				}
				inline.DocPr = docPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "graphic":
				graphic, err := d.parseDrawingGraphic(decoder, t)
				if err != nil {
					return nil, err
				}
				inline.Graphic = graphic
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "inline" {
				return inline, nil
			}
		}
	}
}

// parseAnchorDrawing 解析浮动绘图
func (d *Document) parseAnchorDrawing(decoder *xml.Decoder, startElement xml.StartElement) (*AnchorDrawing, error) {
	anchor := &AnchorDrawing{}

	// parse properties
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case "distT":
			anchor.DistT = attr.Value
		case "distB":
			anchor.DistB = attr.Value
		case "distL":
			anchor.DistL = attr.Value
		case "distR":
			anchor.DistR = attr.Value
		case "simplePos":
			anchor.SimplePos = attr.Value
		case "relativeHeight":
			anchor.RelativeHeight = attr.Value
		case "behindDoc":
			anchor.BehindDoc = attr.Value
		case "locked":
			anchor.Locked = attr.Value
		case "layoutInCell":
			anchor.LayoutInCell = attr.Value
		case "allowOverlap":
			anchor.AllowOverlap = attr.Value
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_anchor_drawing", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "extent":
				extent := &DrawingExtent{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						extent.Cx = attr.Value
					case "cy":
						extent.Cy = attr.Value
					}
				}
				anchor.Extent = extent
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "docPr":
				docPr := &DrawingDocPr{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						docPr.ID = attr.Value
					case "name":
						docPr.Name = attr.Value
					case "descr":
						docPr.Descr = attr.Value
					case "title":
						docPr.Title = attr.Value
					}
				}
				anchor.DocPr = docPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "graphic":
				graphic, err := d.parseDrawingGraphic(decoder, t)
				if err != nil {
					return nil, err
				}
				anchor.Graphic = graphic
			case "wrapNone":
				anchor.WrapNone = &WrapNone{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "wrapSquare":
				wrapSquare := &WrapSquare{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "wrapText":
						wrapSquare.WrapText = attr.Value
					case "distT":
						wrapSquare.DistT = attr.Value
					case "distB":
						wrapSquare.DistB = attr.Value
					case "distL":
						wrapSquare.DistL = attr.Value
					case "distR":
						wrapSquare.DistR = attr.Value
					}
				}
				anchor.WrapSquare = wrapSquare
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "anchor" {
				return anchor, nil
			}
		}
	}
}

// parseDrawingGraphic 解析绘图图形元素
func (d *Document) parseDrawingGraphic(decoder *xml.Decoder, startElement xml.StartElement) (*DrawingGraphic, error) {
	graphic := &DrawingGraphic{}

	// 解析xmlns属性
	for _, attr := range startElement.Attr {
		// 检查xmlns属性（命名空间声明）
		if attr.Name.Space == "xmlns" || (attr.Name.Space == "" && strings.HasPrefix(attr.Name.Local, "xmlns")) {
			if attr.Value == "http://schemas.openxmlformats.org/drawingml/2006/main" {
				graphic.Xmlns = attr.Value
			}
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_drawing_graphic", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "graphicData":
				graphicData, err := d.parseGraphicData(decoder, t)
				if err != nil {
					return nil, err
				}
				graphic.GraphicData = graphicData
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "graphic" {
				return graphic, nil
			}
		}
	}
}

// parseGraphicData 解析图形数据元素
func (d *Document) parseGraphicData(decoder *xml.Decoder, startElement xml.StartElement) (*GraphicData, error) {
	graphicData := &GraphicData{}

	// parse properties
	for _, attr := range startElement.Attr {
		if attr.Name.Local == "uri" {
			graphicData.Uri = attr.Value
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_graphic_data", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pic":
				pic, err := d.parsePicElement(decoder, t)
				if err != nil {
					return nil, err
				}
				graphicData.Pic = pic
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "graphicData" {
				return graphicData, nil
			}
		}
	}
}

// parsePicElement 解析图片元素
func (d *Document) parsePicElement(decoder *xml.Decoder, startElement xml.StartElement) (*PicElement, error) {
	pic := &PicElement{}

	// 解析xmlns属性
	for _, attr := range startElement.Attr {
		// 检查xmlns属性（命名空间声明）
		if attr.Name.Space == "xmlns" || (attr.Name.Space == "" && strings.HasPrefix(attr.Name.Local, "xmlns")) {
			if attr.Value == "http://schemas.openxmlformats.org/drawingml/2006/picture" {
				pic.Xmlns = attr.Value
			}
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_pic_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "nvPicPr":
				nvPicPr, err := d.parseNvPicPr(decoder, t)
				if err != nil {
					return nil, err
				}
				pic.NvPicPr = nvPicPr
			case "blipFill":
				blipFill, err := d.parseBlipFill(decoder, t)
				if err != nil {
					return nil, err
				}
				pic.BlipFill = blipFill
			case "spPr":
				spPr, err := d.parseSpPr(decoder, t)
				if err != nil {
					return nil, err
				}
				pic.SpPr = spPr
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "pic" {
				return pic, nil
			}
		}
	}
}

// parseNvPicPr 解析非可视图片属性
func (d *Document) parseNvPicPr(decoder *xml.Decoder, startElement xml.StartElement) (*NvPicPr, error) {
	nvPicPr := &NvPicPr{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_nv_pic_pr", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cNvPr":
				cNvPr := &CNvPr{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						cNvPr.ID = attr.Value
					case "name":
						cNvPr.Name = attr.Value
					case "descr":
						cNvPr.Descr = attr.Value
					case "title":
						cNvPr.Title = attr.Value
					}
				}
				nvPicPr.CNvPr = cNvPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "cNvPicPr":
				cNvPicPr := &CNvPicPr{}
				// 解析picLocks如果存在
				nvPicPr.CNvPicPr = cNvPicPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "nvPicPr" {
				return nvPicPr, nil
			}
		}
	}
}

// parseBlipFill 解析图片填充
func (d *Document) parseBlipFill(decoder *xml.Decoder, startElement xml.StartElement) (*BlipFill, error) {
	blipFill := &BlipFill{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_blip_fill", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "blip":
				blip := &Blip{}
				for _, attr := range t.Attr {
					if attr.Name.Local == "embed" {
						blip.Embed = attr.Value
					}
				}
				blipFill.Blip = blip
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "stretch":
				blipFill.Stretch = &Stretch{FillRect: &FillRect{}}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "blipFill" {
				return blipFill, nil
			}
		}
	}
}

// parseSpPr 解析形状属性
func (d *Document) parseSpPr(decoder *xml.Decoder, startElement xml.StartElement) (*SpPr, error) {
	spPr := &SpPr{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_sp_pr", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "xfrm":
				xfrm, err := d.parseXfrm(decoder, t)
				if err != nil {
					return nil, err
				}
				spPr.Xfrm = xfrm
			case "prstGeom":
				prstGeom := &PrstGeom{AvLst: &AvLst{}}
				for _, attr := range t.Attr {
					if attr.Name.Local == "prst" {
						prstGeom.Prst = attr.Value
					}
				}
				spPr.PrstGeom = prstGeom
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "spPr" {
				return spPr, nil
			}
		}
	}
}

// parseXfrm 解析变换元素
func (d *Document) parseXfrm(decoder *xml.Decoder, startElement xml.StartElement) (*Xfrm, error) {
	xfrm := &Xfrm{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_xfrm", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "off":
				off := &Off{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "x":
						off.X = attr.Value
					case "y":
						off.Y = attr.Value
					}
				}
				xfrm.Off = off
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "ext":
				ext := &Ext{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						ext.Cx = attr.Value
					case "cy":
						ext.Cy = attr.Value
					}
				}
				xfrm.Ext = ext
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "xfrm" {
				return xfrm, nil
			}
		}
	}
}
