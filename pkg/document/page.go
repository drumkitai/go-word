// Package document provides page settings functionality
package document

import (
	"encoding/xml"
	"errors"
	"fmt"
	"strconv"
)

// PageOrientation page orientation type
type PageOrientation string

const (
	// OrientationPortrait
	OrientationPortrait PageOrientation = "portrait"
	// OrientationLandscape
	OrientationLandscape PageOrientation = "landscape"
)

// DocGridType
type DocGridType string

const (
	// DocGridDefault
	DocGridDefault DocGridType = "default"
	// DocGridLines
	DocGridLines DocGridType = "lines"
	// DocGridSnapToChars
	DocGridSnapToChars DocGridType = "snapToChars"
	// DocGridSnapToLines
	DocGridSnapToLines DocGridType = "snapToLines"
)

// PageSize
type PageSize string

const (
	// PageSizeA4 (210mm x 297mm)
	PageSizeA4 PageSize = "A4"
	// PageSizeLetter US Letter (8.5" x 11")
	PageSizeLetter PageSize = "Letter"
	// PageSizeLegal US Legal (8.5" x 14")
	PageSizeLegal PageSize = "Legal"
	// PageSizeA3 A3 paper (297mm x 420mm)
	PageSizeA3 PageSize = "A3"
	// PageSizeA5 A5 paper (148mm x 210mm)
	PageSizeA5 PageSize = "A5"
	// PageSizeCustom custom size
	PageSizeCustom PageSize = "Custom"
)

// Page settings related errors
var (
	// ErrInvalidPageSettings
	ErrInvalidPageSettings = errors.New("invalid page settings")
)

// SectionProperties contains page settings information
type SectionProperties struct {
	XMLName          xml.Name                 `xml:"w:sectPr"`
	XmlnsR           string                   `xml:"xmlns:r,attr,omitempty"`
	PageSize         *PageSizeXML             `xml:"w:pgSz,omitempty"`
	PageMargins      *PageMargin              `xml:"w:pgMar,omitempty"`
	Columns          *Columns                 `xml:"w:cols,omitempty"`
	HeaderReferences []*HeaderFooterReference `xml:"w:headerReference,omitempty"`
	FooterReferences []*FooterReference       `xml:"w:footerReference,omitempty"`
	TitlePage        *TitlePage               `xml:"w:titlePg,omitempty"`
	PageNumType      *PageNumType             `xml:"w:pgNumType,omitempty"`
	DocGrid          *DocGrid                 `xml:"w:docGrid,omitempty"`
}

// PageSizeXML page size XML structure
type PageSizeXML struct {
	XMLName xml.Name `xml:"w:pgSz"`
	W       string   `xml:"w:w,attr"`      // page width (twips)
	H       string   `xml:"w:h,attr"`      // page height (twips)
	Orient  string   `xml:"w:orient,attr"` // page orientation
}

// PageMargin page margin
type PageMargin struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top     string   `xml:"w:top,attr"`    // top margin (twips)
	Right   string   `xml:"w:right,attr"`  // right margin (twips)
	Bottom  string   `xml:"w:bottom,attr"` // bottom margin (twips)
	Left    string   `xml:"w:left,attr"`   // left margin (twips)
	Header  string   `xml:"w:header,attr"` // header margin (twips)
	Footer  string   `xml:"w:footer,attr"` // footer margin (twips)
	Gutter  string   `xml:"w:gutter,attr"` // gutter margin (twips)
}

// Columns column settings
type Columns struct {
	XMLName xml.Name `xml:"w:cols"`
	Space   string   `xml:"w:space,attr,omitempty"` // column spacing
	Num     string   `xml:"w:num,attr,omitempty"`   // column count
}

// PageNumType page number type
type PageNumType struct {
	XMLName xml.Name `xml:"w:pgNumType"`
	Fmt     string   `xml:"w:fmt,attr,omitempty"`
}

// PageSettings page settings configuration
type PageSettings struct {
	// page size
	Size PageSize
	// custom size (when Size is Custom)
	CustomWidth  float64 // custom width (mm)
	CustomHeight float64 // custom height (mm)
	// page orientation
	Orientation PageOrientation
	// page margins (mm)
	MarginTop    float64
	MarginRight  float64
	MarginBottom float64
	MarginLeft   float64
	// header footer distance (mm)
	HeaderDistance float64
	FooterDistance float64
	// gutter width (mm)
	GutterWidth float64
	// document grid settings
	DocGridType      DocGridType // document grid type
	DocGridLinePitch int         // line grid pitch (1/20 pt)
	DocGridCharSpace int         // character spacing
}

// predefined page sizes (mm)
var predefinedSizes = map[PageSize]struct {
	width  float64
	height float64
}{
	PageSizeA4:     {210, 297},
	PageSizeLetter: {215.9, 279.4}, // US Letter (8.5" x 11")
	PageSizeLegal:  {215.9, 355.6}, // US Legal (8.5" x 14")
	PageSizeA3:     {297, 420},
	PageSizeA5:     {148, 210},
}

// DefaultPageSettings return default page settings (A4 portrait)
func DefaultPageSettings() *PageSettings {
	return &PageSettings{
		Size:             PageSizeA4,
		Orientation:      OrientationPortrait,
		MarginTop:        25.4, // 1 inch
		MarginRight:      25.4, // 1 inch
		MarginBottom:     25.4, // 1 inch
		MarginLeft:       25.4, // 1 inch
		HeaderDistance:   12.7, // 0.5 inch
		FooterDistance:   12.7, // 0.5 inch
		GutterWidth:      0,    // no gutter
		DocGridType:      DocGridLines,
		DocGridLinePitch: 312, // default line grid pitch
		DocGridCharSpace: 0,
	}
}

// SetPageSettings set document page properties
func (d *Document) SetPageSettings(settings *PageSettings) error {
	if settings == nil {
		return WrapError("SetPageSettings", errors.New("page settings cannot be empty"))
	}

	// validate page settings
	if err := validatePageSettings(settings); err != nil {
		return WrapError("SetPageSettings", err)
	}

	// get or create section properties
	sectPr := d.getSectionProperties()

	// set page size
	width, height := getPageDimensions(settings)
	sectPr.PageSize = &PageSizeXML{
		W:      fmt.Sprintf("%.0f", mmToTwips(width)),
		H:      fmt.Sprintf("%.0f", mmToTwips(height)),
		Orient: string(settings.Orientation),
	}

	// 设置页面边距
	sectPr.PageMargins = &PageMargin{
		Top:    fmt.Sprintf("%.0f", mmToTwips(settings.MarginTop)),
		Right:  fmt.Sprintf("%.0f", mmToTwips(settings.MarginRight)),
		Bottom: fmt.Sprintf("%.0f", mmToTwips(settings.MarginBottom)),
		Left:   fmt.Sprintf("%.0f", mmToTwips(settings.MarginLeft)),
		Header: fmt.Sprintf("%.0f", mmToTwips(settings.HeaderDistance)),
		Footer: fmt.Sprintf("%.0f", mmToTwips(settings.FooterDistance)),
		Gutter: fmt.Sprintf("%.0f", mmToTwips(settings.GutterWidth)),
	}

	// 设置文档网格
	if settings.DocGridType != "" {
		sectPr.DocGrid = &DocGrid{
			Type:      string(settings.DocGridType),
			LinePitch: strconv.Itoa(settings.DocGridLinePitch),
		}

		if settings.DocGridCharSpace > 0 {
			sectPr.DocGrid.CharSpace = strconv.Itoa(settings.DocGridCharSpace)
		}
	}

	Infof("页面设置已更新: 尺寸=%s, 方向=%s", settings.Size, settings.Orientation)
	return nil
}

// GetPageSettings 获取当前文档的页面设置
func (d *Document) GetPageSettings() *PageSettings {
	sectPr := d.getSectionProperties()
	settings := DefaultPageSettings()

	if sectPr.PageSize != nil {
		// 解析页面尺寸
		width := twipsToMM(parseFloat(sectPr.PageSize.W))
		height := twipsToMM(parseFloat(sectPr.PageSize.H))

		// 判断是否为预定义尺寸
		settings.Size = identifyPageSize(width, height)
		if settings.Size == PageSizeCustom {
			settings.CustomWidth = width
			settings.CustomHeight = height
		}

		// 设置方向
		if sectPr.PageSize.Orient == string(OrientationLandscape) {
			settings.Orientation = OrientationLandscape
		} else {
			settings.Orientation = OrientationPortrait
		}
	}

	if sectPr.PageMargins != nil {
		// 解析页面边距
		settings.MarginTop = twipsToMM(parseFloat(sectPr.PageMargins.Top))
		settings.MarginRight = twipsToMM(parseFloat(sectPr.PageMargins.Right))
		settings.MarginBottom = twipsToMM(parseFloat(sectPr.PageMargins.Bottom))
		settings.MarginLeft = twipsToMM(parseFloat(sectPr.PageMargins.Left))
		settings.HeaderDistance = twipsToMM(parseFloat(sectPr.PageMargins.Header))
		settings.FooterDistance = twipsToMM(parseFloat(sectPr.PageMargins.Footer))
		settings.GutterWidth = twipsToMM(parseFloat(sectPr.PageMargins.Gutter))
	}

	// 解析文档网格设置
	if sectPr.DocGrid != nil {
		if sectPr.DocGrid.Type != "" {
			settings.DocGridType = DocGridType(sectPr.DocGrid.Type)
		}

		if sectPr.DocGrid.LinePitch != "" {
			settings.DocGridLinePitch = int(parseFloat(sectPr.DocGrid.LinePitch))
		}

		if sectPr.DocGrid.CharSpace != "" {
			settings.DocGridCharSpace = int(parseFloat(sectPr.DocGrid.CharSpace))
		}
	}

	return settings
}

// SetPageSize 设置页面大小
func (d *Document) SetPageSize(size PageSize) error {
	settings := d.GetPageSettings()
	settings.Size = size
	return d.SetPageSettings(settings)
}

// SetCustomPageSize 设置自定义页面大小（毫米）
func (d *Document) SetCustomPageSize(width, height float64) error {
	if width <= 0 || height <= 0 {
		return WrapError("SetCustomPageSize", errors.New("页面尺寸必须大于0"))
	}

	settings := d.GetPageSettings()
	settings.Size = PageSizeCustom
	settings.CustomWidth = width
	settings.CustomHeight = height
	return d.SetPageSettings(settings)
}

// SetPageOrientation 设置页面方向
func (d *Document) SetPageOrientation(orientation PageOrientation) error {
	settings := d.GetPageSettings()
	settings.Orientation = orientation
	return d.SetPageSettings(settings)
}

// SetPageMargins 设置页面边距（毫米）
func (d *Document) SetPageMargins(top, right, bottom, left float64) error {
	if top < 0 || right < 0 || bottom < 0 || left < 0 {
		return WrapError("SetPageMargins", errors.New("页面边距不能为负数"))
	}

	settings := d.GetPageSettings()
	settings.MarginTop = top
	settings.MarginRight = right
	settings.MarginBottom = bottom
	settings.MarginLeft = left
	return d.SetPageSettings(settings)
}

// SetHeaderFooterDistance 设置页眉页脚距离（毫米）
func (d *Document) SetHeaderFooterDistance(header, footer float64) error {
	if header < 0 || footer < 0 {
		return WrapError("SetHeaderFooterDistance", errors.New("页眉页脚距离不能为负数"))
	}

	settings := d.GetPageSettings()
	settings.HeaderDistance = header
	settings.FooterDistance = footer
	return d.SetPageSettings(settings)
}

// SetGutterWidth 设置装订线宽度（毫米）
func (d *Document) SetGutterWidth(width float64) error {
	if width < 0 {
		return WrapError("SetGutterWidth", errors.New("装订线宽度不能为负数"))
	}

	settings := d.GetPageSettings()
	settings.GutterWidth = width
	return d.SetPageSettings(settings)
}

// getSectionProperties get or create section properties
func (d *Document) getSectionProperties() *SectionProperties {
	if d.Body == nil {
		return &SectionProperties{}
	}

	// find existing SectionProperties in Elements (may be in any position)
	for _, element := range d.Body.Elements {
		if sectPr, ok := element.(*SectionProperties); ok {
			return sectPr
		}
	}

	// if not found, create new section properties and add to the end
	sectPr := &SectionProperties{}
	d.Body.Elements = append(d.Body.Elements, sectPr)

	return sectPr
}

// setSectionProperties 替换或设置节属性
func (d *Document) setSectionProperties(sectPr *SectionProperties) {
	if sectPr == nil {
		return
	}

	if d.Body == nil {
		d.Body = &Body{Elements: []interface{}{sectPr}}
		return
	}

	for i, element := range d.Body.Elements {
		if _, ok := element.(*SectionProperties); ok {
			d.Body.Elements[i] = sectPr
			return
		}
	}

	d.Body.Elements = append(d.Body.Elements, sectPr)
}

// ElementType 返回节属性元素类型
func (s *SectionProperties) ElementType() string {
	return "sectionProperties"
}

// validatePageSettings 验证页面设置
func validatePageSettings(settings *PageSettings) error {
	// 验证页面尺寸
	if settings.Size == PageSizeCustom {
		if settings.CustomWidth <= 0 || settings.CustomHeight <= 0 {
			return errors.New("自定义页面尺寸必须大于0")
		}

		// 检查尺寸范围（Word支持的最小和最大尺寸）
		const minSize = 12.7  // 0.5英寸
		const maxSize = 558.8 // 22英寸

		if settings.CustomWidth < minSize || settings.CustomWidth > maxSize ||
			settings.CustomHeight < minSize || settings.CustomHeight > maxSize {
			return fmt.Errorf("页面尺寸必须在%.1f-%.1fmm范围内", minSize, maxSize)
		}
	}

	// 验证方向
	if settings.Orientation != OrientationPortrait && settings.Orientation != OrientationLandscape {
		return errors.New("无效的页面方向")
	}

	return nil
}

// getPageDimensions 获取页面尺寸（毫米）
func getPageDimensions(settings *PageSettings) (width, height float64) {
	if settings.Size == PageSizeCustom {
		width = settings.CustomWidth
		height = settings.CustomHeight
	} else {
		size, exists := predefinedSizes[settings.Size]
		if !exists {
			// 默认使用A4
			size = predefinedSizes[PageSizeA4]
		}
		width = size.width
		height = size.height
	}

	// 如果是横向，交换宽高
	if settings.Orientation == OrientationLandscape {
		width, height = height, width
	}

	return width, height
}

// identifyPageSize 根据尺寸识别页面类型
func identifyPageSize(width, height float64) PageSize {
	// 允许1mm的误差
	const tolerance = 1.0

	for size, dims := range predefinedSizes {
		if (abs(width-dims.width) < tolerance && abs(height-dims.height) < tolerance) ||
			(abs(width-dims.height) < tolerance && abs(height-dims.width) < tolerance) {
			return size
		}
	}

	return PageSizeCustom
}

// mmToTwips 毫米转换为Twips（1毫米 = 56.69 twips）
func mmToTwips(mm float64) float64 {
	return mm * 56.692913385827
}

// twipsToMM Twips转换为毫米
func twipsToMM(twips float64) float64 {
	return twips / 56.692913385827
}

// parseFloat 安全地解析浮点数字符串
func parseFloat(s string) float64 {
	if s == "" {
		return 0
	}

	// 使用strconv.ParseFloat解析浮点数
	if val, err := strconv.ParseFloat(s, 64); err == nil {
		return val
	}

	return 0
}

// abs 返回浮点数的绝对值
func abs(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}

// DocGrid 文档网格设置
type DocGrid struct {
	XMLName   xml.Name `xml:"w:docGrid"`
	Type      string   `xml:"w:type,attr,omitempty"`      // 网格类型
	LinePitch string   `xml:"w:linePitch,attr,omitempty"` // 行网格间距
	CharSpace string   `xml:"w:charSpace,attr,omitempty"` // 字符间距
}

// SetDocGrid 设置文档网格
func (d *Document) SetDocGrid(gridType DocGridType, linePitch int, charSpace int) error {
	if gridType == "" {
		return WrapError("SetDocGrid", errors.New("网格类型不能为空"))
	}

	settings := d.GetPageSettings()
	settings.DocGridType = gridType
	settings.DocGridLinePitch = linePitch
	settings.DocGridCharSpace = charSpace
	return d.SetPageSettings(settings)
}

// ClearDocGrid 清除文档网格设置
func (d *Document) ClearDocGrid() error {
	sectPr := d.getSectionProperties()
	sectPr.DocGrid = nil
	return nil
}
