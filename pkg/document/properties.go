// Package document provides Word document properties functionality
package document

import (
	"encoding/xml"
	"fmt"
	"strings"
	"time"
)

// DocumentProperties document properties
type DocumentProperties struct {
	Title        string
	Subject      string
	Creator      string
	Keywords     string
	Description  string
	Language     string
	Category     string
	Version      string
	Revision     string
	Created      time.Time
	LastModified time.Time
	LastPrinted  time.Time
	Pages        int
	Words        int
	Characters   int
	Paragraphs   int
	Lines        int
}

// CoreProperties core properties XML structure
type CoreProperties struct {
	XMLName       xml.Name `xml:"cp:coreProperties"`
	XmlnsCP       string   `xml:"xmlns:cp,attr"`
	XmlnsDC       string   `xml:"xmlns:dc,attr"`
	XmlnsDCTerms  string   `xml:"xmlns:dcterms,attr"`
	XmlnsDCMIType string   `xml:"xmlns:dcmitype,attr"`
	XmlnsXSI      string   `xml:"xmlns:xsi,attr"`
	Title         *DCText  `xml:"dc:title,omitempty"`
	Subject       *DCText  `xml:"dc:subject,omitempty"`
	Creator       *DCText  `xml:"dc:creator,omitempty"`
	Keywords      *CPText  `xml:"cp:keywords,omitempty"`
	Description   *DCText  `xml:"dc:description,omitempty"`
	Language      *DCText  `xml:"dc:language,omitempty"`
	Category      *CPText  `xml:"cp:category,omitempty"`
	Version       *CPText  `xml:"cp:version,omitempty"`
	Revision      *CPText  `xml:"cp:revision,omitempty"`
	Created       *DCDate  `xml:"dcterms:created,omitempty"`
	Modified      *DCDate  `xml:"dcterms:modified,omitempty"`
	LastPrinted   *DCDate  `xml:"cp:lastPrinted,omitempty"`
}

// AppProperties application properties XML structure
type AppProperties struct {
	XMLName       xml.Name `xml:"Properties"`
	Xmlns         string   `xml:"xmlns,attr"`
	XmlnsVT       string   `xml:"xmlns:vt,attr"`
	Application   string   `xml:"Application,omitempty"`
	DocSecurity   int      `xml:"DocSecurity,omitempty"`
	ScaleCrop     bool     `xml:"ScaleCrop,omitempty"`
	LinksUpToDate bool     `xml:"LinksUpToDate,omitempty"`
	Pages         int      `xml:"Pages,omitempty"`
	Words         int      `xml:"Words,omitempty"`
	Characters    int      `xml:"Characters,omitempty"`
	Paragraphs    int      `xml:"Paragraphs,omitempty"`
	Lines         int      `xml:"Lines,omitempty"`
}

// DCText DC namespace text element
type DCText struct {
	Text string `xml:",chardata"`
}

// CPText CP namespace text element
type CPText struct {
	Text string `xml:",chardata"`
}

// DCDate DC命名空间日期元素
type DCDate struct {
	XSIType string    `xml:"xsi:type,attr"`
	Date    time.Time `xml:",chardata"`
}

// SetDocumentProperties 设置文档属性
func (d *Document) SetDocumentProperties(properties *DocumentProperties) error {
	if properties == nil {
		return fmt.Errorf("文档属性不能为空")
	}

	// 生成核心属性XML
	if err := d.generateCoreProperties(properties); err != nil {
		return fmt.Errorf("生成核心属性失败: %v", err)
	}

	// 生成应用程序属性XML
	if err := d.generateAppProperties(properties); err != nil {
		return fmt.Errorf("生成应用程序属性失败: %v", err)
	}

	// add content type和关系
	d.addPropertiesContentTypes()
	d.addPropertiesRelationships()

	return nil
}

// GetDocumentProperties 获取文档属性
func (d *Document) GetDocumentProperties() (*DocumentProperties, error) {
	properties := &DocumentProperties{
		Created:      time.Now(),
		LastModified: time.Now(),
		Language:     "zh-CN",
	}

	// 从已保存的属性中读取（如果存在）
	if coreData, exists := d.parts["docProps/core.xml"]; exists {
		if err := d.parseCoreProperties(coreData, properties); err != nil {
			return nil, fmt.Errorf("解析核心属性失败: %v", err)
		}
	}

	if appData, exists := d.parts["docProps/app.xml"]; exists {
		if err := d.parseAppProperties(appData, properties); err != nil {
			return nil, fmt.Errorf("解析应用程序属性失败: %v", err)
		}
	}

	return properties, nil
}

// SetTitle 设置文档标题
func (d *Document) SetTitle(title string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Title = title
	return d.SetDocumentProperties(properties)
}

// SetAuthor 设置文档作者
func (d *Document) SetAuthor(author string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Creator = author
	return d.SetDocumentProperties(properties)
}

// SetSubject 设置文档主题
func (d *Document) SetSubject(subject string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Subject = subject
	return d.SetDocumentProperties(properties)
}

// SetKeywords 设置文档关键字
func (d *Document) SetKeywords(keywords string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Keywords = keywords
	return d.SetDocumentProperties(properties)
}

// SetDescription 设置文档描述
func (d *Document) SetDescription(description string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Description = description
	return d.SetDocumentProperties(properties)
}

// SetCategory 设置文档类别
func (d *Document) SetCategory(category string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Category = category
	return d.SetDocumentProperties(properties)
}

// UpdateStatistics 更新文档统计信息
func (d *Document) UpdateStatistics() error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}

	// 计算统计信息
	properties.Paragraphs = len(d.Body.GetParagraphs())
	properties.Words = d.countWords()
	properties.Characters = d.countCharacters()
	properties.Lines = d.countLines()
	properties.Pages = 1 // 简化处理，实际需要复杂计算

	// 更新最后修改时间
	properties.LastModified = time.Now()

	return d.SetDocumentProperties(properties)
}

// generateCoreProperties 生成核心属性XML
func (d *Document) generateCoreProperties(properties *DocumentProperties) error {
	coreProps := &CoreProperties{
		XmlnsCP:       "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		XmlnsDC:       "http://purl.org/dc/elements/1.1/",
		XmlnsDCTerms:  "http://purl.org/dc/terms/",
		XmlnsDCMIType: "http://purl.org/dc/dcmitype/",
		XmlnsXSI:      "http://www.w3.org/2001/XMLSchema-instance",
	}

	// 设置属性值
	if properties.Title != "" {
		coreProps.Title = &DCText{Text: properties.Title}
	}
	if properties.Subject != "" {
		coreProps.Subject = &DCText{Text: properties.Subject}
	}
	if properties.Creator != "" {
		coreProps.Creator = &DCText{Text: properties.Creator}
	}
	if properties.Keywords != "" {
		coreProps.Keywords = &CPText{Text: properties.Keywords}
	}
	if properties.Description != "" {
		coreProps.Description = &DCText{Text: properties.Description}
	}
	if properties.Language != "" {
		coreProps.Language = &DCText{Text: properties.Language}
	}
	if properties.Category != "" {
		coreProps.Category = &CPText{Text: properties.Category}
	}
	if properties.Version != "" {
		coreProps.Version = &CPText{Text: properties.Version}
	}
	if properties.Revision != "" {
		coreProps.Revision = &CPText{Text: properties.Revision}
	}

	// 设置时间属性
	if !properties.Created.IsZero() {
		coreProps.Created = &DCDate{
			XSIType: "dcterms:W3CDTF",
			Date:    properties.Created,
		}
	}
	if !properties.LastModified.IsZero() {
		coreProps.Modified = &DCDate{
			XSIType: "dcterms:W3CDTF",
			Date:    properties.LastModified,
		}
	}
	if !properties.LastPrinted.IsZero() {
		coreProps.LastPrinted = &DCDate{
			XSIType: "dcterms:W3CDTF",
			Date:    properties.LastPrinted,
		}
	}

	// serializeXML
	coreXML, err := xml.MarshalIndent(coreProps, "", "  ")
	if err != nil {
		return err
	}

	// add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["docProps/core.xml"] = append(xmlDeclaration, coreXML...)

	return nil
}

// generateAppProperties generate application properties XML
func (d *Document) generateAppProperties(properties *DocumentProperties) error {
	appProps := &AppProperties{
		Xmlns:         "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		XmlnsVT:       "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		Application:   "go-word/1.0",
		DocSecurity:   0,
		ScaleCrop:     false,
		LinksUpToDate: false,
		Pages:         properties.Pages,
		Words:         properties.Words,
		Characters:    properties.Characters,
		Paragraphs:    properties.Paragraphs,
		Lines:         properties.Lines,
	}

	// serializeXML
	appXML, err := xml.MarshalIndent(appProps, "", "  ")
	if err != nil {
		return err
	}

	// add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["docProps/app.xml"] = append(xmlDeclaration, appXML...)

	return nil
}

// parseCoreProperties 解析核心属性
func (d *Document) parseCoreProperties(data []byte, properties *DocumentProperties) error {
	var coreProps CoreProperties
	if err := xml.Unmarshal(data, &coreProps); err != nil {
		return err
	}

	if coreProps.Title != nil {
		properties.Title = coreProps.Title.Text
	}
	if coreProps.Subject != nil {
		properties.Subject = coreProps.Subject.Text
	}
	if coreProps.Creator != nil {
		properties.Creator = coreProps.Creator.Text
	}
	if coreProps.Keywords != nil {
		properties.Keywords = coreProps.Keywords.Text
	}
	if coreProps.Description != nil {
		properties.Description = coreProps.Description.Text
	}
	if coreProps.Language != nil {
		properties.Language = coreProps.Language.Text
	}
	if coreProps.Category != nil {
		properties.Category = coreProps.Category.Text
	}
	if coreProps.Version != nil {
		properties.Version = coreProps.Version.Text
	}
	if coreProps.Revision != nil {
		properties.Revision = coreProps.Revision.Text
	}

	if coreProps.Created != nil {
		properties.Created = coreProps.Created.Date
	}
	if coreProps.Modified != nil {
		properties.LastModified = coreProps.Modified.Date
	}
	if coreProps.LastPrinted != nil {
		properties.LastPrinted = coreProps.LastPrinted.Date
	}

	return nil
}

// parseAppProperties 解析应用程序属性
func (d *Document) parseAppProperties(data []byte, properties *DocumentProperties) error {
	var appProps AppProperties
	if err := xml.Unmarshal(data, &appProps); err != nil {
		return err
	}

	properties.Pages = appProps.Pages
	properties.Words = appProps.Words
	properties.Characters = appProps.Characters
	properties.Paragraphs = appProps.Paragraphs
	properties.Lines = appProps.Lines

	return nil
}

// addPropertiesContentTypes 添加属性相关的内容类型
func (d *Document) addPropertiesContentTypes() {
	d.addContentType("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
	d.addContentType("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
}

// addPropertiesRelationships 添加属性相关的关系
func (d *Document) addPropertiesRelationships() {
	// 这些关系通常在包级别的 _rels/.rels 中定义
	// 简化处理，实际实现时需要管理包级别的关系
}

// countWords 统计字数
func (d *Document) countWords() int {
	count := 0
	for _, paragraph := range d.Body.GetParagraphs() {
		for _, run := range paragraph.Runs {
			// 简化统计，按空格分割
			words := len(strings.Fields(run.Text.Content))
			count += words
		}
	}
	return count
}

// countCharacters 统计字符数
func (d *Document) countCharacters() int {
	count := 0
	for _, paragraph := range d.Body.GetParagraphs() {
		for _, run := range paragraph.Runs {
			count += len(run.Text.Content)
		}
	}
	return count
}

// countLines 统计行数
func (d *Document) countLines() int {
	count := 0
	for _, paragraph := range d.Body.GetParagraphs() {
		for _, run := range paragraph.Runs {
			// 简化处理，按换行符统计
			lines := strings.Count(run.Text.Content, "\n") + 1
			count += lines
		}
	}
	return count
}
