// Package document provides Word document footnote and endnote functionality
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// FootnoteType footnote type
type FootnoteType string

const (
	// FootnoteTypeFootnote - footnote
	FootnoteTypeFootnote FootnoteType = "footnote"
	// FootnoteTypeEndnote - endnote
	FootnoteTypeEndnote FootnoteType = "endnote"
)

// Footnotes footnotes collection
type Footnotes struct {
	XMLName   xml.Name    `xml:"w:footnotes"`
	Xmlns     string      `xml:"xmlns:w,attr"`
	Footnotes []*Footnote `xml:"w:footnote"`
}

// Endnotes endnotes collection
type Endnotes struct {
	XMLName  xml.Name   `xml:"w:endnotes"`
	Xmlns    string     `xml:"xmlns:w,attr"`
	Endnotes []*Endnote `xml:"w:endnote"`
}

// Footnote footnote structure
type Footnote struct {
	XMLName    xml.Name     `xml:"w:footnote"`
	Type       string       `xml:"w:type,attr,omitempty"`
	ID         string       `xml:"w:id,attr"`
	Paragraphs []*Paragraph `xml:"w:p"`
}

// Endnote endnote structure
type Endnote struct {
	XMLName    xml.Name     `xml:"w:endnote"`
	Type       string       `xml:"w:type,attr,omitempty"`
	ID         string       `xml:"w:id,attr"`
	Paragraphs []*Paragraph `xml:"w:p"`
}

// FootnoteReference footnote reference
type FootnoteReference struct {
	XMLName xml.Name `xml:"w:footnoteReference"`
	ID      string   `xml:"w:id,attr"`
}

// EndnoteReference endnote reference
type EndnoteReference struct {
	XMLName xml.Name `xml:"w:endnoteReference"`
	ID      string   `xml:"w:id,attr"`
}

// FootnoteConfig footnote configuration
type FootnoteConfig struct {
	NumberFormat FootnoteNumberFormat // numbering format
	StartNumber  int                  // starting number
	RestartEach  FootnoteRestart      // restart rule
	Position     FootnotePosition     // position
}

// FootnoteNumberFormat footnote numbering format
type FootnoteNumberFormat string

const (
	// FootnoteFormatDecimal - decimal numbers
	FootnoteFormatDecimal FootnoteNumberFormat = "decimal"
	// FootnoteFormatLowerRoman - lowercase Roman numerals
	FootnoteFormatLowerRoman FootnoteNumberFormat = "lowerRoman"
	// FootnoteFormatUpperRoman - uppercase Roman numerals
	FootnoteFormatUpperRoman FootnoteNumberFormat = "upperRoman"
	// FootnoteFormatLowerLetter - lowercase letters
	FootnoteFormatLowerLetter FootnoteNumberFormat = "lowerLetter"
	// FootnoteFormatUpperLetter - uppercase letters
	FootnoteFormatUpperLetter FootnoteNumberFormat = "upperLetter"
	// FootnoteFormatSymbol - symbols
	FootnoteFormatSymbol FootnoteNumberFormat = "symbol"
)

// FootnoteRestart footnote restart rule
type FootnoteRestart string

const (
	// FootnoteRestartContinuous - continuous numbering
	FootnoteRestartContinuous FootnoteRestart = "continuous"
	// FootnoteRestartEachSection - restart each section
	FootnoteRestartEachSection FootnoteRestart = "eachSect"
	// FootnoteRestartEachPage - restart each page
	FootnoteRestartEachPage FootnoteRestart = "eachPage"
)

// FootnotePosition footnote position
type FootnotePosition string

const (
	// FootnotePositionPageBottom - page bottom
	FootnotePositionPageBottom FootnotePosition = "pageBottom"
	// FootnotePositionBeneathText - beneath text
	FootnotePositionBeneathText FootnotePosition = "beneathText"
	// FootnotePositionSectionEnd - section end
	FootnotePositionSectionEnd FootnotePosition = "sectEnd"
	// FootnotePositionDocumentEnd - document end
	FootnotePositionDocumentEnd FootnotePosition = "docEnd"
)

// FootnoteProperties footnote properties
type FootnoteProperties struct {
	NumberFormat string `xml:"w:numFmt,attr,omitempty"`
	StartNumber  int    `xml:"w:numStart,attr,omitempty"`
	RestartRule  string `xml:"w:numRestart,attr,omitempty"`
	Position     string `xml:"w:pos,attr,omitempty"`
}

// EndnoteProperties endnote properties
type EndnoteProperties struct {
	NumberFormat string `xml:"w:numFmt,attr,omitempty"`
	StartNumber  int    `xml:"w:numStart,attr,omitempty"`
	RestartRule  string `xml:"w:numRestart,attr,omitempty"`
	Position     string `xml:"w:pos,attr,omitempty"`
}

// Settings document settings XML structure
type Settings struct {
	XMLName                 xml.Name                 `xml:"w:settings"`
	Xmlns                   string                   `xml:"xmlns:w,attr"`
	DefaultTabStop          *DefaultTabStop          `xml:"w:defaultTabStop,omitempty"`
	CharacterSpacingControl *CharacterSpacingControl `xml:"w:characterSpacingControl,omitempty"`
	FootnotePr              *FootnotePr              `xml:"w:footnotePr,omitempty"`
	EndnotePr               *EndnotePr               `xml:"w:endnotePr,omitempty"`
}

// DefaultTabStop default tab stop settings
type DefaultTabStop struct {
	XMLName xml.Name `xml:"w:defaultTabStop"`
	Val     string   `xml:"w:val,attr"`
}

// CharacterSpacingControl character spacing control
type CharacterSpacingControl struct {
	XMLName xml.Name `xml:"w:characterSpacingControl"`
	Val     string   `xml:"w:val,attr"`
}

// FootnotePr footnote properties settings
type FootnotePr struct {
	XMLName    xml.Name            `xml:"w:footnotePr"`
	NumFmt     *FootnoteNumFmt     `xml:"w:numFmt,omitempty"`
	NumStart   *FootnoteNumStart   `xml:"w:numStart,omitempty"`
	NumRestart *FootnoteNumRestart `xml:"w:numRestart,omitempty"`
	Pos        *FootnotePos        `xml:"w:pos,omitempty"`
}

// EndnotePr endnote properties settings
type EndnotePr struct {
	XMLName    xml.Name           `xml:"w:endnotePr"`
	NumFmt     *EndnoteNumFmt     `xml:"w:numFmt,omitempty"`
	NumStart   *EndnoteNumStart   `xml:"w:numStart,omitempty"`
	NumRestart *EndnoteNumRestart `xml:"w:numRestart,omitempty"`
	Pos        *EndnotePos        `xml:"w:pos,omitempty"`
}

// FootnoteNumFmt footnote numbering format
type FootnoteNumFmt struct {
	XMLName xml.Name `xml:"w:numFmt"`
	Val     string   `xml:"w:val,attr"`
}

// FootnoteNumStart footnote starting number
type FootnoteNumStart struct {
	XMLName xml.Name `xml:"w:numStart"`
	Val     string   `xml:"w:val,attr"`
}

// FootnoteNumRestart footnote restart rule
type FootnoteNumRestart struct {
	XMLName xml.Name `xml:"w:numRestart"`
	Val     string   `xml:"w:val,attr"`
}

// FootnotePos footnote position
type FootnotePos struct {
	XMLName xml.Name `xml:"w:pos"`
	Val     string   `xml:"w:val,attr"`
}

// EndnoteNumFmt endnote numbering format
type EndnoteNumFmt struct {
	XMLName xml.Name `xml:"w:numFmt"`
	Val     string   `xml:"w:val,attr"`
}

// EndnoteNumStart endnote starting number
type EndnoteNumStart struct {
	XMLName xml.Name `xml:"w:numStart"`
	Val     string   `xml:"w:val,attr"`
}

// EndnoteNumRestart endnote restart rule
type EndnoteNumRestart struct {
	XMLName xml.Name `xml:"w:numRestart"`
	Val     string   `xml:"w:val,attr"`
}

// EndnotePos endnote position
type EndnotePos struct {
	XMLName xml.Name `xml:"w:pos"`
	Val     string   `xml:"w:val,attr"`
}

// Global footnote/endnote manager
var globalFootnoteManager *FootnoteManager

// FootnoteManager footnote manager
type FootnoteManager struct {
	nextFootnoteID int
	nextEndnoteID  int
	footnotes      map[string]*Footnote
	endnotes       map[string]*Endnote
}

// getFootnoteManager gets the global footnote manager
func getFootnoteManager() *FootnoteManager {
	if globalFootnoteManager == nil {
		globalFootnoteManager = &FootnoteManager{
			nextFootnoteID: 1,
			nextEndnoteID:  1,
			footnotes:      make(map[string]*Footnote),
			endnotes:       make(map[string]*Endnote),
		}
	}
	return globalFootnoteManager
}

// DefaultFootnoteConfig returns default footnote configuration
func DefaultFootnoteConfig() *FootnoteConfig {
	return &FootnoteConfig{
		NumberFormat: FootnoteFormatDecimal,
		StartNumber:  1,
		RestartEach:  FootnoteRestartContinuous,
		Position:     FootnotePositionPageBottom,
	}
}

// AddFootnote adds a footnote
func (d *Document) AddFootnote(text string, footnoteText string) error {
	return d.addFootnoteOrEndnote(text, footnoteText, FootnoteTypeFootnote)
}

// AddEndnote adds an endnote
func (d *Document) AddEndnote(text string, endnoteText string) error {
	return d.addFootnoteOrEndnote(text, endnoteText, FootnoteTypeEndnote)
}

// addFootnoteOrEndnote generic method to add footnotes or endnotes
func (d *Document) addFootnoteOrEndnote(text string, noteText string, noteType FootnoteType) error {
	manager := getFootnoteManager()

	// Ensure footnote/endnote system is initialized
	d.ensureFootnoteInitialized(noteType)

	var noteID string
	if noteType == FootnoteTypeFootnote {
		noteID = strconv.Itoa(manager.nextFootnoteID)
		manager.nextFootnoteID++
	} else {
		noteID = strconv.Itoa(manager.nextEndnoteID)
		manager.nextEndnoteID++
	}

	paragraph := &Paragraph{}

	if text != "" {
		textRun := Run{
			Text: Text{Content: text},
		}
		paragraph.Runs = append(paragraph.Runs, textRun)
	}

	// Add footnote/endnote reference
	refRun := Run{
		Properties: &RunProperties{},
	}

	if noteType == FootnoteTypeFootnote {
		// Simplified handling: insert footnote marker in text
		refRun.Text = Text{Content: fmt.Sprintf("[%s]", noteID)}
	} else {
		// Simplified handling: insert endnote marker in text
		refRun.Text = Text{Content: fmt.Sprintf("[endnote%s]", noteID)}
	}

	paragraph.Runs = append(paragraph.Runs, refRun)
	d.Body.Elements = append(d.Body.Elements, paragraph)

	// Create footnote/endnote content
	if err := d.createNoteContent(noteID, noteText, noteType); err != nil {
		return fmt.Errorf("failed to create %s content: %v", noteType, err)
	}

	return nil
}

// AddFootnoteToRun adds a footnote reference to an existing Run
func (d *Document) AddFootnoteToRun(run *Run, footnoteText string) error {
	manager := getFootnoteManager()
	d.ensureFootnoteInitialized(FootnoteTypeFootnote)

	noteID := strconv.Itoa(manager.nextFootnoteID)
	manager.nextFootnoteID++

	// Add footnote reference after current Run
	refText := fmt.Sprintf("[%s]", noteID)
	run.Text.Content += refText

	// Create footnote content
	return d.createNoteContent(noteID, footnoteText, FootnoteTypeFootnote)
}

// SetFootnoteConfig sets footnote configuration
func (d *Document) SetFootnoteConfig(config *FootnoteConfig) error {
	if config == nil {
		config = DefaultFootnoteConfig()
	}

	// Ensure document settings are initialized
	d.ensureSettingsInitialized()

	// Create footnote properties XML structure
	footnoteProps := &FootnoteProperties{
		NumberFormat: string(config.NumberFormat),
		StartNumber:  config.StartNumber,
		RestartRule:  string(config.RestartEach),
		Position:     string(config.Position),
	}

	// Create endnote properties XML structure
	endnoteProps := &EndnoteProperties{
		NumberFormat: string(config.NumberFormat),
		StartNumber:  config.StartNumber,
		RestartRule:  string(config.RestartEach),
		Position:     string(config.Position),
	}

	// Update document settings
	if err := d.updateDocumentSettings(footnoteProps, endnoteProps); err != nil {
		return fmt.Errorf("failed to update footnote configuration: %v", err)
	}

	return nil
}

// ensureFootnoteInitialized ensures footnote/endnote system is initialized
func (d *Document) ensureFootnoteInitialized(noteType FootnoteType) {
	if noteType == FootnoteTypeFootnote {
		if _, exists := d.parts["word/footnotes.xml"]; !exists {
			d.initializeFootnotes()
		}
	} else {
		if _, exists := d.parts["word/endnotes.xml"]; !exists {
			d.initializeEndnotes()
		}
	}
}

// initializeFootnotes initializes the footnote system
func (d *Document) initializeFootnotes() {
	footnotes := &Footnotes{
		Xmlns:     "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Footnotes: []*Footnote{},
	}

	// Add default separator footnote
	separatorFootnote := &Footnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	footnotes.Footnotes = append(footnotes.Footnotes, separatorFootnote)

	// serialize脚注
	footnotesXML, err := xml.MarshalIndent(footnotes, "", "  ")
	if err != nil {
		return
	}

	// add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/footnotes.xml"] = append(xmlDeclaration, footnotesXML...)

	// add content type
	d.addContentType("word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")

	// add relationship
	d.addFootnoteRelationship()
}

// initializeEndnotes initializes the endnote system
func (d *Document) initializeEndnotes() {
	endnotes := &Endnotes{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Endnotes: []*Endnote{},
	}

	// Add default separator endnote
	separatorEndnote := &Endnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	endnotes.Endnotes = append(endnotes.Endnotes, separatorEndnote)

	// serialize尾注
	endnotesXML, err := xml.MarshalIndent(endnotes, "", "  ")
	if err != nil {
		return
	}

	// add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/endnotes.xml"] = append(xmlDeclaration, endnotesXML...)

	// add content type
	d.addContentType("word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")

	// add relationship
	d.addEndnoteRelationship()
}

// createNoteContent creates footnote/endnote content
func (d *Document) createNoteContent(noteID string, noteText string, noteType FootnoteType) error {
	manager := getFootnoteManager()

	// Create footnote/endnote paragraph
	noteParagraph := &Paragraph{
		Runs: []Run{
			{
				Text: Text{Content: noteText},
			},
		},
	}

	if noteType == FootnoteTypeFootnote {
		// Create footnote
		footnote := &Footnote{
			ID:         noteID,
			Paragraphs: []*Paragraph{noteParagraph},
		}
		manager.footnotes[noteID] = footnote

		// Update footnote file
		d.updateFootnotesFile()
	} else {
		// Create endnote
		endnote := &Endnote{
			ID:         noteID,
			Paragraphs: []*Paragraph{noteParagraph},
		}
		manager.endnotes[noteID] = endnote

		// Update endnote file
		d.updateEndnotesFile()
	}

	return nil
}

// updateFootnotesFile updates the footnotes file
func (d *Document) updateFootnotesFile() {
	manager := getFootnoteManager()

	footnotes := &Footnotes{
		Xmlns:     "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Footnotes: []*Footnote{},
	}

	// Add default separator
	separatorFootnote := &Footnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	footnotes.Footnotes = append(footnotes.Footnotes, separatorFootnote)

	// 添加所有脚注
	for _, footnote := range manager.footnotes {
		footnotes.Footnotes = append(footnotes.Footnotes, footnote)
	}

	// serialize
	footnotesXML, err := xml.MarshalIndent(footnotes, "", "  ")
	if err != nil {
		return
	}

	// add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/footnotes.xml"] = append(xmlDeclaration, footnotesXML...)
}

// updateEndnotesFile updates the endnotes file
func (d *Document) updateEndnotesFile() {
	manager := getFootnoteManager()

	endnotes := &Endnotes{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Endnotes: []*Endnote{},
	}

	// Add default separator
	separatorEndnote := &Endnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	endnotes.Endnotes = append(endnotes.Endnotes, separatorEndnote)

	// Add all endnotes
	for _, endnote := range manager.endnotes {
		endnotes.Endnotes = append(endnotes.Endnotes, endnote)
	}

	// serialize
	endnotesXML, err := xml.MarshalIndent(endnotes, "", "  ")
	if err != nil {
		return
	}

	// add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/endnotes.xml"] = append(xmlDeclaration, endnotesXML...)
}

// addFootnoteRelationship adds a footnote relationship
func (d *Document) addFootnoteRelationship() {
	relationshipID := fmt.Sprintf("rId%d", len(d.relationships.Relationships)+1)

	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
		Target: "footnotes.xml",
	}
	d.relationships.Relationships = append(d.relationships.Relationships, relationship)
}

// addEndnoteRelationship adds an endnote relationship
func (d *Document) addEndnoteRelationship() {
	relationshipID := fmt.Sprintf("rId%d", len(d.relationships.Relationships)+1)

	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
		Target: "endnotes.xml",
	}
	d.relationships.Relationships = append(d.relationships.Relationships, relationship)
}

// GetFootnoteCount returns the number of footnotes
func (d *Document) GetFootnoteCount() int {
	manager := getFootnoteManager()
	return len(manager.footnotes)
}

// GetEndnoteCount returns the number of endnotes
func (d *Document) GetEndnoteCount() int {
	manager := getFootnoteManager()
	return len(manager.endnotes)
}

// RemoveFootnote removes a specified footnote
func (d *Document) RemoveFootnote(footnoteID string) error {
	manager := getFootnoteManager()

	if _, exists := manager.footnotes[footnoteID]; !exists {
		return fmt.Errorf("footnote %s does not exist", footnoteID)
	}

	delete(manager.footnotes, footnoteID)
	d.updateFootnotesFile()

	return nil
}

// RemoveEndnote removes a specified endnote
func (d *Document) RemoveEndnote(endnoteID string) error {
	manager := getFootnoteManager()

	if _, exists := manager.endnotes[endnoteID]; !exists {
		return fmt.Errorf("endnote %s does not exist", endnoteID)
	}

	delete(manager.endnotes, endnoteID)
	d.updateEndnotesFile()

	return nil
}

// ensureSettingsInitialized ensures document settings are initialized
func (d *Document) ensureSettingsInitialized() {
	// Check if settings.xml exists, if not create default settings
	if _, exists := d.parts["word/settings.xml"]; !exists {
		d.initializeSettings()
	}
}

// initializeSettings initializes document settings
func (d *Document) initializeSettings() {
	// Create default settings
	settings := d.createDefaultSettings()

	// Save settings
	if err := d.saveSettings(settings); err != nil {
		// If saving fails, use fallback hardcoded method
		settingsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>`
		d.parts["word/settings.xml"] = []byte(settingsXML)
	}

	// add content type
	d.addContentType("word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")

	// add relationship
	d.addSettingsRelationship()
}

// updateDocumentSettings updates footnote and endnote configuration in document settings
func (d *Document) updateDocumentSettings(footnoteProps *FootnoteProperties, endnoteProps *EndnoteProperties) error {
	// Parse existing settings.xml
	settings, err := d.parseSettings()
	if err != nil {
		return fmt.Errorf("failed to parse settings file: %v", err)
	}

	// Update footnote settings
	if footnoteProps != nil {
		footnotePr := &FootnotePr{}

		if footnoteProps.NumberFormat != "" {
			footnotePr.NumFmt = &FootnoteNumFmt{Val: footnoteProps.NumberFormat}
		}

		if footnoteProps.StartNumber > 0 {
			footnotePr.NumStart = &FootnoteNumStart{Val: strconv.Itoa(footnoteProps.StartNumber)}
		}

		if footnoteProps.RestartRule != "" {
			footnotePr.NumRestart = &FootnoteNumRestart{Val: footnoteProps.RestartRule}
		}

		if footnoteProps.Position != "" {
			footnotePr.Pos = &FootnotePos{Val: footnoteProps.Position}
		}

		settings.FootnotePr = footnotePr
	}

	// Update endnote settings
	if endnoteProps != nil {
		endnotePr := &EndnotePr{}

		if endnoteProps.NumberFormat != "" {
			endnotePr.NumFmt = &EndnoteNumFmt{Val: endnoteProps.NumberFormat}
		}

		if endnoteProps.StartNumber > 0 {
			endnotePr.NumStart = &EndnoteNumStart{Val: strconv.Itoa(endnoteProps.StartNumber)}
		}

		if endnoteProps.RestartRule != "" {
			endnotePr.NumRestart = &EndnoteNumRestart{Val: endnoteProps.RestartRule}
		}

		if endnoteProps.Position != "" {
			endnotePr.Pos = &EndnotePos{Val: endnoteProps.Position}
		}

		settings.EndnotePr = endnotePr
	}

	// Save updated settings.xml
	return d.saveSettings(settings)
}

// parseSettings parses the settings.xml file
func (d *Document) parseSettings() (*Settings, error) {
	settingsData, exists := d.parts["word/settings.xml"]
	if !exists {
		// If settings.xml does not exist, return default settings
		return d.createDefaultSettings(), nil
	}

	var settings Settings

	// Using xml.Unmarshal directly may have namespace issues, so we use string replacement
	// Replace w:settings with settings, etc., and then parse using a simplified structure
	settingsStr := string(settingsData)

	// If XML contains w: prefix, it's serialized XML, directly create default settings and update
	// This is a simplified approach to avoid namespace parsing issues
	if len(settingsStr) > 0 {
		// If file exists and is not empty, use default settings as base
		settings = *d.createDefaultSettings()

		// Complex XML parsing logic can be added here later
		// For now, simplified processing, return default settings
		return &settings, nil
	}

	return d.createDefaultSettings(), nil
}

// createDefaultSettings creates default settings
func (d *Document) createDefaultSettings() *Settings {
	return &Settings{
		Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		DefaultTabStop: &DefaultTabStop{
			Val: "708",
		},
		CharacterSpacingControl: &CharacterSpacingControl{
			Val: "doNotCompress",
		},
	}
}

// saveSettings saves the settings.xml file
func (d *Document) saveSettings(settings *Settings) error {
	// Serialize to XML
	settingsXML, err := xml.MarshalIndent(settings, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize settings.xml: %v", err)
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/settings.xml"] = append(xmlDeclaration, settingsXML...)

	return nil
}

// addSettingsRelationship adds settings file relationship
func (d *Document) addSettingsRelationship() {
	relationshipID := fmt.Sprintf("rId%d", len(d.relationships.Relationships)+1)

	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
		Target: "word/settings.xml",
	}
	d.relationships.Relationships = append(d.relationships.Relationships, relationship)
}
