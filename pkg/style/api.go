// Package style provides style application API
package style

import "fmt"

// StyleApplicator defines the interface for applying styles
type StyleApplicator interface {
	ApplyStyle(styleID string) error
	ApplyHeadingStyle(level int) error
	ApplyTitleStyle() error
	ApplySubtitleStyle() error
	ApplyQuoteStyle() error
	ApplyCodeBlockStyle() error
	ApplyListParagraphStyle() error
	ApplyNormalStyle() error
}

// QuickStyleAPI provides convenient style application methods
type QuickStyleAPI struct {
	styleManager *StyleManager
}

// NewQuickStyleAPI creates a new quick style API instance
func NewQuickStyleAPI(styleManager *StyleManager) *QuickStyleAPI {
	return &QuickStyleAPI{
		styleManager: styleManager,
	}
}

// GetStyleInfo retrieves style information for UI display
func (api *QuickStyleAPI) GetStyleInfo(styleID string) (*StyleInfo, error) {
	style := api.styleManager.GetStyle(styleID)
	if style == nil {
		return nil, fmt.Errorf("style %s does not exist", styleID)
	}

	return &StyleInfo{
		ID:          style.StyleID,
		Name:        getStyleDisplayName(style),
		Type:        StyleType(style.Type),
		Description: getStyleDescription(styleID),
		IsBuiltIn:   !style.CustomStyle,
		BasedOn:     getBasedOnStyleID(style),
	}, nil
}

// StyleInfo defines style information for display
type StyleInfo struct {
	ID          string    `json:"id"`
	Name        string    `json:"name"`
	Type        StyleType `json:"type"`
	Description string    `json:"description"`
	IsBuiltIn   bool      `json:"isBuiltIn"`
	BasedOn     string    `json:"basedOn,omitempty"`
}

// GetAllStylesInfo retrieves information for all styles
func (api *QuickStyleAPI) GetAllStylesInfo() []*StyleInfo {
	var stylesInfo []*StyleInfo
	for _, style := range api.styleManager.GetAllStyles() {
		info := &StyleInfo{
			ID:          style.StyleID,
			Name:        getStyleDisplayName(style),
			Type:        StyleType(style.Type),
			Description: getStyleDescription(style.StyleID),
			IsBuiltIn:   !style.CustomStyle,
			BasedOn:     getBasedOnStyleID(style),
		}
		stylesInfo = append(stylesInfo, info)
	}
	return stylesInfo
}

// GetHeadingStylesInfo retrieves information for all heading styles
func (api *QuickStyleAPI) GetHeadingStylesInfo() []*StyleInfo {
	var headingStylesInfo []*StyleInfo
	for i := 1; i <= 9; i++ {
		styleID := fmt.Sprintf("Heading%d", i)
		if info, err := api.GetStyleInfo(styleID); err == nil {
			headingStylesInfo = append(headingStylesInfo, info)
		}
	}
	return headingStylesInfo
}

// GetParagraphStylesInfo retrieves information for all paragraph styles
func (api *QuickStyleAPI) GetParagraphStylesInfo() []*StyleInfo {
	var paragraphStylesInfo []*StyleInfo
	for _, style := range api.styleManager.GetStylesByType(StyleTypeParagraph) {
		info := &StyleInfo{
			ID:          style.StyleID,
			Name:        getStyleDisplayName(style),
			Type:        StyleType(style.Type),
			Description: getStyleDescription(style.StyleID),
			IsBuiltIn:   !style.CustomStyle,
			BasedOn:     getBasedOnStyleID(style),
		}
		paragraphStylesInfo = append(paragraphStylesInfo, info)
	}
	return paragraphStylesInfo
}

// GetCharacterStylesInfo retrieves information for all character styles
func (api *QuickStyleAPI) GetCharacterStylesInfo() []*StyleInfo {
	var characterStylesInfo []*StyleInfo
	for _, style := range api.styleManager.GetStylesByType(StyleTypeCharacter) {
		info := &StyleInfo{
			ID:          style.StyleID,
			Name:        getStyleDisplayName(style),
			Type:        StyleType(style.Type),
			Description: getStyleDescription(style.StyleID),
			IsBuiltIn:   !style.CustomStyle,
			BasedOn:     getBasedOnStyleID(style),
		}
		characterStylesInfo = append(characterStylesInfo, info)
	}
	return characterStylesInfo
}

// CreateQuickStyle creates a new custom style with the given configuration
func (api *QuickStyleAPI) CreateQuickStyle(config QuickStyleConfig) (*Style, error) {
	if api.styleManager.StyleExists(config.ID) {
		return nil, fmt.Errorf("style ID %s already exists", config.ID)
	}

	style := api.styleManager.CreateCustomStyle(
		config.ID,
		config.Name,
		config.Type,
		config.BasedOn,
	)

	if config.ParagraphConfig != nil {
		style.ParagraphPr = createParagraphProperties(config.ParagraphConfig)
	}

	if config.RunConfig != nil {
		style.RunPr = createRunProperties(config.RunConfig)
	}

	return style, nil
}

// QuickStyleConfig defines quick style configuration
type QuickStyleConfig struct {
	ID              string                `json:"id"`
	Name            string                `json:"name"`
	Type            StyleType             `json:"type"`
	BasedOn         string                `json:"basedOn,omitempty"`
	ParagraphConfig *QuickParagraphConfig `json:"paragraphConfig,omitempty"`
	RunConfig       *QuickRunConfig       `json:"runConfig,omitempty"`
}

// QuickParagraphConfig defines quick paragraph configuration
// LineSpacing is a multiplier: 1.0 = single, 1.5 = 1.5x, 2.0 = double (converted to OOXML units: value*240)
// All indent/space values in points
type QuickParagraphConfig struct {
	Alignment       string  `json:"alignment,omitempty"`       // left, center, right, justify
	LineSpacing     float64 `json:"lineSpacing,omitempty"`
	SpaceBefore     int     `json:"spaceBefore,omitempty"`
	SpaceAfter      int     `json:"spaceAfter,omitempty"`
	FirstLineIndent int     `json:"firstLineIndent,omitempty"`
	LeftIndent      int     `json:"leftIndent,omitempty"`
	RightIndent     int     `json:"rightIndent,omitempty"`
	SnapToGrid      *bool   `json:"snapToGrid,omitempty"` // disable grid alignment to enable precise line spacing
}

// QuickRunConfig defines quick character configuration
type QuickRunConfig struct {
	FontName  string `json:"fontName,omitempty"`
	FontSize  int    `json:"fontSize,omitempty"`  // in points
	FontColor string `json:"fontColor,omitempty"` // hex color code
	Bold      bool   `json:"bold,omitempty"`
	Italic    bool   `json:"italic,omitempty"`
	Underline bool   `json:"underline,omitempty"`
	Strike    bool   `json:"strike,omitempty"`
	Highlight string `json:"highlight,omitempty"`
}

// getStyleDisplayName returns the display name of a style
func getStyleDisplayName(style *Style) string {
	if style.Name != nil {
		return style.Name.Val
	}
	return style.StyleID
}

// getStyleDescription returns the description of a style
func getStyleDescription(styleID string) string {
	configs := GetPredefinedStyleConfigs()
	for _, config := range configs {
		if config.StyleID == styleID {
			return config.Description
		}
	}
	return ""
}

// getBasedOnStyleID returns the parent style ID
func getBasedOnStyleID(style *Style) string {
	if style.BasedOn != nil {
		return style.BasedOn.Val
	}
	return ""
}

// createParagraphProperties creates paragraph properties from configuration
func createParagraphProperties(config *QuickParagraphConfig) *ParagraphProperties {
	props := &ParagraphProperties{}

	if config.Alignment != "" {
		props.Justification = &Justification{Val: config.Alignment}
	}

	if config.SnapToGrid != nil && !*config.SnapToGrid {
		props.SnapToGrid = &SnapToGrid{Val: "0"}
	}

	if config.LineSpacing > 0 || config.SpaceBefore > 0 || config.SpaceAfter > 0 {
		spacing := &Spacing{}
		if config.SpaceBefore > 0 {
			spacing.Before = fmt.Sprintf("%d", config.SpaceBefore*20) // convert to twips
		}
		if config.SpaceAfter > 0 {
			spacing.After = fmt.Sprintf("%d", config.SpaceAfter*20) // convert to twips
		}
		if config.LineSpacing > 0 {
			spacing.Line = fmt.Sprintf("%.0f", config.LineSpacing*240) // convert to line spacing units
			spacing.LineRule = "auto"
		}
		props.Spacing = spacing
	}

	if config.FirstLineIndent > 0 || config.LeftIndent > 0 || config.RightIndent > 0 {
		indentation := &Indentation{}
		if config.FirstLineIndent > 0 {
			indentation.FirstLine = fmt.Sprintf("%d", config.FirstLineIndent*20) // convert to twips
		}
		if config.LeftIndent > 0 {
			indentation.Left = fmt.Sprintf("%d", config.LeftIndent*20) // convert to twips
		}
		if config.RightIndent > 0 {
			indentation.Right = fmt.Sprintf("%d", config.RightIndent*20) // convert to twips
		}
		props.Indentation = indentation
	}

	return props
}

// createRunProperties creates run properties from configuration
func createRunProperties(config *QuickRunConfig) *RunProperties {
	props := &RunProperties{}

	if config.FontName != "" {
		props.FontFamily = &FontFamily{
			ASCII:    config.FontName,
			EastAsia: config.FontName,
			HAnsi:    config.FontName,
			CS:       config.FontName,
		}
	}

	if config.FontSize > 0 {
		props.FontSize = &FontSize{Val: fmt.Sprintf("%d", config.FontSize*2)} // Word uses half-point units
	}

	if config.FontColor != "" {
		props.Color = &Color{Val: config.FontColor}
	}

	if config.Bold {
		props.Bold = &Bold{}
	}

	if config.Italic {
		props.Italic = &Italic{}
	}

	if config.Underline {
		props.Underline = &Underline{Val: "single"}
	}

	if config.Strike {
		props.Strike = &Strike{}
	}

	if config.Highlight != "" {
		props.Highlight = &Highlight{Val: config.Highlight}
	}

	return props
}
