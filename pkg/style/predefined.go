// Package style provides predefined style constants
package style

const (
	StyleNormal = "Normal"

	StyleHeading1 = "Heading1"
	StyleHeading2 = "Heading2"
	StyleHeading3 = "Heading3"
	StyleHeading4 = "Heading4"
	StyleHeading5 = "Heading5"
	StyleHeading6 = "Heading6"
	StyleHeading7 = "Heading7"
	StyleHeading8 = "Heading8"
	StyleHeading9 = "Heading9"

	StyleTitle    = "Title"
	StyleSubtitle = "Subtitle"

	StyleEmphasis = "Emphasis"
	StyleStrong   = "Strong"
	StyleCodeChar = "CodeChar"

	StyleQuote         = "Quote"
	StyleListParagraph = "ListParagraph"
	StyleCodeBlock     = "CodeBlock"
)

// GetPredefinedStyleNames returns a mapping of style IDs to display names
func GetPredefinedStyleNames() map[string]string {
	return map[string]string{
		StyleNormal:        "Normal",
		StyleHeading1:      "Heading 1",
		StyleHeading2:      "Heading 2",
		StyleHeading3:      "Heading 3",
		StyleHeading4:      "Heading 4",
		StyleHeading5:      "Heading 5",
		StyleHeading6:      "Heading 6",
		StyleHeading7:      "Heading 7",
		StyleHeading8:      "Heading 8",
		StyleHeading9:      "Heading 9",
		StyleTitle:         "Title",
		StyleSubtitle:      "Subtitle",
		StyleEmphasis:      "Emphasis",
		StyleStrong:        "Strong",
		StyleCodeChar:      "Code Character",
		StyleQuote:         "Quote",
		StyleListParagraph: "List Paragraph",
		StyleCodeBlock:     "Code Block",
	}
}

// StyleConfig defines configuration for a style
type StyleConfig struct {
	StyleID     string
	Name        string
	Description string
	StyleType   StyleType
}

// GetPredefinedStyleConfigs returns all predefined style configurations
func GetPredefinedStyleConfigs() []StyleConfig {
	return []StyleConfig{
		{
			StyleID:     StyleNormal,
			Name:        "普通文本",
			Description: "默认的段落样式，使用Calibri字体，11磅字号",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading1,
			Name:        "标题 1",
			Description: "一级标题，16磅蓝色粗体，段前12磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading2,
			Name:        "标题 2",
			Description: "二级标题，13磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading3,
			Name:        "标题 3",
			Description: "三级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading4,
			Name:        "标题 4",
			Description: "四级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading5,
			Name:        "标题 5",
			Description: "五级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading6,
			Name:        "标题 6",
			Description: "六级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading7,
			Name:        "标题 7",
			Description: "七级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading8,
			Name:        "标题 8",
			Description: "八级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading9,
			Name:        "标题 9",
			Description: "九级标题，12磅蓝色粗体，段前6磅间距",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleTitle,
			Name:        "文档标题",
			Description: "文档标题样式",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleSubtitle,
			Name:        "副标题",
			Description: "副标题样式",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleEmphasis,
			Name:        "强调",
			Description: "斜体文本样式",
			StyleType:   StyleTypeCharacter,
		},
		{
			StyleID:     StyleStrong,
			Name:        "加粗",
			Description: "粗体文本样式",
			StyleType:   StyleTypeCharacter,
		},
		{
			StyleID:     StyleCodeChar,
			Name:        "代码字符",
			Description: "等宽字体，红色文本，适用于代码片段",
			StyleType:   StyleTypeCharacter,
		},
		{
			StyleID:     StyleQuote,
			Name:        "引用",
			Description: "引用段落样式，斜体灰色，左右各缩进0.5英寸",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleListParagraph,
			Name:        "列表段落",
			Description: "列表段落样式",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleCodeBlock,
			Name:        "代码块",
			Description: "代码块样式",
			StyleType:   StyleTypeParagraph,
		},
	}
}
