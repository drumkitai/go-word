package markdown

import (
	"path/filepath"
	"regexp"
	"strings"

	"github.com/drumkitai/go-word/pkg/document"
	"github.com/yuin/goldmark/ast"

	// add goldmark extension AST node support
	extast "github.com/yuin/goldmark/extension/ast"

	// add math formula support
	mathjax "github.com/litao91/goldmark-mathjax"
)

// WordRenderer renders an AST to a Word document
type WordRenderer struct {
	doc       *document.Document
	opts      *ConvertOptions
	source    []byte
	listLevel int // current list nesting level
}

// Render renders the AST to a Word document
func (r *WordRenderer) Render(doc ast.Node) error {
	return ast.Walk(doc, func(node ast.Node, entering bool) (ast.WalkStatus, error) {
		if !entering {
			return ast.WalkContinue, nil
		}

		switch n := node.(type) {
		case *ast.Document:
			return ast.WalkContinue, nil

		case *ast.Heading:
			return r.renderHeading(n)

		case *ast.Paragraph:
			return r.renderParagraph(n)

		case *ast.List:
			return r.renderList(n)

		case *ast.ListItem:
			return r.renderListItem(n)

		case *ast.Blockquote:
			return r.renderBlockquote(n)

		case *ast.FencedCodeBlock:
			return r.renderCodeBlock(n)

		case *ast.CodeBlock:
			return r.renderCodeBlock(n)

		case *ast.ThematicBreak:
			return r.renderThematicBreak(n)

		case *ast.Text:
			return ast.WalkSkipChildren, nil

		case *ast.Emphasis:
			return ast.WalkSkipChildren, nil

		case *ast.Link:
			return ast.WalkSkipChildren, nil

		case *ast.Image:
			return r.renderImage(n)

		case *extast.Table:
			if r.opts.EnableTables {
				return r.renderTable(n)
			}
			return ast.WalkContinue, nil

		case *extast.TableRow:
			return ast.WalkSkipChildren, nil

		case *extast.TableCell:
			return ast.WalkSkipChildren, nil

		case *extast.TaskCheckBox:
			if r.opts.EnableTaskList {
				return r.renderTaskCheckBox(n)
			}
			return ast.WalkContinue, nil

		default:
			if r.opts.EnableMath {
				if node.Kind() == mathjax.KindMathBlock {
					return r.renderMathBlock(node)
				}
				if node.Kind() == mathjax.KindInlineMath {
					return r.renderInlineMath(node)
				}
			}
			if r.opts.ErrorCallback != nil {
				r.opts.ErrorCallback(NewConversionError("UnsupportedNode", "unsupported markdown node type", 0, 0, nil))
			}
			return ast.WalkContinue, nil
		}
	})
}

// renderHeading renders heading elements
func (r *WordRenderer) renderHeading(node *ast.Heading) (ast.WalkStatus, error) {
	text := r.extractTextContent(node)
	level := node.Level

	if level > 6 {
		level = 6
	}

	if r.opts.GenerateTOC && level <= r.opts.TOCMaxLevel {
		r.doc.AddHeadingWithBookmark(text, level, "")
	} else {
		r.doc.AddHeadingParagraph(text, level)
	}

	return ast.WalkSkipChildren, nil
}

// renderParagraph renders paragraph elements
func (r *WordRenderer) renderParagraph(node *ast.Paragraph) (ast.WalkStatus, error) {
	if !node.HasChildren() {
		return ast.WalkSkipChildren, nil
	}

	para := r.doc.AddParagraph("")
	r.renderInlineContent(node, para)

	return ast.WalkSkipChildren, nil
}

// renderInlineContent renders inline content (text, emphasis, links, etc.)
func (r *WordRenderer) renderInlineContent(node ast.Node, para *document.Paragraph) {
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		switch n := child.(type) {
		case *ast.Text:
			text := string(n.Segment.Value(r.source))
			para.AddFormattedText(text, nil)

			if n.SoftLineBreak() {
				para.AddFormattedText(" ", nil)
			}

		case *ast.Emphasis:
			text := r.extractTextContent(n)
			if n.Level == 2 {
				format := &document.TextFormat{Bold: true}
				para.AddFormattedText(text, format)
			} else {
				format := &document.TextFormat{Italic: true}
				para.AddFormattedText(text, format)
			}

		case *ast.CodeSpan:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				FontFamily: "Consolas",
				FontColor:  "D73A49",
			}
			para.AddFormattedText(text, format)

		case *ast.Link:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				FontColor: "0000FF",
			}
			para.AddFormattedText(text, format)

		case *ast.Image:
			r.renderImageInline(n, para)
		case *extast.Strikethrough:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				Strike: true,
			}
			para.AddFormattedText(text, format)

		default:
			if r.opts.EnableMath && child.Kind() == mathjax.KindInlineMath {
				r.renderInlineMathToParagraph(child, para)
				continue
			}
			text := r.extractTextContent(n)
			if text != "" {
				para.AddFormattedText(text, nil)
			}
		}
	}
}

// renderList renders list elements
func (r *WordRenderer) renderList(node *ast.List) (ast.WalkStatus, error) {
	r.listLevel++
	defer func() { r.listLevel-- }()

	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if listItem, ok := child.(*ast.ListItem); ok {
			r.renderListItem(listItem)
		}
	}

	return ast.WalkSkipChildren, nil
}

// renderListItem renders list item elements
func (r *WordRenderer) renderListItem(node *ast.ListItem) (ast.WalkStatus, error) {
	hasTaskCheckBox := false
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if _, ok := child.(*extast.TaskCheckBox); ok {
			hasTaskCheckBox = true
			break
		}
	}

	if hasTaskCheckBox && r.opts.EnableTaskList {
		return ast.WalkContinue, nil
	}

	text := r.extractTextContent(node)
	indent := strings.Repeat("  ", r.listLevel-1)
	bulletText := "• " + text
	r.doc.AddParagraph(indent + bulletText)

	return ast.WalkSkipChildren, nil
}

// renderBlockquote renders blockquote elements
func (r *WordRenderer) renderBlockquote(node *ast.Blockquote) (ast.WalkStatus, error) {
	text := r.extractTextContent(node)
	para := r.doc.AddParagraph(text)
	para.SetStyle("Quote")

	return ast.WalkSkipChildren, nil
}

// renderCodeBlock renders code block elements
func (r *WordRenderer) renderCodeBlock(node ast.Node) (ast.WalkStatus, error) {
	lines := r.extractCodeBlockLines(node)

	for _, line := range lines {
		if strings.TrimSpace(line) == "" {
			para := r.doc.AddParagraph(" ")
			para.SetStyle("CodeBlock")
			r.applyCodeBlockFormatting(para)
			continue
		}

		para := r.doc.AddParagraph(line)
		para.SetStyle("CodeBlock")
		r.applyCodeBlockFormatting(para)
	}

	return ast.WalkSkipChildren, nil
}

// extractCodeBlockLines extracts code block lines preserving formatting
func (r *WordRenderer) extractCodeBlockLines(node ast.Node) []string {
	var lines []string

	for i := 0; i < node.Lines().Len(); i++ {
		line := node.Lines().At(i)
		lineText := string(line.Value(r.source))
		lines = append(lines, lineText)
	}

	return lines
}

// applyCodeBlockFormatting applies code block formatting to a paragraph
func (r *WordRenderer) applyCodeBlockFormatting(para *document.Paragraph) {
	if para.Properties == nil {
		para.Properties = &document.ParagraphProperties{}
	}

	para.Properties.Indentation = &document.Indentation{
		Left: "360",
	}

	para.Properties.Spacing = &document.Spacing{
		Before: "60",
		After:  "60",
	}

	para.Properties.Justification = &document.Justification{
		Val: "left",
	}
}

// renderThematicBreak renders thematic break (horizontal rule) elements
func (r *WordRenderer) renderThematicBreak(node *ast.ThematicBreak) (ast.WalkStatus, error) {
	para := r.doc.AddParagraph("")
	para.SetHorizontalRule(document.BorderStyleSingle, 12, "000000")
	para.SetSpacing(&document.SpacingConfig{
		BeforePara: 6,
		AfterPara:  6,
	})

	return ast.WalkSkipChildren, nil
}

// renderImage renders image elements
func (r *WordRenderer) renderImage(node *ast.Image) (ast.WalkStatus, error) {
	src := string(node.Destination)
	alt := r.extractTextContent(node)

	if !filepath.IsAbs(src) && r.opts.ImageBasePath != "" {
		src = filepath.Join(r.opts.ImageBasePath, src)
	}

	if alt != "" {
		r.doc.AddParagraph("[Image: " + alt + "]")
	} else {
		r.doc.AddParagraph("[Image: " + src + "]")
	}

	return ast.WalkSkipChildren, nil
}

// renderImageInline renders inline image elements
func (r *WordRenderer) renderImageInline(node *ast.Image, para *document.Paragraph) {
	src := string(node.Destination)
	alt := r.extractTextContent(node)

	if !filepath.IsAbs(src) && r.opts.ImageBasePath != "" {
		src = filepath.Join(r.opts.ImageBasePath, src)
	}

	if alt != "" {
		para.AddFormattedText("[Image: "+alt+"]", nil)
	} else {
		para.AddFormattedText("[Image: "+src+"]", nil)
	}
}

// extractTextContent extracts text content from an AST node
func (r *WordRenderer) extractTextContent(node ast.Node) string {
	var buf strings.Builder
	r.extractTextContentRecursive(node, &buf)
	return buf.String()
}

// extractTextContentRecursive recursively extracts text content
func (r *WordRenderer) extractTextContentRecursive(node ast.Node, buf *strings.Builder) {
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		switch n := child.(type) {
		case *ast.Text:
			buf.Write(n.Segment.Value(r.source))
		default:
			r.extractTextContentRecursive(child, buf)
		}
	}
}

// cleanText cleans excess whitespace from text
func (r *WordRenderer) cleanText(text string) string {
	re := regexp.MustCompile(`\s+`)
	text = re.ReplaceAllString(text, " ")
	return strings.TrimSpace(text)
}

// renderTable renders table elements
func (r *WordRenderer) renderTable(node *extast.Table) (ast.WalkStatus, error) {
	var tableData [][]string
	var alignments []extast.Alignment
	var emphases [][]int

	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if row, ok := child.(*extast.TableHeader); ok {
			var rowData []string
			var rowEmphasis []int
			for cellChild := row.FirstChild(); cellChild != nil; cellChild = cellChild.NextSibling() {
				if cell, ok := cellChild.(*extast.TableCell); ok {
					cellText := r.extractTextContent(cell)
					rowData = append(rowData, cellText)
					rowEmphasis = append(rowEmphasis, 2)
				}
			}
			tableData = append(tableData, rowData)
			emphases = append(emphases, rowEmphasis)
		}
	}

	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if row, ok := child.(*extast.TableRow); ok {
			var rowData []string
			var rowEmphasis []int
			if len(alignments) == 0 {
				alignments = row.Alignments
			}

			for cellChild := row.FirstChild(); cellChild != nil; cellChild = cellChild.NextSibling() {
				if cell, ok := cellChild.(*extast.TableCell); ok {
					cellText := r.extractTextContent(cell)
					rowData = append(rowData, cellText)
					emphasis := extractCellEmphasis(cell)
					rowEmphasis = append(rowEmphasis, emphasis)
				}
			}
			tableData = append(tableData, rowData)
			emphases = append(emphases, rowEmphasis)
		}
	}

	if len(tableData) == 0 {
		return ast.WalkSkipChildren, nil
	}

	cols := 0
	for _, row := range tableData {
		if len(row) > cols {
			cols = len(row)
		}
	}

	config := &document.TableConfig{
		Rows:     len(tableData),
		Cols:     cols,
		Width:    9000,
		Data:     tableData,
		Emphases: emphases,
	}

	table, err := r.doc.AddTable(config)
	if err != nil && r.opts.ErrorCallback != nil {
		r.opts.ErrorCallback(NewConversionError("AddTable", err.Error(), 0, 0, err))
	}
	if table != nil {
		if len(tableData) > 0 {
			err := table.SetRowAsHeader(0, true)
			if err != nil && r.opts.ErrorCallback != nil {
				r.opts.ErrorCallback(NewConversionError("TableHeader", "failed to set table header", 0, 0, err))
			}
		}

		for rowIdx, row := range tableData {
			for colIdx := range row {
				if colIdx < len(alignments) {
					var align document.CellAlignment
					switch alignments[colIdx] {
					case extast.AlignLeft:
						align = document.CellAlignLeft
					case extast.AlignCenter:
						align = document.CellAlignCenter
					case extast.AlignRight:
						align = document.CellAlignRight
					default:
						align = document.CellAlignLeft
					}

					format := &document.CellFormat{
						HorizontalAlign: align,
					}
					err := table.SetCellFormat(rowIdx, colIdx, format)
					if err != nil && r.opts.ErrorCallback != nil {
						r.opts.ErrorCallback(NewConversionError("CellFormat", "failed to set cell format", rowIdx, colIdx, err))
					}
				}
			}
		}
	}

	return ast.WalkSkipChildren, nil
}

// extractCellEmphasis extracts emphasis level from a table cell
func extractCellEmphasis(cell *extast.TableCell) int {
	format := 0 // 0 = no format
	// iterate over cell content
	ast.Walk(cell, func(n ast.Node, entering bool) (ast.WalkStatus, error) {
		if !entering {
			return ast.WalkContinue, nil
		}

		switch node := n.(type) {
		case *ast.Emphasis:
			format = node.Level
		}

		return ast.WalkContinue, nil
	})

	return format
}

// renderTaskCheckBox renders task list checkbox elements
func (r *WordRenderer) renderTaskCheckBox(node *extast.TaskCheckBox) (ast.WalkStatus, error) {
	checked := node.IsChecked

	var checkSymbol string
	if checked {
		checkSymbol = "☑"
	} else {
		checkSymbol = "☐"
	}

	para := r.doc.AddParagraph("")
	para.AddFormattedText(checkSymbol+" ", nil)

	parent := node.Parent()
	if parent != nil {
		r.renderTaskItemContent(parent, para, node)
	}

	return ast.WalkSkipChildren, nil
}

// renderTaskItemContent renders task item content (excluding the checkbox)
func (r *WordRenderer) renderTaskItemContent(parent ast.Node, para *document.Paragraph, skipNode ast.Node) {
	for child := parent.FirstChild(); child != nil; child = child.NextSibling() {
		if child == skipNode {
			continue
		}

		switch n := child.(type) {
		case *ast.Text:
			text := string(n.Segment.Value(r.source))
			para.AddFormattedText(text, nil)

			if n.SoftLineBreak() {
				para.AddFormattedText(" ", nil)
			}
		case *ast.Emphasis:
			text := r.extractTextContent(n)
			if n.Level == 2 {
				format := &document.TextFormat{Bold: true}
				para.AddFormattedText(text, format)
			} else {
				format := &document.TextFormat{Italic: true}
				para.AddFormattedText(text, format)
			}
		case *ast.CodeSpan:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				FontFamily: "Consolas",
			}
			para.AddFormattedText(text, format)
		case *ast.Link:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				FontColor: "0000FF", // blue
			}
			para.AddFormattedText(text, format)
		default:
			if r.opts.EnableMath && child.Kind() == mathjax.KindInlineMath {
				r.renderInlineMathToParagraph(child, para)
				continue
			}
			text := r.extractTextContent(n)
			if text != "" {
				para.AddFormattedText(text, nil)
			}
		}
	}
}

// renderInlineMathToParagraph renders inline math formula to a paragraph
func (r *WordRenderer) renderInlineMathToParagraph(node ast.Node, para *document.Paragraph) {
	latex := r.extractMathContent(node)
	para.AddFormattedText(latex, &document.TextFormat{
		FontFamily: "Cambria Math",
	})
}

// renderMathBlock renders block-level math formula
func (r *WordRenderer) renderMathBlock(node ast.Node) (ast.WalkStatus, error) {
	latex := r.extractMathContent(node)
	para := r.doc.AddParagraph("")
	para.SetAlignment(document.AlignCenter)
	para.AddFormattedText(latex, &document.TextFormat{
		FontFamily: "Cambria Math",
		FontSize:   12,
	})

	return ast.WalkSkipChildren, nil
}

// renderInlineMath renders inline math formula
func (r *WordRenderer) renderInlineMath(node ast.Node) (ast.WalkStatus, error) {
	para := r.doc.AddParagraph("")
	r.renderInlineMathToParagraph(node, para)

	return ast.WalkSkipChildren, nil
}

// extractMathContent extracts LaTeX content from a math node
func (r *WordRenderer) extractMathContent(node ast.Node) string {
	var content strings.Builder

	if node.Kind() == mathjax.KindMathBlock {
		if blockNode, ok := node.(ast.Node); ok {
			lines := blockNode.Lines()
			if lines != nil {
				for i := 0; i < lines.Len(); i++ {
					line := lines.At(i)
					content.Write(line.Value(r.source))
				}
			}
		}
	}

	if content.Len() == 0 {
		for child := node.FirstChild(); child != nil; child = child.NextSibling() {
			if text, ok := child.(*ast.Text); ok {
				content.Write(text.Segment.Value(r.source))
			}
		}
	}

	latex := strings.TrimSpace(content.String())
	latex = convertLaTeXToDisplay(latex)

	return latex
}

// convertLaTeXToDisplay converts LaTeX commands to displayable Unicode characters
func convertLaTeXToDisplay(latex string) string {
	fracPattern := regexp.MustCompile(`\\frac\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}`)
	for fracPattern.MatchString(latex) {
		latex = fracPattern.ReplaceAllStringFunc(latex, func(match string) string {
			parts := fracPattern.FindStringSubmatch(match)
			if len(parts) == 3 {
				num := convertLaTeXToDisplay(parts[1])
				den := convertLaTeXToDisplay(parts[2])
				return "(" + num + ")/(" + den + ")"
			}
			return match
		})
	}

	sqrtPattern := regexp.MustCompile(`\\sqrt\s*(?:\[([^\]]*)\])?\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}`)
	for sqrtPattern.MatchString(latex) {
		latex = sqrtPattern.ReplaceAllStringFunc(latex, func(match string) string {
			parts := sqrtPattern.FindStringSubmatch(match)
			if len(parts) == 3 {
				deg := parts[1]
				content := convertLaTeXToDisplay(parts[2])
				if deg == "" {
					return "√(" + content + ")"
				}
				degSup := convertToSuperscript(deg)
				return degSup + "√(" + content + ")"
			}
			return match
		})
	}

	supBracePattern := regexp.MustCompile(`\^\{([^{}]*)\}`)
	latex = supBracePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := supBracePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSuperscript(parts[1])
		}
		return match
	})
	supSimplePattern := regexp.MustCompile(`\^([a-zA-Z0-9])`)
	latex = supSimplePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := supSimplePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSuperscript(parts[1])
		}
		return match
	})

	subBracePattern := regexp.MustCompile(`_\{([^{}]*)\}`)
	latex = subBracePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := subBracePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSubscript(parts[1])
		}
		return match
	})
	subSimplePattern := regexp.MustCompile(`_([a-zA-Z0-9])`)
	latex = subSimplePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := subSimplePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSubscript(parts[1])
		}
		return match
	})

	replacements := map[string]string{
		`\alpha`:   "α",
		`\beta`:    "β",
		`\gamma`:   "γ",
		`\delta`:   "δ",
		`\epsilon`: "ε",
		`\zeta`:    "ζ",
		`\eta`:     "η",
		`\theta`:   "θ",
		`\iota`:    "ι",
		`\kappa`:   "κ",
		`\lambda`:  "λ",
		`\mu`:      "μ",
		`\nu`:      "ν",
		`\xi`:      "ξ",
		`\pi`:      "π",
		`\rho`:     "ρ",
		`\sigma`:   "σ",
		`\tau`:     "τ",
		`\upsilon`: "υ",
		`\phi`:     "φ",
		`\chi`:     "χ",
		`\psi`:     "ψ",
		`\omega`:   "ω",

		// Greek letters (uppercase)
		`\Alpha`:   "Α",
		`\Beta`:    "Β",
		`\Gamma`:   "Γ",
		`\Delta`:   "Δ",
		`\Epsilon`: "Ε",
		`\Zeta`:    "Ζ",
		`\Eta`:     "Η",
		`\Theta`:   "Θ",
		`\Iota`:    "Ι",
		`\Kappa`:   "Κ",
		`\Lambda`:  "Λ",
		`\Mu`:      "Μ",
		`\Nu`:      "Ν",
		`\Xi`:      "Ξ",
		`\Pi`:      "Π",
		`\Rho`:     "Ρ",
		`\Sigma`:   "Σ",
		`\Tau`:     "Τ",
		`\Upsilon`: "Υ",
		`\Phi`:     "Φ",
		`\Chi`:     "Χ",
		`\Psi`:     "Ψ",
		`\Omega`:   "Ω",

		// operators
		`\times`:  "×",
		`\div`:    "÷",
		`\pm`:     "±",
		`\mp`:     "∓",
		`\cdot`:   "·",
		`\ast`:    "∗",
		`\star`:   "⋆",
		`\circ`:   "∘",
		`\bullet`: "•",
		`\oplus`:  "⊕",
		`\ominus`: "⊖",
		`\otimes`: "⊗",

		// relational symbols
		`\leq`:      "≤",
		`\le`:       "≤",
		`\geq`:      "≥",
		`\ge`:       "≥",
		`\neq`:      "≠",
		`\ne`:       "≠",
		`\approx`:   "≈",
		`\equiv`:    "≡",
		`\sim`:      "∼",
		`\simeq`:    "≃",
		`\cong`:     "≅",
		`\propto`:   "∝",
		`\ll`:       "≪",
		`\gg`:       "≫",
		`\subset`:   "⊂",
		`\supset`:   "⊃",
		`\subseteq`: "⊆",
		`\supseteq`: "⊇",
		`\in`:       "∈",
		`\notin`:    "∉",
		`\ni`:       "∋",

		// arrows
		`\rightarrow`:     "→",
		`\leftarrow`:      "←",
		`\leftrightarrow`: "↔",
		`\Rightarrow`:     "⇒",
		`\Leftarrow`:      "⇐",
		`\Leftrightarrow`: "⇔",
		`\uparrow`:        "↑",
		`\downarrow`:      "↓",
		`\to`:             "→",
		`\gets`:           "←",
		`\mapsto`:         "↦",

		// miscellaneous symbols
		`\infty`:      "∞",
		`\partial`:    "∂",
		`\nabla`:      "∇",
		`\forall`:     "∀",
		`\exists`:     "∃",
		`\nexists`:    "∄",
		`\emptyset`:   "∅",
		`\varnothing`: "∅",
		`\neg`:        "¬",
		`\lnot`:       "¬",
		`\land`:       "∧",
		`\lor`:        "∨",
		`\cap`:        "∩",
		`\cup`:        "∪",
		`\int`:        "∫",
		`\iint`:       "∬",
		`\iiint`:      "∭",
		`\oint`:       "∮",
		`\sum`:        "∑",
		`\prod`:       "∏",
		`\coprod`:     "∐",

		// ellipsis
		`\ldots`: "…",
		`\cdots`: "⋯",
		`\vdots`: "⋮",
		`\ddots`: "⋱",

		// whitespace
		`\quad`:  " ",
		`\qquad`: "  ",
		`\,`:     " ",
		`\;`:     " ",
		`\:`:     " ",
		`\ `:     " ",

		// parentheses
		`\{`:      "{",
		`\}`:      "}",
		`\lbrace`: "{",
		`\rbrace`: "}",
		`\langle`: "⟨",
		`\rangle`: "⟩",
		`\lceil`:  "⌈",
		`\rceil`:  "⌉",
		`\lfloor`: "⌊",
		`\rfloor`: "⌋",
		`\left`:   "",
		`\right`:  "",
	}

	type replacement struct {
		cmd     string
		unicode string
	}
	sortedReplacements := make([]replacement, 0, len(replacements))
	for cmd, u := range replacements {
		sortedReplacements = append(sortedReplacements, replacement{cmd, u})
	}
	// sort by command length descending
	for i := 0; i < len(sortedReplacements); i++ {
		for j := i + 1; j < len(sortedReplacements); j++ {
			if len(sortedReplacements[j].cmd) > len(sortedReplacements[i].cmd) {
				sortedReplacements[i], sortedReplacements[j] = sortedReplacements[j], sortedReplacements[i]
			}
		}
	}
	for _, r := range sortedReplacements {
		latex = strings.ReplaceAll(latex, r.cmd, r.unicode)
	}

	latex = strings.ReplaceAll(latex, "{", "")
	latex = strings.ReplaceAll(latex, "}", "")

	return latex
}

// convertToSuperscript converts string to superscript form
func convertToSuperscript(s string) string {
	superscripts := map[rune]rune{
		'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
		'5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
		'+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
		'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
		'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'i': 'ⁱ', 'j': 'ʲ',
		'k': 'ᵏ', 'l': 'ˡ', 'm': 'ᵐ', 'n': 'ⁿ', 'o': 'ᵒ',
		'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ',
		'v': 'ᵛ', 'w': 'ʷ', 'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ',
	}

	var result strings.Builder
	for _, r := range s {
		if sup, ok := superscripts[r]; ok {
			result.WriteRune(sup)
		} else {
			result.WriteRune(r)
		}
	}
	return result.String()
}

// convertToSubscript converts string to subscript form
func convertToSubscript(s string) string {
	subscripts := map[rune]rune{
		'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
		'5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
		'+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
		'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ',
		'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ',
		'p': 'ₚ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ',
		'v': 'ᵥ', 'x': 'ₓ',
	}

	var result strings.Builder
	for _, r := range s {
		if sub, ok := subscripts[r]; ok {
			result.WriteRune(sub)
		} else {
			result.WriteRune(r)
		}
	}
	return result.String()
}
