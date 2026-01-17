// Package markdown provides Markdown to Word document conversion functionality
package markdown

import (
	"encoding/xml"
	"regexp"
	"strings"
)

// OfficeMath represents the root element of an Office math formula (m:oMath in OMML)
type OfficeMath struct {
	XMLName xml.Name      `xml:"m:oMath"`
	Content []interface{} `xml:"-"` // custom serialization
}

// MarshalXML implements custom XML serialization
func (o *OfficeMath) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:oMath"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range o.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// OfficeMathPara represents an Office math formula paragraph (m:oMathPara in OMML, for block-level formulas)
type OfficeMathPara struct {
	XMLName xml.Name    `xml:"m:oMathPara"`
	Math    *OfficeMath `xml:"m:oMath"`
}

// MathRun represents a math run element
type MathRun struct {
	XMLName xml.Name     `xml:"m:r"`
	Text    *MathText    `xml:"m:t,omitempty"`
	RunPr   *MathRunProp `xml:"m:rPr,omitempty"`
}

// MathText represents math text
type MathText struct {
	XMLName xml.Name `xml:"m:t"`
	Content string   `xml:",chardata"`
}

// MathRunProp represents math run properties
type MathRunProp struct {
	XMLName xml.Name `xml:"m:rPr"`
	Sty     *MathSty `xml:"m:sty,omitempty"`
}

// MathSty represents a math style
type MathSty struct {
	XMLName xml.Name `xml:"m:sty"`
	Val     string   `xml:"m:val,attr"`
}

// MathFrac represents a fraction
type MathFrac struct {
	XMLName xml.Name     `xml:"m:f"`
	FracPr  *MathFracPr  `xml:"m:fPr,omitempty"`
	Num     *MathNum     `xml:"m:num"`
	Den     *MathDen     `xml:"m:den"`
}

// MathFracPr represents fraction properties
type MathFracPr struct {
	XMLName xml.Name      `xml:"m:fPr"`
	Type    *MathFracType `xml:"m:type,omitempty"`
}

// MathFracType represents fraction type
type MathFracType struct {
	XMLName xml.Name `xml:"m:type"`
	Val     string   `xml:"m:val,attr"`
}

// MathNum represents the numerator
type MathNum struct {
	XMLName xml.Name      `xml:"m:num"`
	Content []interface{} `xml:"-"`
}

// MarshalXML implements custom XML serialization
func (n *MathNum) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:num"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range n.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// MathDen represents the denominator
type MathDen struct {
	XMLName xml.Name      `xml:"m:den"`
	Content []interface{} `xml:"-"`
}

// MarshalXML implements custom XML serialization
func (d *MathDen) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:den"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range d.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// MathSup represents a superscript
type MathSup struct {
	XMLName xml.Name          `xml:"m:sSup"`
	E       *MathE            `xml:"m:e"`
	Sup     *MathSupElement   `xml:"m:sup"`
}

// MathE represents a base element
type MathE struct {
	XMLName xml.Name      `xml:"m:e"`
	Content []interface{} `xml:"-"`
}

// MarshalXML implements custom XML serialization
func (m *MathE) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:e"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range m.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// MathSupElement represents a superscript element
type MathSupElement struct {
	XMLName xml.Name      `xml:"m:sup"`
	Content []interface{} `xml:"-"`
}

// MarshalXML implements custom XML serialization
func (s *MathSupElement) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:sup"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range s.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// MathSub represents a subscript
type MathSub struct {
	XMLName xml.Name         `xml:"m:sSub"`
	E       *MathE           `xml:"m:e"`
	Sub     *MathSubElement  `xml:"m:sub"`
}

// MathSubElement represents a subscript element
type MathSubElement struct {
	XMLName xml.Name      `xml:"m:sub"`
	Content []interface{} `xml:"-"`
}

// MarshalXML implements custom XML serialization
func (s *MathSubElement) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:sub"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range s.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// MathRad represents a radical (root)
type MathRad struct {
	XMLName xml.Name   `xml:"m:rad"`
	RadPr   *MathRadPr `xml:"m:radPr,omitempty"`
	Deg     *MathDeg   `xml:"m:deg,omitempty"`
	E       *MathE     `xml:"m:e"`
}

// MathRadPr represents radical properties
type MathRadPr struct {
	XMLName xml.Name     `xml:"m:radPr"`
	DegHide *MathDegHide `xml:"m:degHide,omitempty"`
}

// MathDegHide represents whether to hide the radical degree
type MathDegHide struct {
	XMLName xml.Name `xml:"m:degHide"`
	Val     string   `xml:"m:val,attr"`
}

// MathDeg represents the degree of a radical
type MathDeg struct {
	XMLName xml.Name      `xml:"m:deg"`
	Content []interface{} `xml:"-"`
}

// MarshalXML implements custom XML serialization
func (d *MathDeg) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "m:deg"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, content := range d.Content {
		if err := e.Encode(content); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// MathSubSup represents both subscript and superscript
type MathSubSup struct {
	XMLName xml.Name        `xml:"m:sSubSup"`
	E       *MathE          `xml:"m:e"`
	Sub     *MathSubElement `xml:"m:sub"`
	Sup     *MathSupElement `xml:"m:sup"`
}

// MathDelim represents a delimiter (parentheses, brackets, etc.)
type MathDelim struct {
	XMLName xml.Name     `xml:"m:d"`
	DPr     *MathDelimPr `xml:"m:dPr,omitempty"`
	E       *MathE       `xml:"m:e"`
}

// MathDelimPr represents delimiter properties
type MathDelimPr struct {
	XMLName xml.Name       `xml:"m:dPr"`
	BegChr  *MathDelimChar `xml:"m:begChr,omitempty"`
	EndChr  *MathDelimChar `xml:"m:endChr,omitempty"`
}

// MathDelimChar represents a delimiter character
type MathDelimChar struct {
	XMLName xml.Name `xml:"m:begChr"`
	Val     string   `xml:"m:val,attr"`
}

// LaTeXToOMML converts LaTeX formulas to OMML format
// This is a simplified converter supporting common LaTeX math syntax
func LaTeXToOMML(latex string) *OfficeMath {
	latex = strings.TrimSpace(latex)
	omath := &OfficeMath{
		Content: []interface{}{},
	}

	// Parse LaTeX and convert to OMML
	content := parseLatex(latex)
	omath.Content = content

	return omath
}

// parseLatex parses LaTeX string and returns a list of OMML elements
func parseLatex(latex string) []interface{} {
	var result []interface{}
	latex = strings.TrimSpace(latex)

	// Handle empty string
	if latex == "" {
		return result
	}

	// Define regex patterns
	fracPattern := regexp.MustCompile(`^\\frac\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}`)
	sqrtPattern := regexp.MustCompile(`^\\sqrt(?:\[([^\]]*)\])?\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}`)
	supPattern := regexp.MustCompile(`^([a-zA-Z0-9])\^(?:\{([^{}]*)\}|([a-zA-Z0-9]))`)
	subPattern := regexp.MustCompile(`^([a-zA-Z0-9])_(?:\{([^{}]*)\}|([a-zA-Z0-9]))`)
	subSupPattern := regexp.MustCompile(`^([a-zA-Z0-9])_(?:\{([^{}]*)\}|([a-zA-Z0-9]))\^(?:\{([^{}]*)\}|([a-zA-Z0-9]))`)
	cmdPattern := regexp.MustCompile(`^\\([a-zA-Z]+)`)
	textPattern := regexp.MustCompile(`^[a-zA-Z0-9.,;:!?\s\+\-\*\/\=\(\)\[\]]+`)

	i := 0
	for i < len(latex) {
		remaining := latex[i:]

		// Check subscript and superscript combination
		if match := subSupPattern.FindStringSubmatch(remaining); match != nil {
			base := match[1]
			sub := match[2]
			if sub == "" {
				sub = match[3]
			}
			sup := match[4]
			if sup == "" {
				sup = match[5]
			}

			result = append(result, &MathSubSup{
				E:   &MathE{Content: []interface{}{createMathRun(base)}},
				Sub: &MathSubElement{Content: parseLatex(sub)},
				Sup: &MathSupElement{Content: parseLatex(sup)},
			})
			i += len(match[0])
			continue
		}

		if match := fracPattern.FindStringSubmatch(remaining); match != nil {
			num := match[1]
			den := match[2]
			result = append(result, &MathFrac{
				Num: &MathNum{Content: parseLatex(num)},
				Den: &MathDen{Content: parseLatex(den)},
			})
			i += len(match[0])
			continue
		}

		if match := sqrtPattern.FindStringSubmatch(remaining); match != nil {
			deg := match[1]  // may be empty (square root)
			content := match[2]
			rad := &MathRad{
				E: &MathE{Content: parseLatex(content)},
			}
			if deg == "" {
				rad.RadPr = &MathRadPr{
					DegHide: &MathDegHide{Val: "1"},
				}
			} else {
				rad.Deg = &MathDeg{Content: parseLatex(deg)}
			}
			result = append(result, rad)
			i += len(match[0])
			continue
		}

		if match := supPattern.FindStringSubmatch(remaining); match != nil {
			base := match[1]
			sup := match[2]
			if sup == "" {
				sup = match[3]
			}
			result = append(result, &MathSup{
				E:   &MathE{Content: []interface{}{createMathRun(base)}},
				Sup: &MathSupElement{Content: parseLatex(sup)},
			})
			i += len(match[0])
			continue
		}

		if match := subPattern.FindStringSubmatch(remaining); match != nil {
			base := match[1]
			sub := match[2]
			if sub == "" {
				sub = match[3]
			}
			result = append(result, &MathSub{
				E:   &MathE{Content: []interface{}{createMathRun(base)}},
				Sub: &MathSubElement{Content: parseLatex(sub)},
			})
			i += len(match[0])
			continue
		}

		if match := cmdPattern.FindStringSubmatch(remaining); match != nil {
			cmd := match[1]
			cmdText := convertLaTeXCommand(cmd)
			result = append(result, createMathRun(cmdText))
			i += len(match[0])
			continue
		}

		if remaining[0] == '{' {
			depth := 1
			j := 1
			for j < len(remaining) && depth > 0 {
				if remaining[j] == '{' {
					depth++
				} else if remaining[j] == '}' {
					depth--
				}
				j++
			}
			if depth == 0 {
				inner := remaining[1 : j-1]
				innerContent := parseLatex(inner)
				result = append(result, innerContent...)
				i += j
				continue
			}
		}

		if match := textPattern.FindString(remaining); match != "" {
			result = append(result, createMathRun(match))
			i += len(match)
			continue
		}

		if i < len(latex) {
			result = append(result, createMathRun(string(latex[i])))
			i++
		}
	}

	return result
}

// createMathRun creates a math run element
func createMathRun(text string) *MathRun {
	return &MathRun{
		Text: &MathText{Content: text},
	}
}

// convertLaTeXCommand converts LaTeX commands to corresponding Unicode characters
func convertLaTeXCommand(cmd string) string {
	commands := map[string]string{
		"alpha":   "α",
		"beta":    "β",
		"gamma":   "γ",
		"delta":   "δ",
		"epsilon": "ε",
		"zeta":    "ζ",
		"eta":     "η",
		"theta":   "θ",
		"iota":    "ι",
		"kappa":   "κ",
		"lambda":  "λ",
		"mu":      "μ",
		"nu":      "ν",
		"xi":      "ξ",
		"pi":      "π",
		"rho":     "ρ",
		"sigma":   "σ",
		"tau":     "τ",
		"upsilon": "υ",
		"phi":     "φ",
		"chi":     "χ",
		"psi":     "ψ",
		"omega":   "ω",

		"Alpha":   "Α",
		"Beta":    "Β",
		"Gamma":   "Γ",
		"Delta":   "Δ",
		"Epsilon": "Ε",
		"Zeta":    "Ζ",
		"Eta":     "Η",
		"Theta":   "Θ",
		"Iota":    "Ι",
		"Kappa":   "Κ",
		"Lambda":  "Λ",
		"Mu":      "Μ",
		"Nu":      "Ν",
		"Xi":      "Ξ",
		"Pi":      "Π",
		"Rho":     "Ρ",
		"Sigma":   "Σ",
		"Tau":     "Τ",
		"Upsilon": "Υ",
		"Phi":     "Φ",
		"Chi":     "Χ",
		"Psi":     "Ψ",
		"Omega":   "Ω",

		"times":   "×",
		"div":     "÷",
		"pm":      "±",
		"mp":      "∓",
		"cdot":    "·",
		"ast":     "∗",
		"star":    "⋆",
		"circ":    "∘",
		"bullet":  "∙",
		"oplus":   "⊕",
		"ominus":  "⊖",
		"otimes":  "⊗",
		"oslash":  "⊘",
		"odot":    "⊙",

		"leq":     "≤",
		"geq":     "≥",
		"neq":     "≠",
		"approx":  "≈",
		"equiv":   "≡",
		"sim":     "∼",
		"simeq":   "≃",
		"cong":    "≅",
		"propto":   "∝",
		"ll":       "≪",
		"gg":       "≫",
		"subset":   "⊂",
		"supset":   "⊃",
		"subseteq": "⊆",
		"supseteq": "⊇",
		"in":       "∈",
		"notin":    "∉",
		"ni":       "∋",

		"rightarrow":     "→",
		"leftarrow":      "←",
		"leftrightarrow": "↔",
		"Rightarrow":     "⇒",
		"Leftarrow":      "⇐",
		"Leftrightarrow": "⇔",
		"uparrow":        "↑",
		"downarrow":      "↓",
		"to":             "→",
		"gets":           "←",
		"mapsto":         "↦",

		"infty":    "∞",
		"partial":  "∂",
		"nabla":    "∇",
		"forall":   "∀",
		"exists":   "∃",
		"nexists":  "∄",
		"emptyset": "∅",
		"varnothing": "∅",
		"neg":      "¬",
		"lnot":     "¬",
		"land":     "∧",
		"lor":      "∨",
		"cap":      "∩",
		"cup":      "∪",
		"int":      "∫",
		"iint":     "∬",
		"iiint":    "∭",
		"oint":     "∮",
		"sum":      "∑",
		"prod":     "∏",
		"coprod":   "∐",

		"lim":      "lim",
		"limsup":   "lim sup",
		"liminf":   "lim inf",
		"max":      "max",
		"min":      "min",
		"sup":      "sup",
		"inf":      "inf",
		"sin":      "sin",
		"cos":      "cos",
		"tan":      "tan",
		"cot":      "cot",
		"sec":      "sec",
		"csc":      "csc",
		"arcsin":   "arcsin",
		"arccos":   "arccos",
		"arctan":   "arctan",
		"sinh":     "sinh",
		"cosh":     "cosh",
		"tanh":     "tanh",
		"log":      "log",
		"ln":       "ln",
		"exp":      "exp",
		"deg":      "deg",
		"det":      "det",
		"dim":      "dim",
		"ker":      "ker",
		"hom":      "hom",
		"arg":      "arg",
		"gcd":      "gcd",

		"lbrace":   "{",
		"rbrace":   "}",
		"langle":   "⟨",
		"rangle":   "⟩",
		"lceil":    "⌈",
		"rceil":    "⌉",
		"lfloor":   "⌊",
		"rfloor":   "⌋",
		"left":     "",
		"right":    "",

		"ldots":    "…",
		"cdots":    "⋯",
		"vdots":    "⋮",
		"ddots":    "⋱",
		"quad":     " ",
		"qquad":    "  ",
		"space":    " ",
	}

	if result, ok := commands[cmd]; ok {
		return result
	}
	return "\\" + cmd // unknown command keep original
}

// LaTeXToOMMLString converts LaTeX formula to OMML XML string
func LaTeXToOMMLString(latex string, isBlock bool) (string, error) {
	omath := LaTeXToOMML(latex)

	var result interface{}
	if isBlock {
		result = &OfficeMathPara{Math: omath}
	} else {
		result = omath
	}

	data, err := xml.Marshal(result)
	if err != nil {
		return "", err
	}

	return string(data), nil
}
