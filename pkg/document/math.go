// Package document provides Word document core functionality
package document

import (
	"encoding/xml"
)

// OfficeMath represents an Office math formula element
// Corresponds to m:oMath element in OMML
type OfficeMath struct {
	XMLName xml.Name `xml:"m:oMath"`
	Xmlns   string   `xml:"xmlns:m,attr,omitempty"`
	RawXML  string   `xml:",innerxml"` // stores internal XML content
}

// OfficeMathPara represents an Office math formula paragraph (for block-level formulas)
// Corresponds to m:oMathPara element in OMML
type OfficeMathPara struct {
	XMLName xml.Name    `xml:"m:oMathPara"`
	Xmlns   string      `xml:"xmlns:m,attr,omitempty"`
	Math    *OfficeMath `xml:"m:oMath"`
}

// MathParagraph represents a paragraph containing math formulas
// Used to embed math formulas in documents
type MathParagraph struct {
	XMLName    xml.Name             `xml:"w:p"`
	Properties *ParagraphProperties `xml:"w:pPr,omitempty"`
	Math       *OfficeMath          `xml:"m:oMath,omitempty"`
	MathPara   *OfficeMathPara      `xml:"m:oMathPara,omitempty"`
	Runs       []Run                `xml:"w:r"`
}

// MarshalXML custom serialization
func (mp *MathParagraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Start paragraph element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Serialize paragraph properties
	if mp.Properties != nil {
		if err := e.Encode(mp.Properties); err != nil {
			return err
		}
	}

	// Serialize Runs (text before the formula)
	for _, run := range mp.Runs {
		if err := e.Encode(run); err != nil {
			return err
		}
	}

	// Serialize math formula (block-level)
	if mp.MathPara != nil {
		if err := e.Encode(mp.MathPara); err != nil {
			return err
		}
	}

	// Serialize math formula (inline)
	if mp.Math != nil {
		if err := e.Encode(mp.Math); err != nil {
			return err
		}
	}

	// End paragraph element
	return e.EncodeToken(start.End())
}

// ElementType returns the math paragraph element type
func (mp *MathParagraph) ElementType() string {
	return "math_paragraph"
}

// AddMathFormula adds a math formula to the document
// latex: math formula in LaTeX format
// isBlock: whether the formula is block-level (true for block, false for inline)
func (d *Document) AddMathFormula(latex string, isBlock bool) *MathParagraph {
	Debugf("Adding math formula: %s (block: %v)", latex, isBlock)

	mp := &MathParagraph{
		Runs: []Run{},
	}

	// Create formula content
	// Note: using RawXML to store formula content because OMML structure is complex
	// Actual LaTeX to OMML conversion is done by the markdown package's LaTeXToOMML function
	if isBlock {
		mp.MathPara = &OfficeMathPara{
			Xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/math",
			Math: &OfficeMath{
				Xmlns:  "http://schemas.openxmlformats.org/officeDocument/2006/math",
				RawXML: latex, // stores preprocessed OMML content here
			},
		}
	} else {
		mp.Math = &OfficeMath{
			Xmlns:  "http://schemas.openxmlformats.org/officeDocument/2006/math",
			RawXML: latex,
		}
	}

	d.Body.Elements = append(d.Body.Elements, mp)
	return mp
}

// AddInlineMathFormula adds an inline math formula to a paragraph
// This adds a math formula at the end of the current paragraph
func (p *Paragraph) AddInlineMath(ommlContent string) {
	Debugf("Adding inline math formula to paragraph")

	// Create a special Run to contain formula reference
	// Note: In Word, inline formulas are implemented through special oMath elements
	// Here we use a placeholder method, actual implementation needs to modify paragraph serialization logic
	run := Run{
		Text: Text{
			Content: "[Formula]", // Placeholder, actual formula content is processed during serialization
			Space:   "preserve",
		},
	}
	p.Runs = append(p.Runs, run)
}
