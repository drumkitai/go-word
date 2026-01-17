// Package document provides Word document field structures
package document

import (
	"encoding/xml"
	"fmt"
)

// FieldChar field character
type FieldChar struct {
	XMLName       xml.Name `xml:"w:fldChar"`
	FieldCharType string   `xml:"w:fldCharType,attr"`
}

// InstrText field instruction text
type InstrText struct {
	XMLName xml.Name `xml:"w:instrText"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Content string   `xml:",chardata"`
}

// HyperlinkField hyperlink field
type HyperlinkField struct {
	BeginChar    FieldChar
	InstrText    InstrText
	SeparateChar FieldChar
	EndChar      FieldChar
}

// CreateHyperlinkField creates a hyperlink field
func CreateHyperlinkField(anchor string) HyperlinkField {
	return HyperlinkField{
		BeginChar: FieldChar{
			FieldCharType: "begin",
		},
		InstrText: InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" HYPERLINK \\l %s ", anchor),
		},
		SeparateChar: FieldChar{
			FieldCharType: "separate",
		},
		EndChar: FieldChar{
			FieldCharType: "end",
		},
	}
}

// PageRefField page reference field
type PageRefField struct {
	BeginChar    FieldChar
	InstrText    InstrText
	SeparateChar FieldChar
	EndChar      FieldChar
}

// CreatePageRefField creates a page reference field
func CreatePageRefField(anchor string) PageRefField {
	return PageRefField{
		BeginChar: FieldChar{
			FieldCharType: "begin",
		},
		InstrText: InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" PAGEREF %s \\h ", anchor),
		},
		SeparateChar: FieldChar{
			FieldCharType: "separate",
		},
		EndChar: FieldChar{
			FieldCharType: "end",
		},
	}
}
