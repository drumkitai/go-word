package test

import (
	"fmt"
	"testing"

	"github.com/drumkitai/go-word/pkg/document"
)

func TestFootnoteConfig(t *testing.T) {
	doc := document.New()

	config := &document.FootnoteConfig{
		NumberFormat: document.FootnoteFormatDecimal,
		StartNumber:  1,
		RestartEach:  document.FootnoteRestartContinuous,
		Position:     document.FootnotePositionPageBottom,
	}

	err := doc.SetFootnoteConfig(config)
	if err != nil {
		fmt.Printf("Failed to set footnote configuration: %v\n", err)

		parts := doc.GetParts()
		if settingsXML, exists := parts["word/settings.xml"]; exists {
			fmt.Printf("Generated settings.xml content:\n%s\n", string(settingsXML))
		}

		t.Fatalf("Failed to set footnote configuration: %v", err)
	}

	_, exists := doc.GetParts()["word/settings.xml"]
	if !exists {
		t.Error("settings.xml file not created")
	}

	err = doc.AddFootnote("This is the main text", "This is the footnote content")
	if err != nil {
		t.Fatalf("Failed to add footnote: %v", err)
	}

	_, exists = doc.GetParts()["word/footnotes.xml"]
	if !exists {
		t.Error("footnotes.xml file not created")
	}

	count := doc.GetFootnoteCount()
	if count != 1 {
		t.Errorf("Expected footnote count to be 1, got %d", count)
	}
}

func TestEndnoteConfig(t *testing.T) {
	doc := document.New()

	err := doc.AddEndnote("This is the main text", "This is the endnote content")
	if err != nil {
		t.Fatalf("Failed to add endnote: %v", err)
	}

	_, exists := doc.GetParts()["word/endnotes.xml"]
	if !exists {
		t.Error("endnotes.xml file not created")
	}

	count := doc.GetEndnoteCount()
	if count != 1 {
		t.Errorf("Expected endnote count to be 1, got %d", count)
	}
}

func TestFootnoteNumberFormats(t *testing.T) {
	doc := document.New()

	// 测试不同的编号格式
	formats := []document.FootnoteNumberFormat{
		document.FootnoteFormatDecimal,
		document.FootnoteFormatLowerRoman,
		document.FootnoteFormatUpperRoman,
		document.FootnoteFormatLowerLetter,
		document.FootnoteFormatUpperLetter,
		document.FootnoteFormatSymbol,
	}

	for _, format := range formats {
		config := &document.FootnoteConfig{
			NumberFormat: format,
			StartNumber:  1,
			RestartEach:  document.FootnoteRestartContinuous,
			Position:     document.FootnotePositionPageBottom,
		}

		err := doc.SetFootnoteConfig(config)
		if err != nil {
			t.Fatalf("Failed to set footnote format %s: %v", format, err)
		}
	}
}

func TestFootnotePositions(t *testing.T) {
	doc := document.New()

	positions := []document.FootnotePosition{
		document.FootnotePositionPageBottom,
		document.FootnotePositionBeneathText,
		document.FootnotePositionSectionEnd,
		document.FootnotePositionDocumentEnd,
	}

	for _, position := range positions {
		config := &document.FootnoteConfig{
			NumberFormat: document.FootnoteFormatDecimal,
			StartNumber:  1,
			RestartEach:  document.FootnoteRestartContinuous,
			Position:     position,
		}

		err := doc.SetFootnoteConfig(config)
		if err != nil {
			t.Fatalf("Failed to set footnote position %s: %v", position, err)
		}
	}
}

func TestDefaultFootnoteConfig(t *testing.T) {
	config := document.DefaultFootnoteConfig()

	if config.NumberFormat != document.FootnoteFormatDecimal {
		t.Errorf("Default number format error, expected %s, got %s",
			document.FootnoteFormatDecimal, config.NumberFormat)
	}

	if config.StartNumber != 1 {
		t.Errorf("Default start number error, expected 1, got %d", config.StartNumber)
	}

	if config.RestartEach != document.FootnoteRestartContinuous {
		t.Errorf("Default restart rule error, expected %s, got %s",
			document.FootnoteRestartContinuous, config.RestartEach)
	}

	if config.Position != document.FootnotePositionPageBottom {
		t.Errorf("Default position error, expected %s, got %s",
			document.FootnotePositionPageBottom, config.Position)
	}
}
