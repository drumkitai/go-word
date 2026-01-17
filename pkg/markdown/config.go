// Package markdown provides Markdown to Word document conversion functionality
package markdown

import "github.com/drumkitai/go-word/pkg/document"

// ConvertOptions defines markdown to word conversion options
type ConvertOptions struct {
	// Basic configuration
	EnableGFM       bool
	EnableFootnotes bool
	EnableTables    bool
	EnableTaskList  bool
	EnableMath      bool

	// Style configuration
	StyleMapping      map[string]string
	DefaultFontFamily string
	DefaultFontSize   float64

	// Image handling
	ImageBasePath string
	EmbedImages   bool
	MaxImageWidth float64

	// Link handling
	PreserveLinkStyle  bool
	ConvertToBookmarks bool

	// Document settings
	GenerateTOC  bool
	TOCMaxLevel  int
	PageSettings *document.PageSettings

	// Error handling
	StrictMode    bool
	IgnoreErrors  bool
	ErrorCallback func(error)

	// Progress reporting
	ProgressCallback func(int, int)
}

// DefaultOptions returns default conversion options
func DefaultOptions() *ConvertOptions {
	return &ConvertOptions{
		EnableGFM:         true,
		EnableFootnotes:   true,
		EnableTables:      true,
		EnableTaskList:    true,
		EnableMath:        true,
		DefaultFontFamily: "Calibri",
		DefaultFontSize:   11.0,
		EmbedImages:       false,
		MaxImageWidth:     6.0,
		GenerateTOC:       true,
		TOCMaxLevel:       3,
		StrictMode:        false,
		IgnoreErrors:      true,
	}
}

// HighQualityOptions returns high-quality conversion options
func HighQualityOptions() *ConvertOptions {
	opts := DefaultOptions()
	opts.EmbedImages = true
	opts.PreserveLinkStyle = true
	opts.ConvertToBookmarks = true
	opts.StrictMode = true
	opts.IgnoreErrors = false
	return opts
}
