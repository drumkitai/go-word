package markdown

import (
	"os"
	"path/filepath"
	"strings"

	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/extension"
	"github.com/yuin/goldmark/parser"
	"github.com/yuin/goldmark/text"

	mathjax "github.com/litao91/goldmark-mathjax"

	"github.com/drumkitai/go-word/pkg/document"
)

// MarkdownConverter defines the interface for markdown conversion
type MarkdownConverter interface {
	ConvertFile(mdPath, docxPath string, options *ConvertOptions) error
	ConvertBytes(mdContent []byte, options *ConvertOptions) (*document.Document, error)
	ConvertString(mdContent string, options *ConvertOptions) (*document.Document, error)
	BatchConvert(inputs []string, outputDir string, options *ConvertOptions) error
}

// Converter is the default markdown to word converter implementation
type Converter struct {
	md   goldmark.Markdown
	opts *ConvertOptions
}

// NewConverter creates a new converter instance with specified options
func NewConverter(opts *ConvertOptions) *Converter {
	if opts == nil {
		opts = DefaultOptions()
	}

	extensions := []goldmark.Extender{}
	if opts.EnableGFM {
		extensions = append(extensions, extension.GFM)
	}
	if opts.EnableFootnotes {
		extensions = append(extensions, extension.Footnote)
	}
	if opts.EnableMath {
		// Use standard LaTeX math delimiters: $...$ for inline, $$...$$ for block
		extensions = append(extensions, mathjax.NewMathJax(
			mathjax.WithInlineDelim("$", "$"),
			mathjax.WithBlockDelim("$$", "$$"),
		))
	}

	md := goldmark.New(
		goldmark.WithExtensions(extensions...),
		goldmark.WithParserOptions(
			parser.WithAutoHeadingID(),
		),
	)

	return &Converter{md: md, opts: opts}
}

// ConvertString converts string markdown content to a Word document
func (c *Converter) ConvertString(content string, opts *ConvertOptions) (*document.Document, error) {
	return c.ConvertBytes([]byte(content), opts)
}

// ConvertBytes converts byte array markdown content to a Word document
func (c *Converter) ConvertBytes(content []byte, opts *ConvertOptions) (*document.Document, error) {
	if opts != nil {
		c.opts = opts
	}

	doc := document.New()

	if c.opts.PageSettings != nil {
		// Page settings API can be extended here
	}

	reader := text.NewReader(content)
	astDoc := c.md.Parser().Parse(reader)

	renderer := &WordRenderer{
		doc:    doc,
		opts:   c.opts,
		source: content,
	}

	err := renderer.Render(astDoc)
	if err != nil {
		return nil, err
	}

	return doc, nil
}

// ConvertFile converts a markdown file to a Word document
func (c *Converter) ConvertFile(mdPath, docxPath string, options *ConvertOptions) error {
	content, err := os.ReadFile(mdPath)
	if err != nil {
		return NewConversionError("FileRead", "failed to read markdown file", 0, 0, err)
	}

	if options == nil {
		options = c.opts
	}
	if options.ImageBasePath == "" {
		options.ImageBasePath = filepath.Dir(mdPath)
	}

	doc, err := c.ConvertBytes(content, options)
	if err != nil {
		return err
	}

	err = doc.Save(docxPath)
	if err != nil {
		return NewConversionError("FileSave", "failed to save word document", 0, 0, err)
	}

	return nil
}

// BatchConvert converts multiple markdown files to Word documents
func (c *Converter) BatchConvert(inputs []string, outputDir string, options *ConvertOptions) error {
	err := os.MkdirAll(outputDir, 0755)
	if err != nil {
		return NewConversionError("DirectoryCreate", "failed to create output directory", 0, 0, err)
	}

	total := len(inputs)
	for i, input := range inputs {
		if options != nil && options.ProgressCallback != nil {
			options.ProgressCallback(i+1, total)
		}

		base := strings.TrimSuffix(filepath.Base(input), filepath.Ext(input))
		output := filepath.Join(outputDir, base+".docx")

		err := c.ConvertFile(input, output, options)
		if err != nil {
			if options != nil && options.ErrorCallback != nil {
				options.ErrorCallback(err)
			}
			if options == nil || !options.IgnoreErrors {
				return err
			}
		}
	}

	return nil
}
