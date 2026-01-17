package test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/drumkitai/go-word/pkg/markdown"
)

// TestMarkdownMathFormulaConversion
func TestMarkdownMathFormulaConversion(t *testing.T) {
	tests := []struct {
		name        string
		markdown    string
		wantErr     bool
		description string
	}{
		{
			name:        "simple inline formula",
			markdown:    `The energy equation is $E = mc^2$, this is the famous formula of Einstein.`,
			wantErr:     false,
			description: "test simple inline math formula",
		},
		{
			name:        "complex inline formula",
			markdown:    `The quadratic equation root formula is $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$.`,
			wantErr:     false,
			description: "test complex inline math formula",
		},
		{
			name:        "block formula",
			markdown:    "The Pythagorean theorem: \n$$a^2 + b^2 = c^2$$",
			wantErr:     false,
			description: "test block math formula",
		},
		{
			name:        "multiple formulas",
			markdown:    "Let $x$ and $y$ be two variables, then: \n$$x + y = z$$",
			wantErr:     false,
			description: "test multiple math formulas",
		},
		{
			name:        "Greek letters",
			markdown:    `The circumference of a circle is $\pi \approx 3.14159$, the angle is $\theta$ and $\alpha$.`,
			wantErr:     false,
			description: "test Greek letters conversion",
		},
		{
			name:        "subscripts and superscripts",
			markdown:    `The molecular formula of water is $H_2O$, the chemical equation is $x^2 + y^2$.`,
			wantErr:     false,
			description: "test subscripts and superscripts conversion",
		},
		{
			name:        "fraction",
			markdown:    `The fraction $\frac{1}{2}$ represents half.`,
			wantErr:     false,
			description: "test fraction conversion",
		},
		{
			name:        "square root",
			markdown:    `The square root $\sqrt{2}$ and the cube root $\sqrt[3]{8}$.`,
			wantErr:     false,
			description: "test square root conversion",
		},
		{
			name:        "integral and summation",
			markdown:    `The integral $\int_0^1 x dx$ and the summation $\sum_{i=1}^n i$.`,
			wantErr:     false,
			description: "test integral and summation conversion",
		},
		{
			name:        "mathematical operators",
			markdown:    `The mathematical operators: $a \times b$, $a \div b$, $a \pm b$, $a \leq b$, $a \geq b$`,
			wantErr:     false,
			description: "test mathematical operators conversion",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			opts := markdown.DefaultOptions()
			opts.EnableMath = true
			converter := markdown.NewConverter(opts)

			doc, err := converter.ConvertString(tt.markdown, opts)
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertString() error = %v, wantErr %v", err, tt.wantErr)
				return
			}

			if doc == nil && !tt.wantErr {
				t.Error("ConvertString() returned nil document")
				return
			}

			if doc != nil {
				if len(doc.Body.Elements) == 0 {
					t.Error("Expected document to contain at least one element")
				}
			}
		})
	}
}

// TestMarkdownMathDisabled
func TestMarkdownMathDisabled(t *testing.T) {
	markdownContent := `公式 $E = mc^2$ 不应该被特殊处理。`

	opts := markdown.DefaultOptions()
	opts.EnableMath = false
	converter := markdown.NewConverter(opts)

	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		t.Errorf("ConvertString() error = %v", err)
		return
	}

	if doc == nil {
		t.Error("ConvertString() returned nil document")
		return
	}

	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to contain at least one element")
	}
}

// TestMarkdownMathWithOtherElements
func TestMarkdownMathWithOtherElements(t *testing.T) {
	markdownContent := `# Math formula example

## Basic formula

This is a simple formula: $E = mc^2$

## Complex formula

The quadratic equation root formula:
$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$

## Table and formula

| Name | Formula |
|------|------|
| Pythagorean theorem | $a^2 + b^2 = c^2$ |
| Circle area | $S = \pi r^2$ |

## List and formula

- Euler formula: $e^{i\pi} + 1 = 0$
- Newton's law: $F = ma$
`

	opts := markdown.DefaultOptions()
	opts.EnableMath = true
	opts.EnableTables = true
	converter := markdown.NewConverter(opts)

	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		t.Errorf("ConvertString() error = %v", err)
		return
	}

	if doc == nil {
		t.Error("ConvertString() returned nil document")
		return
	}

	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) == 0 {
		t.Error("Expected document to contain paragraphs")
	}

	t.Logf("Document generated with %d paragraphs and %d tables",
		len(paragraphs), len(doc.Body.GetTables()))
}

// TestMarkdownMathSaveToFile
func TestMarkdownMathSaveToFile(t *testing.T) {
	markdownContent := `# Math formula document

## Famous formula

1. Energy equation: $E = mc^2$
2. Pythagorean theorem: $a^2 + b^2 = c^2$
3. 欧拉恒等式：$e^{i\pi} + 1 = 0$

## Quadratic equation root formula

$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$

## Greek letters

- Alpha: $\alpha$
- Beta: $\beta$
- Gamma: $\gamma$
- Delta: $\delta$
- Pi: $\pi$
- Sigma: $\sigma$
- Omega: $\omega$

## Operators

- Multiplication: $a \times b$
- Division: $a \div b$
- Greater than or equal to: $a \geq b$
- Less than or equal to: $a \leq b$
- Not equal to: $a \neq b$
- Approximately equal to: $a \approx b$

## Set symbols

- Belongs to: $x \in A$
- Not belongs to: $x \notin A$
- Subset: $A \subset B$
- Intersection: $A \cap B$
- Union: $A \cup B$

## Calculus

- Integration: $\int_0^1 x dx$
- Summation: $\sum_{i=1}^n i$
- Limit: $\lim_{x \to \infty} f(x)$
`

	opts := markdown.DefaultOptions()
	opts.EnableMath = true
	converter := markdown.NewConverter(opts)

	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		t.Fatalf("ConvertString() error = %v", err)
	}

	outputDir := "test_output"
	err = os.MkdirAll(outputDir, 0755)
	if err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll(outputDir)

	outputPath := filepath.Join(outputDir, "math_formula_test.docx")
	err = doc.Save(outputPath)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Error("Output file was not created")
	}
}

// TestLaTeXToDisplayConversion
func TestLaTeXToDisplayConversion(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{
			name:     "simple superscript",
			input:    "x^2",
			expected: "x²",
		},
		{
			name:     "simple subscript",
			input:    "x_i",
			expected: "xᵢ",
		},
		{
			name:     "Greek letter alpha",
			input:    `\alpha`,
			expected: "α",
		},
		{
			name:     "Greek letter pi",
			input:    `\pi`,
			expected: "π",
		},
		{
			name:     "multiplication symbol",
			input:    `\times`,
			expected: "×",
		},
		{
			name:     "less than or equal to",
			input:    `\leq`,
			expected: "≤",
		},
		{
			name:     "infinity symbol",
			input:    `\infty`,
			expected: "∞",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			md := "$" + tt.input + "$"
			opts := markdown.DefaultOptions()
			opts.EnableMath = true
			converter := markdown.NewConverter(opts)

			doc, err := converter.ConvertString(md, opts)
			if err != nil {
				t.Errorf("ConvertString() error = %v", err)
				return
			}

			if doc == nil {
				t.Error("ConvertString() returned nil document")
				return
			}

			if doc.Body.Elements == nil || len(doc.Body.Elements) == 0 {
				t.Error("Document has no elements")
			}
		})
	}
}

// TestMarkdownBlockMathFormula
func TestMarkdownBlockMathFormula(t *testing.T) {
	markdownContent := `This is a block formula:

$$
\frac{d}{dx} \int_a^x f(t) dt = f(x)
$$

This is the basic theorem of calculus.`

	opts := markdown.DefaultOptions()
	opts.EnableMath = true
	converter := markdown.NewConverter(opts)

	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		t.Errorf("ConvertString() error = %v", err)
		return
	}

	if doc == nil {
		t.Error("ConvertString() returned nil document")
		return
	}

	if len(doc.Body.Elements) < 2 {
		t.Errorf("Expected at least 2 elements, got %d", len(doc.Body.Elements))
	}
}

// TestMarkdownInlineMathInParagraph
func TestMarkdownInlineMathInParagraph(t *testing.T) {
	markdownContent := `In physics, energy $E$ and mass $m$ are related by the speed of light $c$: $E = mc^2$.`

	opts := markdown.DefaultOptions()
	opts.EnableMath = true
	converter := markdown.NewConverter(opts)

	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		t.Errorf("ConvertString() error = %v", err)
		return
	}

	if doc == nil {
		t.Error("ConvertString() returned nil document")
		return
	}

	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) == 0 {
		t.Error("Expected document to contain paragraphs")
	}
}

// TestMathFormulaEdgeCases
func TestMathFormulaEdgeCases(t *testing.T) {
	tests := []struct {
		name        string
		markdown    string
		shouldParse bool
	}{
		{
			name:        "empty formula",
			markdown:    `$$$$`,
			shouldParse: true,
		},
		{
			name:        "formula with only spaces",
			markdown:    `$   $`,
			shouldParse: true,
		},
		{
			name:        "nested curly braces",
			markdown:    `$\frac{\frac{1}{2}}{3}$`,
			shouldParse: true,
		},
		{
			name:        "special characters",
			markdown:    `$a + b = c$`,
			shouldParse: true,
		},
		{
			name:        "unclosed dollar symbol",
			markdown:    `There is an unclosed $ symbol`,
			shouldParse: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			opts := markdown.DefaultOptions()
			opts.EnableMath = true
			converter := markdown.NewConverter(opts)

			doc, err := converter.ConvertString(tt.markdown, opts)
			if tt.shouldParse {
				if err != nil {
					t.Errorf("Expected successful parse, got error: %v", err)
				}
				if doc == nil {
					t.Error("Expected non-nil document")
				}
			}
		})
	}
}

// TestMathDefaultOptionEnabled
func TestMathDefaultOptionEnabled(t *testing.T) {
	opts := markdown.DefaultOptions()
	if !opts.EnableMath {
		t.Error("Expected EnableMath to be true by default")
	}
}

// TestMarkdownMathContentPreservation
func TestMarkdownMathContentPreservation(t *testing.T) {
	formulas := []string{
		`E = mc^2`,
		`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`,
		`\alpha + \beta = \gamma`,
		`a^2 + b^2 = c^2`,
		`\int_0^1 x dx`,
	}

	for _, formula := range formulas {
		t.Run(formula, func(t *testing.T) {
			md := "$" + formula + "$"
			opts := markdown.DefaultOptions()
			opts.EnableMath = true
			converter := markdown.NewConverter(opts)

			doc, err := converter.ConvertString(md, opts)
			if err != nil {
				t.Errorf("Failed to convert formula %q: %v", formula, err)
				return
			}

			if doc == nil {
				t.Errorf("Got nil document for formula %q", formula)
				return
			}

			// 检查文档不为空
			if len(doc.Body.Elements) == 0 {
				t.Errorf("Document has no elements for formula %q", formula)
			}
		})
	}
}

// TestComplexMathDocument
func TestComplexMathDocument(t *testing.T) {
	markdownContent := `# Advanced math formula summary

## 1. Limit

### Important limit
$$\lim_{x \to 0} \frac{\sin x}{x} = 1$$

$$\lim_{x \to \infty} (1 + \frac{1}{x})^x = e$$

## 2. Derivative

基本导数公式：
- $(x^n)' = nx^{n-1}$
- $(e^x)' = e^x$
- $(\sin x)' = \cos x$
- $(\cos x)' = -\sin x$

## 3. Integration

### Indefinite integral
$$\int x^n dx = \frac{x^{n+1}}{n+1} + C \quad (n \neq -1)$$

### Definite integral
$$\int_a^b f(x) dx = F(b) - F(a)$$

## 4. Series

Taylor series:
$$e^x = \sum_{n=0}^{\infty} \frac{x^n}{n!} = 1 + x + \frac{x^2}{2!} + \frac{x^3}{3!} + \cdots$$

## 5. Matrix

Determinant:
$$\det(A) = \sum_{\sigma \in S_n} \text{sgn}(\sigma) \prod_{i=1}^{n} a_{i,\sigma(i)}$$
`

	opts := markdown.DefaultOptions()
	opts.EnableMath = true
	converter := markdown.NewConverter(opts)

	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		t.Fatalf("ConvertString() error = %v", err)
	}

	if doc == nil {
		t.Fatal("ConvertString() returned nil document")
	}

	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) < 3 {
		t.Errorf("Expected at least 3 paragraphs for complex document, got %d", len(paragraphs))
	}

	if len(doc.Body.Elements) == 0 {
		t.Error("Document has no elements")
	}

	t.Logf("Complex math document generated with %d paragraphs", len(paragraphs))
}
