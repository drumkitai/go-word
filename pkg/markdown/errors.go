package markdown

import (
	"errors"
	"fmt"
)

var (
	ErrUnsupportedMarkdown = errors.New("unsupported markdown syntax")
	ErrInvalidImagePath    = errors.New("invalid image path")
	ErrFileNotFound        = errors.New("file not found")
	ErrInvalidMarkdown     = errors.New("invalid markdown content")
	ErrConversionFailed    = errors.New("conversion failed")
	ErrUnsupportedWordElement = errors.New("unsupported word element")
	ErrExportFailed        = errors.New("export failed")
	ErrInvalidDocument     = errors.New("invalid word document")
)

// ConversionError represents a conversion error with detailed information
type ConversionError struct {
	Type    string // error type
	Message string // error message
	Line    int    // error line number (if applicable)
	Column  int    // error column number (if applicable)
	Cause   error  // underlying cause
}

// Error implements the error interface
func (e *ConversionError) Error() string {
	if e.Line > 0 {
		return fmt.Sprintf("%s at line %d, column %d: %s", e.Type, e.Line, e.Column, e.Message)
	}
	return fmt.Sprintf("%s: %s", e.Type, e.Message)
}

// Unwrap returns the underlying cause error, supporting errors.Unwrap
func (e *ConversionError) Unwrap() error {
	return e.Cause
}

// NewConversionError creates a new conversion error
func NewConversionError(errorType, message string, line, column int, cause error) *ConversionError {
	return &ConversionError{
		Type:    errorType,
		Message: message,
		Line:    line,
		Column:  column,
		Cause:   cause,
	}
}

// ExportError represents an export error with detailed information
type ExportError struct {
	Type    string // error type
	Message string // error message
	Cause   error  // underlying cause
}

// Error implements the error interface
func (e *ExportError) Error() string {
	return fmt.Sprintf("%s: %s", e.Type, e.Message)
}

// Unwrap returns the underlying cause error, supporting errors.Unwrap
func (e *ExportError) Unwrap() error {
	return e.Cause
}

// NewExportError creates a new export error
func NewExportError(errorType, message string, cause error) *ExportError {
	return &ExportError{
		Type:    errorType,
		Message: message,
		Cause:   cause,
	}
}
