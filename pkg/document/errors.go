// Package document error handling
package document

import (
	"errors"
	"fmt"
)

// Predefined error types
var (
	// ErrInvalidDocument - invalid document
	ErrInvalidDocument = errors.New("invalid document")

	// ErrDocumentNotFound - document not found
	ErrDocumentNotFound = errors.New("document not found")

	// ErrInvalidFormat - invalid format
	ErrInvalidFormat = errors.New("invalid format")

	// ErrCorruptedFile - corrupted file
	ErrCorruptedFile = errors.New("corrupted file")

	// ErrUnsupportedOperation - unsupported operation
	ErrUnsupportedOperation = errors.New("unsupported operation")
)

// DocumentError document operation error
type DocumentError struct {
	Operation string // operation name
	Cause     error  // cause
	Context   string // context information
}

// Error implements the error interface
func (e *DocumentError) Error() string {
	if e.Context != "" {
		return fmt.Sprintf("document operation failed: %s (%s): %v", e.Operation, e.Context, e.Cause)
	}
	return fmt.Sprintf("document operation failed: %s: %v", e.Operation, e.Cause)
}

// Unwrap unwraps the error, supporting errors.Is and errors.As
func (e *DocumentError) Unwrap() error {
	return e.Cause
}

// NewDocumentError creates a new document error
func NewDocumentError(operation string, cause error, context string) *DocumentError {
	return &DocumentError{
		Operation: operation,
		Cause:     cause,
		Context:   context,
	}
}

// WrapError wraps an error with operation context
func WrapError(operation string, err error) error {
	if err == nil {
		return nil
	}
	return NewDocumentError(operation, err, "")
}

// WrapErrorWithContext wraps an error with operation and context information
func WrapErrorWithContext(operation string, err error, context string) error {
	if err == nil {
		return nil
	}
	return NewDocumentError(operation, err, context)
}

// ValidationError validation error
type ValidationError struct {
	Field   string // field name
	Value   string // error value
	Message string // error message
}

// Error implements the error interface
func (e *ValidationError) Error() string {
	return fmt.Sprintf("validation error for field '%s' with value '%s': %s", e.Field, e.Value, e.Message)
}

// NewValidationError creates a new validation error
func NewValidationError(field, value, message string) *ValidationError {
	return &ValidationError{
		Field:   field,
		Value:   value,
		Message: message,
	}
}
