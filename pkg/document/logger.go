// Package document logging system
package document

import (
	"fmt"
	"io"
	"log"
	"os"
	"time"
)

// LogLevel log level
type LogLevel int

const (
	// LogLevelDebug - debug level
	LogLevelDebug LogLevel = iota
	// LogLevelInfo - info level
	LogLevelInfo
	// LogLevelWarn - warning level
	LogLevelWarn
	// LogLevelError - error level
	LogLevelError
	// LogLevelSilent - silent level
	LogLevelSilent
)

// String returns the string representation of the log level
func (l LogLevel) String() string {
	switch l {
	case LogLevelDebug:
		return "DEBUG"
	case LogLevelInfo:
		return "INFO"
	case LogLevelWarn:
		return "WARN"
	case LogLevelError:
		return "ERROR"
	case LogLevelSilent:
		return "SILENT"
	default:
		return "UNKNOWN"
	}
}

// Logger logger
type Logger struct {
	level  LogLevel    // log level
	output io.Writer   // output destination
	logger *log.Logger // internal logger
}

// defaultLogger default global logger
var defaultLogger = NewLogger(LogLevelInfo, os.Stdout)

// NewLogger creates a new logger
func NewLogger(level LogLevel, output io.Writer) *Logger {
	return &Logger{
		level:  level,
		output: output,
		logger: log.New(output, "", 0),
	}
}

// SetLevel sets the log level
func (l *Logger) SetLevel(level LogLevel) {
	l.level = level
}

// SetOutput sets the output destination
func (l *Logger) SetOutput(output io.Writer) {
	l.output = output
	l.logger.SetOutput(output)
}

// logf formats and outputs a log message
func (l *Logger) logf(level LogLevel, format string, args ...interface{}) {
	if l.level > level {
		return
	}

	timestamp := time.Now().Format("2006-01-02 15:04:05")
	message := fmt.Sprintf(format, args...)
	l.logger.Printf("[%s] %s - %s", timestamp, level.String(), message)
}

// Debugf outputs a debug level log
func (l *Logger) Debugf(format string, args ...interface{}) {
	l.logf(LogLevelDebug, format, args...)
}

// Infof outputs an info level log
func (l *Logger) Infof(format string, args ...interface{}) {
	l.logf(LogLevelInfo, format, args...)
}

// Warnf outputs a warning level log
func (l *Logger) Warnf(format string, args ...interface{}) {
	l.logf(LogLevelWarn, format, args...)
}

// Errorf outputs an error level log
func (l *Logger) Errorf(format string, args ...interface{}) {
	l.logf(LogLevelError, format, args...)
}

// Debug outputs a debug level log message
func (l *Logger) Debug(msg string) {
	l.Debugf("%s", msg)
}

// Info outputs an info level log message
func (l *Logger) Info(msg string) {
	l.Infof("%s", msg)
}

// Warn outputs a warning level log message
func (l *Logger) Warn(msg string) {
	l.Warnf("%s", msg)
}

// Error outputs an error level log message
func (l *Logger) Error(msg string) {
	l.Errorf("%s", msg)
}

// Global logging functions using the default logger

// SetGlobalLevel sets the global log level
func SetGlobalLevel(level LogLevel) {
	defaultLogger.SetLevel(level)
}

// SetGlobalOutput sets the global log output
func SetGlobalOutput(output io.Writer) {
	defaultLogger.SetOutput(output)
}

// Debugf global debug level log
func Debugf(format string, args ...interface{}) {
	defaultLogger.Debugf(format, args...)
}

// Infof global info level log
func Infof(format string, args ...interface{}) {
	defaultLogger.Infof(format, args...)
}

// Warnf global warning level log
func Warnf(format string, args ...interface{}) {
	defaultLogger.Warnf(format, args...)
}

// Errorf global error level log
func Errorf(format string, args ...interface{}) {
	defaultLogger.Errorf(format, args...)
}

// Debug global debug level log message
func Debug(msg string) {
	defaultLogger.Debug(msg)
}

// Info global info level log message
func Info(msg string) {
	defaultLogger.Info(msg)
}

// Warn global warning level log message
func Warn(msg string) {
	defaultLogger.Warn(msg)
}

// Error global error level log message
func Error(msg string) {
	defaultLogger.Error(msg)
}
