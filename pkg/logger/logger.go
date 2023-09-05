package logger

import (
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
	"time"
)

type Field = zapcore.Field

var (
	Int    = zap.Int
	String = zap.String
	Error  = zap.Error
	Bool   = zap.Bool
	Any    = zap.Any
)

type Logger interface {
	Debug(msg string, fields ...Field)
	Info(msg string, fields ...Field)
	Warn(msg string, fields ...Field)
	Error(msg string, fields ...Field)
	Fatal(msg string, fields ...Field)
}

type loggerImplementation struct {
	zap *zap.Logger
}

var customTimeFormat string

func New(level string, namespace string) Logger {
	if level == "" {
		level = LevelInfo
	}

	logger := loggerImplementation{
		zap: newZapLogger(level, time.RFC850),
	}

	logger.zap = logger.zap.Named(namespace)

	zap.RedirectStdLog(logger.zap)

	return &logger
}

func (l *loggerImplementation) Debug(msg string, fields ...Field) {
	l.zap.Debug(msg, fields...)
}

func (l *loggerImplementation) Info(msg string, fields ...Field) {
	l.zap.Info(msg, fields...)
}

func (l *loggerImplementation) Warn(msg string, fields ...Field) {
	l.zap.Warn(msg, fields...)
}

func (l *loggerImplementation) Error(msg string, fields ...Field) {
	l.zap.Error(msg, fields...)
}

func (l *loggerImplementation) Fatal(msg string, fields ...Field) {
	l.zap.Fatal(msg, fields...)
}

func GetNamed(l Logger, name string) Logger {
	switch v := l.(type) {
	case *loggerImplementation:
		v.zap = v.zap.Named(name)
		return v
	default:
		l.Info("logger.GetNamed: invalid logger type")
		return l
	}
}

func WithFields(l Logger, fields ...Field) Logger {
	switch v := l.(type) {
	case *loggerImplementation:
		return &loggerImplementation{
			zap: v.zap.With(fields...),
		}
	default:
		l.Info("logger.WithFields: invalid logger type")
		return l
	}
}

func CleanUp(l Logger) error {
	switch v := l.(type) {
	case *loggerImplementation:
		return v.zap.Sync()
	default:
		l.Info("logger.Cleanup: invalid logger type")
		return nil
	}
}
