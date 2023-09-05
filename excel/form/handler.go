package form

import (
	"github.com/jmoiron/sqlx"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/configs"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/pkg/logger"
)

type HandlerConfig struct {
	Conf   configs.Config
	Logger logger.Logger
	DB     *sqlx.DB
}

type Handler struct {
	config configs.Config
	log    logger.Logger
	db     *sqlx.DB
}

func New(cfg *HandlerConfig) *Handler {
	return &Handler{
		config: cfg.Conf,
		log:    cfg.Logger,
		db:     cfg.DB,
	}
}
