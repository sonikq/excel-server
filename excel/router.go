package excel

import (
	"github.com/jmoiron/sqlx"
	"net/http"

	"github.com/gin-gonic/gin"

	"gitlab.geogracom.com/skdf/skdf-excel-server-go/configs"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel/form"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/pkg/logger"
)

type Handler struct {
	CommonHandler form.Handler
}

type Option struct {
	Conf   configs.Config
	Logger logger.Logger
	DB     *sqlx.DB
}

func NewRouter(option Option) *gin.Engine {
	gin.SetMode(gin.TestMode)

	router := gin.New()

	router.Use(gin.Logger())
	router.Use(gin.Recovery())

	router.MaxMultipartMemory = 8 << 20

	h := &Handler{
		CommonHandler: *form.New(&form.HandlerConfig{
			Conf:   option.Conf,
			Logger: option.Logger,
			DB:     option.DB,
		}),
	}

	router.GET("/ping", func(ctx *gin.Context) {
		ctx.JSON(http.StatusOK, gin.H{
			"message": "Pong!",
		})
	})

	router.POST("/common/getform", h.CommonHandler.GetForm)

	return router
}
