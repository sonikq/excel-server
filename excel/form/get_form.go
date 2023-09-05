package form

import (
	"encoding/json"
	"github.com/gin-gonic/gin"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel/form/internal"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel/form/utils"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/pkg/logger"
	"net/http"
)

const (
	contentType = "application/octet-stream"
)

func (h *Handler) GetForm(c *gin.Context) {
	body, err := utils.GetBody(c.Request.Body)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "internal server error!"})
		h.log.Error("Error in reading body: ", logger.Error(err))
		return
	}

	var form internal.GetFormRequest

	if err = json.Unmarshal(body, &form); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "invalid json! error: " + err.Error()})
		h.log.Error("Error in parsing body: ", logger.Error(err))
		return //filename := output + ext
	}

	nsi := internal.NSI{
		Address: h.config.NSIAddress + h.config.NSIEndpoint,
		Profile: h.config.NSIProfile,
	}

	data, err := internal.PrintForm(form, body, nsi, h.db)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "internal server error! error: " + err.Error()})
		h.log.Error("Error in printing form: ", logger.Error(err))
		return
	}

	c.Data(http.StatusOK, contentType, data)
}
