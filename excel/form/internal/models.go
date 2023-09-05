package internal

// NSI Нормативная справочная информация
type NSI struct {
	Address string
	Profile string
}

type SheetOptions struct {
	Protected bool
	HasErrors bool
}

type GetFormRequest struct {
	Title *TitlePage `json:"TITLE_PAGE,omitempty"`
	DOCS  []DOC      `json:"DOCS,omitempty"`
}

// TitlePage титульный лист документа
type TitlePage struct {
	OKUD         string     `json:"OKUD,omitempty"`
	OKPO         string     `json:"OKPO,omitempty"`
	Title        string     `json:"TITLE,omitempty"`
	StateByDate  string     `json:"STATE_BY_DATE,omitempty"`
	Responses    []Response `json:"RESPONSES,omitempty"`
	FormName     string     `json:"FORM_NAME,omitempty"`
	OrderDate    string     `json:"ORDER_DATE,omitempty"`
	OrderNum     string     `json:"ORDER_NUM,omitempty"`
	ReportType   string     `json:"REPORT_TYPE,omitempty"`
	Organization string     `json:"ORGANIZATION,omitempty"`
}

type Response struct {
	ORGResponder string   `json:"ORG_RESPONDER,omitempty"`
	ORGReviewers []string `json:"ORG_REVIEWERS,omitempty"`
	Deadline     string   `json:"DEADLINE,omitempty"`
}

// DOC раздел документа
type DOC struct {
	ERROR   bool   `json:"ERROR,omitempty"`
	ID      int    `json:"ID,omitempty"`
	TITLE   string `json:"TITLE,omitempty"`
	OKUD    string `json:"OKUD,omitempty"`
	CODE    string `json:"CODE,omitempty"`
	LABEL   string `json:"LABEL,omitempty"`
	CUSTOM  bool   `json:"CUSTOM,omitempty"`
	CHAPTER struct {
		TEXT  string `json:"TEXT,omitempty"`
		STYLE STYLE  `json:"STYLE,omitempty"`
	} `json:"CHAPTER,omitempty"`
	PERIOD         string `json:"PERIOD,omitempty"`
	ORGRESPONDERID int    `json:"ORG_RESPONDER_ID,omitempty"`
	ORGRESPONDER   string `json:"ORG_RESPONDER,omitempty"`
	ORGREVIEWERID  int    `json:"ORG_REVIEWER_ID,omitempty"`
	EXECUTORID     int    `json:"EXECUTOR_ID,omitempty"`
	EXECUTORFIO    string `json:"EXECUTOR_FIO,omitempty"`
	REGION         string `json:"REGION,omitempty"`
	REGIONOKATO    string `json:"REGION_OKATO,omitempty"`
	REGIONID       int    `json:"REGION_ID,omitempty"`
	PREHEADER      struct {
		TEXT  string `json:"TEXT,omitempty"`
		STYLE STYLE  `json:"STYLE,omitempty"`
	} `json:"PRE_HEADER,omitempty"`
	HEADER    HEADER                   `json:"HEADER,omitempty"`
	COMMON    COMMON                   `json:"COMMON,omitempty"`
	COLUMNS   []COLUMN                 `json:"COLUMNS"`
	ROWUNIONS []ROWUNION               `json:"ROW_UNIONS,omitempty"`
	ROWBREAKS []ROWBREAK               `json:"ROW_BREAKS,omitempty"`
	READONLY  [][]int                  `json:"READONLY"`
	STYLES    []ExtraStyle             `json:"STYLES,omitempty"`
	DATA      []map[string]interface{} `json:"DATA,omitempty"`
	FOOTER    FOOTER                   `json:"FOOTER,omitempty"`
}

type ExtraStyle struct {
	GRID       []int `json:"GRID,omitempty"`
	CELL_STYLE STYLE `json:"CELL_STYLE,omitempty"`
	PROTECTED  bool  `json:"PROTECTED,omitempty"`
}

type HEADER struct {
	HEADERCOLUMNS int   `json:"HEADER_COLUMNS,omitempty"`
	HEADERROWS    int   `json:"HEADER_ROWS,omitempty"`
	STYLE         STYLE `json:"STYLE,omitempty"`
	HEADERS       []struct {
		HEADERID      int    `json:"HEADER_ID,omitempty"`
		HEADERCAPTION string `json:"HEADER_CAPTION,omitempty"`
		HEADERGRID    []int  `json:"HEADER_GRID,omitempty"`
	} `json:"HEADERS,omitempty"`
}

type COMMON struct {
	ROWUNITWIDTH int    `json:"ROW_UNIT_WIDTH,omitempty"`
	STYLE        STYLE  `json:"STYLE,omitempty"`
	OKEIID       int    `json:"OKEI_ID,omitempty"`
	OKEINAME     string `json:"OKEI_NAME,omitempty"`
	STATICROWS   bool   `json:"STATIC_ROWS,omitempty"`
}

type COLUMN struct {
	COLUMNID      int       `json:"COLUMN_ID,omitempty"`
	COLUMNNUM     string    `json:"COLUMN_NUM,omitempty"`
	COLUMNWIDTHPX int       `json:"COLUMN_WIDTH_PX,omitempty"`
	COLUMNWIDTHMM int       `json:"COLUMN_WIDTH_MM,omitempty"`
	STYLE         STYLE     `json:"STYLE"`
	VALUEFIELD    string    `json:"VALUE_FIELD,omitempty"`
	DATATYPE      string    `json:"DATA_TYPE,omitempty"`
	ISEDITABLE    *bool     `json:"IS_EDITABLE,omitempty"`
	VALUELABEL    string    `json:"VALUE_LABEL,omitempty"`
	ALLOWMULTI    bool      `json:"ALLOW_MULTI,omitempty"`
	SEPARATOR     string    `json:"SEPARATOR,omitempty"`
	SELECTOR      *SELECTOR `json:"SELECTOR,omitempty"`
}

type SELECTOR struct {
	DATASOURCE string `json:"DATA_SOURCE,omitempty"`
	VALUEFIELD string `json:"VALUE_FIELD,omitempty"`
	VALUELABEL string `json:"VALUE_LABEL,omitempty"`
}

type ROWUNION struct {
	GRID []int `json:"GRID,omitempty"`
}

type ROWBREAK struct {
	PREVROW      int    `json:"PREV_ROW"`
	BREAKCAPTION string `json:"BREAK_CAPTION"`
	INDENT       int    `json:"INDENT"`
	STYLE        STYLE  `json:"STYLE"`
	GRID         []int  `json:"GRID,omitempty"`
}

type FOOTER struct {
	CAPTION string `json:"CAPTION,omitempty"`
	DATA    []struct {
		FOOTERID  int    `json:"FOOTER_ID,omitempty"`
		ROWNUM    int    `json:"ROW_NUM,omitempty"`
		COLUMNNUM string `json:"COLUMN_NUM,omitempty"`
		INDENT    int    `json:"INDENT,omitempty"`
		VALUE     string `json:"VALUE,omitempty"`
	} `json:"DATA,omitempty"`
}

type STYLE struct {
	FONT       string  `json:"FONT,omitempty"`
	SIZE       float64 `json:"SIZE,omitempty"`
	FONT_STYLE struct {
		BOLD       bool `json:"BOLD,omitempty"`
		ITALIC     bool `json:"ITALIC,omitempty"`
		UNDERLINED bool `json:"UNDERLINED,omitempty"`
	} `json:"FONT_STYLE,omitempty"`
	INDENT        int    `json:"INDENT,omitempty"`
	HORALIGNMENT  string `json:"HOR_ALIGNMENT,omitempty"`
	VERTALIGNMENT string `json:"VERT_ALIGNMENT,omitempty"`
}
