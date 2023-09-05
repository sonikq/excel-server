package db

import (
	"context"
	"fmt"
	"github.com/jmoiron/sqlx"
	_ "github.com/lib/pq"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/configs"
	"time"
)

type ConnectionResponse struct {
	db  *sqlx.DB
	err error
}

func ConnectContext(ctx context.Context, config configs.PostgresConfig) (*sqlx.DB, error) {
	ctx, cancel := context.WithTimeout(ctx, 5000*time.Millisecond)
	defer cancel()

	ch := make(chan ConnectionResponse)

	connInfo := fmt.Sprintf("host=%s port=%s user=%s password=%s dbname=%s sslmode=%s",
		config.Host, config.Port, config.Username, config.Password, config.DBName, config.SSLMode)

	go func() {
		db, err := connect("postgres", connInfo)
		ch <- ConnectionResponse{
			db:  db,
			err: err,
		}
	}()

	for {
		select {
		case resp := <-ch:
			return resp.db, resp.err
		case <-ctx.Done():
			return nil, fmt.Errorf("connecting to database took long time")
		}
	}
}

func connect(driverName string, connInfo string) (_db *sqlx.DB, err error) {
	_db, err = sqlx.Connect(driverName, connInfo)
	return
}
