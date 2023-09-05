package main

import (
	"context"
	"flag"
	"fmt"
	lg "log"
	"os"
	"os/signal"
	"syscall"
	"time"

	"gitlab.geogracom.com/skdf/skdf-excel-server-go/configs"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/pkg/db"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/pkg/logger"
)

func main() {

	modePtr := flag.String("mode", "debug", "mode defines which env file to use")
	flag.Parse()
	var config configs.Config

	// Configurations
	switch *modePtr {
	case "debug":
		config = configs.Load(".env.debug")
		lg.Printf("running on %s port\n", config.Port)
	case "release":
		config = configs.Load(".env")
	default:
		lg.Fatalf("invalid mode %s, please check mode is valid", *modePtr)

	}

	// Logger
	log := logger.New("info", "excel-server")
	defer func() {
		err := logger.CleanUp(log)
		log.Error("failed to cleanup logs", logger.Error(err))
	}()

	start := time.Now()
	ctx := context.Background()
	_db, err := db.ConnectContext(ctx, config.DB)
	if err != nil {
		log.Fatal("failed to connect to database", logger.Error(err))
	}
	fmt.Println("connection to database took: ", time.Since(start))

	options := excel.Option{
		Conf:   config,
		Logger: log,
		DB:     _db,
	}

	server := NewServer(config, excel.NewRouter(options))

	go func() {
		if err := server.Run(); err != nil {
			log.Fatal("failed to run http server", logger.Error(err))
		}
	}()

	log.Info("Server started...")

	quit := make(chan os.Signal, 1)
	signal.Notify(quit, syscall.SIGTERM, syscall.SIGINT)

	<-quit

	ctx, shutdown := context.WithTimeout(context.Background(), 5*time.Second)
	defer shutdown()

	if err := server.Shutdown(ctx); err != nil {
		log.Error("failed to stop server", logger.Error(err))
	}

}
