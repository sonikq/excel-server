package configs

import (
	"github.com/joho/godotenv"
	"github.com/spf13/cast"
	"log"
	"os"
)

type Config struct {
	Port        string
	NSIAddress  string
	NSIEndpoint string
	NSIProfile  string
	DB          PostgresConfig
}

type PostgresConfig struct {
	Host     string
	Port     string
	Username string
	Password string
	DBName   string
	SSLMode  string
}

func Load(envFilename string) (cfg Config) {

	if err := godotenv.Load(envFilename); err != nil {
		log.Fatal(err)
		return
	}

	cfg.Port = cast.ToString(os.Getenv("HTTP_PORT"))
	cfg.NSIEndpoint = cast.ToString(os.Getenv("NSI_ENDPOINT"))
	cfg.NSIProfile = cast.ToString(os.Getenv("NSI_PROFILE"))

	cfg.DB.Host = cast.ToString(os.Getenv("POSTGRES_HOST"))
	cfg.DB.Port = cast.ToString(os.Getenv("POSTGRES_PORT"))
	cfg.DB.Username = cast.ToString(os.Getenv("POSTGRES_USERNAME"))
	cfg.DB.Password = cast.ToString(os.Getenv("POSTGRES_PASSWORD"))
	cfg.DB.DBName = cast.ToString(os.Getenv("POSTGRES_DB"))
	cfg.DB.SSLMode = cast.ToString(os.Getenv("POSTGRES_SSLMODE"))

	value, ok := os.LookupEnv("PG_ENDPOINT")
	if !ok {
		log.Fatal("not found env variable: ")
	}

	cfg.NSIAddress = value

	return cfg
}
