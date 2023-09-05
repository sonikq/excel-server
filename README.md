# excel-server.

### Примерные креды
>$USER = imomali.ramazonov

### На локале
```
docker login -u $USER https://registry.geogracom.com/skdf
docker build -t registry.geogracom.com/skdf/skdf-excel-server-go:latest .
docker push registry.geogracom.com/skdf/skdf-excel-server-go:latest
```

### На 45ом хосте

#### С docker compose
```
docker login -u $USER https://registry.geogracom.com/skdf
docker compose down --remove-orphans
docker compose pull
docker compose up -d
```

#### Без docker compose
```
docker login -u $USER https://registry.geogracom.com/skdf
docker pull registry.geogracom.com/skdf/skdf-excel-server-go:latest
docker container rm -f skdf-excel-server-go || true
docker run -d -p 3001:3000 --name skdf-excel-server-go registry.geogracom.com/skdf/skdf-excel-server-go:latest
```