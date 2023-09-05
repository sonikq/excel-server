run:
	go run .

login:
	docker login -u imomali.ramazonov https://registry.geogracom.com/skdf

push:
	docker build --build-arg=SOURCE_COMMIT=$(git rev-parse --short HEAD) -t registry.geogracom.com/skdf/skdf-excel-server-go:0.1 . && \
	docker push registry.geogracom.com/skdf/skdf-excel-server-go:0.1