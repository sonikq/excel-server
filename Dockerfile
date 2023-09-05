# BUILD STAGE
FROM golang:latest as builder

WORKDIR /app

COPY go.mod go.sum ./
RUN go mod download && go mod verify

COPY . .

RUN CGO_ENABLED=0 GOOS=linux go build -a -installsuffix cgo -o /app/bin/main .


# RUN STAGE
FROM --platform=linux/x86_64 alpine:latest

ARG SOURCE_COMMIT="${SOURCE_COMMIT:-unknown}"
ENV SOURCE_COMMIT=${SOURCE_COMMIT}

COPY --from=builder /app/bin/main .
COPY --from=builder /app/.env .

EXPOSE 3001
CMD ["/main", "-mode=release"]
