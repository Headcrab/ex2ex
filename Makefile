.PHONY: help build run stop clean docker-build docker-up docker-down docker-logs test

help: ## Показать справку
	@echo "Available commands:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

build: ## Собрать приложение
	go build -o ex2ex main.go

run: ## Запустить приложение локально
	go run main.go

test: ## Запустить тесты
	go test -v ./...

clean: ## Очистить временные файлы
	rm -rf uploads/* output/* ex2ex

docker-build: ## Собрать Docker образ
	docker-compose build

docker-up: ## Запустить приложение в Docker
	docker-compose up -d
	@echo "Application is running at http://localhost:8080"

docker-down: ## Остановить Docker контейнеры
	docker-compose down

docker-logs: ## Показать логи Docker контейнера
	docker-compose logs -f

docker-restart: ## Перезапустить Docker контейнеры
	docker-compose restart

docker-rebuild: ## Пересобрать и перезапустить Docker контейнеры
	docker-compose down
	docker-compose build --no-cache
	docker-compose up -d
	@echo "Application is running at http://localhost:8080"

deps: ## Установить зависимости
	go mod download
	go mod tidy

fmt: ## Форматировать код
	go fmt ./...

lint: ## Проверить код линтером
	golangci-lint run

setup: ## Первоначальная настройка проекта
	go mod download
	mkdir -p uploads output templates
	@echo "Setup complete!"

dev: ## Запустить в режиме разработки
	go run main.go

all: clean build ## Полная сборка проекта
