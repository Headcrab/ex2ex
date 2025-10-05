# Build stage
FROM golang:1.24-alpine AS builder

WORKDIR /app

# Copy go mod files
COPY go.mod go.sum ./

# Download dependencies
RUN go mod download

# Copy source code
COPY . .

# Build the application
RUN CGO_ENABLED=0 GOOS=linux go build -a -installsuffix cgo -o ex2ex .

# Runtime stage
FROM alpine:latest

RUN apk --no-cache add ca-certificates tzdata

WORKDIR /app

# Copy the binary from builder
COPY --from=builder /app/ex2ex .

# Copy templates and config
COPY templates ./templates
COPY config.yaml .

# Create directories for uploads and output
RUN mkdir -p /app/uploads /app/output

# Expose port
EXPOSE 8080

# Set environment variables
ENV PORT=8080
ENV UPLOAD_DIR=/app/uploads
ENV OUTPUT_DIR=/app/output
ENV CONFIG_FILE=/app/config.yaml

# Run the application
CMD ["./ex2ex"]
