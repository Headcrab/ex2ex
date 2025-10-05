package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
	"gopkg.in/yaml.v3"
)

type Config struct {
	OutputFilename string        `yaml:"output_filename" json:"output_filename"`
	Mappings       []Mapping     `yaml:"mappings" json:"mappings"`
	OutputSheets   []OutputSheet `yaml:"output_sheets" json:"output_sheets"`
}

type Mapping struct {
	Source      string `yaml:"source" json:"source"`
	Destination string `yaml:"destination" json:"destination"`
}

type OutputSheet struct {
	Name              string `yaml:"name" json:"name"`
	CreateIfNotExists bool   `yaml:"create_if_not_exists" json:"create_if_not_exists"`
}

type Response struct {
	Success     bool   `json:"success"`
	DownloadURL string `json:"download_url,omitempty"`
	Error       string `json:"error,omitempty"`
}

var (
	uploadDir  string
	outputDir  string
	configFile string
	port       string
)

func init() {
	// Load environment variables
	uploadDir = getEnv("UPLOAD_DIR", "./uploads")
	outputDir = getEnv("OUTPUT_DIR", "./output")
	configFile = getEnv("CONFIG_FILE", "./config.yaml")
	port = getEnv("PORT", "8080")

	// Create directories if they don't exist
	os.MkdirAll(uploadDir, 0755)
	os.MkdirAll(outputDir, 0755)
}

func getEnv(key, defaultValue string) string {
	if value := os.Getenv(key); value != "" {
		return value
	}
	return defaultValue
}

func main() {
	// Create logging middleware
	loggedMux := http.NewServeMux()

	loggedMux.HandleFunc("/", indexHandler)
	loggedMux.HandleFunc("/admin", adminHandler)
	loggedMux.HandleFunc("/upload", uploadHandler)
	loggedMux.HandleFunc("/download/", downloadHandler)
	loggedMux.HandleFunc("/api/config", configAPIHandler)

	// Wrap with logging
	handler := loggingMiddleware(loggedMux)

	log.Printf("Server starting on port %s...", port)
	log.Printf("Open http://localhost:%s in your browser", port)
	log.Printf("Admin panel: http://localhost:%s/admin", port)
	log.Printf("Config file: %s", configFile)

	if err := http.ListenAndServe(":"+port, handler); err != nil {
		log.Fatal(err)
	}
}

func loggingMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		log.Printf("%s %s from %s", r.Method, r.URL.Path, r.RemoteAddr)
		next.ServeHTTP(w, r)
	})
}

func indexHandler(w http.ResponseWriter, r *http.Request) {
	http.ServeFile(w, r, "./templates/index.html")
}

func adminHandler(w http.ResponseWriter, r *http.Request) {
	http.ServeFile(w, r, "./templates/admin.html")
}

func configAPIHandler(w http.ResponseWriter, r *http.Request) {
	// Enable CORS
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

	// Handle preflight request
	if r.Method == http.MethodOptions {
		w.WriteHeader(http.StatusOK)
		return
	}

	switch r.Method {
	case http.MethodGet:
		log.Printf("Loading configuration from: %s", configFile)

		// Get current configuration
		config, err := loadConfig(configFile)
		if err != nil {
			log.Printf("Error loading config: %v", err)
			sendError(w, "Failed to load config: "+err.Error(), http.StatusInternalServerError)
			return
		}

		log.Printf("Configuration loaded successfully: %d mappings, %d sheets",
			len(config.Mappings), len(config.OutputSheets))

		w.Header().Set("Content-Type", "application/json")
		if err := json.NewEncoder(w).Encode(config); err != nil {
			log.Printf("Error encoding config: %v", err)
		}

	case http.MethodPost:
		log.Printf("Saving configuration to: %s", configFile)

		// Save configuration
		var config Config
		if err := json.NewDecoder(r.Body).Decode(&config); err != nil {
			log.Printf("Error decoding JSON: %v", err)
			sendError(w, "Invalid JSON: "+err.Error(), http.StatusBadRequest)
			return
		}

		log.Printf("Received config: filename=%s, mappings=%d, sheets=%d",
			config.OutputFilename, len(config.Mappings), len(config.OutputSheets))

		// Validate configuration
		if config.OutputFilename == "" {
			sendError(w, "output_filename is required", http.StatusBadRequest)
			return
		}

		// Convert to YAML and save
		yamlData, err := yaml.Marshal(&config)
		if err != nil {
			log.Printf("Error marshaling YAML: %v", err)
			sendError(w, "Failed to marshal config: "+err.Error(), http.StatusInternalServerError)
			return
		}

		// Add header comment
		yamlWithComments := "# Конфигурация для трансформации Excel файлов\n" + string(yamlData)

		if err := os.WriteFile(configFile, []byte(yamlWithComments), 0644); err != nil {
			log.Printf("Error writing file: %v", err)
			sendError(w, "Failed to save config: "+err.Error(), http.StatusInternalServerError)
			return
		}

		log.Printf("Configuration saved successfully to: %s", configFile)

		response := Response{
			Success: true,
		}
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(response)

	default:
		sendError(w, "Method not allowed", http.StatusMethodNotAllowed)
	}
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		sendError(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	// Parse multipart form
	err := r.ParseMultipartForm(32 << 20) // 32 MB max
	if err != nil {
		sendError(w, "Failed to parse form: "+err.Error(), http.StatusBadRequest)
		return
	}

	file, header, err := r.FormFile("file")
	if err != nil {
		sendError(w, "Failed to get file: "+err.Error(), http.StatusBadRequest)
		return
	}
	defer file.Close()

	// Validate file extension
	ext := filepath.Ext(header.Filename)
	if ext != ".xlsx" && ext != ".xls" {
		sendError(w, "Invalid file type. Only .xlsx and .xls files are allowed", http.StatusBadRequest)
		return
	}

	// Save uploaded file
	timestamp := time.Now().Format("20060102_150405")
	uploadedFilePath := filepath.Join(uploadDir, timestamp+"_"+header.Filename)

	dst, err := os.Create(uploadedFilePath)
	if err != nil {
		sendError(w, "Failed to save file: "+err.Error(), http.StatusInternalServerError)
		return
	}
	defer dst.Close()

	if _, err := io.Copy(dst, file); err != nil {
		sendError(w, "Failed to save file: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Process the Excel file
	outputFilePath, err := processExcel(uploadedFilePath)
	if err != nil {
		sendError(w, "Failed to process Excel file: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Generate download URL
	downloadURL := "/download/" + filepath.Base(outputFilePath)

	// Send success response
	response := Response{
		Success:     true,
		DownloadURL: downloadURL,
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(response)
}

func downloadHandler(w http.ResponseWriter, r *http.Request) {
	filename := filepath.Base(r.URL.Path)
	filePath := filepath.Join(outputDir, filename)

	// Check if file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		http.NotFound(w, r)
		return
	}

	// Set headers for download
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// Serve file
	http.ServeFile(w, r, filePath)
}

func processExcel(inputFilePath string) (string, error) {
	// Load configuration
	config, err := loadConfig(configFile)
	if err != nil {
		return "", fmt.Errorf("failed to load config: %w", err)
	}

	// Open source Excel file
	sourceFile, err := excelize.OpenFile(inputFilePath)
	if err != nil {
		return "", fmt.Errorf("failed to open source file: %w", err)
	}
	defer sourceFile.Close()

	// Create new Excel file for output
	destFile := excelize.NewFile()
	defer destFile.Close()

	// Create output sheets if needed
	for _, sheet := range config.OutputSheets {
		if sheet.CreateIfNotExists {
			index, err := destFile.NewSheet(sheet.Name)
			if err != nil {
				return "", fmt.Errorf("failed to create sheet %s: %w", sheet.Name, err)
			}
			// Set as active sheet if it's the first one
			if index == 1 {
				destFile.SetActiveSheet(index)
			}
		}
	}

	// Delete default Sheet1 if we created custom sheets
	if len(config.OutputSheets) > 0 {
		destFile.DeleteSheet("Sheet1")
	}

	// Apply mappings
	for _, mapping := range config.Mappings {
		if err := applyMapping(sourceFile, destFile, mapping); err != nil {
			log.Printf("Warning: failed to apply mapping %s -> %s: %v",
				mapping.Source, mapping.Destination, err)
			// Continue with other mappings even if one fails
		}
	}

	// Save output file
	timestamp := time.Now().Format("20060102_150405")
	outputFilePath := filepath.Join(outputDir, timestamp+"_"+config.OutputFilename)

	if err := destFile.SaveAs(outputFilePath); err != nil {
		return "", fmt.Errorf("failed to save output file: %w", err)
	}

	return outputFilePath, nil
}

func applyMapping(sourceFile, destFile *excelize.File, mapping Mapping) error {
	// Parse source (sheet!cell or sheet!range)
	sourceSheet, sourceRange := parseReference(mapping.Source)
	destSheet, destCell := parseReference(mapping.Destination)

	// Check if source is a range or single cell
	if isRange(sourceRange) {
		return copyRange(sourceFile, destFile, sourceSheet, sourceRange, destSheet, destCell)
	}
	return copyCellValue(sourceFile, destFile, sourceSheet, sourceRange, destSheet, destCell)
}

// parseFloat attempts to parse a string as a float64
func parseFloat(s string) (float64, error) {
	return strconv.ParseFloat(s, 64)
}

func copyCellValue(sourceFile, destFile *excelize.File, sourceSheet, sourceCell, destSheet, destCell string) error {
	// Get cell type
	cellType, err := sourceFile.GetCellType(sourceSheet, sourceCell)
	if err != nil {
		return fmt.Errorf("failed to get cell type: %w", err)
	}

	// Copy value based on type
	switch cellType {
	case excelize.CellTypeNumber:
		// Get as float to preserve number type
		floatValue, err := sourceFile.GetCellValue(sourceSheet, sourceCell)
		if err != nil {
			return fmt.Errorf("failed to get cell value: %w", err)
		}
		// Try to parse as float
		if numValue, err := parseFloat(floatValue); err == nil {
			if err := destFile.SetCellFloat(destSheet, destCell, numValue, -1, 64); err != nil {
				return fmt.Errorf("failed to set cell float: %w", err)
			}
		} else {
			// Fallback to string if parsing fails
			if err := destFile.SetCellValue(destSheet, destCell, floatValue); err != nil {
				return fmt.Errorf("failed to set cell value: %w", err)
			}
		}

	case excelize.CellTypeBool:
		// Get as bool
		boolValue, err := sourceFile.GetCellValue(sourceSheet, sourceCell)
		if err != nil {
			return fmt.Errorf("failed to get cell value: %w", err)
		}
		if err := destFile.SetCellValue(destSheet, destCell, boolValue == "TRUE" || boolValue == "true" || boolValue == "1"); err != nil {
			return fmt.Errorf("failed to set cell bool: %w", err)
		}

	case excelize.CellTypeFormula:
		// Get formula
		formula, err := sourceFile.GetCellFormula(sourceSheet, sourceCell)
		if err == nil && formula != "" {
			if err := destFile.SetCellFormula(destSheet, destCell, formula); err != nil {
				return fmt.Errorf("failed to set cell formula: %w", err)
			}
		} else {
			// Fallback to calculated value
			value, _ := sourceFile.GetCellValue(sourceSheet, sourceCell)
			if numValue, err := parseFloat(value); err == nil {
				destFile.SetCellFloat(destSheet, destCell, numValue, -1, 64)
			} else {
				destFile.SetCellValue(destSheet, destCell, value)
			}
		}

	default:
		// String or other types
		value, err := sourceFile.GetCellValue(sourceSheet, sourceCell)
		if err != nil {
			return fmt.Errorf("failed to get cell value: %w", err)
		}
		// Try to detect if it's actually a number
		if numValue, err := parseFloat(value); err == nil && value != "" {
			if err := destFile.SetCellFloat(destSheet, destCell, numValue, -1, 64); err != nil {
				return fmt.Errorf("failed to set cell float: %w", err)
			}
		} else {
			if err := destFile.SetCellValue(destSheet, destCell, value); err != nil {
				return fmt.Errorf("failed to set cell value: %w", err)
			}
		}
	}

	// Copy cell style if possible
	styleID, err := sourceFile.GetCellStyle(sourceSheet, sourceCell)
	if err == nil && styleID != 0 {
		destFile.SetCellStyle(destSheet, destCell, destCell, styleID)
	}

	return nil
}

func copyRange(sourceFile, destFile *excelize.File, sourceSheet, sourceRange, destSheet, destCell string) error {
	// Get rows from source range
	rows, err := sourceFile.GetRows(sourceSheet)
	if err != nil {
		return fmt.Errorf("failed to get rows: %w", err)
	}

	// Parse the range
	startCol, startRow, endCol, endRow, err := parseRangeCoords(sourceRange)
	if err != nil {
		return fmt.Errorf("failed to parse range: %w", err)
	}

	// Parse destination cell
	destCol, destRow, err := excelize.CellNameToCoordinates(destCell)
	if err != nil {
		return fmt.Errorf("failed to parse destination cell: %w", err)
	}

	// Copy data
	rowOffset := 0
	for r := startRow; r <= endRow && r <= len(rows); r++ {
		if r > len(rows) {
			break
		}
		row := rows[r-1]

		colOffset := 0
		for c := startCol; c <= endCol && c <= len(row); c++ {
			if c > len(row) {
				break
			}

			sourceCellName, _ := excelize.CoordinatesToCellName(c, r)
			destCellName, _ := excelize.CoordinatesToCellName(destCol+colOffset, destRow+rowOffset)

			// Copy cell with type preservation
			copyCellValue(sourceFile, destFile, sourceSheet, sourceCellName, destSheet, destCellName)

			colOffset++
		}
		rowOffset++
	}

	return nil
}

func parseReference(ref string) (sheet, cellOrRange string) {
	// Split by '!'
	parts := splitReference(ref)
	if len(parts) == 2 {
		return parts[0], parts[1]
	}
	return "Sheet1", ref
}

func splitReference(ref string) []string {
	for i, char := range ref {
		if char == '!' {
			return []string{ref[:i], ref[i+1:]}
		}
	}
	return []string{ref}
}

func isRange(cellRef string) bool {
	for _, char := range cellRef {
		if char == ':' {
			return true
		}
	}
	return false
}

func parseRangeCoords(rangeRef string) (startCol, startRow, endCol, endRow int, err error) {
	// Split range by ':'
	parts := []string{}
	colonIndex := -1
	for i, char := range rangeRef {
		if char == ':' {
			colonIndex = i
			break
		}
	}

	if colonIndex == -1 {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format")
	}

	parts = []string{rangeRef[:colonIndex], rangeRef[colonIndex+1:]}

	if len(parts) != 2 {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format")
	}

	startCol, startRow, err = excelize.CellNameToCoordinates(parts[0])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	endCol, endRow, err = excelize.CellNameToCoordinates(parts[1])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	return startCol, startRow, endCol, endRow, nil
}

func loadConfig(configPath string) (*Config, error) {
	data, err := os.ReadFile(configPath)
	if err != nil {
		return nil, err
	}

	var config Config
	if err := yaml.Unmarshal(data, &config); err != nil {
		return nil, err
	}

	return &config, nil
}

func sendError(w http.ResponseWriter, message string, statusCode int) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(statusCode)

	response := Response{
		Success: false,
		Error:   message,
	}

	json.NewEncoder(w).Encode(response)
}
