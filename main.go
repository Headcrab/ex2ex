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
	"sync"
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
	Source       string `yaml:"source" json:"source"`
	Destination  string `yaml:"destination" json:"destination"`
	FilterColumn string `yaml:"filter_column,omitempty" json:"filter_column,omitempty"`
	FilterMask   string `yaml:"filter_mask,omitempty" json:"filter_mask,omitempty"`
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

// Validate checks if the configuration is valid
func (c *Config) Validate() error {
	if c.OutputFilename == "" {
		return fmt.Errorf("output_filename is required")
	}

	if len(c.Mappings) == 0 {
		return fmt.Errorf("at least one mapping is required")
	}

	for i, m := range c.Mappings {
		if m.Source == "" {
			return fmt.Errorf("mapping %d: source is required", i)
		}
		if m.Destination == "" {
			return fmt.Errorf("mapping %d: destination is required", i)
		}
	}

	for i, sheet := range c.OutputSheets {
		if sheet.CreateIfNotExists {
			if err := validateSheetName(sheet.Name); err != nil {
				return fmt.Errorf("output_sheet %d: %w", i, err)
			}
		}
	}

	return nil
}

var (
	uploadDir     string
	outputDir     string
	configFile    string
	port          string
	configMutex   sync.RWMutex
	cachedConfig  *Config
	configLastMod time.Time
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

	// Start cleanup goroutine for old files (24 hours retention)
	go startCleanupRoutine(outputDir, 24)
	go startCleanupRoutine(uploadDir, 24)

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

	// Set max upload size limit (100 MB)
	maxUploadSize := int64(100 << 20)
	r.Body = http.MaxBytesReader(w, r.Body, maxUploadSize)

	// Parse multipart form
	err := r.ParseMultipartForm(maxUploadSize)
	if err != nil {
		if err.Error() == "http: request body too large" {
			sendError(w, "File size exceeds maximum limit of 100 MB", http.StatusBadRequest)
		} else {
			sendError(w, "Failed to parse form: "+err.Error(), http.StatusBadRequest)
		}
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

	// Security: prevent path traversal attacks
	if !isPathSafe(filePath, outputDir) {
		http.NotFound(w, r)
		return
	}

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

	// Check if template file exists in templates folder
	templatePath := filepath.Join("./templates", config.OutputFilename)
	var destFile *excelize.File

	if _, err := os.Stat(templatePath); err == nil {
		// Template exists - use it as base
		log.Printf("Using template file: %s", templatePath)
		destFile, err = excelize.OpenFile(templatePath)
		if err != nil {
			return "", fmt.Errorf("failed to open template file: %w", err)
		}
	} else {
		// No template - create new file
		log.Printf("No template found, creating new file")
		destFile = excelize.NewFile()

		// Create output sheets if needed
		for _, sheet := range config.OutputSheets {
			if sheet.CreateIfNotExists {
				// Validate sheet name
				if err := validateSheetName(sheet.Name); err != nil {
					return "", fmt.Errorf("invalid sheet name: %w", err)
				}

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
	}
	defer destFile.Close()

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
		return copyRange(sourceFile, destFile, sourceSheet, sourceRange, destSheet, destCell, mapping.FilterColumn, mapping.FilterMask)
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

func copyRange(sourceFile, destFile *excelize.File, sourceSheet, sourceRange, destSheet, destCell, filterColumn, filterMask string) error {
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

	// Parse filter column if specified (e.g., "B" -> column 2)
	var filterColNum int
	if filterColumn != "" {
		filterColNum, _, err = excelize.CellNameToCoordinates(filterColumn + "1")
		if err != nil {
			return fmt.Errorf("failed to parse filter column: %w", err)
		}
	}

	// Copy data with filtering
	rowOffset := 0
	for r := startRow; r <= endRow && r <= len(rows); r++ {
		if r > len(rows) {
			break
		}
		row := rows[r-1]

		// Apply filter if specified
		if filterColumn != "" && filterMask != "" {
			// Get value from filter column
			if filterColNum > len(row) {
				continue // skip row if filter column doesn't exist
			}
			filterValue := ""
			if filterColNum <= len(row) {
				filterValue = row[filterColNum-1]
			}

			// Check if value matches mask
			if !matchesMask(filterValue, filterMask) {
				continue // skip this row
			}
		}

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

// matchesMask checks if a string matches a pattern with wildcards (*)
// Example: matchesMask("abc123", "*3*") returns true
func matchesMask(value, mask string) bool {
	if mask == "" {
		return true // empty mask matches everything
	}

	// Split mask by wildcards
	parts := []string{}
	current := ""
	for _, char := range mask {
		if char == '*' {
			if current != "" {
				parts = append(parts, current)
				current = ""
			}
		} else {
			current += string(char)
		}
	}
	if current != "" {
		parts = append(parts, current)
	}

	// If no wildcards, do exact match
	if len(parts) == 0 {
		return true
	}
	if len(parts) == 1 && mask[0] != '*' && mask[len(mask)-1] != '*' {
		return value == mask
	}

	// Check if all parts exist in order
	position := 0
	for _, part := range parts {
		index := -1
		for i := position; i <= len(value)-len(part); i++ {
			if value[i:i+len(part)] == part {
				index = i
				break
			}
		}
		if index == -1 {
			return false
		}
		position = index + len(part)
	}

	// Check prefix and suffix
	if len(mask) > 0 && mask[0] != '*' && len(parts) > 0 {
		if len(value) < len(parts[0]) || value[:len(parts[0])] != parts[0] {
			return false
		}
	}
	if len(mask) > 0 && mask[len(mask)-1] != '*' && len(parts) > 0 {
		lastPart := parts[len(parts)-1]
		if len(value) < len(lastPart) || value[len(value)-len(lastPart):] != lastPart {
			return false
		}
	}

	return true
}

func loadConfig(configPath string) (*Config, error) {
	// Check if file was modified
	info, err := os.Stat(configPath)
	if err != nil {
		return nil, err
	}

	configMutex.RLock()
	// If cache exists and file hasn't been modified, return cached version
	if cachedConfig != nil && info.ModTime() == configLastMod {
		defer configMutex.RUnlock()
		return cachedConfig, nil
	}
	configMutex.RUnlock()

	// Load and parse config
	data, err := os.ReadFile(configPath)
	if err != nil {
		return nil, err
	}

	var config Config
	if err := yaml.Unmarshal(data, &config); err != nil {
		return nil, err
	}

	// Validate configuration
	if err := config.Validate(); err != nil {
		return nil, fmt.Errorf("config validation failed: %w", err)
	}

	// Update cache
	configMutex.Lock()
	cachedConfig = &config
	configLastMod = info.ModTime()
	configMutex.Unlock()

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

// isPathSafe checks if the given path is within the baseDir to prevent path traversal attacks
func isPathSafe(filePath, baseDir string) bool {
	absPath, err1 := filepath.Abs(filePath)
	absBase, err2 := filepath.Abs(baseDir)

	if err1 != nil || err2 != nil {
		return false
	}

	// Ensure baseDir ends with separator for proper comparison
	if !filepath.HasPrefix(absPath, filepath.Clean(absBase)+string(os.PathSeparator)) &&
		absPath != filepath.Clean(absBase) {
		return false
	}

	return true
}

// validateSheetName checks if the sheet name is valid for Excel
// Sheet names must be 1-31 characters and cannot contain: [ ] : * ? / \
func validateSheetName(name string) error {
	if name == "" {
		return fmt.Errorf("sheet name cannot be empty")
	}

	if len(name) > 31 {
		return fmt.Errorf("sheet name cannot exceed 31 characters, got %d", len(name))
	}

	invalidChars := []rune{'[', ']', ':', '*', '?', '/', '\\'}
	for _, char := range name {
		for _, invalid := range invalidChars {
			if char == invalid {
				return fmt.Errorf("sheet name contains invalid character: '%c'", char)
			}
		}
	}

	return nil
}

// startCleanupRoutine starts a background goroutine that deletes files older than maxAgeHours
func startCleanupRoutine(dir string, maxAgeHours int) {
	ticker := time.NewTicker(1 * time.Hour)
	defer ticker.Stop()

	for range ticker.C {
		cleanupOldFiles(dir, maxAgeHours)
	}
}

// cleanupOldFiles removes files older than maxAgeHours from the directory
func cleanupOldFiles(dir string, maxAgeHours int) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		log.Printf("Error reading directory for cleanup: %v", err)
		return
	}

	now := time.Now()
	maxAge := time.Duration(maxAgeHours) * time.Hour
	deletedCount := 0

	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		info, err := entry.Info()
		if err != nil {
			log.Printf("Error getting file info: %v", err)
			continue
		}

		age := now.Sub(info.ModTime())
		if age > maxAge {
			filePath := filepath.Join(dir, entry.Name())
			if err := os.Remove(filePath); err != nil {
				log.Printf("Error deleting old file %s: %v", filePath, err)
			} else {
				log.Printf("Cleaned up old file: %s (age: %v)", filePath, age)
				deletedCount++
			}
		}
	}

	if deletedCount > 0 {
		log.Printf("Cleanup complete: deleted %d old files from %s", deletedCount, dir)
	}
}
