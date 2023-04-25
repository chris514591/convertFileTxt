package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	configFile, err := os.Open("config.json")
	if err != nil {
		log.Fatalf("Error opening config file: %s", err)
	}
	defer configFile.Close()

	var config struct {
		InputDir  string `json:"input_dir"`
		OutputDir string `json:"output_dir"`
	}

	err = json.NewDecoder(configFile).Decode(&config)
	if err != nil {
		log.Fatalf("Error decoding config file: %s", err)
	}

	errorLogFile, err := os.OpenFile("errors.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		log.Fatalf("Error opening error log file: %s", err)
	}
	defer errorLogFile.Close()

	err = filepath.Walk(config.InputDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			errorMsg := fmt.Sprintf("Error walking directory: %s\n", err)
			errorLogFile.WriteString(errorMsg)
			return err
		}

		if info.IsDir() {
			return nil
		}

		if filepath.Ext(path) == ".txt" {
			txtFile, err := os.Open(path)
			if err != nil {
				errorMsg := fmt.Sprintf("Error opening %s: %s\n", path, err)
				errorLogFile.WriteString(errorMsg)
				return err
			}
			defer txtFile.Close()

			docxFile, err := os.Create(filepath.Join(config.OutputDir, info.Name()+".docx"))
			if err != nil {
				errorMsg := fmt.Sprintf("Error creating %s: %s\n", docxFile.Name(), err)
				errorLogFile.WriteString(errorMsg)
				return err
			}
			defer docxFile.Close()

			f := excelize.NewFile()
			sheetName := "Sheet1"
			f.SetCellValue(sheetName, "A1", txtFile.Name())

			err = f.SaveAs(docxFile.Name())
			if err != nil {
				errorMsg := fmt.Sprintf("Error converting %s to %s: %s\n", path, docxFile.Name(), err)
				errorLogFile.WriteString(errorMsg)
				return err
			}
		}

		return nil
	})
	if err != nil {
		errorMsg := fmt.Sprintf("Error walking directory: %s\n", err)
		errorLogFile.WriteString(errorMsg)
	}
}
