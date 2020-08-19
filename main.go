package main

import (
	"fmt"
	"os"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tfquirk/excel-ship/helpers"
)

func main() {
	start := time.Now()

	// get current working directory
	wd, _ := os.Getwd()

	excelFiles := helpers.GetAllExcelFiles(wd, start)

	// read each Excel file
	// Review 'AN' tab and gather expected shipment counts, file numbers, and a sorted version of the file numbers
	// count number of shipment references and validate against IPI names
	// save counts as new Excel sheet
	for _, excelFileName := range excelFiles {
		file, err := excelize.OpenFile(excelFileName)
		if err != nil {
			fmt.Println(err)
			return
		}

		IPINames := helpers.GetIPINames(file, start)
		expectedShipmentCounts, fileNumbers, sortedFileNumbers := helpers.ExpectedShipmentCounts(file)
		countOfShipmentReferences, numerosSinFacturas := helpers.CountShipmentReferences(file, expectedShipmentCounts, fileNumbers, IPINames)
		helpers.CreateAnalysisSheet(file, countOfShipmentReferences, expectedShipmentCounts, numerosSinFacturas, sortedFileNumbers)
	}

	// Log performance to command line for general interest purposes
	elapsed := time.Since(start)
	fmt.Println("Execution completed.")
	fmt.Printf("Operation took %s\n\n", elapsed)
}
