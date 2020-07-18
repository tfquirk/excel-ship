package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
	"time"

	excelize2 "github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx/v3"
)

func main() {
	start := time.Now()

	// get current working directory
	wd, _ := os.Getwd()

	// track all files in current working directory
	files, _ := ioutil.ReadDir(wd)

	// get XLSX files from list of files
	excelFiles := make([]string, 0)
	for _, f := range files {
		match, _ := regexp.MatchString("xlsx", f.Name())
		if match {
			excelFiles = append(excelFiles, f.Name())
		}
	}

	// read each Excel file, and count number of shipment references
	for _, excelFileName := range excelFiles {
		file, err := excelize2.OpenFile(excelFileName)
		if err != nil {
			fmt.Println(err)
			return
		}

		// Read each company name in the IPI tab, and create map
		IPIRows, err := file.GetRows("IPI")
		IPINames := make(map[string]bool)
		for name := range IPIRows {
			currentRow := xlsx.RowIndexToString(name)
			companyName, _ := file.GetCellValue("IPI", "A"+currentRow)
			IPINames[companyName] = true
		}

		// instantiate a map of fileNumbers to keep track of multiple files
		mapOfFileNumbers := make(map[string]map[string]int)

		// Get all rows in the Complete Summary tab
		CompleteSummaryRows, err := file.GetRows("Complete Summary")
		for id := range CompleteSummaryRows {
			currentRow := xlsx.RowIndexToString(id)
			fileNumber, _ := file.GetCellValue("Complete Summary", "B"+currentRow)
			accountsPayable, _ := file.GetCellValue("Complete Summary", "E"+currentRow)
			company, _ := file.GetCellValue("Complete Summary", "J"+currentRow)

			// only track items if col J = A/P and if company is in the list of IPI companies
			if accountsPayable == "A/P" && IPINames[company] {
				shipmentReferenceCell := "O" + currentRow
				clientID, _ := file.GetCellValue("Complete Summary", shipmentReferenceCell)

				// only track shipment ref numbers that are 12 numbers in length
				if len(clientID) == 12 {

					// if the file number is already tracking the shipment ref num, increase it by one
					if mapOfFileNumbers[fileNumber][clientID] >= 1 {
						mapOfFileNumbers[fileNumber][clientID]++
					} else {
						// instantiate shipment ref number
						if _, err := mapOfFileNumbers[fileNumber]; !err {
							mapOfFileNumbers[fileNumber] = make(map[string]int)
						}

						// count the first shipment ref after instantiation
						mapOfFileNumbers[fileNumber][clientID] = 1
					}
				}
			} else {
				continue
			}

		}

		// uniquely name and create new sheet
		newSheet := "Analysis " + time.Now().Local().Format(time.Stamp)
		file.NewSheet(newSheet)

		// track position so we can dynamically write new rows
		xCoordinate := 0
		yCoordinate := 1
		file.SetSheetRow(newSheet, "A1", &[]interface{}{"FILE", "SET", "ZSSL", "CONT", "COST", "IPI", "#"})
		for fileNum, items := range mapOfFileNumbers {
			for refNum, count := range items {
				newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
				file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{fileNum, "", "", "", "", refNum, count})
				yCoordinate++
			}

			newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
			file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{"", "", "", "", "", "", ""})
			yCoordinate++
		}

		file.Save()
	}
	elapsed := time.Since(start)
	fmt.Printf("Execution completed.\n")
	fmt.Printf("Operation took %s\n", elapsed)
}
