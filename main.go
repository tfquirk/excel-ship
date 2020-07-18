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

		IPIRows, err := file.GetRows("IPI")
		IPINames := make(map[string]bool)
		for name := range IPIRows {
			currentRow := xlsx.RowIndexToString(name)
			companyName, _ := file.GetCellValue("IPI", "A"+currentRow)
			IPINames[companyName] = true
		}

		shipmentRefCounts := make(map[string]int)
		CompleteSummaryRows, err := file.GetRows("Complete Summary")
		for id := range CompleteSummaryRows {
			currentRow := xlsx.RowIndexToString(id)
			accountsPayable, _ := file.GetCellValue("Complete Summary", "E"+currentRow)
			company, _ := file.GetCellValue("Complete Summary", "J"+currentRow)

			if accountsPayable == "A/P" && IPINames[company] {
				shipmentReferenceCell := "O" + currentRow
				clientID, _ := file.GetCellValue("Complete Summary", shipmentReferenceCell)

				if len(clientID) == 12 {
					if shipmentRefCounts[clientID] >= 1 {
						shipmentRefCounts[clientID]++
					} else {
						shipmentRefCounts[clientID] = 1
					}
				}
			} else {
				continue
			}

		}

		newSheet := "Analysis " + time.Now().Local().Format(time.Stamp)

		file.NewSheet(newSheet)
		xCoordinate := 0
		yCoordinate := 1
		file.SetSheetRow(newSheet, "A1", &[]interface{}{"FILE", "SET", "ZSSL", "CONT", "COST", "IPI", "#"})
		for ref, count := range shipmentRefCounts {
			newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
			file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{"", "", "", "", "", ref, count})
			yCoordinate++
		}

		file.Save()
	}
	elapsed := time.Since(start)
	fmt.Printf("Execution completed. Operation took %s\n", elapsed)
}
