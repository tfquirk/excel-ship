package helpers

import (
	"fmt"
	"sort"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// CreateAnalysisSheet creates a new Excel sheet
// with counts of each shipment reference id
func CreateAnalysisSheet(file *excelize.File, shipmentCounts map[string]map[string]int, expectedCounts map[string]map[string]int, numerosSinFacturas map[string]bool, sortedFileNumbers []string) {
	fmt.Println("Counting all shipment reference numbers.")
	// uniquely name and create new sheet
	newSheet := "Analysis " + time.Now().Local().Format(time.Stamp)
	file.NewSheet(newSheet)

	// set row header and styles
	file.SetSheetRow(newSheet, "A1", &[]interface{}{"FILE", "SET", "ZSSL", "CONT", "COST", "IPI", "COUNT", "EXPECTED", "MISSING"})
	titleStyle, _ := file.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Color: "#FFFFFF", Bold: true, Family: "Calibri (Body)"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#4169E1"}, Pattern: 1},
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"},
		Border:    []excelize.Border{{Type: "top", Style: 2, Color: "1f7f3b"}},
	})
	file.SetCellStyle(newSheet, "A1", "I1", titleStyle)

	// define missing styling
	missingStyle, _ := file.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#DB7093"}, Pattern: 1},
	})

	// define missing file numbers styling
	missingFileNumStyles, _ := file.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Family: "Calibri (Body)"},
		Alignment: &excelize.Alignment{Horizontal: "right"},
	})

	// track position so we can dynamically write new rows
	xCoordinate := 0
	yCoordinate := 1

	// loop over each file and nested shipments
	for _, fileNum := range sortedFileNumbers {
		if _, ok := shipmentCounts[fileNum]; ok {

			// for legibility, and consitency sort ship ref numbers before iterating over them
			sortedShipmentReferences := []string{}
			for key := range shipmentCounts[fileNum] {
				sortedShipmentReferences = append(sortedShipmentReferences, key)
			}
			sort.Strings(sortedShipmentReferences)

			for _, refNum := range sortedShipmentReferences {
				expectedCount := expectedCounts[fileNum][refNum]
				count := shipmentCounts[fileNum][refNum]
				missing := expectedCount - count
				newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
				file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{fileNum, "", "", "", "", refNum, count, expectedCount, missing})

				if missing != 0 {
					addErrStyle := xlsx.GetCellIDStringFromCoordsWithFixed(8, yCoordinate, false, false)
					file.SetCellStyle(newSheet, addErrStyle, addErrStyle, missingStyle)
				}

				yCoordinate++
			}
		}

		// Add empty line between different files
		newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
		file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{"", "", "", "", "", "", "", "", ""})
		yCoordinate++
	}

	// track position so we can dynamically write new rows
	sinFacturasXCoordinate := 12
	sinFacturasYCoordinate := 1

	file.SetCellValue(newSheet, "M1", "File numbers with any invoices:")
	file.SetCellStyle(newSheet, "M1", "M1", missingFileNumStyles)

	// loop over each of the file numbers that were not utilized
	for unusedFileNum := range numerosSinFacturas {
		newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(sinFacturasXCoordinate, sinFacturasYCoordinate, false, false)
		file.SetCellValue(newSheet, newRowCoordinates, unusedFileNum)
		file.SetCellStyle(newSheet, newRowCoordinates, newRowCoordinates, missingFileNumStyles)

		sinFacturasYCoordinate++
	}

	// set column widths
	file.SetColWidth(newSheet, "A", "E", 11)
	file.SetColWidth(newSheet, "F", "F", 13)
	file.SetColWidth(newSheet, "G", "I", 12)
	file.SetColWidth(newSheet, "M", "M", 25)

	file.Save()
}
