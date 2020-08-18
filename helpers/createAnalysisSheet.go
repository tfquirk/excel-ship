package helpers

import (
	"fmt"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// CreateAnalysisSheet creates a new Excel sheet
// with counts of each shipment reference id
func CreateAnalysisSheet(file *excelize.File, shipmentCounts map[string]map[string]int, expectedShipmentCounts map[string]map[string]int) {
	fmt.Print("Creating new tab with updated shipment counts.\n\n")
	// uniquely name and create new sheet
	newSheet := "Analysis " + time.Now().Local().Format(time.Stamp)
	file.NewSheet(newSheet)

	// track position so we can dynamically write new rows
	xCoordinate := 0
	yCoordinate := 1

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

	// loop over each file and nested shipments
	for fileNum, items := range shipmentCounts {
		for refNum, count := range items {
			expectedCount := expectedShipmentCounts[fileNum][refNum]
			missing := expectedCount - count
			newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
			file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{fileNum, "", "", "", "", refNum, count, expectedCount, missing})

			if missing != 0 {
				addErrStyle := xlsx.GetCellIDStringFromCoordsWithFixed(8, yCoordinate, false, false)
				file.SetCellStyle(newSheet, addErrStyle, addErrStyle, missingStyle)
			}

			yCoordinate++
		}

		// Add empty line between different files
		newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
		file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{"", "", "", "", "", "", "", "", ""})
		yCoordinate++
	}

	// set column widths
	file.SetColWidth(newSheet, "A", "E", 11)
	file.SetColWidth(newSheet, "F", "F", 12.33)
	file.SetColWidth(newSheet, "G", "I", 11.33)

	file.Save()
}
