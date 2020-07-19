package helpers

import (
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// CreateAnalysisSheet creates a new Excel sheet
// with counts of each shipment reference id
func CreateAnalysisSheet(file *excelize.File, shipmentCounts map[string]map[string]int) {
	// uniquely name and create new sheet
	newSheet := "Analysis " + time.Now().Local().Format(time.Stamp)
	file.NewSheet(newSheet)

	// track position so we can dynamically write new rows
	xCoordinate := 0
	yCoordinate := 1

	// set row header
	file.SetSheetRow(newSheet, "A1", &[]interface{}{"FILE", "SET", "ZSSL", "CONT", "COST", "IPI", "#"})

	// loop over each file and nested shipments
	for fileNum, items := range shipmentCounts {
		for refNum, count := range items {
			newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
			file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{fileNum, "", "", "", "", refNum, count})
			yCoordinate++
		}

		// Add empty line between different files
		newRowCoordinates := xlsx.GetCellIDStringFromCoordsWithFixed(xCoordinate, yCoordinate, false, false)
		file.SetSheetRow(newSheet, newRowCoordinates, &[]interface{}{"", "", "", "", "", "", ""})
		yCoordinate++
	}

	file.Save()
}
