package helpers

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// ExpectedShipmentCounts maps shipment ids to their expected counts
func ExpectedShipmentCounts(file *excelize.File) map[string]map[string]int {
	fmt.Println("Counting all tracked shipment reference numbers, and expected invoice counts.")
	// instantiate a map of fileNumbers to keep track of multiple files
	expectedShipmentRefCounts := make(map[string]map[string]int)

	// Get all rows in the AN tab
	ANRows, _ := file.GetRows("AN")
	for id := range ANRows {
		currentRow := xlsx.RowIndexToString(id)
		fileNumber, _ := file.GetCellValue("AN", "AB"+currentRow)
		accountNumber, _ := file.GetCellValue("AN", "F"+currentRow)
		expectedShipRefCount, _ := file.GetCellValue("AN", "AG"+currentRow)

		// only track shipment ref numbers that are 12 numbers in length
		if len(accountNumber) == 12 {

			// instantiate shipment ref number
			if _, err := expectedShipmentRefCounts[fileNumber]; !err {
				expectedShipmentRefCounts[fileNumber] = make(map[string]int)
			}

			expectedCountToInt, _ := strconv.Atoi(expectedShipRefCount)
			// assign the expected shipment count after instantiation
			expectedShipmentRefCounts[fileNumber][accountNumber] = expectedCountToInt
		}
	}

	fmt.Printf("COUNTED: %v tracked file numbers.\n\n", len(expectedShipmentRefCounts))

	return expectedShipmentRefCounts
}
