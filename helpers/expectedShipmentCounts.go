package helpers

import (
	"fmt"
	"sort"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// ExpectedShipmentCounts maps shipment ids to their expected counts
func ExpectedShipmentCounts(file *excelize.File) (map[string]map[string]int, map[string]bool, []string) {
	fmt.Print("Counting all tracked shipment reference numbers, and expected invoice counts.\n\n")
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

	// create a map of just the file numbers, which can then
	// updated once shipments are verified
	// also create a sorted version, so that list can be used
	// to print from in the same order on various program executions
	fileNumbers := make(map[string]bool)
	sortedFileNumbers := []string{}
	for key := range expectedShipmentRefCounts {
		fileNumbers[key] = true
		sortedFileNumbers = append(sortedFileNumbers, key)
	}
	sort.Strings(sortedFileNumbers)

	return expectedShipmentRefCounts, fileNumbers, sortedFileNumbers
}
