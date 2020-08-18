package helpers

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// CountShipmentReferences tracks all file names,
// and counts the shipment refs in each
func CountShipmentReferences(file *excelize.File, expectedCounts map[string]map[string]int, fileNumbers map[string]bool, IPINames map[string]bool) (map[string]map[string]int, map[string]bool) {
	fmt.Println("Counting all shipment reference numbers.")
	// instantiate a map of fileNumbers to keep track of multiple files
	countsOfShipmentRefIds := make(map[string]map[string]int)
	numerosSinFacturas := fileNumbers

	// Get all rows in the Complete Summary tab
	CompleteSummaryRows, _ := file.GetRows("Complete Summary")
	for id := range CompleteSummaryRows {
		currentRow := xlsx.RowIndexToString(id)
		fileNumber, _ := file.GetCellValue("Complete Summary", "B"+currentRow)
		accountsPayable, _ := file.GetCellValue("Complete Summary", "E"+currentRow)
		company, _ := file.GetCellValue("Complete Summary", "J"+currentRow)

		// only gather info if we have tracked the
		// file number in the previous step
		if _, ok := expectedCounts[fileNumber]; ok {

			// only track items if col J = A/P and if company is in the list of IPI companies
			if accountsPayable == "A/P" {
				if IPINames[company] {

					shipmentReferenceCell := "O" + currentRow
					clientID, _ := file.GetCellValue("Complete Summary", shipmentReferenceCell)

					// only track shipment ref numbers that are 12 numbers in length
					if len(clientID) == 12 {

						// if the file number is already tracking the shipment ref num, increase it by one
						if countsOfShipmentRefIds[fileNumber][clientID] >= 1 {
							countsOfShipmentRefIds[fileNumber][clientID]++
						} else {
							// instantiate shipment ref number
							if _, err := countsOfShipmentRefIds[fileNumber]; !err {
								countsOfShipmentRefIds[fileNumber] = make(map[string]int)
							}
							// count the first shipment ref after instantiation
							countsOfShipmentRefIds[fileNumber][clientID] = 1
						}

						// after successfully logging a file number, delete it from the duplicate map
						// to keep track of any tracked file numbers that are not utilized
						delete(numerosSinFacturas, fileNumber)
					}
				}
			}
		}
	}

	fmt.Printf("COUNTED: %v tracked file numbers.\n", len(expectedCounts))
	fmt.Printf("COUNTED: %v tracked file numbers without invoices.\n\n", len(numerosSinFacturas))

	return countsOfShipmentRefIds, numerosSinFacturas
}
