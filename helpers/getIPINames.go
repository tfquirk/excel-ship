package helpers

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// GetIPINames gets all company names from the IPI tab
// return a map of company names to be referenced
func GetIPINames(file *excelize.File) map[string]bool {
	IPIRows, _ := file.GetRows("IPI")
	IPINames := make(map[string]bool)
	for name := range IPIRows {
		currentRow := xlsx.RowIndexToString(name)
		companyName, _ := file.GetCellValue("IPI", "A"+currentRow)
		IPINames[companyName] = true
	}

	return IPINames
}
