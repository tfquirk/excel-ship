package helpers

import (
	"fmt"
	"os"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

// GetIPINames gets all company names from the IPI tab
// return a map of company names to be referenced
func GetIPINames(file *excelize.File, start time.Time) map[string]bool {
	IPIRows, _ := file.GetRows("IPI")
	IPINames := make(map[string]bool)
	for name := range IPIRows {
		currentRow := xlsx.RowIndexToString(name)
		companyName, _ := file.GetCellValue("IPI", "A"+currentRow)
		IPINames[companyName] = true
	}

	if len(IPINames) == 0 {
		elapsed := time.Since(start)
		fmt.Println("")
		fmt.Println("ERROR, Pipe!")
		fmt.Println("")
		fmt.Println("Execution ceased because your file does not have an IPI tab.")
		fmt.Printf("Operation exited after %s\n", elapsed)
		os.Exit(1)
	}

	return IPINames
}
