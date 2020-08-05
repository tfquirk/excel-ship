package helpers

import (
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
	"time"
)

// GetAllExcelFiles gets all Excel files in the working directory
func GetAllExcelFiles(wd string, start time.Time) []string {
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

	// If no Excel files, exit with failure status
	if len(excelFiles) == 0 {
		elapsed := time.Since(start)
		fmt.Println("")
		fmt.Println("ERROR, Pipe!")
		fmt.Println("")
		fmt.Println("Execution ceased because no Excel files were found.")
		fmt.Printf("Operation exited after %s\n", elapsed)
		os.Exit(1)
	}

	return excelFiles
}
