/*
Copyright © 2024 NAME HERE <EMAIL ADDRESS>
*/
package cmd

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "tsv2xlsx",
	Short: "convert tsv to xlsx",
	Run: func(cmd *cobra.Command, args []string) {
		input, _ := cmd.PersistentFlags().GetString("input")
		output, _ := cmd.PersistentFlags().GetString("output")
		shouldSetFilter, _ := cmd.PersistentFlags().GetBool("filter")
		createXlsxFile(input, output, shouldSetFilter)
	},
}

func createXlsxFile(input string, output string, shouldSetFilter bool) {
	tsv, _ := os.Open(input)
	defer tsv.Close()

	book := excelize.NewFile()
	defer book.Close()
	sheetName := "Sheet1"
	book.SetDefaultFont("Meiryo UI")

	row := 0
	headerCount := 0
	scanner := bufio.NewScanner(tsv)
	for scanner.Scan() {
		row++
		line := scanner.Text()
		texts := strings.Split(line, "\t")
		setRow(book, sheetName, row, texts)
		if headerCount == 0 {
			headerCount = len(texts)
		}
	}

	if shouldSetFilter {
		endCell, _ := excelize.CoordinatesToCellName(headerCount, 1)
		book.AutoFilter(sheetName, "A1:"+endCell, []excelize.AutoFilterOptions{})
	}

	err := book.SaveAs(output)
	if err != nil {
		fmt.Println(err)
	}
}

func setRow(book *excelize.File, sheet string, row int, texts []string) {
	startCell, _ := excelize.CoordinatesToCellName(1, row)
	cellTexts := make([]interface{}, len(texts))
	for i := range texts {
		text := texts[i]
		if number, err := strconv.Atoi(text); err == nil {
			cellTexts[i] = number
		} else {
			cellTexts[i] = text
		}
	}
	err := book.SetSheetRow(sheet, startCell, &cellTexts)
	if err != nil {
		fmt.Println(err)
	}
}

func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	rootCmd.PersistentFlags().StringP("input", "i", "", "")
	rootCmd.PersistentFlags().StringP("output", "o", "", "")
	rootCmd.PersistentFlags().BoolP("filter", "f", false, "")

	rootCmd.MarkPersistentFlagRequired("input")
	rootCmd.MarkPersistentFlagRequired("output")
}
