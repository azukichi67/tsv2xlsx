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

	"github.com/pkg/errors"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "tsv2xlsx",
	Short: "convert tsv to xlsx",
	RunE: func(cmd *cobra.Command, args []string) error {
		input, _ := cmd.PersistentFlags().GetString("input")
		output, _ := cmd.PersistentFlags().GetString("output")
		columnWidth, _ := cmd.PersistentFlags().GetString("column-width")
		shouldSetFilter, _ := cmd.PersistentFlags().GetBool("filter")

		err := createXlsxFile(input, output, columnWidth, shouldSetFilter)
		if err != nil {
			return err
		}

		return nil
	},
}

func createXlsxFile(input string, output string, columnWidth string, shouldSetFilter bool) error {
	tsv, _ := os.Open(input)
	defer tsv.Close()

	book := excelize.NewFile()
	defer book.Close()
	sheetName := "Sheet1"
	book.SetDefaultFont("Meiryo UI")

	if columnWidth != "" {
		for _, x := range strings.Split(columnWidth, ",") {

			settings := strings.Split(x, ":")
			if len(settings) != 2 {
				return fmt.Errorf("[%s] is invalid format", x)
			}

			column := settings[0]
			width, err := strconv.Atoi(settings[1])
			if err != nil {
				return fmt.Errorf("[%s] width is should be number", x)
			}
			book.SetColWidth(sheetName, column, column, float64(width))
		}
	}

	row := 0
	headerCount := 0
	scanner := bufio.NewScanner(tsv)
	for scanner.Scan() {
		row++
		line := scanner.Text()
		texts := strings.Split(line, "\t")
		err := setRow(book, sheetName, row, texts)
		if err != nil {
			return fmt.Errorf("failed to set text to row. line=%d", row)
		}
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
		return errors.Wrap(err, "failed to save a file")
	}

	return nil
}

func setRow(book *excelize.File, sheet string, row int, texts []string) error {
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
		return err
	}
	return nil
}

func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	rootCmd.PersistentFlags().StringP("input", "i", "", "input tsv file path")
	rootCmd.PersistentFlags().StringP("output", "o", "", "output xlsx file path")
	rootCmd.PersistentFlags().BoolP("filter", "f", false, "set filter to header")
	rootCmd.PersistentFlags().StringP("column-width", "c", "", "change columns width (e.g. A:50,B100)")

	rootCmd.MarkPersistentFlagRequired("input")
	rootCmd.MarkPersistentFlagRequired("output")
}
