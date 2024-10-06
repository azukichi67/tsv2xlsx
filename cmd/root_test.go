package cmd

import (
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestExecute_正常(t *testing.T) {

	// Arrange
	input := "../test/sample.tsv"
	output := "../test/result.xlsx"
	defer os.Remove(output)

	rootCmd.SetArgs([]string{
		"-i", input,
		"-o", output,
		"-f",
		"-c", "A:20,C:50",
	})

	// Act
	err := rootCmd.Execute()

	// Assert
	assert.NoError(t, err)
}

func TestExecute_colmunWidth_フォーマット不正_区切り文字が不正(t *testing.T) {

	// Arrange
	input := "../test/sample.tsv"
	output := "../test/result.xlsx"
	defer os.Remove(output)

	rootCmd.SetArgs([]string{
		"-i", input,
		"-o", output,
		"-c", "A20",
	})

	// Act
	err := rootCmd.Execute()

	// Assert
	assert.EqualError(t, err, "[A20] is invalid format")
}

func TestExecute_colmunWidth_フォーマット不正_幅が数値でない(t *testing.T) {

	// Arrange
	input := "../test/sample.tsv"
	output := "../test/result.xlsx"
	defer os.Remove(output)

	rootCmd.SetArgs([]string{
		"-i", input,
		"-o", output,
		"-c", "A:test",
	})

	// Act
	err := rootCmd.Execute()

	// Assert
	assert.EqualError(t, err, "[A:test] width is should be number")
}
