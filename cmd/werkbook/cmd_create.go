package main

import (
	"encoding/json"
	"io"
	"os"

	werkbook "github.com/werkbook/werkbook"
)

type createSpec struct {
	Sheets []string  `json:"sheets"`
	Cells  []patchOp `json:"cells"`
}

type createData struct {
	File   string `json:"file"`
	Sheets int    `json:"sheets"`
	Cells  int    `json:"cells"`
}

func cmdCreate(args []string, globals globalFlags) int {
	cmd := "create"
	var specFlag string

	i := 0
	var filePath string
	for i < len(args) {
		switch args[i] {
		case "--spec":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--spec requires a value"), globals)
				return ExitUsage
			}
			specFlag = args[i+1]
			i += 2
		default:
			if filePath == "" && len(args[i]) > 0 && args[i][0] != '-' {
				filePath = args[i]
				i++
			} else {
				writeError(cmd, errUsage("unknown flag: "+args[i]), globals)
				return ExitUsage
			}
		}
	}

	if filePath == "" {
		writeError(cmd, errUsage("file path required"), globals)
		return ExitUsage
	}

	// Read spec from flag or stdin.
	var specBytes []byte
	if specFlag != "" {
		specBytes = []byte(specFlag)
	} else {
		var err error
		specBytes, err = io.ReadAll(os.Stdin)
		if err != nil {
			writeError(cmd, errInternal(err), globals)
			return ExitInternal
		}
	}

	var spec createSpec
	if len(specBytes) > 0 {
		if err := json.Unmarshal(specBytes, &spec); err != nil {
			writeError(cmd, errInvalidSpec(err), globals)
			return ExitValidate
		}
	}

	// Create workbook.
	var f *werkbook.File
	if len(spec.Sheets) > 0 {
		f = werkbook.New(werkbook.FirstSheet(spec.Sheets[0]))
		for _, name := range spec.Sheets[1:] {
			f.NewSheet(name)
		}
	} else {
		f = werkbook.New()
	}

	// Apply cell operations.
	defaultSheet := f.SheetNames()[0]
	cellsApplied := 0
	if len(spec.Cells) > 0 {
		_, cellsApplied = applyPatches(f, spec.Cells, defaultSheet)
	}

	if err := f.SaveAs(filePath); err != nil {
		writeError(cmd, errFileSave(filePath, err), globals)
		return ExitFileIO
	}

	data := createData{
		File:   filePath,
		Sheets: len(f.SheetNames()),
		Cells:  cellsApplied,
	}
	writeSuccess(cmd, data, globals)
	return ExitSuccess
}
