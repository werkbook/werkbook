package main

import (
	"encoding/json"
	"fmt"
	"strings"

	werkbook "github.com/werkbook/werkbook"
)

// patchOp represents a single operation in a patch array.
type patchOp struct {
	Cell        string           `json:"cell,omitempty"`
	Row         int              `json:"row,omitempty"`
	Sheet       string           `json:"sheet,omitempty"`
	Value       json.RawMessage  `json:"value,omitempty"`
	Formula     *string          `json:"formula,omitempty"`
	Style       json.RawMessage  `json:"style,omitempty"`
	ColumnWidth *float64         `json:"column_width,omitempty"`
	RowHeight   *float64         `json:"row_height,omitempty"`
	AddSheet    string           `json:"add_sheet,omitempty"`
	DeleteSheet string           `json:"delete_sheet,omitempty"`
}

// opResult reports the outcome of a single patch operation.
type opResult struct {
	Index  int    `json:"index"`
	Cell   string `json:"cell,omitempty"`
	Action string `json:"action"`
	Status string `json:"status"`
	Error  string `json:"error,omitempty"`
}

// parsePatchOps parses a JSON array of patch operations.
func parsePatchOps(data []byte) ([]patchOp, error) {
	var ops []patchOp
	if err := json.Unmarshal(data, &ops); err != nil {
		return nil, err
	}
	return ops, nil
}

// applyPatches applies a list of patch operations to the workbook.
// It returns per-operation results and the count of successes.
func applyPatches(f *werkbook.File, ops []patchOp, defaultSheet string) ([]opResult, int) {
	var results []opResult
	applied := 0

	for i, op := range ops {
		res := applyOnePatch(f, op, defaultSheet, i)
		if res.Status == "ok" {
			applied++
		}
		results = append(results, res)
	}

	return results, applied
}

func applyOnePatch(f *werkbook.File, op patchOp, defaultSheet string, index int) opResult {
	// Handle sheet-level operations first.
	if op.AddSheet != "" {
		_, err := f.NewSheet(op.AddSheet)
		if err != nil {
			return opResult{Index: index, Action: "add_sheet", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Action: "add_sheet", Status: "ok"}
	}
	if op.DeleteSheet != "" {
		err := f.DeleteSheet(op.DeleteSheet)
		if err != nil {
			return opResult{Index: index, Action: "delete_sheet", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Action: "delete_sheet", Status: "ok"}
	}

	// Handle row height.
	if op.Row > 0 && op.RowHeight != nil {
		sheetName := op.Sheet
		if sheetName == "" {
			sheetName = defaultSheet
		}
		s := f.Sheet(sheetName)
		if s == nil {
			return opResult{Index: index, Action: "set_row_height", Status: "error", Error: fmt.Sprintf("sheet %q not found", sheetName)}
		}
		err := s.SetRowHeight(op.Row, *op.RowHeight)
		if err != nil {
			return opResult{Index: index, Action: "set_row_height", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Action: "set_row_height", Status: "ok"}
	}

	if op.Cell == "" {
		return opResult{Index: index, Action: "unknown", Status: "error", Error: "missing 'cell' field"}
	}

	sheetName := op.Sheet
	if sheetName == "" {
		sheetName = defaultSheet
	}
	s := f.Sheet(sheetName)
	if s == nil {
		return opResult{Index: index, Cell: op.Cell, Action: "unknown", Status: "error", Error: fmt.Sprintf("sheet %q not found", sheetName)}
	}

	// Column width: cell should be just a column letter like "B".
	if op.ColumnWidth != nil {
		err := s.SetColumnWidth(op.Cell, *op.ColumnWidth)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_column_width", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_column_width", Status: "ok"}
	}

	// Style on a range (contains colon).
	if op.Style != nil && len(op.Style) > 0 && string(op.Style) != "null" {
		style, err := jsonToStyle(op.Style)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_style", Status: "error", Error: err.Error()}
		}
		if strings.Contains(op.Cell, ":") {
			err = s.SetRangeStyle(op.Cell, style)
		} else {
			err = s.SetStyle(op.Cell, style)
		}
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_style", Status: "error", Error: err.Error()}
		}
		// If only style was set (no value or formula), return now.
		if op.Value == nil && op.Formula == nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_style", Status: "ok"}
		}
	}

	// Formula.
	if op.Formula != nil {
		err := s.SetFormula(op.Cell, *op.Formula)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_formula", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_formula", Status: "ok"}
	}

	// Value (including null to clear).
	if op.Value != nil {
		val, err := jsonValueToGo(op.Value)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: err.Error()}
		}
		err = s.SetValue(op.Cell, val)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "ok"}
	}

	return opResult{Index: index, Cell: op.Cell, Action: "noop", Status: "ok"}
}

// jsonValueToGo converts a JSON raw value to a Go value suitable for SetValue.
func jsonValueToGo(raw json.RawMessage) (any, error) {
	if string(raw) == "null" {
		return nil, nil
	}

	// Try number.
	var num float64
	if err := json.Unmarshal(raw, &num); err == nil {
		return num, nil
	}

	// Try bool.
	var b bool
	if err := json.Unmarshal(raw, &b); err == nil {
		return b, nil
	}

	// Try string.
	var s string
	if err := json.Unmarshal(raw, &s); err == nil {
		return s, nil
	}

	return nil, fmt.Errorf("unsupported JSON value: %s", string(raw))
}
