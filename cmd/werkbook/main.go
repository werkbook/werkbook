package main

import (
	"fmt"
	"os"
)

type globalFlags struct {
	format string
	quiet  bool
}

func main() {
	os.Exit(run(os.Args[1:]))
}

func run(args []string) int {
	globals := globalFlags{format: FormatJSON}

	// Extract global flags from args.
	var remaining []string
	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--format":
			if i+1 >= len(args) {
				writeError("", errUsage("--format requires a value"), globals)
				return ExitUsage
			}
			globals.format = args[i+1]
			i++
		case "--quiet":
			globals.quiet = true
		default:
			remaining = append(remaining, args[i])
		}
	}

	// Validate format.
	switch globals.format {
	case FormatJSON, FormatMarkdown, FormatCSV:
		// ok
	default:
		writeError("", &ErrorInfo{
			Code:    ErrCodeInvalidFormat,
			Message: fmt.Sprintf("unknown format %q", globals.format),
			Hint:    "Supported formats: json, markdown, csv.",
		}, globals)
		return ExitUsage
	}

	if len(remaining) == 0 {
		printUsage()
		return ExitUsage
	}

	command := remaining[0]
	cmdArgs := remaining[1:]

	switch command {
	case "info":
		return cmdInfo(cmdArgs, globals)
	case "read":
		return cmdRead(cmdArgs, globals)
	case "edit":
		return cmdEdit(cmdArgs, globals)
	case "create":
		return cmdCreate(cmdArgs, globals)
	case "calc":
		return cmdCalc(cmdArgs, globals)
	case "formula":
		return cmdFormula(cmdArgs, globals)
	case "version":
		return cmdVersion(globals)
	default:
		writeError("", errUsage(fmt.Sprintf("unknown command %q", command)), globals)
		return ExitUsage
	}
}

func writeSuccess(command string, data any, globals globalFlags) {
	resp := successResponse(command, data)
	out, err := marshalJSON(resp)
	if err != nil {
		fmt.Fprintf(os.Stderr, `{"ok":false,"error":{"code":"INTERNAL","message":%q}}`+"\n", err.Error())
		return
	}
	fmt.Println(string(out))
}

func writeError(command string, ei *ErrorInfo, globals globalFlags) {
	resp := errorResponse(command, ei)
	out, err := marshalJSON(resp)
	if err != nil {
		fmt.Fprintf(os.Stderr, `{"ok":false,"error":{"code":"INTERNAL","message":%q}}`+"\n", err.Error())
		return
	}
	fmt.Fprintln(os.Stderr, string(out))
}

func printUsage() {
	fmt.Fprintln(os.Stderr, `Usage: werkbook <command> [flags] <file>

Commands:
  info      Show sheet metadata (dimensions, cell counts)
  read      Read cell data for a range or full sheet
  edit      Apply JSON patch array of cell changes
  create    Create new workbook from JSON spec
  calc      Force recalculation and return results
  formula   Formula-related subcommands (e.g. 'formula list')
  version   Print version info

Global flags:
  --format <json|markdown|csv>   Output format (default: json)
  --quiet                        Suppress non-essential output`)
}
