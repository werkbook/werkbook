package formula

func init() {
	Register("COLUMN", fnCOLUMN)
	Register("IFNA", NoCtx(fnIFNA))
	Register("ROW", fnROW)
}

func fnCOLUMN(args []Value, ctx *EvalContext) (Value, error) {
	switch len(args) {
	case 0:
		if ctx == nil {
			return ErrorVal(ErrValVALUE), nil
		}
		return NumberVal(float64(ctx.CurrentCol)), nil
	case 1:
		if args[0].Type == ValueRef {
			col := int(args[0].Num) % 100_000
			return NumberVal(float64(col)), nil
		}
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}
}

func fnROW(args []Value, ctx *EvalContext) (Value, error) {
	switch len(args) {
	case 0:
		if ctx == nil {
			return ErrorVal(ErrValVALUE), nil
		}
		return NumberVal(float64(ctx.CurrentRow)), nil
	case 1:
		if args[0].Type == ValueRef {
			row := int(args[0].Num) / 100_000
			return NumberVal(float64(row)), nil
		}
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}
}

func fnIFNA(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError && args[0].Err == ErrValNA {
		return args[1], nil
	}
	return args[0], nil
}
