package formula

import (
	"testing"
)

func TestIFNA(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFNA(#N/A,"default")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "default" {
		t.Errorf("IFNA(#N/A) = %v, want default", got)
	}

	cf = evalCompile(t, `IFNA(42,"default")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("IFNA(42) = %v, want 42", got)
	}
}

func TestROW(t *testing.T) {
	resolver := &mockResolver{}

	// ROW(A1) should return 1
	cf := evalCompile(t, "ROW(A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(ROW(A1)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("ROW(A1) = %v, want 1", got)
	}

	// ROW(B3) should return 3
	cf = evalCompile(t, "ROW(B3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(ROW(B3)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("ROW(B3) = %v, want 3", got)
	}

	// ROW() with context should return current row
	ctx := &EvalContext{CurrentRow: 5, CurrentCol: 3, CurrentSheet: "Sheet1"}
	cf = evalCompile(t, "ROW()")
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval(ROW()): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("ROW() = %v, want 5", got)
	}
}

func TestCOLUMN(t *testing.T) {
	resolver := &mockResolver{}

	// COLUMN(A1) should return 1
	cf := evalCompile(t, "COLUMN(A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(COLUMN(A1)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COLUMN(A1) = %v, want 1", got)
	}

	// COLUMN(C5) should return 3
	cf = evalCompile(t, "COLUMN(C5)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(COLUMN(C5)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COLUMN(C5) = %v, want 3", got)
	}

	// COLUMN() with context should return current column
	ctx := &EvalContext{CurrentRow: 5, CurrentCol: 3, CurrentSheet: "Sheet1"}
	cf = evalCompile(t, "COLUMN()")
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval(COLUMN()): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COLUMN() = %v, want 3", got)
	}
}
