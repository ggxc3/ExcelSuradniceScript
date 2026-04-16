package main

import "testing"

func TestParseColumnsSpec(t *testing.T) {
	columns, err := ParseColumnsSpec("a-n, b-e, c-v")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(columns) != 3 {
		t.Fatalf("expected 3 items, got %d", len(columns))
	}
	if columns[0][0] != "A" || columns[0][1] != "N" {
		t.Fatalf("unexpected first pair: %#v", columns[0])
	}
}

func TestParseColumnsSpecRejectsInvalid(t *testing.T) {
	_, err := ParseColumnsSpec("A-X")
	if err == nil {
		t.Fatal("expected error for invalid type")
	}
}

func TestParsePositiveInt(t *testing.T) {
	value, err := ParsePositiveInt("42", "Od")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if value != 42 {
		t.Fatalf("expected 42, got %d", value)
	}
}

func TestParsePositiveIntRejectsZero(t *testing.T) {
	_, err := ParsePositiveInt("0", "Od")
	if err == nil {
		t.Fatal("expected error")
	}
}
