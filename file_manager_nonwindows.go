//go:build !windows

package main

import "fmt"

func LoadFile() (string, bool, error) {
	return "", false, fmt.Errorf("táto Go verzia podporuje file dialógy iba na Windows")
}

func SaveFile() (string, bool, error) {
	return "", false, fmt.Errorf("táto Go verzia podporuje file dialógy iba na Windows")
}
