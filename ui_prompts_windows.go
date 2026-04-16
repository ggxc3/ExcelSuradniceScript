//go:build windows

package main

import (
	"fmt"
	"os/exec"
	"strconv"
	"strings"
)

func PromptText(title, message, defaultValue string) (string, error) {
	script := fmt.Sprintf(`
Add-Type -AssemblyName Microsoft.VisualBasic
$result = [Microsoft.VisualBasic.Interaction]::InputBox(%q, %q, %q)
if ($result -eq "") { Write-Output "__CANCELLED__" } else { Write-Output $result }
`, message, title, defaultValue)

	out, err := exec.Command("powershell", "-NoProfile", "-Command", script).Output()
	if err != nil {
		return "", err
	}
	result := strings.TrimSpace(string(out))
	if result == "__CANCELLED__" {
		return "", fmt.Errorf("operácia zrušená používateľom")
	}
	return result, nil
}

func PromptChoice(title, message string, options []string) (string, error) {
	if len(options) == 0 {
		return "", fmt.Errorf("nie sú dostupné žiadne možnosti")
	}

	list := make([]string, 0, len(options))
	for i, opt := range options {
		list = append(list, fmt.Sprintf("%d - %s", i+1, strings.ReplaceAll(opt, "\"", "'")))
	}
	prompt := message + "\n" + strings.Join(list, "\n") + "\n\nZadaj číslo možnosti:"

	value, err := PromptText(title, prompt, "1")
	if err != nil {
		return "", err
	}

	index, err := strconv.Atoi(strings.TrimSpace(value))
	if err != nil || index < 1 || index > len(options) {
		return "", fmt.Errorf("neplatný výber hárku")
	}
	return options[index-1], nil
}
