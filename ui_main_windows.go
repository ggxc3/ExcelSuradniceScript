//go:build windows

package main

import (
	"encoding/base64"
	"fmt"
	"os/exec"
	"strings"
	"syscall"
)

type WindowConfig struct {
	InputPath  string
	OutputPath string
	Sheet      string
	Columns    string
	StartRow   int
	EndRow     int
}

func ShowMainWindow(appVersion string) (*WindowConfig, bool, error) {
	script := fmt.Sprintf(`
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Excel Súradnice Script %s'
$form.Size = New-Object System.Drawing.Size(760, 500)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245,248,252)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Fill'
$panel.Padding = New-Object System.Windows.Forms.Padding(18)
$form.Controls.Add($panel)

function Add-Label($text, $x, $y, $w = 220) {
  $lbl = New-Object System.Windows.Forms.Label
  $lbl.Text = $text
  $lbl.Location = New-Object System.Drawing.Point($x, $y)
  $lbl.Size = New-Object System.Drawing.Size($w, 24)
  $panel.Controls.Add($lbl)
  return $lbl
}

function Add-Text($x, $y, $w = 420, $text = '') {
  $tb = New-Object System.Windows.Forms.TextBox
  $tb.Location = New-Object System.Drawing.Point($x, $y)
  $tb.Width = $w
  $tb.Multiline = $false
  $tb.AutoSize = $true
  $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  $tb.Font = New-Object System.Drawing.Font('Segoe UI', 10)
  $tb.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Left
  $tb.Text = $text
  $tb.SelectionStart = 0
  $tb.SelectionLength = 0
  $panel.Controls.Add($tb)
  return $tb
}

$title = Add-Label 'Excel Súradnice Script' 12 8 500
$title.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$subtitle = Add-Label 'Všetko nastavíš v jednom okne. Potom klikni na Spracovať.' 12 40 680

Add-Label 'Vstupný Excel:' 12 84
$input = Add-Text 220 82
$btnInput = New-Object System.Windows.Forms.Button
$btnInput.Text = 'Prehľadávať...'
$btnInput.Location = New-Object System.Drawing.Point(650, 80)
$btnInput.Size = New-Object System.Drawing.Size(90, 30)
$panel.Controls.Add($btnInput)

Add-Label 'Výstupný Excel:' 12 124
$output = Add-Text 220 122
$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Text = 'Uložiť ako...'
$btnOutput.Location = New-Object System.Drawing.Point(650, 120)
$btnOutput.Size = New-Object System.Drawing.Size(90, 30)
$panel.Controls.Add($btnOutput)

Add-Label 'Hárok:' 12 184
$sheet = Add-Text 220 182 250
Add-Label 'Stĺpce (A-N,B-E,C-V):' 12 224
$cols = Add-Text 220 222 250 'A-N'
Add-Label 'Od riadku:' 12 264
$start = Add-Text 220 262 120 '1'
Add-Label 'Do riadku:' 12 304
$end = Add-Text 220 302 120 '1'

$status = Add-Label 'Tip: pri výbere vstupu sa výstup nastaví automaticky.' 12 352 720
$status.ForeColor = [System.Drawing.Color]::FromArgb(70, 90, 120)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Zavrieť'
$btnCancel.Location = New-Object System.Drawing.Point(540, 400)
$btnCancel.Size = New-Object System.Drawing.Size(90, 34)
$panel.Controls.Add($btnCancel)

$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = 'Spracovať'
$btnProcess.Location = New-Object System.Drawing.Point(640, 400)
$btnProcess.Size = New-Object System.Drawing.Size(100, 34)
$btnProcess.BackColor = [System.Drawing.Color]::FromArgb(45, 125, 255)
$btnProcess.ForeColor = [System.Drawing.Color]::White
$panel.Controls.Add($btnProcess)

function DefaultOut([string]$path) {
  if ([string]::IsNullOrWhiteSpace($path)) { return '' }
  if ($path.ToLower().EndsWith('.xlsx')) { return $path.Substring(0, $path.Length-5) + '_formatted.xlsx' }
  return $path + '_formatted.xlsx'
}

$btnInput.Add_Click({
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = 'Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls'
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $input.Text = $dlg.FileName
    $input.SelectionStart = 0
    $input.SelectionLength = 0
    if ([string]::IsNullOrWhiteSpace($output.Text) -or $output.Text.ToLower().EndsWith('_formatted.xlsx')) {
      $output.Text = DefaultOut $dlg.FileName
      $output.SelectionStart = 0
      $output.SelectionLength = 0
    }
  }
})

$btnOutput.Add_Click({
  $dlg = New-Object System.Windows.Forms.SaveFileDialog
  $dlg.Filter = 'Excel files (*.xlsx)|*.xlsx'
  $dlg.FileName = $output.Text
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $output.Text = $dlg.FileName
    $output.SelectionStart = 0
    $output.SelectionLength = 0
  }
})

$btnCancel.Add_Click({
  Write-Output '__CANCELLED__'
  $form.Close()
})

$btnProcess.Add_Click({
  $payload = @($input.Text, $output.Text, $sheet.Text, $cols.Text, $start.Text, $end.Text) -join [Environment]::NewLine
  $encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload))
  Write-Output $encoded
  $form.Close()
})

[void]$form.ShowDialog()
`, appVersion)

	cmd := exec.Command("powershell", "-NoProfile", "-STA", "-Command", script)
	// CREATE_NO_WINDOW hides only the PowerShell console host while still allowing
	// WinForms to create and show its own application window.
	cmd.SysProcAttr = &syscall.SysProcAttr{CreationFlags: 0x08000000}
	out, err := cmd.Output()
	if err != nil {
		return nil, false, err
	}

	result := strings.TrimSpace(string(out))
	if result == "" || strings.Contains(result, "__CANCELLED__") {
		return nil, true, nil
	}

	decoded, err := base64.StdEncoding.DecodeString(lastLine(result))
	if err != nil {
		return nil, false, fmt.Errorf("nepodarilo sa načítať údaje z formulára: %w", err)
	}

	parts := strings.Split(strings.ReplaceAll(string(decoded), "\r", ""), "\n")
	if len(parts) < 6 {
		return nil, false, fmt.Errorf("neúplné údaje z formulára")
	}

	start, err := ParsePositiveInt(parts[4], "Od riadku")
	if err != nil {
		return nil, false, err
	}
	end, err := ParsePositiveInt(parts[5], "Do riadku")
	if err != nil {
		return nil, false, err
	}

	return &WindowConfig{
		InputPath:  strings.TrimSpace(parts[0]),
		OutputPath: strings.TrimSpace(parts[1]),
		Sheet:      strings.TrimSpace(parts[2]),
		Columns:    strings.TrimSpace(parts[3]),
		StartRow:   start,
		EndRow:     end,
	}, false, nil
}

func lastLine(text string) string {
	items := strings.Split(strings.ReplaceAll(text, "\r", ""), "\n")
	for i := len(items) - 1; i >= 0; i-- {
		if strings.TrimSpace(items[i]) != "" {
			return strings.TrimSpace(items[i])
		}
	}
	return ""
}
