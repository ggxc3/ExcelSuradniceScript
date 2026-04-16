//go:build !windows

package main

import "fmt"

func main() {
	fmt.Println("Táto verzia je desktop GUI aplikácia určená primárne pre Windows release build.")
}
