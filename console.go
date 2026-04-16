package main

import (
	"bufio"
	"fmt"
	"os"

	"golang.org/x/term"
)

func waitForAnyKey(reader *bufio.Reader) {
	fmt.Println("Pre zatvorenie konzoly stlačte ľubovoľné tlačidlo.")

	fd := int(os.Stdin.Fd())
	if !term.IsTerminal(fd) {
		_, _ = reader.ReadByte()
		return
	}

	oldState, err := term.MakeRaw(fd)
	if err != nil {
		_, _ = reader.ReadByte()
		return
	}
	defer func() {
		_ = term.Restore(fd, oldState)
	}()

	var buffer [1]byte
	_, _ = os.Stdin.Read(buffer[:])
}
