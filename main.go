package main

import (
	"bufio"
	"fmt"
	"os"
)

func main() {
	reader := bufio.NewReader(os.Stdin)

	app, err := NewApp(reader)
	if err != nil {
		fmt.Println(err)
		waitForAnyKey(reader)
		return
	}

	if err := app.Start(); err != nil {
		fmt.Println(err)
	}

	waitForAnyKey(reader)
}
