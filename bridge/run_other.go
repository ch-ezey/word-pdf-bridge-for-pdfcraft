//go:build !windows

package main

import "errors"

// runWord needs Microsoft Word, which only exists on Windows here.
func runWord(script, input, output string) error {
	return errors.New("the Word bridge only runs on Windows")
}

func startDetached(program string, args []string) error { return errors.New("Windows only") }

func showError(title, text string) {}
