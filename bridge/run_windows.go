package main

import (
	"bytes"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"
	"unsafe"
)

// startDetached starts program with args and doesn't wait for it.
func startDetached(program string, args []string) error {
	cmd := exec.Command(program, args...)
	return cmd.Start()
}

// showError shows a Windows message box (MB_OK | MB_ICONERROR | MB_SETFOREGROUND).
func showError(title, text string) {
	t, _ := syscall.UTF16PtrFromString(title)
	m, _ := syscall.UTF16PtrFromString(text)
	proc := syscall.NewLazyDLL("user32.dll").NewProc("MessageBoxW")
	proc.Call(0, uintptr(unsafe.Pointer(m)), uintptr(unsafe.Pointer(t)), 0x10|0x10000)
}

// runWord runs word-to-pdf.ps1 with Windows' own PowerShell, with no console window.
func runWord(script, input, output string) error {
	ps := "powershell.exe"
	if root := os.Getenv("SystemRoot"); root != "" {
		ps = filepath.Join(root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
	}
	cmd := exec.Command(ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
		"-File", script, "-InputPath", input, "-OutputPath", output)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: 0x08000000} // CREATE_NO_WINDOW
	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		msg := strings.TrimSpace(stderr.String())
		if msg == "" {
			msg = err.Error()
		}
		return fmt.Errorf("Word couldn't convert %s: %s", filepath.Base(input), msg)
	}
	return nil
}
