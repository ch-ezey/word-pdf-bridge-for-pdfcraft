package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// createPdfFlag is the Explorer menu mode: soffice.exe --pdfcraft-create-pdf <file>...
// Each Word file becomes <name>.pdf beside it (or <name> (2).pdf and so on, never replacing a
// file), made by Word, and the PDF opens in PdfCraft.
const createPdfFlag = "--pdfcraft-create-pdf"

// uniquePdfPath is dir\name.pdf, or the first of dir\name (2).pdf, (3)... that doesn't exist.
func uniquePdfPath(input string, exists func(string) bool) string {
	dir := filepath.Dir(input)
	base := filepath.Base(input)
	stem := strings.TrimSuffix(base, filepath.Ext(base))
	out := filepath.Join(dir, stem+".pdf")
	for n := 2; exists(out); n++ {
		out = filepath.Join(dir, fmt.Sprintf("%s (%d).pdf", stem, n))
	}
	return out
}

func fileExists(p string) bool {
	_, err := os.Stat(p)
	return err == nil
}

// pdfcraftExe is PdfCraft's program: PDFCRAFT_EXE, else the standard install folder.
func pdfcraftExe() string {
	if p := os.Getenv("PDFCRAFT_EXE"); p != "" && fileExists(p) {
		return p
	}
	for _, root := range []string{os.Getenv("ProgramW6432"), os.Getenv("ProgramFiles"), `C:\Program Files`} {
		if root == "" {
			continue
		}
		if p := filepath.Join(root, "ArtCraft", "PdfCraft", "pdfcraft.exe"); fileExists(p) {
			return p
		}
	}
	return ""
}

// createPdfs converts each file and opens the results in PdfCraft. Problems are collected and
// shown together in one message box.
func createPdfs(script string, files []string) int {
	var problems []string
	var made []string
	for _, in := range files {
		abs, err := filepath.Abs(in)
		if err == nil {
			in = abs
		}
		if !wordExtensions[strings.ToLower(filepath.Ext(in))] {
			problems = append(problems, filepath.Base(in)+": not a Word document")
			continue
		}
		out := uniquePdfPath(in, fileExists)
		if err := runWord(script, in, out); err != nil {
			problems = append(problems, err.Error())
			logf("error: %v", err)
			continue
		}
		if !fileExists(out) {
			problems = append(problems, filepath.Base(in)+": Word finished but no PDF was written")
			continue
		}
		logf("ok (menu): %s -> %s", in, out)
		made = append(made, out)
	}
	if len(made) > 0 {
		if exe := pdfcraftExe(); exe != "" {
			if err := startDetached(exe, made); err != nil {
				problems = append(problems, "PdfCraft couldn't be started: "+err.Error())
			}
		} else {
			problems = append(problems, "Saved "+strings.Join(made, ", ")+", but PdfCraft wasn't found to open it.")
		}
	}
	if len(problems) > 0 {
		showError("Convert to PDF with PdfCraft", strings.Join(problems, "\n\n"))
		return 1
	}
	return 0
}
