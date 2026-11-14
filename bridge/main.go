// PdfCraft Word bridge: a stand-in for LibreOffice's soffice.exe that converts Word documents
// to PDF with the installed Microsoft Word, so PdfCraft's "Office documents to PDF" works
// without LibreOffice.
//
// PdfCraft runs (crates/create/src/office.rs):
//
//	soffice.exe -env:UserInstallation=... --headless ... --convert-to pdf --outdir <out> <in>
//
// and reads <out>/<stem of in>.pdf. This program accepts that command line, runs
// word-to-pdf.ps1 (beside this exe) through Windows PowerShell, and exits 0 when the PDF exists.
// Point PdfCraft at it with the PDFCRAFT_SOFFICE environment variable.
package main

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// Extensions Word opens. PdfCraft also sends Excel and PowerPoint files here; those are refused.
var wordExtensions = map[string]bool{
	".doc": true, ".docx": true, ".docm": true, ".dot": true, ".dotx": true, ".dotm": true,
	".odt": true, ".ott": true, ".rtf": true,
}

type request struct {
	input  string
	outDir string
}

// parseArgs reads the soffice-style command line: --convert-to pdf, --outdir DIR, and one
// input file (the last argument that isn't an option). Other options are ignored.
func parseArgs(args []string) (request, error) {
	var req request
	format := ""
	for i := 0; i < len(args); i++ {
		a := args[i]
		switch {
		case a == "--convert-to" || a == "-convert-to":
			if i+1 >= len(args) {
				return req, errors.New("--convert-to needs a format")
			}
			i++
			format = args[i]
		case a == "--outdir" || a == "-outdir":
			if i+1 >= len(args) {
				return req, errors.New("--outdir needs a folder")
			}
			i++
			req.outDir = args[i]
		case strings.HasPrefix(a, "-"):
			// --headless, --nologo, -env:UserInstallation=..., and the rest: not needed for Word.
		default:
			req.input = a
		}
	}
	// "pdf" or a filter-qualified form such as "pdf:writer_pdf_Export".
	if f := strings.ToLower(format); f != "pdf" && !strings.HasPrefix(f, "pdf:") {
		return req, fmt.Errorf("only --convert-to pdf is supported (got %q)", format)
	}
	if req.input == "" {
		return req, errors.New("no input file given")
	}
	if req.outDir == "" {
		req.outDir = filepath.Dir(req.input)
	}
	if !wordExtensions[strings.ToLower(filepath.Ext(req.input))] {
		return req, fmt.Errorf("%s: only Word documents are converted by this bridge", filepath.Base(req.input))
	}
	return req, nil
}

// outputPath is where soffice would write: the input's name with .pdf, inside outDir.
func outputPath(req request) string {
	base := filepath.Base(req.input)
	return filepath.Join(req.outDir, strings.TrimSuffix(base, filepath.Ext(base))+".pdf")
}

// logf appends a line to %TEMP%\pdfcraft-word-bridge.log, since PdfCraft discards our output.
func logf(format string, a ...any) {
	f, err := os.OpenFile(filepath.Join(os.TempDir(), "pdfcraft-word-bridge.log"), os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0o600)
	if err != nil {
		return
	}
	defer f.Close()
	fmt.Fprintf(f, "%s %s\n", time.Now().Format(time.RFC3339), fmt.Sprintf(format, a...))
}

func fail(err error) {
	logf("error: %v", err)
	fmt.Fprintln(os.Stderr, "pdfcraft-word-bridge:", err)
	os.Exit(1)
}

func main() {
	if len(os.Args) > 1 && (os.Args[1] == "--version" || os.Args[1] == "--help") {
		fmt.Println("PdfCraft Word bridge 1.0: soffice-compatible Word to PDF conversion with Microsoft Word")
		return
	}
	exe, err := os.Executable()
	if err != nil {
		fail(fmt.Errorf("can't find own folder: %w", err))
	}
	script := filepath.Join(filepath.Dir(exe), "word-to-pdf.ps1")
	if _, err := os.Stat(script); err != nil {
		if len(os.Args) > 1 && os.Args[1] == createPdfFlag {
			showError("Convert to PDF with PdfCraft", "word-to-pdf.ps1 is missing beside "+exe+". Reinstall the PdfCraft Word bridge.")
		}
		fail(fmt.Errorf("word-to-pdf.ps1 is missing beside %s", exe))
	}
	if len(os.Args) > 1 && os.Args[1] == createPdfFlag {
		os.Exit(createPdfs(script, os.Args[2:]))
	}
	req, err := parseArgs(os.Args[1:])
	if err != nil {
		fail(err)
	}
	out := outputPath(req)
	if err := runWord(script, req.input, out); err != nil {
		fail(err)
	}
	info, err := os.Stat(out)
	if err != nil || info.Size() == 0 {
		fail(fmt.Errorf("Word finished but %s wasn't written", out))
	}
	logf("ok: %s -> %s (%d bytes)", req.input, out, info.Size())
}
