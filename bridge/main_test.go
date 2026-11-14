package main

import (
	"path/filepath"
	"testing"
)

// The exact command line PdfCraft builds in crates/create/src/office.rs.
func pdfcraftArgs(out, in string) []string {
	return []string{"-env:UserInstallation=file:///C:/Temp/p/profile", "--headless", "--invisible", "--nologo",
		"--norestore", "--nodefault", "--nolockcheck", "--convert-to", "pdf", "--outdir", out, in}
}

func TestPdfCraftCommandLine(t *testing.T) {
	out := filepath.Join("work", "out")
	in := filepath.Join("work", "document.docx")
	req, err := parseArgs(pdfcraftArgs(out, in))
	if err != nil {
		t.Fatal(err)
	}
	if req.input != in || req.outDir != out {
		t.Fatalf("got %+v", req)
	}
	if got, want := outputPath(req), filepath.Join(out, "document.pdf"); got != want {
		t.Fatalf("output %q, want %q", got, want)
	}
}

func TestWordExtensionsAccepted(t *testing.T) {
	for _, name := range []string{"document.doc", "document.DOCX", "document.rtf", "document.odt", "document.dotm"} {
		if _, err := parseArgs(pdfcraftArgs("o", name)); err != nil {
			t.Errorf("%s: %v", name, err)
		}
	}
}

func TestRefusals(t *testing.T) {
	cases := map[string][]string{
		"excel":      pdfcraftArgs("o", "document.xlsx"),
		"powerpoint": pdfcraftArgs("o", "document.pptx"),
		"not pdf":    {"--convert-to", "docx", "--outdir", "o", "a.doc"},
		"no input":   {"--convert-to", "pdf", "--outdir", "o"},
		"no format":  {"--outdir", "o", "a.docx"},
		"dangling":   {"a.docx", "--outdir"},
	}
	for name, args := range cases {
		if _, err := parseArgs(args); err == nil {
			t.Errorf("%s: accepted", name)
		}
	}
}

func TestFilterQualifiedPdf(t *testing.T) {
	if _, err := parseArgs([]string{"--convert-to", "pdf:writer_pdf_Export", "a.docx"}); err != nil {
		t.Fatal(err)
	}
}

func TestUniquePdfPathNeverReplaces(t *testing.T) {
	in := filepath.Join("docs", "Report.docx")
	taken := map[string]bool{}
	exists := func(p string) bool { return taken[p] }
	if got := uniquePdfPath(in, exists); got != filepath.Join("docs", "Report.pdf") {
		t.Fatalf("got %q", got)
	}
	taken[filepath.Join("docs", "Report.pdf")] = true
	taken[filepath.Join("docs", "Report (2).pdf")] = true
	if got := uniquePdfPath(in, exists); got != filepath.Join("docs", "Report (3).pdf") {
		t.Fatalf("got %q", got)
	}
}

func TestUniquePdfPathKeepsDotsInName(t *testing.T) {
	got := uniquePdfPath(filepath.Join("d", "v1.2 notes.docx"), func(string) bool { return false })
	if got != filepath.Join("d", "v1.2 notes.pdf") {
		t.Fatalf("got %q", got)
	}
}
