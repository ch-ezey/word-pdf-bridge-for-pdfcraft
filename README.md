# Word PDF bridge for PdfCraft

Turn Word documents into PDFs with [PdfCraft](https://github.com/storytold/pdfcraft) on Windows,
using the Microsoft Word you already have instead of installing LibreOffice.

> **Unofficial.** This project is not made by, affiliated with or endorsed by the ArtCraft Team or
> the PdfCraft contributors. "ArtCraft" and "PdfCraft" are their trademarks and are used here only
> to say what this works with.

## What you get

| Part | What it does |
| --- | --- |
| **Bridge** | PdfCraft's *File ▸ Open*, *Create PDF from file* and *Combine files* accept Word documents (.doc, .docx, .docm, .dot/.dotx/.dotm, .rtf, .odt). |
| **Right-click menu** | In File Explorer, Word documents get **Convert to PDF with PdfCraft** (saves `name.pdf` beside the document and opens it) and **Combine files in PdfCraft**. |
| **Word add-in** | A **PdfCraft** tab in Word's ribbon with **Create PDF** and **Create PDF as…**, like Acrobat's. |

All PDFs are made by Word's own exporter, so layout, fonts, tables and links come out as Word
draws them, and headings become PDF bookmarks.

## Requirements

- Windows 10 or 11
- Microsoft Word (desktop, Microsoft 365 or Office 2016 and later)
- PdfCraft installed in `C:\Program Files\ArtCraft\PdfCraft`

## Install

Download the zips from the [latest release](../../releases/latest). Nothing needs admin rights.

**Bridge and right-click menu** (`word-pdf-bridge-<version>.zip`)

1. Unzip, then double-click `install.cmd`. The result is written to `install-log.txt`.
2. Fully close PdfCraft **and the ArtCraft Launcher**, then reopen them. Programs that were
   already running don't see the new setting.

On Windows 11 the right-click entries are under **Show more options** (or Shift+F10), like
PdfCraft's own entries.

**Word add-in** (`pdfcraft-for-word-<version>.zip`)

1. Unzip, then double-click `install-word-addin.cmd`.
2. Restart Word. The **PdfCraft** tab appears in the ribbon.

### Windows may warn you

The executable is not code-signed, so SmartScreen may say "Windows protected your PC" (choose
*More info ▸ Run anyway*), and some antivirus tools are wary of small programs that start
PowerShell. Every release is built from this repository's source by
[GitHub Actions](.github/workflows/release.yml); you can check the workflow run that produced it,
or build it yourself (below).

## Uninstall

- Bridge and menu: `powershell -ExecutionPolicy Bypass -File .\install.ps1 -Uninstall`
- Word add-in: delete `%APPDATA%\Microsoft\Word\STARTUP\PdfCraft for Word.dotm`

## How it works

PdfCraft converts Office documents by running LibreOffice's `soffice.exe`
(`soffice --headless --convert-to pdf --outdir <dir> <file>`), and it lets you choose which
`soffice.exe` with the `PDFCRAFT_SOFFICE` environment variable. The bridge is a small Go program
that accepts that command line and has Word export the PDF through
[`word-to-pdf.ps1`](bridge/word-to-pdf.ps1). PdfCraft itself is not modified.

- Word runs hidden, in its own instance; documents open in your Word windows are not touched.
- Macros in converted documents are disabled (`AutomationSecurity = ForceDisable`).
- Password-protected documents are refused instead of waiting on a hidden password prompt.
- Problems are logged to `%TEMP%\pdfcraft-word-bridge.log`.

The right-click menu uses per-user registry keys under
`HKCU\Software\Classes\SystemFileAssociations\.<ext>\shell`. The Word add-in is a global template
in Word's STARTUP folder (a trusted location, so no macro settings change); its macros are in
[`PdfCraftForWord.bas`](word-addin/PdfCraftForWord.bas) and its ribbon in
[`customUI14.xml`](word-addin/customUI14.xml).

## Limitations

- Word documents only. Excel and PowerPoint files still need LibreOffice.
- If PdfCraft stops a very slow conversion (after 120 seconds), a hidden `WINWORD.EXE` can be left
  running; end it in Task Manager.
- The right-click menu shows on the classic menu only (Windows 11 "Show more options").

## Building

```sh
cd bridge
go test ./...
GOOS=windows GOARCH=amd64 go build -trimpath -ldflags "-s -w -H windowsgui" -o soffice.exe .
```

`word-addin/PdfCraft for Word.dotm` can't be built without Word. To change it, open it in Word with
*File ▸ Open*, edit the macros (*View ▸ Macros ▸ Edit*), save, and copy the module back into
`PdfCraftForWord.bas` so the source stays readable.

## Licence

[MIT](LICENSE).
