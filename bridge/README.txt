PdfCraft Word bridge
====================
1. PdfCraft opens Word documents (.doc, .docx, .docm, .dot/.dotx/.dotm, .rtf, .odt) as PDFs using
   the Microsoft Word you already have, instead of LibreOffice. PdfCraft itself is unchanged:
   it converts Office files by running a program called soffice.exe, chosen with the
   PDFCRAFT_SOFFICE environment variable, and this soffice.exe has Word do the conversion.
2. Right-click menu on Word documents (.doc, .docx, .docm, .rtf, .odt) in File Explorer:
   - Convert to PDF with PdfCraft: saves <name>.pdf next to the document (or <name> (2).pdf if
     that exists; nothing is replaced) and opens it in PdfCraft.
   - Combine files in PdfCraft: select several documents; they go to PdfCraft's Combine tab.
   On Windows 11 these are under "Show more options" (or Shift+F10), like PdfCraft's own entries.

Install:   double-click install.cmd (result is written to install-log.txt), then fully close
           PdfCraft AND the ArtCraft Launcher and reopen them. A program started before the
           install doesn't see the new PDFCRAFT_SOFFICE setting and passes the old settings on.
Uninstall: powershell -ExecutionPolicy Bypass -File .\install.ps1 -Uninstall
           (removes the files, the setting and the menu entries)

Notes
- Word runs hidden, in its own instance; your open Word windows aren't affected. Macros in the
  converted document are disabled. Password-protected documents are refused.
- Excel and PowerPoint files are not handled by this bridge.
- Problems are logged to %TEMP%\pdfcraft-word-bridge.log; the menu command also shows a message.
- If PdfCraft gives up on a very slow conversion (after 120 s), a hidden WINWORD.EXE may be left
  running; end it in Task Manager.
- Source: https://github.com/ch-ezey/word-pdf-bridge-for-pdfcraft
