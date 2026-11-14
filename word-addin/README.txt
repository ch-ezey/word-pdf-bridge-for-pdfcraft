PdfCraft for Word
=================
Adds a PdfCraft tab to Word's ribbon with two buttons:

  Create PDF       Saves <document name>.pdf next to the document and opens it in PdfCraft.
                   Asks before replacing an existing PDF. A never-saved document asks where to save.
                   Documents in OneDrive are saved to the matching folder of your local OneDrive.
  Create PDF as... Asks where to save the PDF, then opens it in PdfCraft.

The PDF is made by Word's own exporter: headings become PDF bookmarks, links and document
structure tags are kept. PdfCraft opens it as a new tab if it is already running.

Install:   double-click install-word-addin.cmd (copies the template into
           %APPDATA%\Microsoft\Word\STARTUP), then restart Word.
Uninstall: delete %APPDATA%\Microsoft\Word\STARTUP\PdfCraft for Word.dotm

Templates in Word's STARTUP folder are trusted, so no macro security settings need changing.
The macro source is PdfCraftForWord.bas. To edit it, open the .dotm with
File > Open in Word, then View > Macros > Edit.
