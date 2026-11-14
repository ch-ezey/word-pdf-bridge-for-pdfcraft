<#
  PdfCraft Word bridge: converts one Word document to PDF with the installed Microsoft Word.
  Called by soffice.exe (the bridge) in this folder; you don't run it yourself.

  Usage: word-to-pdf.ps1 -InputPath <document> -OutputPath <file.pdf>
  Exit code 0 on success, 1 on failure (the reason goes to stderr and the bridge's log).
#>
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$wdExportFormatPDF        = 17
$wdExportOptimizeForPrint = 0
$wdDoNotSaveChanges       = 0
$msoAutomationSecurityForceDisable = 3

# Word's Open, Close and Quit take their arguments by reference, which PowerShell's COM binder
# refuses for plain values ("Cannot convert the True value of type bool to type Object").
# Calling through IDispatch with named arguments avoids that and documents each one.
function Invoke-Com($target, [string]$method, [System.Collections.Specialized.OrderedDictionary]$named) {
    $names  = [string[]]@($named.Keys)
    $values = [object[]]@($named.Values)
    [System.__ComObject].InvokeMember($method, [Reflection.BindingFlags]::InvokeMethod, $null, $target, $values, $null, $null, $names)
}

$word = $null
$doc  = $null
$code = 1
try {
    # Word's class is single-use, so this starts a separate, hidden Word: documents you have
    # open in your own Word window are not touched.
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0                                       # wdAlertsNone
    $word.AutomationSecurity = $msoAutomationSecurityForceDisable # never run macros in the document

    # A dummy password makes a password-protected document fail at once instead of waiting
    # on a password prompt nobody can see; unprotected documents ignore it.
    $doc = Invoke-Com $word.Documents 'Open' ([ordered]@{
        FileName           = $InputPath
        ConfirmConversions = $false
        ReadOnly           = $true
        AddToRecentFiles   = $false
        PasswordDocument   = 'pdfcraft-no-password'
        Visible            = $false
        OpenAndRepair      = $false
        NoEncodingDialog   = $true
    })

    # CreateBookmarks 1 = from headings, so PdfCraft's bookmark panel gets the outline.
    [void](Invoke-Com $doc 'ExportAsFixedFormat' ([ordered]@{
        OutputFileName     = $OutputPath
        ExportFormat       = $wdExportFormatPDF
        OpenAfterExport    = $false
        OptimizeFor        = $wdExportOptimizeForPrint
        IncludeDocProps    = $true
        KeepIRM            = $true
        CreateBookmarks    = 1
        DocStructureTags   = $true
        BitmapMissingFonts = $true
        UseISO19005_1      = $false
    }))

    $code = 0
}
catch {
    $e = $_.Exception
    while ($e.InnerException) { $e = $e.InnerException }
    [Console]::Error.WriteLine($e.Message)
}
finally {
    if ($doc) {
        try { [void](Invoke-Com $doc 'Close' ([ordered]@{ SaveChanges = $wdDoNotSaveChanges })) } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($word) {
        # Only quit a Word with nothing else open in it.
        try {
            if ($word.Documents.Count -eq 0) { [void](Invoke-Com $word 'Quit' ([ordered]@{ SaveChanges = $wdDoNotSaveChanges })) }
        } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit $code
