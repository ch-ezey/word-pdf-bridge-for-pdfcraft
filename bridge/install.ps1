# Installs the PdfCraft Word bridge for the current user (no admin rights needed):
# - copies soffice.exe and word-to-pdf.ps1 to %LOCALAPPDATA%\PdfCraftWordBridge and sets
#   PDFCRAFT_SOFFICE, so PdfCraft converts Word documents with Microsoft Word;
# - adds two entries to the right-click menu of Word documents in File Explorer:
#   "Convert to PDF with PdfCraft" and "Combine files in PdfCraft".
#   (On Windows 11 they are under "Show more options", like PdfCraft's own entries.)
# Run:  install.cmd, or powershell -ExecutionPolicy Bypass -File .\install.ps1
# Undo: powershell -ExecutionPolicy Bypass -File .\install.ps1 -Uninstall
param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$target = Join-Path $env:LOCALAPPDATA 'PdfCraftWordBridge'
$wordExtensions = 'doc', 'docx', 'docm', 'rtf', 'odt'
$verbs = 'PdfCraft.ConvertToPdf', 'PdfCraft.CombineWord'

function Remove-MenuEntries {
    foreach ($ext in $wordExtensions) {
        foreach ($verb in $verbs) {
            $key = "HKCU:\Software\Classes\SystemFileAssociations\.$ext\shell\$verb"
            if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse -Force }
        }
    }
}

function Set-MenuEntry([string]$ext, [string]$verb, [string]$label, [string]$icon, [string]$command) {
    $key = "HKCU:\Software\Classes\SystemFileAssociations\.$ext\shell\$verb"
    New-Item -Path "$key\command" -Force | Out-Null
    Set-Item -LiteralPath $key -Value $label
    New-ItemProperty -LiteralPath $key -Name 'Icon' -Value $icon -PropertyType String -Force | Out-Null
    # Each selected file starts one command (up to Explorer's limit of 15).
    New-ItemProperty -LiteralPath $key -Name 'MultiSelectModel' -Value 'Document' -PropertyType String -Force | Out-Null
    Set-Item -LiteralPath "$key\command" -Value $command
}

if ($Uninstall) {
    Remove-MenuEntries
    [Environment]::SetEnvironmentVariable('PDFCRAFT_SOFFICE', $null, 'User')
    if (Test-Path $target) { Remove-Item $target -Recurse -Force }
    Write-Host 'Removed the bridge and its right-click menu entries. Restart PdfCraft.'
    exit 0
}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
New-Item -ItemType Directory -Force -Path $target | Out-Null
foreach ($f in 'soffice.exe', 'word-to-pdf.ps1') {
    Copy-Item (Join-Path $here $f) $target -Force
    Unblock-File (Join-Path $target $f)   # clear the "downloaded from the internet" mark
}
$exe = Join-Path $target 'soffice.exe'
[Environment]::SetEnvironmentVariable('PDFCRAFT_SOFFICE', $exe, 'User')

# Right-click menu entries for Word documents.
$pdfcraft = $null
foreach ($root in $env:ProgramW6432, $env:ProgramFiles) {
    if ($root -and (Test-Path (Join-Path $root 'ArtCraft\PdfCraft\pdfcraft.exe'))) {
        $pdfcraft = Join-Path $root 'ArtCraft\PdfCraft\pdfcraft.exe'; break
    }
}
Remove-MenuEntries
if ($pdfcraft) {
    $icon = "`"$pdfcraft`",0"
    foreach ($ext in $wordExtensions) {
        Set-MenuEntry $ext 'PdfCraft.ConvertToPdf' 'Convert to PDF with PdfCraft' $icon "`"$exe`" --pdfcraft-create-pdf `"%1`""
        Set-MenuEntry $ext 'PdfCraft.CombineWord' 'Combine files in PdfCraft' $icon "`"$pdfcraft`" --combine `"%1`""
    }
    $check = Get-Item -LiteralPath 'HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\PdfCraft.ConvertToPdf\command'
    Write-Host "Right-click menu added for: .$($wordExtensions -join ', .')"
    Write-Host "  Convert command: $($check.GetValue(''))"
} else {
    Write-Warning 'PdfCraft was not found in Program Files\ArtCraft\PdfCraft, so the right-click menu entries were not added.'
}

# Check Word answers through COM before declaring success.
try {
    $w = New-Object -ComObject Word.Application
    $v = $w.Version
    $w.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    Write-Host "Installed to $target (Microsoft Word $v found)."
} catch {
    Write-Warning "Installed to $target, but Microsoft Word couldn't be started: $($_.Exception.Message)"
}
Write-Host 'Close PdfCraft (and the ArtCraft Launcher) completely and reopen them so they see the new setting.'
