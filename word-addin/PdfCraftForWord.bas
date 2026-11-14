Option Explicit

' PdfCraft for Word: the PdfCraft ribbon tab's two buttons.
' Create PDF     saves <document name>.pdf next to the document and opens it in PdfCraft.
' Create PDF as  asks where to save the PDF first.
' The PDF is made by Word's own exporter: headings become bookmarks, links and tags are kept.

Private Const PDFCRAFT_EXE As String = "C:\Program Files\ArtCraft\PdfCraft\pdfcraft.exe"

Public Sub PdfCraft_CreatePdf(control As IRibbonControl)
    CreatePdf False
End Sub

Public Sub PdfCraft_CreatePdfAs(control As IRibbonControl)
    CreatePdf True
End Sub

Private Sub CreatePdf(ByVal askWhere As Boolean)
    Dim doc As Document
    Dim target As String

    If Documents.Count = 0 Then
        MsgBox "Open a document first.", vbInformation, "Create PDF"
        Exit Sub
    End If
    Set doc = ActiveDocument

    target = DefaultPdfPath(doc)
    If askWhere Or target = "" Then
        target = AskForPdfPath(target)
        If target = "" Then Exit Sub
    ElseIf Dir(target) <> "" Then
        If MsgBox(Mid(target, InStrRev(target, "\") + 1) & " already exists." & vbCrLf & _
                  "Do you want to replace it?", vbYesNo + vbExclamation, "Create PDF") <> vbYes Then
            target = AskForPdfPath(target)
            If target = "" Then Exit Sub
        End If
    End If

    On Error GoTo Failed
    Application.StatusBar = "Creating " & target & " ..."
    doc.ExportAsFixedFormat OutputFileName:=target, ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    Application.StatusBar = "Created " & target
    OpenInPdfCraft target
    Exit Sub

Failed:
    Application.StatusBar = False
    MsgBox "The PDF couldn't be created:" & vbCrLf & Err.Description, vbExclamation, "Create PDF"
End Sub

' <folder>\<name>.pdf for a saved document; "" for one that has never been saved.
Private Function DefaultPdfPath(ByVal doc As Document) As String
    Dim folder As String, baseName As String
    folder = LocalFolder(doc.Path)
    If folder = "" Then Exit Function
    baseName = doc.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    DefaultPdfPath = folder & "\" & baseName & ".pdf"
End Function

' A document's folder on this PC. Files in OneDrive report a web address
' (https://d.docs.live.net/<id>/folder/...), mapped here to the local OneDrive folder.
Private Function LocalFolder(ByVal docPath As String) As String
    Dim root As String, rest As String, slash As Long
    If docPath = "" Then Exit Function
    If LCase(Left(docPath, 4)) <> "http" Then
        LocalFolder = docPath
        Exit Function
    End If
    If InStr(1, docPath, "d.docs.live.net/", vbTextCompare) = 0 Then Exit Function
    root = Environ("OneDriveConsumer")
    If root = "" Then root = Environ("OneDrive")
    If root = "" Then Exit Function
    rest = Mid(docPath, InStr(1, docPath, "d.docs.live.net/", vbTextCompare) + Len("d.docs.live.net/"))
    slash = InStr(rest, "/")
    If slash = 0 Then
        rest = ""
    Else
        rest = Mid(rest, slash + 1)
    End If
    rest = Replace(UrlDecode(rest), "/", "\")
    If rest = "" Then
        LocalFolder = root
    Else
        LocalFolder = root & "\" & rest
    End If
    If Dir(LocalFolder, vbDirectory) = "" Then LocalFolder = ""
End Function

Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        If Mid(s, i, 1) = "%" And i + 2 <= Len(s) Then
            out = out & Chr(CLng("&H" & Mid(s, i + 1, 2)))
            i = i + 3
        Else
            out = out & Mid(s, i, 1)
            i = i + 1
        End If
    Loop
    UrlDecode = out
End Function

' Word's Save As box, set to PDF. Returns "" when cancelled.
Private Function AskForPdfPath(ByVal suggested As String) As String
    Dim dlg As FileDialog, i As Long, chosen As String
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    dlg.Title = "Create PDF"
    If suggested <> "" Then
        dlg.InitialFileName = suggested
    ElseIf ActiveDocument.Name <> "" Then
        dlg.InitialFileName = Options.DefaultFilePath(wdDocumentsPath) & "\" & ActiveDocument.Name
    End If
    For i = 1 To dlg.Filters.Count
        If InStr(1, dlg.Filters(i).Extensions, "pdf", vbTextCompare) > 0 Then
            dlg.FilterIndex = i
            Exit For
        End If
    Next i
    If dlg.Show <> -1 Then Exit Function
    chosen = dlg.SelectedItems(1)
    If LCase(Right(chosen, 4)) <> ".pdf" Then
        If InStrRev(chosen, ".") > InStrRev(chosen, "\") Then chosen = Left(chosen, InStrRev(chosen, ".") - 1)
        chosen = chosen & ".pdf"
    End If
    AskForPdfPath = chosen
End Function

' Opens the PDF in PdfCraft (as a new tab when PdfCraft is already open).
Private Sub OpenInPdfCraft(ByVal pdfPath As String)
    If Dir(PDFCRAFT_EXE) = "" Then
        MsgBox "The PDF was saved as " & pdfPath & ", but PdfCraft wasn't found at " & PDFCRAFT_EXE & ".", _
               vbInformation, "Create PDF"
        Exit Sub
    End If
    Shell """" & PDFCRAFT_EXE & """ """ & pdfPath & """", vbNormalFocus
End Sub
