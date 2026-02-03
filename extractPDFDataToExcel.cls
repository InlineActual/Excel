VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr) As LongPtr

#End If

Const WM_CLOSE = &H10

Sub ExtractPDFDataToExcel()
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Sheets("Sources")
    Dim sourceRange As Range: Set sourceRange = sourceWorkbook.Sheets("Formated").Range("C4:J7")
    Dim lastRow As Long: lastRow = WorksheetFunction.CountA(sourceSheet.Columns("D"))
    Dim AcrobatPath As String: AcrobatPath = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    
    Dim i As Long, rowData As Variant
    Dim PDFPath As String, ShellCommand As String
    Dim targetWorkbook As Workbook, targetSheet As Worksheet
    Dim targetRange As Range
    
    Application.CutCopyMode = False
    
    If sourceSheet.Range("F2").Value = 1 Then
        sourceSheet.Range("A9").Value = "Started: " & Now
    End If
    
    For i = 2 To lastRow
    
        With sourceWorkbook.Sheets("DataDrop")
            .Columns(1).ClearContents
        End With
        rowData = sourceSheet.Range("D" & i & ":G" & i).Value
        PDFPath = rowData(1, 1)

        If Len(Dir(PDFPath)) = 0 Then
            Debug.Print "PDF not found: " & PDFPath
            GoTo SkipRow
        End If

        ShellCommand = """" & AcrobatPath & """ """ & PDFPath & """"
        Shell ShellCommand, vbNormalFocus
        Application.Wait Now + TimeValue("0:00:03")

        SendKeys "^a", True: SendKeys "^c", True
        Application.Wait Now + TimeValue("0:00:02")
        SendKeys "^w", True  ' CTRL + W closes current PDF tab
        Application.Wait Now + TimeValue("0:00:01")
        AppActivate Application.Caption
        sourceWorkbook.Activate
        sourceWorkbook.Sheets("DataDrop").Activate

        With sourceWorkbook.Sheets("DataDrop")
            .Range("A1").Select
            .PasteSpecial Format:="Text"
        End With

        Call CloseCurrentPDF(ExtractFileName(PDFPath))

        
        Set targetWorkbook = Workbooks.Open(CStr(rowData(1, 2)))
        Set targetSheet = targetWorkbook.Sheets(CStr(rowData(1, 3)))
        'top left cell in desination range
        Set targetRange = targetSheet.Range(CStr(rowData(1, 4)))
        
        sourceRange.Copy
        targetRange.PasteSpecial Paste:=xlPasteValues

        Application.CutCopyMode = False
        'targetWorkbook.Save

SkipRow:
    Next i
    sourceSheet.Activate
    sourceSheet.Range("A10").Value = "Completed: " & Now
    
End Sub

Function ExtractFileName(ByVal fullPath As String) As String
    Dim parts() As String: parts = Split(fullPath, "\")
    ExtractFileName = Replace(parts(UBound(parts)), ".pdf", "")
End Function

Sub CloseCurrentPDF(ByVal pdfTitle As String)
    Dim hwnd As LongPtr
    hwnd = FindWindow(vbNullString, pdfTitle & " - Adobe Acrobat Reader DC")
    If hwnd <> 0 Then SendMessage hwnd, WM_CLOSE, 0, 0
End Sub




