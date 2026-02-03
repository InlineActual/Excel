Attribute VB_Name = "Module20"
Sub ProtectSheet()
Attribute ProtectSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ProtectSheet Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
    ' Right-click cells you want unlocked before running:
    ' Format Cells... -> Protection -> (uncheck) Locked
    ' Save "Personal Macro Workbook" when you close Excel to keep this macro.

    Dim ws As Worksheet
    Dim pwd As String
    Dim errors As String
    Dim errorCount As Long

    pwd = InputBox("Enter the sheet password to apply to all worksheets:", "Protect Sheets")
    If Len(pwd) = 0 Then
        MsgBox "No password entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    
    ' Protect workbook
    On Error Resume Next
    ActiveWorkbook.Protect Password:=pwd, structure:=True, Windows:=False
    If Err.Number <> 0 Then
        errors = errors & "Workbook structure protection failed for '" & ActiveWorkbook.Name & _
                          "': [" & Err.Number & "] " & Err.Description & vbCrLf
        errorCount = errorCount + 1
        Err.Clear
    End If
    On Error GoTo 0

    ' Protect each worksheet with the given options
    Dim protectSucceeded As Boolean
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        ws.Protect _
            Password:=pwd, _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            AllowFormattingCells:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, _
            AllowDeletingColumns:=False, _
            AllowDeletingRows:=False, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=True

        If Err.Number <> 0 Then
            errors = errors & "Sheet protection failed - '" & ws.Name & "': [" & Err.Number & _
                              "] " & Err.Description & vbCrLf
            errorCount = errorCount + 1
            Err.Clear
        Else
            protectSucceeded = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
            If Not protectSucceeded Then
                errors = errors & "Sheet appears not protected after Protect - '" & ws.Name & "'" & vbCrLf
                errorCount = errorCount + 1
            End If
        End If
        On Error GoTo 0
    Next ws

    Application.ScreenUpdating = True

    ' Show result
    If errorCount = 0 Then
        MsgBox "Protected all sheets in '" & ActiveWorkbook.Name & "'.", vbInformation
    Else
        MsgBox "Completed with " & errorCount & " error(s):" & vbCrLf & vbCrLf & errors, vbExclamation
    End If
End Sub

Sub Unprotectsheet()
Attribute Unprotectsheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Unprotectsheet Macro
'

    Dim ws As Worksheet
    Dim pwd As String
    Dim errors As String
    Dim errorCount As Long

    pwd = InputBox("Enter the sheet password to remove from all worksheets:", "Unprotect Sheets")
    If Len(pwd) = 0 Then
        MsgBox "No password entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Try to unprotect workbook
    On Error Resume Next
    ActiveWorkbook.Unprotect Password:=pwd
    If Err.Number <> 0 Then
        errors = errors & "Workbook structure: [" & Err.Number & "] " & Err.Description & vbCrLf
        errorCount = errorCount + 1
        Err.Clear
    End If
    On Error GoTo 0

    ' Try to unprotect each sheet
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=pwd
        If Err.Number <> 0 Then
            errors = errors & "Sheet '" & ws.Name & "': [" & Err.Number & "] " & Err.Description & vbCrLf
            errorCount = errorCount + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next ws

    Application.ScreenUpdating = True

    ' Show result
    If errorCount = 0 Then
        MsgBox "Removed password from all sheets in '" & ActiveWorkbook.Name & "'.", vbInformation
    Else
        MsgBox "Completed with " & errorCount & " error(s):" & vbCrLf & vbCrLf & errors, vbExclamation
    End If
End Sub

