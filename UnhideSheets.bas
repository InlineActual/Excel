Attribute VB_Name = "Module18"
Sub UnhideSheets()
Attribute UnhideSheets.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' UnhideSheets Macro
' Keyboard Shortcut: Ctrl+Shift+T
'
    Dim ws As Worksheet
    Dim rng As Range
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    
    For Each ws In Worksheets
        With ws
            ' Clear filters if present
            If .AutoFilterMode Then
                On Error Resume Next
                .ShowAllData
                On Error GoTo 0
            End If
            
            ' Unhide all rows and columns
            .Rows.Hidden = False
            .Columns.Hidden = False
            
            ' Format and autofit only used range
            Set rng = .UsedRange
            With rng
                .WrapText = False
                .Font.Name = "Calibri"
                .Font.Size = 11
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .EntireColumn.AutoFit
                .EntireRow.AutoFit
            End With
        End With
    Next ws
End Sub
