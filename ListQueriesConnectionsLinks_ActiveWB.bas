Attribute VB_Name = "Module1"
Sub ListQueriesConnectionsLinks_ActiveWB()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim rowNum As Long
    
    ' Use the workbook that is currently active (not the personal macro workbook)
    Set wb = ActiveWorkbook
    
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If
    
    ' Create or clear output sheet in the active workbook
    On Error Resume Next
    Set ws = wb.Sheets("Workbook_Inventory")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add
        ws.Name = "Workbook_Inventory"
    Else
        ws.Cells.Clear
    End If
    
    rowNum = 1
    ws.Cells(rowNum, 1).Value = "Type"
    ws.Cells(rowNum, 2).Value = "Name"
    ws.Cells(rowNum, 3).Value = "Details"
    ws.Cells(rowNum, 4).Value = "Last Refresh"
    ws.Cells(rowNum, 5).Value = "Source/Connection String"
    ws.Rows(1).Font.Bold = True
    
    ' --- List Power Query Queries ---
    On Error Resume Next
    If Not wb.Queries Is Nothing Then
        For i = 1 To wb.Queries.Count
            rowNum = rowNum + 1
            ws.Cells(rowNum, 1).Value = "Power Query"
            ws.Cells(rowNum, 2).Value = wb.Queries(i).Name
            ws.Cells(rowNum, 3).Value = "Power Query M Script"
            ws.Cells(rowNum, 5).Value = wb.Queries(i).Formula
        Next i
    End If
    On Error GoTo 0
    
    ' --- List Workbook Connections ---
    If wb.Connections.Count > 0 Then
        For i = 1 To wb.Connections.Count
            rowNum = rowNum + 1
            ws.Cells(rowNum, 1).Value = "Connection"
            ws.Cells(rowNum, 2).Value = wb.Connections(i).Name
            ws.Cells(rowNum, 3).Value = wb.Connections(i).Description
            
            ' Try to get last refresh date
            On Error Resume Next
            ws.Cells(rowNum, 4).Value = wb.Connections(i).ODBCConnection.RefreshDate
            If Err.Number <> 0 Then Err.Clear
            ws.Cells(rowNum, 4).Value = wb.Connections(i).OLEDBConnection.RefreshDate
            On Error GoTo 0
            
            ' Try to get connection string
            On Error Resume Next
            ws.Cells(rowNum, 5).Value = wb.Connections(i).ODBCConnection.Connection
            If Err.Number <> 0 Then Err.Clear
            ws.Cells(rowNum, 5).Value = wb.Connections(i).OLEDBConnection.Connection
            On Error GoTo 0
        Next i
    End If
    
    ' --- List External Links ---
    Dim linkArr As Variant
    linkArr = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(linkArr) Then
        For i = LBound(linkArr) To UBound(linkArr)
            rowNum = rowNum + 1
            ws.Cells(rowNum, 1).Value = "External Link"
            ws.Cells(rowNum, 2).Value = "Link " & i
            ws.Cells(rowNum, 5).Value = linkArr(i)
        Next i
    End If
    
    ' Autofit columns
    ws.Columns("A:E").AutoFit
    
    MsgBox "Workbook inventory created in '" & wb.Name & "' on sheet 'Workbook_Inventory'.", vbInformation
End Sub

