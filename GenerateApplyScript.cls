VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GenerateGPApplyScript()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim output As String
    Dim header As String, batch As String
    Dim rowData As Variant
    Dim i As Integer
    Dim lastRow As Long
    
    ' Set worksheet and range
    Set ws = ThisWorkbook.Worksheets("Apply")
    ws.Columns("S").ClearContents
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    Set rng = ws.Range("M1:Q" & lastRow)

    'Create the apply transactions
    ' Loop through each row
    For i = 2 To lastRow Step 1
        rowData = rng.Rows(i).Value
            output = "ActivateWindow dictionary 'default'  form 'PM_Apply_To_Maintenance' window 'PM_Apply_Document'" & vbCrLf
            output = output & "  MoveTo field 'Document Number' " & vbCrLf
            output = output & "  TypeTo field 'Document Number' , '" & rowData(1, 1) & "'" & vbCrLf
            output = output & "  CommandExec dictionary 'default'  form 'PM_Apply_To_Maintenance' command 'Find Button_w_PM_Apply_Document_f_PM_Apply_To_Maintenance'" & vbCrLf
            output = output & "NewActiveWin dictionary 'default'  form 'PM_Apply_To_Maintenance' window 'PM_Document_Apply_Find'" & vbCrLf
            output = output & "  TypeTo field 'Document Number' , '" & rowData(1, 4) & "'" & vbCrLf
            output = output & "  MoveTo field 'Find Button NONE'" & vbCrLf
            output = output & "  ClickHit field 'Find Button NONE' " & vbCrLf
            output = output & "NewActiveWin dictionary 'default'  form 'PM_Apply_To_Maintenance' window 'PM_Apply_Document'" & vbCrLf
            output = output & "ActivateWindow dictionary 'default'  form 'PM_Apply_To_Maintenance' window 'PM_Apply_Document' " & vbCrLf
            output = output & " # Key 2: '1', '" & rowData(1, 5) & "'" & vbCrLf
            output = output & "  MoveTo line 1 scrollwin 'PM_Applied_Document_Scroll' field '(L) Applied Amount'" & vbCrLf
            output = output & " # Key 2: '1', '" & rowData(1, 5) & "'" & vbCrLf
            output = output & "  TypeTo line 1 scrollwin 'PM_Applied_Document_Scroll' field '(L) Applied Amount' , '" & rowData(1, 3) & "'" & vbCrLf
            output = output & " # Key 2: '1', '" & rowData(1, 5) & "'" & vbCrLf
            output = output & "  MoveTo line 1 scrollwin 'PM_Applied_Document_Scroll' field 'Discount Taken Amount'"
            output = output & " # Key 2: '1', '" & rowData(1, 5) & "'" & vbCrLf
            output = output & "  TypeTo line 1 scrollwin 'PM_Applied_Document_Scroll' field 'Discount Taken Amount' , '0'" & vbCrLf
            output = output & " # Key 2: '1', '" & rowData(1, 5) & "'" & vbCrLf
            output = output & "  MoveTo line 1 scrollwin 'PM_Applied_Document_Scroll' field '(L) Applied Amount'" & vbCrLf
        ws.Activate
        ws.Cells(i, 19).Value = output
    Next i

    'create the header
    header = "# DEXVERSION=14.00.0085.000 2 2" & vbCrLf
    header = header & "ClickHitIE , 'http://dexterity/product=0/script=ExecuteAreaPageLink/args=0,1566,26/' # http://dexterity will be pre-pended" & vbCrLf
    header = header & "ActivateWindow dictionary 'default'  form 'PM_Apply_To_Maintenance' window 'PM_Apply_Document'" & vbCrLf
    header = header & "MoveTo field 'Document Type' item 3  # 'Payment' " & vbCrLf
    header = header & "ClickHit field 'Document Type' item 2  # 'Credit Memo'" & vbCrLf
    ws.Range("S1").Value = header
    ws.Columns("S").WrapText = False

End Sub




