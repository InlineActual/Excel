
Sub UnclearedCheckReport()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim bottomRow As Long
    Dim dataStart As Long
    Dim helperCol As Long
    Dim startRow As Long
    Dim i As Long
    Dim leftCount As Long
    Dim rightCount As Long
    Dim bottomLeft As Long
    Dim bottomRight As Long
    Dim startingRight As Long
    Dim startingLeft As Long
    Dim sumFormula As Long
    
    'Headers[ A:"Post Date", B:"Check", C:"Description", D:"Bank Total", E:"Total", F:"Voucher", G:"Description", H:"Post Date"]
    'ws.Range("O2").Value = positivepayUnclearedTotal
        

    On Error GoTo CleanExit
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("Working Copy")
    ws.Cells.FormatConditions.Delete

    dataStart = 3                 ' first data row
    helperCol = 9                 ' column I for helper formulas

    ' --- Determine boundaries ---
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    bottomRow = lastRow * 2
    If bottomRow > ws.Rows.Count Then bottomRow = ws.Rows.Count

    ' --- Create a small buffer row with "blankzzz" to ensure sorts/formatting have a bottom anchor ---
    ws.Range("A" & bottomRow & ":H" & bottomRow).Value = "blankzzz"

    ' --- Apply UniqueValues conditional formatting to columns B and F ---
    With ws.Range("B:B,F:F")
        .FormatConditions.Delete
        .FormatConditions.AddUniqueValues
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlDuplicate
        With .FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With

    ' --- Sort E:H by cell color in column F (rows dataStart to bottomRow) ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("F" & dataStart & ":F" & bottomRow), _
            SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("E" & dataStart & ":H" & bottomRow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' --- Sort A:D by cell color in column B (rows dataStart to bottomRow) ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("B" & dataStart & ":B" & bottomRow), _
            SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A" & dataStart & ":D" & bottomRow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' --- sum formula in a safe, dynamic row ---
    sumFormula = lastRow - 2
    If bottomRow < ws.Rows.Count Then
        ws.Range("D" & sumFormula).Formula = "=SUM(D" & sumFormula & ":D" & bottomRow & ")"
        ws.Range("E" & sumFormula).Formula = "=SUM(E" & sumFormula & ":E" & bottomRow & ")"
    End If
    
    ' --- Sort A:D by values in column B ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("B" & lastRow & ":B" & bottomRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SetRange ws.Range("A" & lastRow & ":D" & bottomRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' --- Sort E:H by values in column F ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("F" & lastRow & ":F" & bottomRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("E" & lastRow & ":H" & bottomRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' --- Create helper column I with a comparison formula (D = E) ---
    ws.Range(ws.Cells(lastRow, helperCol), ws.Cells(bottomRow, helperCol)).FormulaR1C1 = "=RC[-5]=RC[-4]"

    ' --- Sort entire block A:I by helper column (TRUE/FALSE) so matching rows group together ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(lastRow, helperCol), ws.Cells(bottomRow, helperCol)), _
            Order:=xlAscending
        .SetRange ws.Range("A" & lastRow & ":I" & bottomRow)
        .Header = xlNo
        .Apply
    End With
    
    startRow = lastRow - 3
    helperCol = 9 ' column I

    ' --- Apply UniqueValues conditional formatting to D & E ---
    With ws.Range("D" & dataStart & ":E" & startRow)
        .FormatConditions.Delete
        .FormatConditions.AddUniqueValues
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlDuplicate
        With .FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    
    

    ' --- Final color-based sorts on E and D ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("E" & dataStart & ":E" & startRow), _
            SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("E" & dataStart & ":H" & startRow)
        .Header = xlGuess
        .Apply
    End With

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("D" & dataStart & ":D" & startRow), _
            SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A" & dataStart & ":D" & startRow)
        .Header = xlGuess
        .Apply
    End With
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    startRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    ws.Range("I" & startRow & ":I" & lastRow).ClearContents
    
    ' --- Replace "blankzzz" with empty string across the sheet ---
    ws.Cells.Replace What:="blankzzz", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    startingRight = FirstBlockEnd(ws, "E", 3)
    startingLeft = FirstBlockEnd(ws, "D", 3)

    If startingRight > startingLeft Then
        startRow = startingRight + 10
    Else
        startRow = startingLeft + 10
    End If

    ' --- Sort A:D by values in column B (numeric/text as numbers) for the data block ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("D" & startRow & ":D" & sumFormula - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SetRange ws.Range("A" & startRow & ":D" & sumFormula - 1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' --- Sort E:H by values in column E for the data block ---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("E" & startRow & ":E" & sumFormula - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("E" & startRow & ":H" & sumFormula - 1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    bottomRight = FirstBlockEnd(ws, "E", startRow)
    bottomLeft = FirstBlockEnd(ws, "D", startRow)
    If bottomRight > bottomLeft Then
        lastRow = bottomRight
    Else
        lastRow = bottomLeft
    End If
    
    ws.Range("I" & lastRow).Value = "End sort"
    
    

    leftCount = 0
    rightCount = 0
    
    ' --- Loop once and collect rows to move into arrays ---
    For i = startRow To lastRow
        ' Only act when there is a numeric or non-empty value in the key columns
        If Len(Trim(ws.Cells(i, "E").Value & "")) > 0 Or Len(Trim(ws.Cells(i, "D").Value & "")) > 0 Then
            If IsNumeric(ws.Cells(i, "E").Value) And IsNumeric(ws.Cells(i, "D").Value) Then
                If ws.Cells(i, 5).Value < ws.Cells(i, 4).Value Then
                   ' move E:H to right buffer area
                    rightCount = rightCount + 1
                    ws.Range("W" & rightCount & ":Z" & rightCount).Value = ws.Range("E" & i & ":H" & i).Value
                    ws.Range("E" & i & ":H" & i).ClearContents
                    With ws.Sort
                        .SortFields.Clear
                        .SortFields.Add2 Key:=ws.Range("E" & i & ":E" & sumFormula - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                        .SetRange ws.Range("E" & i & ":H" & sumFormula - 1)
                        .Header = xlNo
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                    i = i - 1

                ElseIf ws.Cells(i, 5).Value > ws.Cells(i, 4).Value Then
                    leftCount = leftCount + 1
                     ' move A:D to left buffer area
                    ws.Range("S" & leftCount & ":V" & leftCount).Value = ws.Range("A" & i & ":D" & i).Value
                    ws.Range("A" & i & ":D" & i).ClearContents
                    With ws.Sort
                        .SortFields.Clear
                        .SortFields.Add2 Key:=ws.Range("E" & i & ":E" & sumFormula - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                        .SetRange ws.Range("E" & i & ":H" & sumFormula - 1)
                        .Header = xlNo
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                     i = i - 1
                End If
            End If
        End If
    Next i
    
    
    bottomRight = FirstBlockEnd(ws, "E", startRow)
    ws.Range(ws.Cells(startRow, helperCol), ws.Cells(bottomRight, helperCol)).FormulaR1C1 = "=RC[-5]=RC[-4]"
    
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    ws.Range("A" & lastRow + 1 & ":I" & (lastRow + bottomRight - startRow + 1)).Value = ws.Range("A" & startRow & ":I" & bottomRight).Value
    
    ws.Range("A" & startRow & ":I" & bottomRight).ClearContents
    ws.Cells.FormatConditions.Delete
    
    startingRight = FirstBlockEnd(ws, "E", 3) + 1
    startingLeft = FirstBlockEnd(ws, "D", 3) + 1
    rightCount = FirstBlockEnd(ws, "W", 1)
    leftCount = FirstBlockEnd(ws, "S", 1)
    
    
    'insert data from S:V and W:Z back into their original columns and cleanup
    If rightCount > 0 Then
        ws.Range("E" & startingRight & ":H" & startingRight + rightCount - 1).Value = ws.Range("W1:Z" & rightCount).Value
        ws.Range("W1:Z" & rightCount).ClearContents
    End If
    If leftCount > 0 Then
        ws.Range("A" & startingLeft & ":D" & startingLeft + leftCount - 1).Value = ws.Range("S1:V" & leftCount).Value
        ws.Range("S1:V" & leftCount).ClearContents
    End If
    
    startingRight = FirstBlockEnd(ws, "E", 3)
    startingLeft = FirstBlockEnd(ws, "D", 3)
    If startingRight > startingLeft Then
        startRow = startingRight + 10
    Else
        startRow = startingLeft + 10
    End If
    ws.Range("A" & startRow & ":A" & sumFormula - 1).EntireRow.Delete
    
    
    If startingRight < 3 Then startingRight = 3
    ws.Range("E2").FormulaR1C1 = "=SUMIF(R3C6:R" & startingRight & "C6,""<999999"",R3C5:R" & startingRight & "C5)"
    
    ws.Range("F2").Value = "Not Cleared"
    ws.Range("O3").FormulaR1C1 = "=R[-1]C+R[-1]C[-10]"
    With ws.Range("E2:F2").Interior
        .Pattern = xlSolid
        .Color = 13421619
    End With

    ' --- Sort E:H by values in column F---
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("F3:F" & startingRight), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("E3:H" & startingRight)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    With ws.Range("F3:F" & startingRight)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=1", Formula2:="=999999"
        With .FormatConditions(.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13421619
        End With
    End With
   
    ' --- Apply number format to columns D:E (accounting style) ---
    ws.Columns("D:E").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    ' --- Return view to top-left cell A1 ---
    Application.Goto ws.Range("A1"), True

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Check Report encountered an error: " & Err.Description, vbExclamation
    End If
End Sub




Function FirstBlockEnd(ws As Worksheet, colLetter As String, startRow As Long) As Long
' Returns the end row of the first contiguous non-blank block that begins at startRow.
' If startRow is blank, it searches down to find the first non-blank and then returns that block's end.
    Dim r As Long, lastRow As Long
    Dim inBlock As Boolean
    Dim blockStart As Long

    ' Determine a safe last row to iterate to
    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
    If lastRow < startRow Then
        FirstBlockEnd = startRow - 1
        Exit Function
    End If

    inBlock = False
    For r = startRow To lastRow
        If Len(Trim(ws.Cells(r, colLetter).Value2 & "")) > 0 Then
            If Not inBlock Then
                blockStart = r
                inBlock = True
            End If
        Else
            If inBlock Then
                FirstBlockEnd = r - 1
                Exit Function
            End If
        End If
    Next r

    ' If we reached the end while still in a block, return lastRow
    If inBlock Then
        FirstBlockEnd = lastRow
    Else
        FirstBlockEnd = startRow - 1 ' no block found
    End If
End Function





