VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub FetchJSONContentPOST()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim url As String
    Dim jsonResponse As String
    Dim lastRow As Long
    Dim http As Object
    Dim status As Long
    Dim msg As String
    Dim startedAt As Date
    Dim finishedAt As Date
    Dim token As String
    Dim postBody As String
    Dim address As String
    Dim productId As Long
    
    token = ""
	'Colorado Department of revenue
    url = "https://api.ttr.services/v1/automation.rates.list"
	'Sales Tax    
productId = "401"
    
    On Error GoTo CleanFail
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set ws = ThisWorkbook.Sheets("Customers")
    
    ' Last row based on column O (address)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    If lastRow < 2 Then GoTo CleanExit
    
    Set rng = ws.Range("I76:I" & 77)
    
    startedAt = Now()
    
    For Each cell In rng
        address = Trim(cell.value)
        
        If Len(address) > 0 Then
            
            ' Build JSON body
            postBody = "{""address"":""" & Replace(address, """", "\""") & """,""productServiceId"":" & productId & "}"
            
            ' Create HTTP object
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
            If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
            
            If http Is Nothing Then
                ws.Cells(cell.Row, "J").value = "Error: No HTTP component available"
                GoTo NextCell
            End If
            
            ' POST request
            http.Open "POST", url, False
            http.setRequestHeader "Accept", "application/json"
            http.setRequestHeader "Content-Type", "application/json"
            http.setRequestHeader "Authorization", "Bearer " & token
            http.Send postBody
            
            status = http.status
            jsonResponse = http.responseText
            
            If status = 429 Then
                Application.Wait Now + TimeValue("00:00:01")
                http.Send postBody ' retry once
                status = http.status
                jsonResponse = http.responseText
            End If
                
            If status = 200 Then
                ws.Cells(cell.Row, "J").value = jsonResponse
            Else
            
                'msg = ExtractJsonMessage(jsonResponse)
                If Len(jsonResponse) > 0 Then
                    'ws.Cells(cell.Row, "J").value = msg
                    ws.Cells(cell.Row, "J").value = jsonResponse
                Else
                    ws.Cells(cell.Row, "J").value = "HTTP " & status & " - " & Left(jsonResponse, 256)
                End If
            End If
            ws.Cells(cell.Row, "K").value = " "
            Application.Wait Now + TimeValue("00:00:03") 'new rate limiter set to 3 seconds just to be safe. Cannot find documentation on new limits.
            
        End If
        
NextCell:
        Set http = Nothing
        DoEvents
    Next cell
    ws.Columns("J").WrapText = False
    
    finishedAt = Now()
    
    MsgBox "Tax Update Completed:" & vbCrLf & _
           FlexRow("Started  ", Format$(startedAt, "yyyy-mm-dd hh:nn:ss")) & vbCrLf & _
           FlexRow("Completed", Format$(finishedAt, "yyyy-mm-dd hh:nn:ss")) & vbCrLf & _
           FlexRow("Total Time", FormatElapsed(startedAt, finishedAt)), _
           vbInformation, "FetchJSONContent"

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:
    MsgBox "An error occurred: " & Err.Number & " - " & Err.Description, vbExclamation, "FetchJSONContent"
    Resume CleanExit
End Sub


' --- Helper: extract the first "message": "..." from a JSON error payload ---
Private Function ExtractJsonMessage(ByVal s As String) As String
    Dim p As Long, q As Long, startPos As Long, endPos As Long
    Dim keyPos As Long, quotePos As Long
    Dim msg As String
    
    ' Find the "message" key (case-insensitive)
    keyPos = InStr(1, s, Chr(34) & "message" & Chr(34), vbTextCompare)
    If keyPos = 0 Then
        ExtractJsonMessage = ""
        Exit Function
    End If
    
    ' Find the first colon after "message"
    p = InStr(keyPos, s, ":")
    If p = 0 Then
        ExtractJsonMessage = ""
        Exit Function
    End If
    
    ' Find the opening quote of the value
    quotePos = InStr(p + 1, s, Chr(34))
    If quotePos = 0 Then
        ExtractJsonMessage = ""
        Exit Function
    End If
    
    ' Find the closing quote of the value
    endPos = InStr(quotePos + 1, s, Chr(34))
    If endPos = 0 Then
        ExtractJsonMessage = ""
        Exit Function
    End If
    
    msg = Mid$(s, quotePos + 1, endPos - (quotePos + 1))
    ExtractJsonMessage = msg
End Function


Private Function FormatElapsed(ByVal startTime As Date, ByVal endTime As Date) As String
    Dim totalSeconds As Long
    totalSeconds = DateDiff("s", startTime, endTime)
    FormatElapsed = Format$(totalSeconds \ 3600, "00") & ":" & _
                    Format$((totalSeconds Mod 3600) \ 60, "00") & ":" & _
                    Format$(totalSeconds Mod 60, "00")
End Function


Private Function FlexRow(ByVal label As String, ByVal value As String, Optional ByVal width As Long = 44) As String
    ' Ensure we don't exceed total width. Truncate safely if too long.
    Dim lbl As String, val As String, gap As Long
    lbl = label
    val = value
    If Len(lbl) > width - 2 Then lbl = Left$(lbl, width - 2)
    If Len(val) > width - 2 Then val = Left$(val, width - 2)
    
    gap = width - (Len(lbl) + Len(val))
    If gap < 1 Then gap = 1
    
    FlexRow = lbl & Space$(gap) & val
End Function





