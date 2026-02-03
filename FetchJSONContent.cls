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

Sub FetchJSONContent()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim url As String
    Dim jsonResponse As String
    Dim lastRow As Long
    Dim http As Object ' WinHttpRequest preferred; fallback to XMLHTTP
    Dim status As Long
    Dim msg As String
    Dim startedAt As Date
    Dim finishedAt As Date
    
    On Error GoTo CleanFail
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("Conversion")
    
    ' Find the last used row in column V (URLs)
    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row
    
    If lastRow < 3 Then GoTo CleanExit
    
    Set rng = ws.Range("V3:V" & lastRow)
    startedAt = Now()
    For Each cell In rng
        If Len(Trim(cell.value)) > 0 Then
            url = CStr(cell.value)
            
            ' Create HTTP object
            On Error Resume Next
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
            If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
            On Error GoTo CleanFail
            
            If http Is Nothing Then
                cell.Offset(0, 1).value = "Error: No HTTP component available"
                GoTo NextCell
            End If
            
            ' Send request (synchronous)
            If TypeName(http) = "IWinHttpRequest" Or TypeName(http) = "WinHttpRequest" Then
                On Error Resume Next
                http.SetTimeouts 5000, 5000, 10000, 15000
                On Error GoTo CleanFail
                http.Open "GET", url, False
                http.setRequestHeader "Accept", "application/json"
                http.Send
            Else
                http.Open "GET", url, False
                On Error Resume Next
                http.setRequestHeader "Accept", "application/json"
                On Error GoTo CleanFail
                http.Send
            End If
            
            ' Check status
            On Error Resume Next
            status = 0
            status = http.status
            jsonResponse = http.responseText
            On Error GoTo CleanFail
            
            If status = 200 Then
                ' Success: write raw JSON to column W
                cell.Offset(0, 1).value = jsonResponse
                cell.Offset(0, 1).WrapText = False
            Else
                ' Non-200: try to extract "message" from JSON error body
                msg = ExtractJsonMessage(jsonResponse)
                If Len(msg) > 0 Then
                    ' Just the human-readable error message
                    cell.Offset(0, 1).value = msg
                Else
                    ' Fallback: show status and a short snippet of the response
                    cell.Offset(0, 1).value = "HTTP " & status & " - " & Left(Replace(jsonResponse, vbCrLf, " "), 256)
                End If
                cell.Offset(0, 1).WrapText = False
            End If
        End If
NextCell:
        Set http = Nothing
        DoEvents
    Next cell
    finishedAt = Now()
    

Dim msgboxMessage As String
msgboxMessage = "Tax Update Completed:" & vbCrLf & _
      FlexRow("Started  ", Format$(startedAt, "yyyy-mm-dd hh:nn:ss")) & vbCrLf & _
      FlexRow("Completed", Format$(finishedAt, "yyyy-mm-dd hh:nn:ss")) & vbCrLf & _
      FlexRow("Total Time", FormatElapsed(startedAt, finishedAt))

MsgBox msgboxMessage, vbInformation, "FetchJSONContent"



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

