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

' Helper: get a reliable SMTP address for Exchange senders
Private Function GetSenderSmtpAddress(ByVal mi As Object) As String
    On Error GoTo CleanFail
    
    Dim addr As String
    addr = mi.SenderEmailAddress
    
    If LCase$(mi.SenderEmailType) = "smtp" And Len(addr) > 0 Then
        GetSenderSmtpAddress = addr
        Exit Function
    End If
    
    Dim ae As Object, exUser As Object
    Set ae = mi.Sender
    If Not ae Is Nothing Then
        On Error Resume Next
        Set exUser = ae.GetExchangeUser
        If Not exUser Is Nothing Then
            addr = exUser.PrimarySmtpAddress
            If Len(addr) > 0 Then
                GetSenderSmtpAddress = addr
                Exit Function
            End If
        End If
        On Error GoTo CleanFail
    End If
    
CleanFail:
    If Len(addr) = 0 Then addr = mi.SenderEmailAddress
    GetSenderSmtpAddress = addr
End Function

' Helper: get Inbox for a specific Store display name
Private Function GetInboxByStoreName(ByVal app As Object, ByVal storeName As String) As Object
    Dim st As Object
    For Each st In app.Session.Stores
        If LCase$(st.DisplayName) = LCase$(storeName) Then
            Set GetInboxByStoreName = st.GetDefaultFolder(6) ' 6 = olFolderInbox
            Exit Function
        End If
    Next st
    Set GetInboxByStoreName = Nothing
End Function

Public Sub RetrieveUnreadInvoices_Backwards_SavePDFs()
    Dim outlookApp As Object, outlookNamespace As Object
    Dim inbox As Object, olItems As Object, unreadItems As Object
    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, rowIndex As Long
    Dim mi As Object
    Dim senderSmtp As String, subj As String, recvTime As Variant
    Dim saveFolder As String


    
    ' ==== Workbook & sheet ====
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Emails")
    saveFolder = ws.Range("D11").Value
    If Right$(saveFolder, 1) <> "\" Then saveFolder = saveFolder & "\"
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With ws
        .Range("D7").Value = "Start: " & Now
        .Range("D8").Value = "Running"
        .Range("I1:L1").Value = Array("Subject", "Received", "Count", "New Name")
        .Range("I2:J2000").ClearContents
        .Range("B20:D30").ClearContents
    End With
    
    ' ==== Outlook ====
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    
    
    Dim targetInbox As Object
    Dim targetInboxName As String
    targetInboxName = ws.Range("D12").Value
    Set targetInbox = GetInboxByStoreName(outlookApp, targetInboxName)
    If targetInbox Is Nothing Then
        MsgBox "Inbox for 'Wesco Group, LLC - Tax' not found. Check the store name.", vbCritical
        GoTo WrapUp
    End If
    Set inbox = targetInbox
    
    ' Now work within that Inbox
    Set olItems = inbox.Items
    On Error Resume Next
    olItems.Sort "[ReceivedTime]", True
    Set unreadItems = olItems.Restrict("[Unread] = True")
    On Error GoTo 0
    
    rowIndex = 2
    
    ' Iterate backwards to avoid skipping when we mark items as read
    For i = unreadItems.Count To 1 Step -1
        Set mi = unreadItems(i)
        If Not mi Is Nothing Then
            If mi.Class = 43 Then ' olMailItem
                ' Read properties
                On Error Resume Next
                subj = mi.Subject
                senderSmtp = GetSenderSmtpAddress(mi)
                recvTime = mi.ReceivedTime
                On Error GoTo 0
                
                ' --- Filters (case-insensitive) ---
                If InStr(LCase$(senderSmtp), "email@email.com") = 0 Then GoTo SkipThis
                If InStr(LCase$(subj), "invoice for") = 0 And InStr(LCase$(subj), "credit memo for") = 0 Then GoTo SkipThis
                
                ' --- Write base info ---
                ws.Cells(rowIndex, "I").Value = subj
                ws.Cells(rowIndex, "J").Value = recvTime
             
                ' === Save PDF attachments using the name from column L (New Name), with -counter for multiples ===
                Dim att As Object
                Dim origName As String
                Dim desiredStem As String
                Dim newName As String
                Dim newPath As String
                Dim savedNames As String
                Dim savedCounters As String
                Dim counter As Long
                
                savedNames = ""
                savedCounters = ""
                counter = 1   ' per-email counter starts at 1
                
                ' READ the desired name from column L for this row (your "New Name")
                desiredStem = Trim$(CStr(ws.Cells(rowIndex, "L").Value))
                
                ' Fallback if L is empty: use a generic name
                If Len(desiredStem) = 0 Then
                    desiredStem = "Attachment-" & CStr(rowIndex - 1)
                End If
                
                ' Sanitize desiredStem (remove invalid Windows filename characters)
                desiredStem = Replace(desiredStem, ":", "-")
                desiredStem = Replace(desiredStem, "\", "-")
                desiredStem = Replace(desiredStem, "/", "-")
                desiredStem = Replace(desiredStem, "*", "-")
                desiredStem = Replace(desiredStem, "?", "-")
                desiredStem = Replace(desiredStem, Chr(34), "-") ' "
                desiredStem = Replace(desiredStem, "<", "-")
                desiredStem = Replace(desiredStem, ">", "-")
                desiredStem = Replace(desiredStem, "|", "-")
                
                If mi.Attachments.Count > 0 Then
                    For Each att In mi.Attachments
                        origName = att.fileName
                        
                        ' Only handle PDFs
                        If LCase$(Right$(origName, 4)) = ".pdf" Then
 
                            If counter = 1 Then
                                newName = desiredStem & ".pdf"
                            Else
                                newName = desiredStem & "-" & CStr(counter) & ".pdf"
                            End If
                            
                            newPath = saveFolder & newName
                            
                            ' Ensure unique file name: if exists, bump suffix until unique
                            Do While Dir$(newPath) <> ""
                                counter = counter + 1
                                newName = desiredStem & "-" & CStr(counter) & ".pdf"
                                newPath = saveFolder & newName
                            Loop
                            
                            ' Save attachment
                            On Error Resume Next
                            att.SaveAsFile newPath
                            If Err.Number = 0 Then
                                ' Track names and counters (comma-separated, aligned)
                                If Len(savedNames) > 0 Then savedNames = savedNames & ", "
                                savedNames = savedNames & newName
                                
                                If Len(savedCounters) > 0 Then savedCounters = savedCounters & ", "
                                savedCounters = savedCounters & CStr(counter)
                                
                                ' Next attachment counter
                                counter = counter + 1
                            Else
                                If Len(savedNames) > 0 Then savedNames = savedNames & ", "
                                savedNames = savedNames & "[Save failed: " & origName & "]"
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                    Next att
                End If
                
                ' Mark as read AFTER writing
                On Error Resume Next
                mi.UnRead = False
                mi.Save
                On Error GoTo 0
                
                rowIndex = rowIndex + 1
            End If
        End If
SkipThis:
    Next i
    
WrapUp:
    ws.Range("D8").Value = "Completed: " & Now
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Processed unread items in Inbox; PDFs saved per New Name.", vbInformation
End Sub

Sub ListOutlookAccountsAndStores()
    Dim app As Object, ses As Object, acct As Object, st As Object
    Dim wb As Workbook, ws As Worksheet
    Set app = CreateObject("Outlook.Application")
    Set ses = app.Session
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Emails")
    Dim rowIndex As Long
    
    rowIndex = 20
    
    For Each acct In ses.Accounts
        ws.Range("B" & rowIndex).Value = "Accounts:"
        ws.Range("D" & rowIndex).Value = acct.DisplayName
        rowIndex = rowIndex + 1
    Next acct
    
    Debug.Print "Stores:"
    For Each st In ses.Stores
        ws.Range("B" & rowIndex).Value = "Stores:"
        ws.Range("D" & rowIndex).Value = st.DisplayName
        rowIndex = rowIndex + 1
    Next st
    MsgBox "All Email accounts listed", vbInformation
End Sub




