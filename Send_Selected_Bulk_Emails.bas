Attribute VB_Name = "Module1"
Sub SendSelectedBulkEmails()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim selectedRange As Range
    Dim cell As Range
    Dim rowNumber As Long
    Dim signatureBody As String
    Dim emailCount As Long
    Dim response As VbMsgBoxResult
    
    ' Work with the active sheet
    Set ws = ActiveSheet
    
    ' Initialize Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    If OutlookApp Is Nothing Then
        MsgBox "Outlook could not be started.", vbExclamation
        Exit Sub
    End If
    
    ' Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the rows you want to send emails for.", vbExclamation
        Exit Sub
    End If
    
    Set selectedRange = Selection
    
    ' Count how many emails are about to be sent (excluding header row)
    emailCount = 0
    For Each cell In selectedRange.Columns(1).Cells
        If cell.Row > 1 Then ' Only process if NOT in header row
            If ws.Cells(cell.Row, 3).Value <> "" Then ' Check if EMAIL field is not empty
                emailCount = emailCount + 1
            End If
        End If
    Next cell
    
    If emailCount = 0 Then
        MsgBox "No valid emails found in the selected range.", vbExclamation
        Exit Sub
    End If
    
    ' Ask for confirmation
    response = MsgBox("You are about to send " & emailCount & " emails." & vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion, "Confirm Sending Emails")
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Loop through each selected row
    For Each cell In selectedRange.Columns(1).Cells
        rowNumber = cell.Row
        
        ' Skip header row
        If rowNumber > 1 Then
            ' Only if EMAIL is not empty
            If ws.Cells(rowNumber, 3).Value <> "" Then
                Set OutlookMail = OutlookApp.CreateItem(0) ' 0 = olMailItem
                
                With OutlookMail
                    .Display ' Load signature first
                    signatureBody = .HTMLBody ' Grab default signature
                    
                    .To = ws.Cells(rowNumber, 3).Value ' EMAIL (column C)
                    .CC = ws.Cells(rowNumber, 4).Value ' CC (column D)
                    .BCC = ws.Cells(rowNumber, 5).Value ' BCC (column E)
                    .Subject = ws.Cells(rowNumber, 6).Value ' SUBJECT (column F)
                    
                    ' Only signature (no body text)
                    .HTMLBody = signatureBody
                    
                    ' Add attachment if available
                    If ws.Cells(rowNumber, 7).Value <> "" Then
                        If Dir(ws.Cells(rowNumber, 7).Value) <> "" Then
                            .Attachments.Add ws.Cells(rowNumber, 7).Value
                        Else
                            MsgBox "Attachment not found for " & ws.Cells(rowNumber, 3).Value, vbExclamation
                        End If
                    End If
                    
                    .Send ' Automatically send
                End With
                
                Set OutlookMail = Nothing
            End If
        End If
    Next cell
    
    Set OutlookApp = Nothing
    
    MsgBox "All selected emails have been sent successfully!", vbInformation
End Sub

