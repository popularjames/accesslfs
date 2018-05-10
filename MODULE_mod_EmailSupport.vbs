Option Compare Database
Option Explicit



Private Const ClassName As String = "mod_EmailSupport"



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function KDSendSQLMail() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO


    strProcName = ClassName & ".SendSQLMail"
    Set oAdo = New clsADO
    
    With oAdo
        .ConnectionString = GetConnectString("V_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "__SendNotification"
        .Parameters.Refresh
        .Parameters("@p_strTO") = "kevin.dearing@connolly.com;kdearing01@comcast.net"
        .Parameters("@p_strSubject") = "Test email"
        .Parameters("@p_strMB") = "This, I'm guessing is the message body"
        .Execute
    End With



Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This has the problems with Outlook security..
''' user will be prompted like 2 times
'''
Public Function SendOutlookEmail(sRecipients As String, sMsgBody As String, sSubject As String, _
        Optional sAttachPath As String) As Boolean
On Error GoTo Block_Exit
Dim strProcName As String
'    Dim oApp As Outlook.Application
'    Dim oInbox As Outlook.Folder
'    Dim oNpsace As Outlook.Namespace
'    Dim oMail As Outlook.MailItem
'    Dim oRecip As Outlook.Recipient
'    Dim oAttach As Outlook.Attachment

Dim oApp As Object
Dim oInbox As Object
Dim oNpsace As Object
Dim oMail As Object
Dim oRecip As Object
Dim oAttach As Object

Dim sAtchPath As String
Dim saryAttachments() As String
Dim saryRecips() As String
Dim iIdx As Integer

    strProcName = ClassName & ".SendOutlookEmail"
    
    If MsgBox(Application.Name & " is going to prepare an email for you. Please grant permission", vbOKCancel, "Preparing email") = vbCancel Then
        LogMessage strProcName, , "User canceled preparing the email"
        SendOutlookEmail = True
        GoTo Block_Exit
    End If
    

    '' If outlook is already opened, then we don't have much to do, otherwise:
    If IsAppRunning(, , "OUTLOOK") = True Then
        Set oApp = GetObject(, "Outlook.Application")
    Else
        Set oApp = CreateObject("Outlook.Application")
        Sleep 1000  ' I hate building this stuff in but I mean, kind of have to to make sure the above finished ..
        Set oNpsace = oApp.GetNamespace("MAPI")
        Set oInbox = oNpsace.GetDefaultFolder(6)    ' olFolderInbox = 6
        oInbox.Display  '' makes it visible
    End If
    
    
    saryRecips = Split(sRecipients, ",")
    
    
    Set oMail = oApp.CreateItem(0)  '' olMailItem = 0
        
    With oMail
        For iIdx = 0 To UBound(saryRecips)
            Set oRecip = .Recipients.Add(saryRecips(iIdx))
            oRecip.Type = 1 ' olTo = 1
        Next
        
        .Subject = sSubject
        .Body = sMsgBody
        
        .Importance = 1 ' olImportanceNormal = 1
        
            ' Resolve each Recipient's name.
        For Each oRecip In .Recipients
            oRecip.Resolve
        Next

        ' Attachment..
        If sAttachPath <> "" Then
            If InStr(1, sAttachPath, ",", vbTextCompare) > 0 Then
                saryAttachments = Split(sAttachPath, ",")
                
                For iIdx = 0 To UBound(saryAttachments)
                    sAtchPath = saryAttachments(iIdx)
                    If FileExists(sAtchPath) Then
                        Set oAttach = .Attachments.Add(sAtchPath)
                    End If
                Next
            Else
                If FileExists(sAttachPath) Then
                    Set oAttach = .Attachments.Add(sAttachPath)
                End If
            End If
            
            If .Attachments.Count < 1 Then
                LogMessage strProcName, "ERROR", "Seems there was a problem attaching file(s).. This may need to be done manually", sAttachPath, True
            End If
        End If
    
            ' not going to send it, just display it
        oMail.Display False
        oMail.Save      ' save it in drafts, just in case..
    End With
    
    '' Only question I have is, should I make sure the mail item is the topmost window or just leave it be?
    
    SendOutlookEmail = True
        '' now just activate it
    ActivateApplicationWindow , , "OUTLOOK"
    
Block_Exit:
    Set oApp = Nothing
    Set oInbox = Nothing
    Set oNpsace = Nothing
'    Set oMail = Nothing
    Set oRecip = Nothing
    Set oAttach = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function