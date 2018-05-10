Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' DLC 05/20/2010
'
' Description: Email Class to allow emails to be sent from Decipher
'
' Dependencies: Identity object
'
' Sample Usage:
'
' Dim oMsg As New CT_clsEmail
' With oMsg
'     .ErrorHandler = RaiseError
'     .RecipientTo = "abc@xyz.com"
'     .Subject = "test"
'     .Body = "TEST BODY"
'     .AttachmentAdd "C:\1b.txt"
'     .AttachmentAdd "C:\1kb.txt2"
'     .Send
' End With
' Set oMsg = Nothing
'
' DLC 08/23/10
'
'    Updated public enums names (prefixed with email)
'    Added enum for Body Type (test / HTML)
'    Updated the way that HTML or Text body is handled
'
Private Const FROM_DOMAIN As String = "connolly.com"
Private Const DEFAULT_MAIL_SERVER As String = "audit.smtp.ccaintranet.net"
Private Const CDO_SCHEMA As String = "http://schemas.microsoft.com/cdo/configuration/"

'***** ENUMS & TYPES *****
Public Enum emailSendUsing
    PickUp = 1
    Port = 2
    Exchange = 3
End Enum

Public Enum emailAuthType
    Anonymous = 0
    Basic = 1
    NTLM = 2
End Enum

Public Enum emailBodyType
    Text = 0
    HTML = 1
End Enum

'This is to allows clsIcon to be used in batch mode with the errors raised to the caller
Public Enum emailErrorHandling
    SuppressError = 0
    DisplayError = 1
    RaiseError = 2
End Enum

Private Type Msg
    From As String
    FromDisplayName As String
    To As String
    CC As String
    BCC As String
    Subject As String
    Body As String
    BodyType As emailBodyType
'   DLC 05/20/10 Removed Attachment as this is now handled though mvAttachments collection
'   Attachment As String
    ReplyTo As String
End Type

'***** PRIVATE VARIABLES *****
Private mvErrorHandling As emailErrorHandling
Private mvAuthType As emailAuthType
Private mvMSG As Msg
Private mvServer As String
Private mvTimeout As Integer
Private mvSendUsing As Byte
' K.Tanacea 12/2/2009 - Allow multiple attachments
Private mvAttachments As New Collection
Private mvAttachmentsSize As Double

'***** PROPERTIES *****
Public Property Get AuthenticationType() As emailAuthType
    AuthenticationType = mvAuthType
End Property
Public Property Let AuthenticationType(ByVal AT As emailAuthType)
    mvAuthType = AT
End Property
Public Property Get Server() As String
    Server = mvServer
End Property
Public Property Let Server(ByVal Value As String)
    mvServer = Value
End Property
Public Property Get Subject() As String
    Subject = mvMSG.Subject
End Property
Public Property Let Subject(ByVal Value As String)
     mvMSG.Subject = Value
End Property
Public Property Get Body() As String
    Body = mvMSG.Body
End Property
Public Property Let Body(ByVal Value As String)
     mvMSG.Body = Value
End Property
Public Property Get BodyType() As emailBodyType
    BodyType = mvMSG.BodyType
End Property
Public Property Let BodyType(ByVal Value As emailBodyType)
     mvMSG.BodyType = Value
End Property
Public Property Get From() As String
     From = mvMSG.From
End Property
'   DPS 08/20/10 added and implemented as override to From
Public Property Get SendOnBehalf() As String
     SendOnBehalf = mvMSG.From
End Property
Public Property Let SendOnBehalf(ByVal Value As String)
    mvMSG.From = Value
End Property
Public Property Let ErrorHandler(ByVal Value As emailErrorHandling)
    mvErrorHandling = Value
End Property
Public Property Get ErrorHandler() As emailErrorHandling
    ErrorHandler = mvErrorHandling
End Property
Public Property Get ReplyTo() As String
     ReplyTo = mvMSG.ReplyTo
End Property
Public Property Let ReplyTo(ByVal Value As String)
    mvMSG.ReplyTo = Value
End Property
Public Property Get FromDisplayName() As String
     FromDisplayName = mvMSG.FromDisplayName
End Property
Public Property Let FromDisplayName(ByVal Value As String)
    mvMSG.FromDisplayName = Value
End Property
Public Property Get RecipientTo() As String
     RecipientTo = mvMSG.To
End Property
Public Property Let RecipientTo(ByVal Value As String)
    mvMSG.To = Value
End Property
Public Property Get RecipientCC() As String
     RecipientCC = mvMSG.CC
End Property
Public Property Let RecipientCC(ByVal Value As String)
    mvMSG.CC = Value
End Property
Public Property Get RecipientBCC() As String
     RecipientBCC = mvMSG.BCC
End Property
Public Property Let RecipientBCC(ByVal Value As String)
    mvMSG.BCC = Value
End Property
Public Property Get AttachmentsCount()
    AttachmentsCount = mvAttachments.Count
End Property
Public Property Get AttachmentName(Optional ByVal AttachId As Integer = 1)
    If AttachId > 0 And AttachId <= mvAttachments.Count Then
        AttachmentName = mvAttachments.Item(AttachId)
    Else
        AttachmentName = ""
    End If
End Property
Public Property Get AttachmentsTotSize()
    'Returns the total attachment size in MB
    AttachmentsTotSize = Round(mvAttachmentsSize / 1048576, 2)
End Property
Public Property Get timeout() As Integer
     timeout = mvTimeout
End Property
Public Property Let timeout(ByVal Value As Integer)
    mvTimeout = Value
End Property

Public Sub AttachmentAdd(ByVal AttachmentFileName As String)
On Error GoTo ErrorHandler
    Dim fso As Object
    Dim fs As Object
    If Trim(AttachmentFileName) <> vbNullString Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(AttachmentFileName) = True Then
            mvAttachments.Add AttachmentFileName
            Set fs = fso.GetFile(AttachmentFileName)
            mvAttachmentsSize = mvAttachmentsSize + fs.Size
        Else
            'If this raises the error rather than displaying or suppressing it, the error will be caught by the ErrorHandler below which
            'will detect the err.source as "AttachmentAdd" and rethrow it.
            DisplayOrRaiseError 9001, "AttachmentAdd", "File Attachment does not exist (" & AttachmentFileName & ")"
        End If
    End If
Exit_ErrorHandler:
    On Error Resume Next
    Set fs = Nothing
    Set fso = Nothing
    Exit Sub
ErrorHandler:
    'If the error was thrown within this method, rethrow it
    If Err.Source = "AttachmentAdd" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        DisplayOrRaiseError 9002, "AttachmentAdd", Err.Description & " (" & Err.Number & ")"
    End If
    Resume Exit_ErrorHandler
End Sub

Public Sub Send()
    On Error GoTo ErrorHappened
    Dim oMessage As Object
    Dim thisAttachment As Variant
    Set oMessage = CreateObject("CDO.Message")
    With oMessage
        .Subject = mvMSG.Subject
        
        '-- MM - Added to change display text
        If mvMSG.FromDisplayName <> vbNullString Then
            .From = mvMSG.FromDisplayName & " <" + mvMSG.From & "> "
        Else
            .From = mvMSG.From
        End If
        
        '-- MM Added for redirection
        .ReplyTo = mvMSG.ReplyTo
        'DLC 05/24/10 Add the .HTMLBody if set, otherwise use .Body
        If Nz(mvMSG.BodyType, emailBodyType.Text) = emailBodyType.HTML Then
            .HTMLBody = mvMSG.Body
        Else
            .TextBody = mvMSG.Body
        End If
        .To = mvMSG.To
        .CC = mvMSG.CC
        .BCC = mvMSG.BCC
        For Each thisAttachment In mvAttachments
            .AddAttachment thisAttachment
        Next
        .Keywords = Identity.UserName & "@" & FROM_DOMAIN ' Capture the actual sender
        'Set Mail Server
        .CONFIGURATION.Fields.Item(CDO_SCHEMA & "smtpserver") = mvServer
        'Type of authentication, NONE, Basic (Base64 encoded), NTLM
        .CONFIGURATION.Fields.Item(CDO_SCHEMA & "smtpauthenticate") = mvAuthType
        .CONFIGURATION.Fields.Item(CDO_SCHEMA & "sendusing") = mvSendUsing
        'Use SSL for the connection (False or True)
        .CONFIGURATION.Fields.Item(CDO_SCHEMA & "smtpusessl") = False
        .CONFIGURATION.Fields.Item(CDO_SCHEMA & "smtpconnectiontimeout") = mvTimeout
        .CONFIGURATION.Fields.Update
        .Send
    End With
ExitNow:
    On Error Resume Next
    Set oMessage = Nothing
    Exit Sub
ErrorHappened:
    On Error GoTo ExitNow
    DisplayOrRaiseError 9003, "Send Mail", "Error Sending Email: " & Err.Description & " (" & Err.Number & ")"
    Resume ExitNow
End Sub

Private Sub Class_Initialize()
    '**** INIT TO DEFAULTS ****
    mvAuthType = emailAuthType.NTLM
    mvSendUsing = emailSendUsing.Port
    mvServer = DEFAULT_MAIL_SERVER
    mvTimeout = 20
    mvMSG.From = Identity.UserName & "@" & FROM_DOMAIN
    mvAttachmentsSize = 0
    'Display Error Messages by default
    mvErrorHandling = emailErrorHandling.DisplayError
End Sub

Private Sub DisplayOrRaiseError(ByVal ErrorNumber As Integer, ByVal Source As String, ByVal ErrorMessage As String)
    Select Case mvErrorHandling
        Case emailErrorHandling.DisplayError
            MsgBox ErrorMessage, vbCritical + vbOKOnly, "Error in " & Source
        Case emailErrorHandling.RaiseError
            Err.Raise ErrorNumber, Source, ErrorMessage
        Case emailErrorHandling.SuppressError
            'Ignore the error
    End Select
End Sub

Public Sub LoadBodyFromFile(ByVal FileName As String, ByVal BodyType As emailBodyType)
' DLC 08/25/2010
' Description: Replace the Body of the email with the contents of the specified file.
'
' Parameters:  FileName - The filename whose contents will replace the body of the email
'              BodyType - Specify whether the file contents are to be treated as text or HTML
'
' Sample Usage:
'
' Dim oMsg As New clsEmail
' With oMsg
'     .ErrorHandler = RaiseError
'     .RecipientTo = "abc@xyz.com"
'     .Subject = "test"
'     .LoadBodyFromFile "c:\email.txt", text
'     .Send
' End With
' Set oMsg = Nothing
'
    mvMSG.Body = ReadFileContent(FileName)
    mvMSG.BodyType = BodyType
End Sub

Private Function ReadFileContent(ByVal FileName As String) As String
On Error GoTo ErrorHandler
    Dim fileContentsText As String
    Dim fso As Object
    Dim txtFile As Object
    If Trim(FileName) <> vbNullString Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(FileName) = True Then
            Set txtFile = fso.OpenTextFile(FileName, 1) ' 1 = ForReading
            'The ReadAll method reads the entire file into the variable BodyText
            fileContentsText = txtFile.ReadAll
            'Close the file
            txtFile.Close
        Else
            'If this raises the error rather than displaying or suppressing it, the error will be caught by the ErrorHandler below which
            'will detect the err.source as "AttachmentAdd" and rethrow it.
            DisplayOrRaiseError 9001, "ReadFileContent", "File does not exist (" & FileName & ")"
        End If
    End If
Exit_ErrorHandler:
    On Error Resume Next
    Set txtFile = Nothing
    Set fso = Nothing
    ReadFileContent = fileContentsText
    Exit Function
ErrorHandler:
    'If the error was thrown within this method, rethrow it
    If Err.Source = "ReadFileContent" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        DisplayOrRaiseError 9002, "ReadFileContent", Err.Description & " (" & Err.Number & ")"
    End If
    Resume Exit_ErrorHandler
End Function