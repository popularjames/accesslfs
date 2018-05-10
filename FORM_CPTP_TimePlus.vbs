Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' DPS 5/2/2012 changed to use Time Track website for form

Private Sub Form_Load()
On Error GoTo ErrorHandler
    DoCmd.Hourglass True

    ' needed to put a URL in the control source of this would not work
    Me.webTimeTrack.Object.Navigate GetTimeTrackURL()
    Me.webTimeTrack.SetFocus
   
    DoCmd.Hourglass False
    
Exit_ErrorHandler:
    On Error Resume Next
    DoCmd.Maximize
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    MsgBox Err.Description
    Resume Exit_ErrorHandler
End Sub

Private Sub Form_Resize()
On Error Resume Next

    ' make web browser as big as possible
    Me.webTimeTrack.Height = Me.Form.InsideHeight - 100
    Me.webTimeTrack.Width = Me.Form.InsideWidth - 200
End Sub

Private Function GetTimeTrackURL() As String
On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim strUser As String
    Dim myUrl As String
    Dim genUtils As New CT_ClsGeneralUtilities
        
    'Set db = CurrentDb
    '' get URL for time track application
    'Set rs = db.OpenRecordSet("SELECT URL FROM vCPuApplicationType WHERE ApplicationTypeID=1", dbOpenSnapshot, dbReadOnly)
        
    'If Not (rs.EOF) Then
        'Send default values on the URL
        myUrl = "https://timeplus.myconnolly.com/" & "?AuditID=" & genUtils.URLEncode(Nz(Identity.AuditNum, ""))
    'End If

    'Debug.Print myUrl
    GetTimeTrackURL = myUrl

Exit_ErrorHandler:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set genUtils = Nothing
    Exit Function

ErrorHandler:
    MsgBox genUtils.CompleteDBExecuteError, vbCritical, "Error: Time Entry --> Getting URL"
    Resume Exit_ErrorHandler
End Function
