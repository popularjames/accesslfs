Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvFileName As String
Property Let FileName(data As String)
    MvFileName = data
End Property

' DLC 05/20/2010 - Fixed bug where Icon was never created and added error handling
Public Sub SaveFile(ByVal AppName As String, Optional ByVal FileName As String = "")
On Error GoTo ErrorHandler
    Dim fso
    Dim oFile
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim Exists As Boolean
    
    If FileName <> "" Then
        MvFileName = FileName
    End If
    ' determine if the icon should be updated
    Set fso = CreateObject("Scripting.FileSystemObject")
    Exists = fso.FileExists(MvFileName)
    If Exists Then
        Set oFile = fso.GetFile(MvFileName)
        Set db = CurrentDb
        Set rs = db.OpenRecordSet("Select updateDate from CT_Icons Where AppName = '" & AppName & "'", dbOpenSnapshot)
        If rs.recordCount > 0 Then
            Exists = (oFile.DateLastModified >= rs!UpdateDate)
        End If
    End If
    If Not Exists Then
        WriteBinaryFile MvFileName, AppName
    End If
Exit_ErrorHandler:
    On Error Resume Next
    Set fso = Nothing
    Set oFile = Nothing
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
ErrorHandler:
    'Ignore any errors
    Resume Exit_ErrorHandler
End Sub

Public Sub LoadFile(AppName As String, Optional FileName As String)
On Error GoTo ErrorHappened
Dim fso
If FileName <> "" Then
    MvFileName = FileName
End If

If MvFileName = "" Then
    MsgBox "No File Name Specified for Icon Load"
    GoTo ExitNow
End If


Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(MvFileName) = False Then
    MsgBox "Specified File Does Not Exist:" & vbCrLf & vbCrLf & MvFileName
    GoTo ExitNow
End If

ReadBinaryFile MvFileName, AppName

ExitNow:
    On Error Resume Next
    Set fso = Nothing
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical, "Load Icon File"
    Resume ExitNow
    Resume
End Sub

Public Sub ReadBinaryFile(ByVal FileName As String, Optional AppName As String = CnlyAppName)
Dim cn 'As ADODB.Connection
Dim rs 'As ADODB.Recordset
Dim BinaryStream 'As ADODB.Stream

Const adTypeText = 2
Const adTypeBinary = 1
Const adOpenKeyset = 1
Const adLockOptimistic = 3

Set BinaryStream = CreateObject("ADODB.Stream") 'Create Stream object
BinaryStream.Type = adTypeBinary 'Specify stream type - we want To get binary data.
BinaryStream.Open 'Open the stream
BinaryStream.LoadFromFile FileName 'Load the file data from disk To stream object
  
Set cn = CreateObject("ADODB.Connection") 'Create a Connection
' updated to use the variable for access
cn.Open LINK_SRC_ACCESS & "Data Source=" & CurrentDb.Name & ";Mode=Share Deny None;Jet OLEDB:Database Locking Mode=1;"


Set rs = CreateObject("ADODB.Recordset")
rs.Open "Select * from CT_Icons Where AppName = '" & AppName & "'", cn, adOpenKeyset, adLockOptimistic

If rs.EOF And rs.BOF Then
    rs.AddNew
    rs.Fields("AppName").Value = AppName
End If
   
' HC - update the file date
rs.Fields("UpdateDate").Value = Now
rs.Fields("Image").Value = BinaryStream.Read
rs.Update

rs.Close
cn.Close


BinaryStream.Close
Set BinaryStream = Nothing
  
End Sub



Private Sub WriteBinaryFile(ByVal FileName As String, Optional AppName As String = CnlyAppName)
Dim cn 'As ADODB.Connection
Dim rs 'As ADODB.Recordset
Dim BinaryStream 'As ADODB.Stream

Const adTypeText = 2
Const adTypeBinary = 1
Const adOpenKeyset = 1
Const adLockOptimistic = 3
Const adSaveCreateOverWrite = 2

Set BinaryStream = CreateObject("ADODB.Stream") 'Create Stream object
BinaryStream.Type = adTypeBinary 'Specify stream type - we want To get binary data.
BinaryStream.Open 'Open the stream

  
Set cn = CreateObject("ADODB.Connection") 'Create a Connection
' Updated to use the variable for access
cn.Open LINK_SRC_ACCESS & "Data Source=" & CurrentDb.Name & ";Mode=Share Deny None;Jet OLEDB:Database Locking Mode=1;"

Set rs = CreateObject("ADODB.Recordset")
rs.Open "Select * from CT_Icons Where AppName = '" & AppName & "'", cn, adOpenKeyset, adLockOptimistic

If rs.EOF And rs.BOF Then
    GoTo ExitNow
End If

BinaryStream.Write rs.Fields("Image").Value
BinaryStream.SaveToFile FileName, adSaveCreateOverWrite

ExitNow:
    On Error Resume Next
    BinaryStream.Close
    rs.Close
    cn.Close
    Set BinaryStream = Nothing
    Set rs = Nothing
    Set cn = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description
    Resume ExitNow
    Resume
End Sub