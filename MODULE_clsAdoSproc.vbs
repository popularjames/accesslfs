Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private oConn As Variant
Private oCmd As Variant

Private msConnect As String
Private msCmdText As String
Private miTimeout As Integer

'Public Params As Collection


Private Sub Class_Initialize()

    Set oConn = CreateObject("ADODB.Connection")
    Set oCmd = CreateObject("ADODB.Command")
    
    miTimeout = 600

End Sub

Private Sub Class_Terminate()

    Set oConn = Nothing
    Set oCmd = Nothing

End Sub

Public Property Let CommandText(CommandText As String)

    msCmdText = CommandText

End Property

Public Property Let ConnectString(ConnectString As String)
    
    msConnect = ConnectString
    
End Property


Public Property Get GetParam(ParamName As String) As Variant

  GetParam = oCmd.Parameters(ParamName)
  
    
End Property

Public Property Let RefTable(ReferenceTable As String)

    msConnect = GetConnectString(ReferenceTable)

End Property

Public Property Let timeout(Seconds As Integer)

    miTimeout = Seconds

End Property

Public Function AddParam(Item As String, pValue As Variant)

oCmd.Parameters(Item) = pValue


End Function

Public Function Exec()

oCmd.Execute


End Function

Public Property Get ReturnValue() As Integer
Dim strName As String

'*Return parameter is always first so it will be 0 in our collection

    If oCmd.Parameters(0).Direction = 4 Then '*adParamReturnValue

        ReturnValue = oCmd.Parameters("@Return_Value")

    Else
        MsgBox "Sproc does not have a return parameter"
    End If

End Property

Public Function Setup()

    
    If oConn.State = 1 Then '* adStateOpen
        oConn.Close
    End If
    
    oConn.Open msConnect
    
    oCmd.ActiveConnection = oConn
    oCmd.CommandTimeout = miTimeout
    oCmd.commandType = 4 '* does this need to be a variable??
    oCmd.CommandText = msCmdText
    oCmd.Parameters.Refresh

End Function

Private Function GetConnStr(TableName As String) As String

Dim strTableConnect As String
Dim intStartPos As Integer
Dim intLen As Integer
Dim strServer As String
Dim strDb As String

    strTableConnect = CurrentDb.TableDefs(TableName).Connect

    'get server and database name from connectstring from linked table
    intStartPos = InStr(strTableConnect, "SERVER=") + 7
    intLen = InStr(intStartPos, strTableConnect, ";") - intStartPos
    strServer = Mid(strTableConnect, intStartPos, intLen)

    intStartPos = InStr(strTableConnect, "DATABASE=") + 9
    intLen = InStr(intStartPos, strTableConnect, ";") - intStartPos
    strDb = Mid(strTableConnect, intStartPos, intLen)


    GetConnStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                "Data Source=" & strServer & ";" & _
                "Initial Catalog=" & strDb & ";"

End Function