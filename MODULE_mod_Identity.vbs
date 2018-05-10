Option Compare Database
Option Explicit

Public Function GetAppKey(strAppID As String) As Long

    Dim MyAdo As clsADO
    Dim cmd As ADODB.Command

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = MyAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_ADMIN_Get_App_Key"
    cmd.Parameters.Refresh
    cmd.Parameters("@pAppID") = strAppID
    cmd.Execute
    GetAppKey = cmd.Parameters("@pAppKey")
     
End Function



Public Function GetInstanceID() As String

    Dim MyAdo As clsADO
    Dim cmd As ADODB.Command
    Dim strInstanceID As String
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = MyAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_GetInstanceID"
    cmd.Parameters.Refresh
    cmd.Parameters("@InstanceID") = strInstanceID
    cmd.Execute
    GetInstanceID = cmd.Parameters("@InstanceID")
     
End Function


Public Function GetAppPermission(AppID As String) As Integer

    Dim rs As ADODB.RecordSet
    Dim strUserName As String
    Dim iAppPermissions As Integer
    
    Dim myCode_ADO As New clsADO
    Dim cmd As ADODB.Command
    
    iAppPermissions = 0
    
    strUserName = Identity.UserName
    
    
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_ADMIN_Get_User_Permissions"
    myCode_ADO.SQLTextType = StoredProc
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_ADMIN_Get_User_Permissions"

    cmd.Parameters.Refresh
    cmd.Parameters("@pUserID") = strUserName
    cmd.Parameters("@pAppID") = AppID
    cmd.Parameters("@pAccountID") = gintAccountID
    Set rs = myCode_ADO.ExecuteRS(cmd.Parameters)
    
    If rs.EOF = True And rs.BOF = True Then
    Else
        rs.MoveFirst
        
        While rs.EOF <> True
            Select Case UCase(rs("ActionID"))
                Case "LOCKED"
                    iAppPermissions = gcLocked
                Case "ADD"
                    iAppPermissions = iAppPermissions Or gcAllowAdd
                Case "CHANGE"
                    iAppPermissions = iAppPermissions Or gcAllowChange
                Case "DELETE"
                    iAppPermissions = iAppPermissions Or gcAllowDelete
                Case "VIEW"
                    iAppPermissions = iAppPermissions Or gcAllowView
                Case "REASSIGN"
                    iAppPermissions = iAppPermissions Or gcAllowReAssign
                Case "FORWARD"
                    iAppPermissions = iAppPermissions Or gcAllowForward
                Case "RELEASE"
                    iAppPermissions = iAppPermissions Or gcReleaseClaim
                Case "PRINTLTR"
                    iAppPermissions = iAppPermissions Or gcPrintLetter
                Case Else
            End Select
            rs.MoveNext
        Wend
    End If
    
    GetAppPermission = iAppPermissions
    
    Set rs = Nothing
    Set myCode_ADO = Nothing
End Function


Public Function GetUserProfile() As String

    Dim rs As ADODB.RecordSet
    Dim strUserName As String
    
    Dim myCode_ADO As New clsADO
    Dim cmd As ADODB.Command
    
    
    strUserName = Identity.UserName
    
    
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_ADMIN_Get_User_Profile"
    myCode_ADO.SQLTextType = StoredProc
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_ADMIN_Get_User_Profile"

    cmd.Parameters.Refresh
    cmd.Parameters("@pUserID") = strUserName
    cmd.Parameters("@pAccountID") = gintAccountID
    
    Set rs = myCode_ADO.ExecuteRS(cmd.Parameters)
    
    If rs.EOF = True And rs.BOF = True Then
        GetUserProfile = ""
    Else
        GetUserProfile = rs("ProfileID")
    End If
    
    Set rs = Nothing
    Set myCode_ADO = Nothing
End Function


Public Function GetConnectString(TableName As String) As String

Dim strTableConnect As String
Dim intStartPos As Integer
Dim intLen As Integer
Dim strServer As String
Dim strDatabase As String

    If IsTable(TableName) = False Then
        GetConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                "Data Source=" & CurrentCMSServer() & ";" & _
                "Initial Catalog=" & TableName & ";"
        Exit Function
    Else
        strTableConnect = CurrentDb.TableDefs(TableName).Connect & ";"
    End If


    'get server and database name from connectstring in workfile linked table
    intStartPos = InStr(strTableConnect, "SERVER=") + 7
    intLen = InStr(intStartPos, strTableConnect, ";") - intStartPos
    strServer = Mid(strTableConnect, intStartPos, intLen)

    intStartPos = InStr(strTableConnect, "DATABASE=") + 9
    intLen = InStr(intStartPos, strTableConnect, ";") - intStartPos
    strDatabase = Mid(strTableConnect, intStartPos, intLen)


    GetConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                "Data Source=" & strServer & ";" & _
                "Initial Catalog=" & strDatabase & ";"

End Function


Public Function GetPCName() As String
    Dim WshNetwork
    
    On Error GoTo Error_Handler
    
    Set WshNetwork = CreateObject("WScript.Network")
    GetPCName = WshNetwork.ComputerName
    
ExitNow:
    Set WshNetwork = Nothing
    Exit Function

Error_Handler:
    MsgBox Err.Description, vbInformation, Application.CodeContextObject.Name
    Resume ExitNow
End Function

Public Function GetUserName() As String
    Dim WshNetwork
    
    On Error GoTo Error_Handler
    
    Set WshNetwork = CreateObject("WScript.Network")
    GetUserName = WshNetwork.UserName
    
ExitNow:
    Set WshNetwork = Nothing
    Exit Function

Error_Handler:
    MsgBox Err.Description, vbInformation, Application.CodeContextObject.Name
    Resume ExitNow
End Function



 