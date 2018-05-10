Option Compare Database
Option Explicit

Private Const ConfigUrl As String = "http://connollyconfig.svc.ccaintranet.com/Service.asmx"
Private ConfigBlock As Variant

Private MvIsDcUser As String  'values will be either: "" / "Y" / "N"

Const adVarChar = 200
Const adChar = 129
Const adInteger = 3
Const adDate = 7
Const adParamInput = 1
Const adCmdStoredProc = 4
Const adParamReturnValue = 4


Private Function GetConfigSoapEnvelope() As String
    Dim SoapReq$
    
    SoapReq = "<?xml version=""1.0"" encoding=""utf-8""?>"
    SoapReq = SoapReq & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    SoapReq = SoapReq & " <soap12:Body>"
    SoapReq = SoapReq & "<GetConfigBlock xmlns=""http://tempuri.org/"">"
    SoapReq = SoapReq & "  <AppName>GroupSecurity</AppName>"
    SoapReq = SoapReq & "    </GetConfigBlock>"
    SoapReq = SoapReq & "  </soap12:Body>"
    SoapReq = SoapReq & "</soap12:Envelope>"
    GetConfigSoapEnvelope = SoapReq
    
End Function

Private Function LoadConfigBlock()
    Dim xReq 'As New XMLHTTP
    Dim xDoc 'As DOMDocument
    Dim xDocSub 'As DOMDocument
    Dim SoapEnvelope 'As String
    
    Set xReq = CreateObject("MSXML2.XMLHTTP")
    xReq.Open "POST", ConfigUrl, False, Nothing, Nothing
    xReq.setRequestHeader "Content-Type", "Application/soap+xml; charset=utf-8"
    SoapEnvelope = GetConfigSoapEnvelope
    xReq.Send SoapEnvelope
    Set xDoc = CreateObject("MSXML2.DOMDocument")
    Set xDocSub = CreateObject("MSXML2.DOMDocument")
    xDoc.loadXML (xReq.ResponseText)
    xDocSub.loadXML (xDoc.Text)
    Set ConfigBlock = xDocSub
End Function
Private Function GetMemberConnectString() As String
    Dim Server$, Database$, Result$
    Server = ConfigBlock.selectSingleNode("/Application/DataServer").Text
    Database = ConfigBlock.selectSingleNode("/Application/DataBase").Text
    Result = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=" & Server & ";Initial Catalog=" & Database & ";"
    GetMemberConnectString = Result
End Function
Public Function isDcUser() As Boolean
On Error GoTo IsDcUser_Err
    Dim cn 'As ADODB.Connection
    Dim cmd 'as ADODB.Command
    Dim prm 'As ADODB.Parameter
    Dim retval 'As ADODB.Parameter
    
    
    'DPR - Running into iussues with MY AD GRoup
    If Identity.UserName = "Damon.Ramaglia" Then
     isDcUser = True
        Exit Function
    End If
    
    
    
    'If we have already calculated this, return the result here
    If MvIsDcUser <> vbNullString Then
        isDcUser = (MvIsDcUser = "Y")
        Exit Function
    End If
    
    LoadConfigBlock
    
    Set cn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    Set prm = CreateObject("ADODB.Parameter")
    cn.ConnectionString = GetMemberConnectString
    cn.Open
    With prm
        .Direction = adParamInput
        .Name = "GroupName"
        .Type = adVarChar
        .Size = 1000
        .Value = "DataCenter"
    End With
    With cmd
        .ActiveConnection = cn
        .CommandText = ConfigBlock.selectSingleNode("/Application/Procedure").Text
        .commandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("Retval", adInteger, adParamReturnValue)
        .Parameters.Append prm
        '.Parameters.Append retval
        .Execute
        retval = (.Parameters(0).Value = 0)
    End With
    
  
IsDcUser_Exit:
    
    cn.Close
    Set cmd = Nothing
    Set cn = Nothing
    MvIsDcUser = IIf(retval, "Y", "N")
    isDcUser = retval
    Exit Function
    
IsDcUser_Err:
    On Error Resume Next
    retval = False
    Resume IsDcUser_Exit
End Function