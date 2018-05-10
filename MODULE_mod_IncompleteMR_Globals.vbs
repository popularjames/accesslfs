Option Compare Database
'Option Explicit

Global Const chBoxName As String = "ch"
Global Const strMRID = "UAT_INC_MR"

Sub CleanFormLeavingIncompletesOnly()

Dim MyCodeAdo As clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String

Set MyCodeAdo = New clsADO

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_AUDITCLM_Incomplete_MR_Requested_Clean"
                cmd.Parameters.Refresh
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@pErrMsg")
End If

Set MyCodeAdo = Nothing
Set cmd = Nothing

End Sub

Public Sub AddMRToRefTable(clmNum As String, DocID As String)

    Dim dbs As Database
    Dim refInsert As String
    Dim strFileLocation As String
    Dim strNewFilePath As String
    Dim strLTTRID As String
      
    strFileLocation = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & strMRID & "'")
    strLTTRID = DLookup("[ExactTime]", "INC_MR_Insert_Ref_Info", "[CnlyClaimNum] ='" & clmNum & "' and DocID = '" & DocID & "'")
    strNewFilePath = "FAX_" & DocID & "_" & strLTTRID
                                
    refInsert = "INSERT INTO AUDITCLM_References ( CnlyClaimNum, CreateDt, RefType, RefSubType, RefLink, LastUpdateUser )" & _
                                " Select '" & clmNum & "','" & Now() & "', ""ATTACH"", ""CS_MRReq"" ,'" & strFileLocation & strNewFilePath & ".TIF' ,'" & GetUserName & "'"
    Set dbs = CurrentDb
    dbs.Execute refInsert
    dbs.Close

End Sub

Public Sub DeleteMRToRefTable()

    Dim dbs As Database
    Dim refDelete As String
    Dim checkStatus As String
    Dim notSend As Long
    Dim DocIDRs As DAO.RecordSet
    
    checkStatus = "select * from FAX_WORK_Queue where Status <> 'Sent' and Client_ext_Ref_ID = '3'"
                                
    refDelete = "DELETE FROM INC_MR_Insert_Ref_Info"
    
    Set dbs = CurrentDb
    
    Set DocIDRs = dbs.OpenRecordSet(checkStatus)

    If (DocIDRs.EOF And DocIDRs.BOF) Then
    dbs.Execute refDelete
    End If

    Set DocIDRs = Nothing
    dbs.Close

End Sub

Public Function GetRequestInfoAndNotes(CnlyClaimNum As String) As Variant

    Dim myCode_ADO As New clsADO
    Dim rs As ADODB.RecordSet
    Dim rs_ED As ADODB.RecordSet
    Dim strSQL As String
    Dim requestText As String
    Dim other As String

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.sqlString = "usp_Incomplete_MR_Get_Request_Info"
    myCode_ADO.Parameters("@pCnlyClaimNum") = CnlyClaimNum
    Set rs = myCode_ADO.ExecuteRS
        
    If rs.BOF = True And rs.EOF = True Then
         ' Were not able to retrieve data for Incomplete Records
             GetRequestInfoAndNotes = Array("", "")
             GoTo Exit_Function
    Else
        rs.MoveFirst
        audNotes = Nz(rs("AuditorNotes").Value, "")
        other = Nz(rs("other").Value, "")

        While rs.EOF <> True
              requestText = requestText & vbCrLf & rs("value").Value
              rs.MoveNext
        Wend
        End If

        requestText = requestText & vbCrLf & other
        GetRequestInfoAndNotes = Array(requestText, audNotes)
        
Exit_Function:

        Set myCode_ADO = Nothing
        Set rs = Nothing
        Set rs_ED = Nothing
        Exit Function
        
Err_handler:
        MsgBox "Error retrieving selected missing MRs : " & Err.Description
        Resume Exit_Function
    
End Function