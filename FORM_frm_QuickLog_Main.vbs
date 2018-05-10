Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub start()

Dim MyAdo As clsADO
Dim strSQL As String
Dim rsFieldNames As ADODB.RecordSet
    
    
    CurrentDb.Execute "Delete * from tbl_QuickLog"
    Forms!frm_QuickLog_Main!.tbl_QuickLog.Requery
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
            
    strSQL = "select name from sys.columns where object_id = object_id('CMS_AUDITORS_CODE.dbo.v_SCANNING_Quick_Image_Log')"
    Me.SCANNING_Quick_Image_Log.Form.RecordSource = "Select * from v_SCANNING_Quick_Image_Log where 1=2"
    'Me.SCANNING_Quick_Image_Log.Form.Requery
    
    Me.lstFieldName.RowSource = ""
    MyAdo.sqlString = strSQL
    Set rsFieldNames = MyAdo.OpenRecordSet
    
    Do Until rsFieldNames.EOF
        Me.lstFieldName.AddItem (rsFieldNames!Name)
     rsFieldNames.MoveNext
    Loop
Me.lstFieldName = Me.lstFieldName.ItemData(0)

MyAdo.DisConnect
Set MyAdo = Nothing
Set rsFieldNames = Nothing

End Sub


Private Sub MoveSelected()

If IsNull(Me.lstFieldName.Value) Then
    Exit Sub
End If

CurrentDb.Execute "Insert Into tbl_QuickLog (fieldname) VALUES ('" & Me.lstFieldName.Value & "')"
Me.lstFieldName.RemoveItem (Me.lstFieldName.Value)
        
Me.lstFieldName.Requery
Me.lstFieldName = Me.lstFieldName.ItemData(0)

End Sub


Private Sub chkSQL_Click()

If Me.chkSQL.Value = 0 Then
    Me.txtSQL.Enabled = False
 Else
    Me.txtSQL.Enabled = True
End If

End Sub

Private Sub cmdClear_Click()

start

End Sub

Private Sub cmdExec_Click()

'Local table so I'm using DAO - Curlan Johnson 11/7/12

Dim strSelectCriteria As String
Dim strSelectTable As String
Dim strDiv As String
Dim intRec As Integer
Dim myDAO As DAO.RecordSet
Dim strValue As String

On Error GoTo ErrHandler
Me.SCANNING_Quick_Image_Log.Form.RecordSource = "Select * from v_SCANNING_Quick_Image_Log where 1=2"
If Me.chkSQL = -1 Then
    If Len(Me.txtSQL) > 0 Then
'        strSQL = Me.txtSQL.Value
'        Me.SCANNING_Quick_Image_Log.Form.RecordSource = strSQL
'        Me.SCANNING_Quick_Image_Log.Form.Requery
    Else
        MsgBox "Please enter a valid Sql Statment.", vbCritical + vbOKOnly, "Invalid SQL"
    End If
Else
    strDiv = "AND"
    strSelectCriteria = ""
    strSelectTable = "Select * from v_SCANNING_Quick_Image_Log WHERE"
    intRec = 1
    strValue = ""
    
    
    Set myDAO = CurrentDb.OpenRecordSet("SELECT * FROM tbl_QuickLog")
    
    'Check to see if the recordset actually contains rows
    If Not (myDAO.EOF And myDAO.BOF) Then
        myDAO.MoveLast
        myDAO.MoveFirst
        Do Until myDAO.EOF = True
            If myDAO!Criteria Like "*LIKE*" Then strValue = "%" & myDAO!Value & "%" Else strValue = myDAO!Value
            strSelectCriteria = strSelectCriteria & " " & myDAO!FieldName & " " & myDAO!Criteria & " '" & strValue & "'"
            If myDAO.recordCount = intRec Then
                Exit Do
            Else
                intRec = intRec + 1
            End If
            strSelectCriteria = strSelectCriteria & " " & strDiv
            myDAO.MoveNext
        Loop
    Else
        GoTo Cleanup
    End If

    Me.txtSQL = ""
    Me.txtSQL = strSelectTable & " " & strSelectCriteria
End If
'

    Dim oAdo As clsADO
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = sqltext
        .sqlString = Me.txtSQL
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Me.SCANNING_Quick_Image_Log.Form.RecordSource = "Select * from v_SCANNING_Quick_Image_Log where 1=2"
            GoTo Cleanup
        End If
    End With

    Set Me.SCANNING_Quick_Image_Log.Form.RecordSet = oRs


'Me.SCANNING_Quick_Image_Log.Form.RecordSource = strSelectTable & " " & strSelectCriteria
'Me.SCANNING_Quick_Image_Log.Form.Requery

Cleanup:
'If myDAO <> "" Then
    
    myDAO.Close
    Set myDAO = Nothing
'End If
Exit Sub

ErrHandler:
    If Err.Number = 3075 Or Err.Number = 94 Then
        MsgBox "Please enter a valid criteria.", vbCritical + vbOKOnly, "Missing Criteria"
    Else
        MsgBox Err.Number & " - " & Err.Description, vbCritical + vbOKOnly, "Lookup Error"
    End If
      
    Resume Cleanup

End Sub

Private Sub cmdMoveAll_Click()

Do Until IsNull(Me.lstFieldName.Value) = True
         Call MoveSelected
Loop
Forms!frm_QuickLog_Main!.tbl_QuickLog.Requery

End Sub

Private Sub cmdRemove_Click()
Dim strSQL As String


If Me.tbl_QuickLog.Form.RecordSet.recordCount = 0 Then
    Exit Sub
End If

Me.lstFieldName.AddItem (Me.tbl_QuickLog("FieldName").Value)

strSQL = "Delete * from tbl_QuickLog where fieldName = '" & Me.tbl_QuickLog("FieldName").Value & "'"

CurrentDb.Execute strSQL

Forms!frm_QuickLog_Main!.tbl_QuickLog.Requery

Me.lstFieldName = Me.lstFieldName.ItemData(0)
Me.lstFieldName.Requery

End Sub

Private Sub cmdSelect_Click()

Call MoveSelected
Forms!frm_QuickLog_Main!.tbl_QuickLog.Requery

End Sub


Private Sub Form_Load()

Me.chkSQL.Value = 0
Me.txtSQL.Enabled = False
start

End Sub
