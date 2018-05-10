Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_AfterUpdate()
    Dim myPortalADO As clsADO
    Set myPortalADO = New clsADO
    myPortalADO.ConnectionString = GetConnectString("v_DATA_Database")
    myPortalADO.sqlString = "update cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp SET PageCnt = " & Me.PageCnt & " where icn = '" & Me.Icn & "' AND scanneddt = '" & charScannedDt & "'" & _
                            "update cms_auditors_claims.dbo.SCANNING_Quick_Image_Log set importflag = '" & Me.ImportFlag & "' where icn = '" & Me.Icn & "' AND seqno = " & Me.SeqNo
    myPortalADO.SQLTextType = sqltext
    myPortalADO.Execute
End Sub

Private Sub Form_Load()


    Dim strBaseSQL As String
    
    Me.AllowAdditions = False
    Me.AllowDeletions = False
    Me.AllowEdits = True
    
    

'    If Not IsSubForm(Me) Then
'        strBaseSQL = "SELECT t1.*, t2.PageCnt, t2.ErrMsg,t2.LocalPath " & _
'                    " FROM SCANNING_Quick_Image_Log as t1 " & _
'                    " INNER JOIN SCANNING_Image_Log_Tmp as t2 " & _
'                    " ON t1.ScannedDt=t2.ScannedDt " & _
'                    " where 1=2"
'        'Me.RecordSource = strBaseSQL
'    End If
    
End Sub


Private Sub Form_Current()
' Call our redraw function.
' We have to do this here because of a bug using
' Withevents to sink a Form's events from a Class module.
'CF.Redraw
End Sub
