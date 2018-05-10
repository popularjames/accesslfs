Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrCalledFrom As String
Private mstrBaseSQL As String
Private mstrWhereClause As String
Private mstrFilterBy As String

Private MyAdo As clsADO
Private myCode_ADO As clsADO

Const CstrFrmAppID As String = "ImageLog"




Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cmdDeleteRecord_Click()
    If (mstrFilterBy & "" <> "") And (Me.txtSearchParam & "" <> "") Then
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.sqlString = "exec usp_SCANNING_Quick_Image_Log_Purge '" & Me.txtSearchParam & "','" & mstrFilterBy & "'"
        myCode_ADO.SQLTextType = sqltext
        myCode_ADO.Execute
        Set myCode_ADO = Nothing
    
        Refresh_Screen
    End If
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close
End Sub


Private Sub cmdGenerateCoverPage_Click()
    Dim strParams As String
    
    
        If Me.SCANNING_Image_Log_Update.Form.RecordSet.recordCount > 0 Then
            If UCase(Me.SCANNING_Image_Log_Update.Form.RecordSet("ImportFlag")) = "Y" Then
                strParams = CStr(Me.SCANNING_Image_Log_Update.Form.RecordSet("ScannedDt")) & ";" & CStr(Me.SCANNING_Image_Log_Update.Form.RecordSet("ReceivedDt")) & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("CnlyProvID") & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("ICN") & ";" & _
                            Me.SCANNING_Image_Log_Update.Form("CnlyClaimNum") & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("ImageName") & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("ImageType") & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("MemberName") & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("DOB") & _
                            ";" & CStr(Me.SCANNING_Image_Log_Update.Form.RecordSet("ClmFromDt")) & ";" & CStr(Me.SCANNING_Image_Log_Update.Form.RecordSet("ClmThruDt")) & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("UserID") & ";QuickEntry" & ";" & _
                            Me.SCANNING_Image_Log_Update.Form.RecordSet("RequestNum") & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("TrackingNum") & ";" & CStr(Me.SCANNING_Image_Log_Update.Form.RecordSet("SeqNo")) & ";" & Me.SCANNING_Image_Log_Update.Form.RecordSet("SessionID")
                'DoCmd.OpenReport "rpt_Scanning_Cover_page", acViewNormal, , , acWindowNormal, strParams
                DoCmd.OpenReport "rpt_Scanning_Cover_page", acViewPreview, , , acWindowNormal, strParams
            End If
        End If
    End Sub

Private Sub cmdSearch_Click()
    Go_Search
End Sub


Private Sub cmdValidate_Click()
    Dim strImagePath As String
    Dim strImageName As String
    
    Set MyAdo = New clsADO
    
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    
    With Me.SCANNING_Image_Log_Update.Form.RecordSet
        If .recordCount > 0 Then
            .MoveFirst
            While Not .EOF
                strImageName = !ImageName
                strImagePath = !LocalPath
                
                
                !ErrMsg = ""
                
                If Right(strImagePath, Len(strImageName)) = strImageName Then
                    strImageName = strImagePath
                ElseIf Right(strImagePath, 1) <> "\" Then
                    strImageName = strImagePath & "\" & strImageName
                Else
                    strImageName = strImagePath & strImageName
                End If
                
                If Not fso.FileExists(strImageName) Then
                    !ErrMsg = "File " & strImageName & " does not exists"
                End If
                Set MyAdo = New clsADO
                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                MyAdo.SQLTextType = sqltext
                If !ErrMsg <> Me.SCANNING_Image_Log_Update.Form.ErrMsg Then
                        MyAdo.sqlString = "update cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp set ErrMsg = '" & !ErrMsg & "' where ScannedDt = '" & Me.SCANNING_Image_Log_Update.Form.charScannedDt & "'"
                        MyAdo.Execute
                End If
                Set MyAdo = Nothing
                .MoveNext
            Wend
        End If
    End With
    
    Refresh_Screen
End Sub

Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub


Private Sub Form_Load()
    Dim strErrMsg As String
    Dim strErrSource As String
    Dim iAppPermission As String
    Dim iTempAccountID As String
    Dim iFormType As Integer
    Dim strSQL As String
    

    
    
    On Error GoTo Err_handler
      
    Me.Caption = "Quick Image Log Update"
        
    Call Account_Check(Me)
    
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub

    
    strErrSource = "Quick_SCANNING_Data_Update.Load"
    
    If IsSubForm(Me) Then
        cmdExit.visible = False
        lblAppTitle.visible = False
    End If
    
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs
    End If
    
    mstrBaseSQL = "SELECT t1.*, t2.PageCnt, t2.ErrMsg, t2.LocalPath, Rownum = ROW_NUMBER() OVER(ORDER BY t1.SeqNo), charScannedDt = convert(varchar(23),t1.scanneddt,121)" & _
                " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                " ON t1.ScannedDt=t2.ScannedDt"

    mstrWhereClause = " where 1=2"
    

    ' init screen display
    Me.txtSearchParam = ""
    Me.SearchType = 1
    Refresh_Screen
    
    

Exit_Sub:
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then
        Err.Raise vbObjectError + 513, strErrSource, strErrMsg
    Else
        Err.Raise Err.Number, strErrSource, Err.Description
    End If
End Sub


Private Sub Refresh_Screen()
    Dim rs As ADODB.RecordSet
    Dim strErrMsg As String
    Dim iRecordSelected As Integer
    
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = mstrBaseSQL & mstrWhereClause
    
    Set Me.SCANNING_Image_Log_Update.Form.RecordSet = MyAdo.OpenRecordSet()
    
    
    
    'Me.SCANNING_Image_Log_Update.Form.RecordSource = mstrBaseSQL & mstrWhereClause
    'Me.SCANNING_Image_Log_Update.Form.Requery
    
    If Me.txtSearchParam & "" <> "" Then
        Set rs = Me.SCANNING_Image_Log_Update.Form.RecordSet
        If rs.recordCount > 0 Then
            Me.lblProviderName1.visible = True
            Me.lblProviderName2.visible = True
            Me.lblProviderName2.Caption = rs("ProvName")
            Me.lblLetterSentDt1.visible = True
            Me.lblLetterSentDt2.visible = True
            Me.lblLetterSentDt2.Caption = rs("LetterReqDt")
            Me.lblRecordSelected1.visible = True
            Me.lblRecordSelected2.visible = True
            Me.lblRecordSelected2.Caption = rs.recordCount & " "
        Else
            Me.lblProviderName2.Caption = ""
            Me.lblProviderName2.Caption = ""
            Me.lblRecordSelected2.Caption = ""
            Me.lblProviderName1.visible = False
            Me.lblProviderName2.visible = False
            Me.lblLetterSentDt1.visible = False
            Me.lblLetterSentDt2.visible = False
            Me.lblRecordSelected1.visible = False
            Me.lblRecordSelected2.visible = False
            MsgBox "No record matching criteria"
            Me.txtSearchParam.SetFocus
        End If
    Else
        ' reset screen and variables
        Me.lblProviderName1.visible = False
        Me.lblProviderName2.visible = False
        Me.lblLetterSentDt1.visible = False
        Me.lblLetterSentDt2.visible = False
        Me.lblRecordSelected1.visible = False
        Me.lblRecordSelected2.visible = False
    End If



End Sub




Private Sub optOrderBy_Click()


    Select Case optOrderBy
        Case 1 'SeqNo
            mstrBaseSQL = "SELECT t1.*, t2.PageCnt, t2.ErrMsg, t2.LocalPath,  Rownum = ROW_NUMBER() OVER(ORDER BY t1.SeqNo), charScannedDt = convert(varchar(23),t1.scanneddt,121)" & _
                        " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                        " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                        " ON t1.ScannedDt=t2.ScannedDt"
            
        Case 2 'ClmFromDt
        
            mstrBaseSQL = "SELECT t1.*, t2.PageCnt, t2.ErrMsg, t2.LocalPath,  Rownum = ROW_NUMBER() OVER(ORDER BY t1.ClmFromDt), charScannedDt = convert(varchar(23),t1.scanneddt,121)" & _
                        " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                        " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                        " ON t1.ScannedDt=t2.ScannedDt"
        
        Case 3 'ICN
        
            mstrBaseSQL = "SELECT t1.*, t2.PageCnt, t2.ErrMsg, t2.LocalPath, Rownum = ROW_NUMBER() OVER(ORDER BY t1.ICN), charScannedDt = convert(varchar(23),t1.scanneddt,121)" & _
                        " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                        " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                        " ON t1.ScannedDt=t2.ScannedDt"
        
    End Select
    
    Go_Search
End Sub

Private Sub SearchType_Click()
    Select Case SearchType
        Case 1
            lblSearchType.Caption = "Enter Batch ID"
        Case 2
            lblSearchType.Caption = "Enter Tracking Number"
        Case 3
            lblSearchType.Caption = "Enter Request Number"
    End Select
    
       
End Sub




Private Sub Go_Search()
    mstrWhereClause = " where 1=2"
    
    Select Case SearchType
        Case 1
            mstrWhereClause = " where SessionID = '" & txtSearchParam & "'"
            mstrFilterBy = "SessionID"
        Case 2
            mstrWhereClause = " where TrackingNum = '" & txtSearchParam & "'"
            mstrFilterBy = "TrackingNum"
        Case 3
            mstrWhereClause = " where RequestNum = '" & txtSearchParam & "'"
            mstrFilterBy = "RequestNum"
    End Select
    
    Refresh_Screen
End Sub
