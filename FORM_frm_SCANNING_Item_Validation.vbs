Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strSQL As String
Dim strWhere As String
Dim strImageType As String
Dim strStartDate As String
Dim strCalledFrom As String

Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdSearch_Click()
    Call txtClaimNum_AfterUpdate
End Sub

Private Sub Form_Close()
    If strCalledFrom <> "" Then
        DoCmd.OpenForm strCalledFrom
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Claim Image Validation"
    
    Me.txtClaimNum.SetFocus
        
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    If Me.OpenArgs() & "" <> "" Then
        strCalledFrom = Me.OpenArgs()
    End If
    
    If UCase(left(GetPCName(), 3)) <> "TS-" Then
        strSQL = "select *, '#' + ImagePath as WiltonFile, '#' + LocalPath as PhillyFile from SCANNING_Image_Log"
        '"select CnlyClaimNum, ImageName, ReceivedDt, ScannedDt, PageCnt, TIFCnt, ValidationDt, '#' + LocalPath as ImageLink from SCANNING_Image_Log"
    Else
        strSQL = "select *, '#' + ImagePath as WiltonFile, '#' + LocalPath as PhillyFile from SCANNING_Image_Log"
        'strSQL = "select CnlyClaimNum, ImageName, ReceivedDt, ScannedDt, PageCnt, TIFCnt, ValidationDt, '#' + ImagePath as ImageLink from SCANNING_Image_Log"
    End If
    
    Me.SubForm.Form.RecordSource = ""
End Sub

Private Sub txtClaimNum_AfterUpdate()
    Dim strRecSource As String
    strWhere = ""
    
    If txtImageType & "" <> "" Then
        strWhere = " where ImageType = '" & txtImageType & "'"
    End If
    
    If txtStartDate & "" <> "" Then
        If strWhere <> "" Then
            strWhere = strWhere & " and ScannedDt >= #" & txtStartDate & "#"
        Else
            strWhere = " where ScannedDt >= #" & txtStartDate & "#"
        End If
    End If
    
    strRecSource = strSQL & strWhere
    strRecSource = strRecSource & " and CnlyClaimNum = '" & txtClaimNum & "'"
    
    Me.SubForm.Form.RecordSource = strRecSource
    Me.SubForm.Form.Requery
    
End Sub


Private Sub txtImageType_AfterUpdate()
    strWhere = ""
    
    If strImageType <> txtImageType Then
        If txtImageType & "" <> "" Then
            strWhere = " where ImageType = '" & txtImageType & "'"
        End If
    
        If txtStartDate & "" <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " and ScannedDt >= '" & txtStartDate & "'"
            Else
                strWhere = " where ScannedDt >= '" & txtStartDate & "'"
            End If
        End If
        
        MyAdo.sqlString = strSQL & strWhere
        Set rs = MyAdo.OpenRecordSet
        Set txtClaimNum.RecordSet = rs
        strImageType = txtImageType
    End If
End Sub

Private Sub txtStartDate_AfterUpdate()
    strWhere = ""
    
    If txtStartDate & "" <> "" Then
        If IsDate(txtStartDate) = False Then
            MsgBox "Please enter a valid date"
            txtStartDate.SetFocus
            Exit Sub
        End If
    End If
    
    If strStartDate <> txtStartDate Then
        If txtImageType & "" <> "" Then
            strWhere = " where ImageType = '" & txtImageType & "'"
        End If
    
        If txtStartDate & "" <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " and ScannedDt >= '" & txtStartDate & "'"
            Else
                strWhere = " where ScannedDt >= '" & txtStartDate & "'"
            End If
        End If
        
        MyAdo.sqlString = strSQL & strWhere
        Set rs = MyAdo.OpenRecordSet
        Set txtClaimNum.RecordSet = rs
        strStartDate = txtStartDate
    End If
End Sub
