Option Compare Database
Option Explicit
'Damon 06/03/08
'These are placeholders to handle actions associated with Claim records
'The information that routes to these functions is stored in the AUDITCLM_Action table
'They are called from the AUDITCLM_Main form
Public Function CloseClaim(strClaimNum As String) As String
    MsgBox "Close Claim was Called for - " & strClaimNum, vbOKOnly
    CloseClaim = "Close Claim was Called for - " & strClaimNum
End Function
Public Function VoidClaim(strClaimNum As String) As String
    MsgBox "Void Claim was Called for - " & strClaimNum, vbOKOnly
    VoidClaim = "Void Claim was Called for - " & strClaimNum
End Function
Public Function SendToPayer(strClaimNum As String) As String
    MsgBox "Send to Payer Claim was Called for - " & strClaimNum, vbOKOnly
    SendToPayer = "SendToPayer Claim was Called for - " & strClaimNum
End Function

Public Function LaunchProviderManagement(strCnlyProvID As String) As String
    NewProvider strCnlyProvID, strCnlyProvID
    'LaunchProviderManagement = "LaunchProviderManagement was Called for - " & strcnlyProvID
End Function
Public Function LaunchChartReview(strCnlyClaimNum As String) As String
    
    DoCmd.OpenReport "rpt_AUDITCLM_ChartReview", acViewPreview, , "cnlyClaimNum = '" & strCnlyClaimNum & "'"
End Function

Public Function LaunchClaimDetailReport(strCnlyClaimNum As String) As String
    
    DoCmd.OpenReport "rptClaimDetail", acViewPreview, , "cnlyClaimNum = '" & strCnlyClaimNum & "'"
End Function

Public Sub ResizeControls(ByRef frm As Form)
On Error GoTo Err_handler

    Dim rstScreenControl As DAO.RecordSet
    Dim strSQL As String
    Dim ctl As Control
    ' 20130417 KD: CHanged to Long's
    Dim intWidth As Long
    Dim intHeight As Long
    
    Dim strParameters() As String
    ' 20130417 KD: CHanged to Long's
    Dim intDistanceFromRight As Long
    Dim intDistanceFromBottom As Long
    
    frm.Form.ScrollBars = 0
    
    strSQL = "SELECT * FROM GENERAL_Screen_Control " + _
    "WHERE Form = '" & frm.Name & "' AND Event = 'Form_Resize' AND Function = 'ResizeControls' AND Enabled = TRUE;"

    Set rstScreenControl = CurrentDb.OpenRecordSet(strSQL)

    'Alex C 2/12/2012 - wrapped check for no records around this for screens with no resize entry
    If Not (rstScreenControl.BOF And rstScreenControl.EOF) Then
        rstScreenControl.MoveFirst
    End If

    Do While Not (rstScreenControl.BOF Or rstScreenControl.EOF)
        For Each ctl In frm.Controls
        
            If ctl.Name = rstScreenControl!Control Then
            
                strParameters = Split(rstScreenControl!Parameters, ",")
                
                'Get DistanceFromRight parameter
                If UBound(strParameters) >= 0 Then
                    intDistanceFromRight = val(strParameters(0))
                Else
                    intDistanceFromRight = -1
                End If

                'Get DistanceFromBottom parameter
                If UBound(strParameters) >= 1 Then
                    intDistanceFromBottom = val(strParameters(1))
                Else
                    intDistanceFromBottom = -1
                End If
            
                If intDistanceFromRight >= 0 Then
                    intWidth = (frm.InsideWidth - ctl.left) - intDistanceFromRight
                    If intWidth >= 0 Then
                        ctl.Width = intWidth
                    End If
                End If
                
                If intDistanceFromBottom >= 0 Then
                    intHeight = (frm.InsideHeight - ctl.top) - intDistanceFromBottom
                    If intHeight >= 0 Then
                        ctl.Height = intHeight
                    End If
                End If
                
                Exit For
            End If
        Next ctl

        Erase strParameters
        rstScreenControl.MoveNext
    Loop
    rstScreenControl.Close

EXIT_HERE:
    Exit Sub
    
Err_handler:
    MsgBox Err.Description, vbExclamation

End Sub

Public Sub GetImagePageCounts(frm As Form)
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim iCountAllowance As Integer
    Dim strSQL As String
    
    Dim strFileName As String
    Dim strFileExt As String
    Dim strErrMsg As String
    
    
    Dim strErrSource As String
    strErrSource = "CountImagePages"
    
    Dim bResult As Boolean
    Dim iResult As Long
    
    On Error GoTo Err_handler
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
   
    iCountAllowance = 5
    strSQL = "select * from SCANNING_Image_Log where ValidationDt = '1/1/1900' or isnull(ValidationDt,'') = ''"
    Set rs = MyAdo.OpenRecordSet(strSQL)
   
    If (rs.BOF = True And rs.EOF = True) Then
        MsgBox "Nothing to do"
    Else
        rs.MoveFirst
        With rs
            While Not .EOF
                'get page count
                strFileName = !ImagePath
                
                'set default
                !ValidationDt = "1/1/1900"
                
                If UCase(Right(strFileName, 3)) = "PDF" Then
                    !PDFCnt = Count_PDF_Pages(!ImagePath)
                    !TIFCnt = !PDFCnt
                Else
                    !TIFCnt = Count_TIF_Pages(!ImagePath)
                    !PDFCnt = !TIFCnt
                End If
                
                ' update validation
                If !TIFCnt >= !PageCnt Then
                    !ValidationDt = Date
                End If
                    
                ' update display
                If frm.lstFiles.ListCount > 30 Then
                    frm.lstFiles.RemoveItem (0)
                End If
                frm.lstFiles.AddItem !TIFCnt & ";" & !LocalPath
                    
                .Update
                MyAdo.BatchUpdate rs
                .MoveNext
            Wend
        End With
    End If
    
    
Exit_Sub:
    Set MyAdo = Nothing
    Set rs = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & strErrSource & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Sub

Public Sub RunScheduledJobs()
    Dim MyAdo As New clsADO
    Dim myCode_ADO As New clsADO
    Dim rs As ADODB.RecordSet
    
    On Error GoTo Err_handler
    
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = "select 1 from sysobjects where name = 'ADMIN_Scheduled_Jobs' and type = 'U'"
    Set rs = MyAdo.OpenRecordSet
    
    If rs.BOF = True And rs.EOF = True Then
        ' database is not setup for running scheduled jobs
        Exit Sub
    End If
    
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.sqlString = "exec usp_ADMIN_Run_Scheduled_Jobs"
    myCode_ADO.Execute

Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Exit Sub
    
Err_handler:
    MsgBox "Error executing scheduled jobs: " & vbCrLf & Err.Description
    Resume Exit_Sub
End Sub