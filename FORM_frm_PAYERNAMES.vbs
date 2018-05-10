Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private cbShowPastPayers As Boolean



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get ShowPastPayers() As Boolean
    ShowPastPayers = cbShowPastPayers
End Property
Public Property Let ShowPastPayers(bShowPreviousPayers As Boolean)

    cbShowPastPayers = bShowPreviousPayers
    If bShowPreviousPayers = True Then
        If InStr(1, Me.filter, "EndDate >", vbTextCompare) < 1 Then
            Me.filter = ""
            Me.FilterOn = False
        End If
    Else
    
        If Me.filter = "" Then
            Me.filter = " EndDate >#" & Format(Now, "m/d/yyyy") & "#"
        Else
            If InStr(1, Me.filter, "EndDate >", vbTextCompare) < 1 Then
                Me.filter = Me.filter & " AND EndDate >#" & Format(Now, "m/d/yyyy") & "#"
            End If
        End If
    Me.FilterOn = True
    End If
End Property


Public Property Get GetSelectedPayerNameIDs(Optional ByRef sNames As String) As String
On Error GoTo Block_Exit
Dim strProcName As String
Dim oRs As RecordSet
Dim sRet As String

    strProcName = ClassName & ".GetSelectedPayerNameIDs"
    sNames = ""

    Set oRs = Me.RecordsetClone
    
    If Not oRs.EOF And Not oRs.BOF Then
    
    With oRs
        .MoveFirst
        While Not .EOF
            If oRs("Selected").Value = True Then
                sRet = sRet & CStr("" & oRs("PayerNameId").Value) & ","
                sNames = sNames & CStr("" & oRs("PayerName").Value) & ","
            End If
            .MoveNext
        Wend
    End With
    End If
    
    If Right(sRet, 1) = "," Then sRet = left(sRet, Len(sRet) - 1)
    If Right(sNames, 1) = "," Then sNames = left(sNames, Len(sNames) - 1)
    
    
Block_Exit:
    GetSelectedPayerNameIDs = sRet
    Exit Property
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Property




Public Property Get AtLeastOnePayerSelected() As Boolean
Dim oRs As RecordSet

    Set oRs = Me.RecordsetClone
    oRs.MoveFirst
    While Not oRs.EOF
        ' Ignore "all"
        If oRs("PayerName").Value <> "All" Then
            If oRs("Selected").Value = True Then
                AtLeastOnePayerSelected = True
                GoTo Block_Exit
            End If
        End If
        oRs.MoveNext
    Wend

Block_Exit:
    Set oRs = Nothing
    Exit Property
Block_Err:
    ReportError Err, ClassName & ".AtLeastOnePayerSelected"
    GoTo Block_Exit
End Property






Private Sub ckPayer_Click()
Dim bChecked As Boolean

    If Me.PayerName = "All" Then
        bChecked = Me.RecordSet("Selected").Value
        
        Call CheckAll(bChecked)

    End If
End Sub

Private Sub CheckAll(bAllSelected As Boolean)
Dim oRs As RecordSet



    Set oRs = Me.RecordSet

    While Not oRs.EOF
        If oRs("PayerName").Value <> "All" Then
            If oRs("ExcludeFromAll").Value <> 0 Then
                Me.ckPayer.Value = Not bAllSelected
            Else
                Me.ckPayer.Value = bAllSelected
            End If
        End If
        oRs.MoveNext
    Wend

End Sub

Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oDb As DAO.Database

Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sLocalTableName As String

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = "SELECT * FROM XREF_PAYERNAMES WHERE ForUserDisplay = 1 ORDER BY PayerName "
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "There was a problem retrieving payer details"
            GoTo Block_Exit
        End If
    End With

    sLocalTableName = CopyDataToLocalTmpTable(oRs, False)

    Me.RecordSource = sLocalTableName
    
'    Set oDb = CurrentDb()
'    oDb.Execute ("DELETE * FROM tmp_SelectPayerNames")
''    oDb.Execute ("TRUNCATE TABLE tmp_SelectPayerNames ")
'    ' populate it:
'    oDb.Execute ("INSERT INTO tmp_SelectPayerNames (Selected, PayerNameId, PayerName, ExcludeFromAll) SELECT False, * FROM XREF_PAYERNAMES")
'
'    Me.RecordSource = "tmp_SelectPayerNames"

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
