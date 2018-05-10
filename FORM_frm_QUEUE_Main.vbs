Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrUserProfile As String

Const CstrFrmAppID As String = "QueueMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Close()
    Me.sub_form.SourceObject = Me.sub_form.Tag
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Queue Maintenance"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    
    mstrUserProfile = GetUserProfile()
    
    lstAppPanel.RowSource = GetListBoxSQL(Me.Name, mstrUserProfile)
    lstAppPanel.Requery
    Me.sub_form.SourceObject = Me.sub_form.Tag
    Me.sub_form.visible = False
    Me.lblSubAppTitle.visible = False
End Sub

Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub


Private Sub lstAppPanel_Click()
'    Dim rs As DAo.Recordset        '' KD 20120910 - Change to ADO
    Dim rs As ADODB.RecordSet
    Dim oAdo As clsADO
    
            '' KD 20120910 - Change to ADO
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = GetListBoxRowSQL(lstAppPanel, Me.Name)
        Set rs = .ExecuteRS
    End With
    
'    Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstAppPanel, Me.Name), dbOpenSnapshot, dbSeeChanges)
    '' / KD 20120910 - Change to ADO
    
    If Not (rs.BOF And rs.EOF) Then
        If UCase(left(rs("SQLValue"), 4)) = "USER" Then
            Me.OperMode = OperationMode.User
        Else
            Me.OperMode = OperationMode.Manager
        End If
        Me.Parameter = rs("SQLValue")
        Me.sub_form.SourceObject = ""
        lblSubAppTitle.Caption = rs("TabName")
        Me.lblSubAppTitle.visible = True
        Me.sub_form.visible = True
        Me.sub_form.SourceObject = rs("FormName")
        Select Case rs("FormName")
            Case "frm_RPT_Generic_Report_OLDVERSION"
                Me.sub_form.Form.StoredProcName = rs("FormValue")
        End Select
    Else
        MsgBox "Application form has not been defined"
    End If
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close '' KD 20120910 - Change to ADO
    End If
    Set rs = Nothing
    Set oAdo = Nothing
    
End Sub
