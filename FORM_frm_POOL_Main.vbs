Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "PoolMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cmdClearStatus_Click()
    Me.lstStatus.RowSource = vbNullString
End Sub

Private Sub Form_Close()
    Me.sub_form.SourceObject = Me.sub_form.Tag
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Pool Main Screen"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    lstAppPanel.RowSource = GetListBoxSQL(Me.Name)
    lstAppPanel.Requery
    Me.sub_form.SourceObject = Me.sub_form.Tag
    Me.sub_form.visible = False
    Me.lblSubAppTitle.visible = False
    
End Sub


Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub

Private Sub lstAppPanel_Click()
    Dim rs As DAO.RecordSet
    
    Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstAppPanel, Me.Name), dbOpenSnapshot, dbSeeChanges)
    
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
        
    Else
        MsgBox "Application form has not been defined"
    End If
    
End Sub
