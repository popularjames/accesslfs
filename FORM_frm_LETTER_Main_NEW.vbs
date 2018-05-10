Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "LetterQueueMain"

Private ciTabToLoad As Integer
Private ciBatchToSelect As Integer



Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get LoadNextTab() As Integer
    LoadNextTab = ciTabToLoad
End Property
Public Property Let LoadNextTab(iLoadNextTab As Integer)
    ciTabToLoad = iLoadNextTab
    Me.TimerInterval = 500
End Property

Public Property Get BatchIdToSelect() As Integer
    BatchIdToSelect = ciBatchToSelect
End Property
Public Property Let BatchIdToSelect(iBatchIdToSelect As Integer)
    ciBatchToSelect = iBatchIdToSelect
End Property



Private Sub Form_Close()
    Me.sub_form.SourceObject = Me.sub_form.Tag
End Sub


Private Sub Form_Load()
Dim iAppPermission As Integer
Dim sSupervisor As String
    
    Me.Caption = "Letter Maintenance"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then
        Stop
        Me.visible = False
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    '' Ok, if we get here and we are the 'Scanning" supervisorid then
    '' open the new BOLD dashboard
    '' Which we should have a button on it to open the old version..
    '' so, if there are openargs then continue on - otherwise go for the new form
'    If Nz(Me.OpenArgs, "") = "" Then
'        sSupervisor = Identity.UserSupervisorId()
'        Select Case UCase(sSupervisor)
'        Case "ADMIN", "DATA CENTER", "MANAGEMENT", "SCANNING"
'
'            DoCmd.OpenForm "frm_BOLD_Mail_Dashboard", acNormal
'            Me.visible = False
'            DoCmd.Close acForm, Me.Name
'            Exit Sub
'        End Select
'    End If
    
    lstAppPanel.RowSource = GetListBoxSQL(Me.Name)
    lstAppPanel.Requery
    Me.sub_form.SourceObject = Me.sub_form.Tag
    Me.sub_form.visible = False
    Me.lblSubAppTitle.visible = False
End Sub


Private Sub Form_Resize()
    Me.sub_form.Width = Me.InsideWidth - Me.sub_form.left - 300
    Me.sub_form.Height = Me.InsideHeight - Me.sub_form.top - 300
End Sub

'Private Sub Form_Resize()
''    ResizeControls Me.Form
'End Sub

Private Sub Form_Timer()
Dim vItem As Variant
Dim iItmCnt As Integer

    Me.TimerInterval = 0
    iItmCnt = Me.LoadNextTab
    
    If iItmCnt <> 0 Then
        If iItmCnt <= Me.lstAppPanel.ListCount Then
            Me.lstAppPanel.Selected(iItmCnt - 1) = True
'            Stop
            Me.lstAppPanel = Me.lstAppPanel.ItemData(iItmCnt - 1)
        End If
        ' now call the function...
        
        Call lstAppPanel_Click
    End If
    
End Sub



Private Sub lstAppPanel_Click()
Dim rs As DAO.RecordSet
Dim oFrm As Form_frm_LETTER_PrintLabels

    Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstAppPanel, Me.Name), dbOpenSnapshot, dbSeeChanges)
    If Not (rs.BOF And rs.EOF) Then
        Me.sub_form.SourceObject = ""
        lblSubAppTitle.Caption = rs("TabName")
        Me.sub_form.SourceObject = rs("FormName")
        Me.sub_form.visible = True
        Me.lblSubAppTitle.visible = True
        
        If Me.BatchIdToSelect <> 0 Then
            If UCase(rs("TabName")) = UCase("03 Print Labels") Then
                Set oFrm = Me.sub_form.Form
                oFrm.BatchToSelect = BatchIdToSelect
                
                Me.BatchIdToSelect = 0  ' reset..
            End If
        End If
        
    Else
        MsgBox "Application form has not been defined"
    End If
End Sub
