Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()
    Me.sub_form.SourceObject = Me.sub_form.Tag
End Sub

Private Sub Form_Load()
    Me.Caption = "Cross Reference Code Maintenance"
    Call Account_Check(Me)
    
    lstAppPanel.RowSource = GetListBoxSQL(Me.Name)
    lstAppPanel.Requery
    Me.sub_form.SourceObject = Me.sub_form.Tag
    Me.sub_form.visible = False
    Me.lblSubAppTitle.visible = False
End Sub

Private Sub lstAppPanel_Click()
    Dim rs As DAO.RecordSet
    
    Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstAppPanel, Me.Name), dbOpenSnapshot, dbSeeChanges)

    If Not (rs.BOF And rs.EOF) Then
        lblSubAppTitle.Caption = rs("TabName")
        Me.sub_form.visible = True
        Me.lblSubAppTitle.visible = True
        Me.sub_form.SourceObject = rs("FormName")
    Else
        MsgBox "Application form as not been defined"
    End If
End Sub
