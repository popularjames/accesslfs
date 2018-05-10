Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AdminMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Close()
    Me.sub_form.SourceObject = Me.sub_form.Tag
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Administration Maintenance"
    
    Call Account_Check(Me)
    
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    lstAppPanel.RowSource = GetListBoxSQL(Me.Name, gstrProfileID)
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

        lblSubAppTitle.Caption = rs("TabName")
        Me.sub_form.visible = True
        Me.lblSubAppTitle.visible = True
        Me.sub_form.SourceObject = rs("FormName")

        Select Case rs("FormName")
            Case "frm_AUDIT_TRACKING_Main"
                Me.sub_form.Form.AppTitle = rs("TabName")
                
                Select Case rs("RowSource")
                    Case "ADMIN_Action_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "Action"
                        Me.sub_form.Form.AuditTableName = "ADMIN_Action_Audit_Hist"

                    Case "ADMIN_Application_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "AppID"
                        Me.sub_form.Form.AuditTableName = "ADMIN_Application_Audit_Hist"

                    Case "ADMIN_App_Keys_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "AppKey"
                        Me.sub_form.Form.AuditTableName = "ADMIN_App_Keys_Audit_Hist"

                    Case "ADMIN_Audit_Number_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "AuditNum"
                        Me.sub_form.Form.AuditTableName = "ADMIN_Audit_Number_Audit_Hist"
                    
                    Case "ADMIN_Profile_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "Profile"
                        Me.sub_form.Form.AuditTableName = "ADMIN_Profile_Audit_Hist"
                        Me.sub_form.Form.AuditKey = "AccountID = " & gintAccountID
                    
                    Case "ADMIN_User_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "User"
                        Me.sub_form.Form.AuditTableName = "ADMIN_User_Audit_Hist"
                    
                    Case "ADMIN_User_Exception_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "UserException"
                        Me.sub_form.Form.AuditTableName = "ADMIN_User_Exception_Audit_Hist"

                    Case "ADMIN_User_Profile_Audit_Hist"
                        Me.sub_form.Form.frmAppID = "UserProfile"
                        Me.sub_form.Form.AuditTableName = "ADMIN_User_Profile_Audit_Hist"
                                    
                End Select

                Me.sub_form.Form.RefreshData
        End Select
    
    Else
        MsgBox "Application form as not been defined"
    End If
End Sub
