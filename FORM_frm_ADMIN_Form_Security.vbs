Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "FormSecurity"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property



Private Sub CtrlName_Enter()
    Dim frm As AccessObject
    Dim ctrl As Control
    Dim bCtrlFound As Boolean
    
    Me.Caption = "Form Security Control"
    
    If Me.Form_Name & "" = "" Then
        Me.CtrlName.RowSource = ""
        Exit Sub
    End If
    
'    If Me.Form_Name.Value = Me.Form_Name.OldValue Then
'        If Me.CtrlName.RowSource <> "" Then Exit Sub
'    End If
    
    Me.CtrlName.RowSource = ""
    
    
    Set frm = CurrentProject.AllForms(Me.Form_Name)
    If frm.IsLoaded = False Then
        DoCmd.OpenForm frm.Name, acDesign
        Forms(frm.Name).visible = False
    End If
    
    For Each ctrl In Forms(frm.Name).Controls
        If DLookup("CtrlName", "ADMIN_Form_Security", "FormName = '" & Form_Name & "' and CtrlName = '" & ctrl.Name & "'") & "" = "" Then
            Me.CtrlName.AddItem ctrl.Name
        End If
        If ctrl.Name = Me.CtrlName Then bCtrlFound = True
    Next
    
    If Not bCtrlFound Then
        Me.CtrlName = ""
    End If
    
    On Error Resume Next
    DoCmd.Close acForm, frm.Name, acSaveNo
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Form_Name & "" = "" Then
        MsgBox "Form_Name can not be blank", vbCritical
        Cancel = True
        Exit Sub
    End If
    
    If CtrlName & "" = "" Then
        MsgBox "CtrlName can not be blank", vbCritical
        Cancel = True
        Exit Sub
    End If
    
    If UserAccess & "" = "" Then
        MsgBox "User access can not be blank", vbCritical
        Cancel = True
        Exit Sub
    End If
    
    If IsNull(Action) Then
        MsgBox "Action can not be blank.", vbCritical
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim Action(8, 2)
    Dim i As Integer
    
    Call Account_Check(Me)
    
    Dim iAppPermission As Integer
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    Action(1, 1) = "Allow Add"
    Action(1, 2) = 1
    Action(2, 1) = "Allow Change"
    Action(2, 2) = 2
    Action(3, 1) = "Allow Delete"
    Action(3, 2) = 4
    Action(4, 1) = "Allow View"
    Action(4, 2) = 8
    Action(5, 1) = "Allow Reassign"
    Action(5, 2) = 16
    Action(6, 1) = "Allow Forward"
    Action(6, 2) = 32
    Action(7, 1) = "Allow Release Claim"
    Action(7, 2) = 64
    Action(8, 1) = "Allow Print Letter"
    Action(8, 2) = 128
    
    Me.Action.RowSource = ""
    Me.CtrlName.RowSource = ""
    For i = 1 To UBound(Action)
        Me.Action.AddItem (Action(i, 2) & ";" & Action(i, 1))
    Next i
    

End Sub

Private Sub Form_Name_AfterUpdate()
    If Me.Form_Name.Value & "" <> Me.Form_Name.OldValue & "" Then
        Me.CtrlName = ""
    End If
    
    If Me.Form_Name & "" <> "" Then
        Me.Form_Name.DefaultValue = Chr(34) & Me.Form_Name.Value & Chr(34)
    End If
End Sub
