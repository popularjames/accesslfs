Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AcctSel"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cboAccountID_AfterUpdate()
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim strParentFormCaption As String
    
    ' reset display
    lblAccountAbreviation.visible = False
    lblAccountDesc.visible = False
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from ADMIN_Client_Account where AccountID = " & cboAccountID
    Set rs = MyAdo.OpenRecordSet
    
    ' reset parent form caption
    If IsSubForm(Me) Then
        strParentFormCaption = Me.Parent.Form.Caption
        If strParentFormCaption = "" Then
            strParentFormCaption = rs("AcctDesc")
        ElseIf gstrAcctDesc <> "" And left(strParentFormCaption, Len(gstrAcctDesc)) = gstrAcctDesc Then
            strParentFormCaption = rs("AcctDesc") & Mid(strParentFormCaption, Len(gstrAcctDesc) + 1)
        Else
            strParentFormCaption = rs("AcctDesc") & " - " & strParentFormCaption
        End If
        Me.Parent.Form.Caption = strParentFormCaption
    End If
    
    gintAccountID = cboAccountID
    gstrAcctAbbrev = rs("AcctAbbrev")
    gstrAcctDesc = rs("AcctDesc")
    
    lblAccountAbreviation.Caption = "Abbreviation:" & Space(7) & gstrAcctAbbrev
    lblAccountDesc.Caption = "Description:" & Space(10) & gstrAcctDesc
    lblAccountAbreviation.visible = True
    lblAccountDesc.visible = True
    
    
    Set rs = Nothing
    Set MyAdo = Nothing
End Sub

Private Sub cmdOk_Click()
    Dim frm As Form
    Dim i As Integer
    
    If gintAccountID = 0 Then
        MsgBox "Please select an Account first"
    Else
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
End Sub

Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub

Private Sub Form_Load()
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    
    Me.Caption = "Account Selection"
    
   
    ' check if user is associated with more than 1 account.
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select AccountID from ADMIN_User_Account where UserID='" & Identity.UserName() & "'"
    Set rs = MyAdo.OpenRecordSet
    If rs.recordCount = 1 And gintAccountID <> rs("AccountID") Then
        gintAccountID = rs("AccountID")
        MyAdo.sqlString = "select * from ADMIN_Client_Account where AccountID = " & gintAccountID
        Set rs = MyAdo.OpenRecordSet
        gstrAcctAbbrev = rs("AcctAbbrev")
        gstrAcctDesc = rs("AcctDesc")
    End If
    
    If gintAccountID > 0 Then
        cboAccountID = gintAccountID
        lblAccountAbreviation.Caption = "Abbreviation:" & Space(7) & gstrAcctAbbrev
        lblAccountDesc.Caption = "Description:" & Space(10) & gstrAcctDesc
        lblAccountAbreviation.visible = True
        lblAccountDesc.visible = True
    End If
    
    If IsSubForm(Me) Then
        lblAppTitle.visible = False
        CmdOK.visible = False
    End If
    
    Set MyAdo = Nothing
End Sub
