Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2014-07-16 VS: Added ALJ Fax Queue prototype
Option Compare Database
Option Explicit
Private Const PACKAGE As String = "Package Name"
Private Const Icn As String = "ICN"
Private Const AP_NUM As String = "ALJ Appeal Number"
Private Const RESP_DATE As String = "NOI Response Date"
Private Const HEAR_DATE As String = "Hearing Date"
Private Const Judge As String = "Judge Name"

Private Sub Form_Load()
   Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where PackageName = '" & Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr & "'"
End Sub

Private Sub cmdSearch_Click()
  
  On Error GoTo Block_Err
  If Me.cboSearchBy = PACKAGE Or Me.cboSearchBy = Judge Or Me.cboSearchBy = HEAR_DATE Then
    Me.frm_ALJ_Fax_Queue_Pkg.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg where PackageName like " & "'*" & Replace(Me.txtSearchFor, "'", "''") & "*'"
    Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where PackageName = '" & Replace(Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr, "'", "''") & "'"
  End If
  
 If Me.cboSearchBy = RESP_DATE Then
    Me.frm_ALJ_Fax_Queue_Pkg.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg where NOIResponseDate like " & "'*" & Me.txtSearchFor & "*'"
    Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where PackageName = '" & Replace(Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr, "'", "''") & "'"
 End If
  
 If Me.cboSearchBy = Icn Then
    Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where ICN like " & "'*" & Me.txtSearchFor & "*'"
    Me.frm_ALJ_Fax_Queue_Pkg.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg where PackageName = '" & Replace(Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr, "'", "''") & "'"
 End If
  
  If Me.cboSearchBy = AP_NUM Then
    Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where ALJAppealNumber like " & "'*" & Me.txtSearchFor & "*'"
    Me.frm_ALJ_Fax_Queue_Pkg.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg where PackageName = '" & Replace(Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr, "'", "''") & "'"
 End If
  
Block_Exit:
    Exit Sub
Block_Err:
    MsgBox ("Invalid Search Results!")
    Me.frm_ALJ_Fax_Queue_Pkg.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg"
    Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where PackageName = '" & Replace(Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr, "'", "''") & "'"
    GoTo Block_Exit
  
End Sub

Private Sub cmdRefresh_Click()
Me.frm_ALJ_Fax_Queue_Pkg.Requery
Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Requery
End Sub

Private Sub cmdClear_Click()
Me.frm_ALJ_Fax_Queue_Pkg.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg"
Me.frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where PackageName = '" & Replace(Me.frm_ALJ_Fax_Queue_Pkg.Form.txtPackageNameHdr, "'", "''") & "'"

End Sub
