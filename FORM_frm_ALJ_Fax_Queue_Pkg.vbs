Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2014-07-16 VS: Added ALJ Fax Queue Package Level prototype
Option Compare Database

Private Sub Form_Click()
Me.Parent!frm_ALJ_Fax_Queue_Pkg_Dtl.Form.RecordSource = "select * from v_ALJ_Fax_Queue_Pkg_Dtl where PackageName = '" & Replace(Me.txtPackageNameHdr, "'", "''") & "'"

End Sub
