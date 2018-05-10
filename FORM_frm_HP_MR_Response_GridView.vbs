Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const ColorRed = 8421631
Const ColorNormal = -2147483643


Private Sub CnlyClaimNum_DblClick(Cancel As Integer)
    
    If Me.CnlyClaimNum & "" <> "" Then
        Me.Parent.DisplayClaimScreen Me.CnlyClaimNum
    End If
    
End Sub

Private Sub Form_Click()
    If IsSubForm(Me) Then
        If Me.SelHeight > 1 Then
            Me.Parent.StartRow = Me.SelTop
            Me.Parent.RowsSelected = Me.SelHeight
        Else
            Me.Parent.StartRow = Me.SelTop
            Me.Parent.RowsSelected = 1
        End If
    End If
End Sub

Public Sub Form_Current()
    Dim strSQL As String
    
    If IsSubForm(Me) Then
        If Me.Parent.Loaded = True Then
            If Me.MRRID & "" <> "" Then
                'strSQL = "select * from v_HP_MR_Consolidated_View where MRRID = " & Me.MRRID
                strSQL = "select * from v_HP_MR_Consolidated_View where MRRID = " & Me.Parent.frm_HP_MR_Response_GridView.Form.MRRID  'imagename  = '" & Me.ImageName & "'"
            Else
                strSQL = "select * from v_HP_MR_Consolidated_View where 1=2"
            End If
            Me.Parent.Form.frm_HP_MR_Response.Form.RecordSource = strSQL
            If Me.Parent.Form.frm_HP_MR_Response.Form.AckCode <> "1" Then
                Me.Parent.Form.frm_HP_MR_Response.Form("AckCode").BackColor = ColorRed
                Me.Parent.Form.frm_HP_MR_Response.Form("AckDesc").BackColor = ColorRed
            Else
                Me.Parent.Form.frm_HP_MR_Response.Form("AckCode").BackColor = ColorNormal
                Me.Parent.Form.frm_HP_MR_Response.Form("AckDesc").BackColor = ColorNormal
            End If
        End If
    End If
    

        
    
End Sub


Private Sub Form_DblClick(Cancel As Integer)

    If Me.CnlyClaimNum & "" <> "" Then
        Me.Parent.DisplayClaimScreen Me.CnlyClaimNum
    End If
    

End Sub

Private Sub Form_Load()
    Call CreatePK("v_HP_MR_Consolidated_View", "MRRID")
    Me.RecordSource = "select * from v_HP_MR_Consolidated_View order by MRRID "
    Me.RecordSet.MoveFirst
End Sub

Private Sub CreatePK(ByVal TableName As String, ByVal Fields As String)
On Error Resume Next
    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & TableName & " On " & TableName & "(" & Fields & ")"
End Sub
