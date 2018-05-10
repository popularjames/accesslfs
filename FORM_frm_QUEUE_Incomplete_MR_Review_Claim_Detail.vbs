Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim ctl As Control

Private Sub Form_Load()
        Me.RecordSource = ""
        Call Get_Data
End Sub

Public Sub Clean_Fields()
            
            For Each ctl In Me.Controls
               If Controls(ctl.Name).ControlType = 109 Then
                    Me.Controls.Item(ctl.Name).Value = ""
               End If
            Next
            
End Sub


Public Sub Get_Data()
        Dim myCode_ADO As New clsADO
        Dim rs As ADODB.RecordSet
        Dim strSQL As String
        'Dim ctl As Control

        On Error GoTo Err_handler

        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.SQLTextType = StoredProc
        myCode_ADO.sqlString = "usp_Incomplete_MR_Claim_Details"
        myCode_ADO.Parameters("@pCnlyClaimNum") = [Forms]![frm_QUEUE_Incomplete_MR]![frm_QUEUE_MR_Request_Sub].Form![txtCnlyClaimNum]
        Set rs = myCode_ADO.ExecuteRS

         If rs.EOF = True Then
            Call Clean_Fields
         Else
         
         rs.MoveFirst
         
            For Each ctl In Me.Controls
               If Controls(ctl.Name).ControlType = 109 Then
                   Me.Controls(ctl.Name) = rs(ctl.Name).Value
               End If
            Next
            
         End If
         
Exit_Sub:

        Set myCode_ADO = Nothing
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
        Exit Sub

Err_handler:
        MsgBox "Error populating Queue Incomplete MR Review Claim Detail Form: " & Err.Description
        Resume Exit_Sub
End Sub
