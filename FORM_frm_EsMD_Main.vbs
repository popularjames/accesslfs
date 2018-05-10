Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Sub DisplayClaimScreen(CnlyClaimNum As String)
    Dim frm_AUDITCLM_Main As Form_frm_AUDITCLM_Main
    Dim iWindowHandle As Long
    Dim f As clsWindowHandles
    Dim bFound As Boolean
    
   
    bFound = False
    
    If CnlyClaimNum & "" <> "" Then
        For Each f In ColWindows
            'Debug.Print f.WindowName
            'Debug.Print f.WindowHandle
            iWindowHandle = f.WindowHandle
            SetForegroundWindow iWindowHandle
            If f.WindowName = "ClaimMain" & CnlyClaimNum Then
                SetForegroundWindow iWindowHandle
                
                If screen.ActiveForm.hwnd = iWindowHandle Then
                    bFound = True
                    Exit For
                End If
            End If
        Next
    
        If Not bFound Then
            Set f = New clsWindowHandles
            Set frm_AUDITCLM_Main = New Form_frm_AUDITCLM_Main
            f.WindowHandle = frm_AUDITCLM_Main.hwnd
            f.WindowName = "ClaimMain" & CnlyClaimNum
            ColWindows.Add f, f.WindowHandle & ""
            ColObjectInstances.Add Item:=frm_AUDITCLM_Main, Key:=frm_AUDITCLM_Main.hwnd & " "
            
            frm_AUDITCLM_Main.Caption = "CMS: ClaimNum : " & CnlyClaimNum
    
            frm_AUDITCLM_Main.visible = True
            frm_AUDITCLM_Main.CnlyClaimNum = CnlyClaimNum
            frm_AUDITCLM_Main.LoadData
        End If
    End If
    
    Set f = Nothing

End Sub

Sub getTransID(filter As String)
  
  Forms!frm_esMD_main!.v_esMD_Consolidated_View.Requery
  
  Select Case filter
        Case "ID"
            If Me.cboTransID <> "" Then
                Forms!frm_esMD_main!.v_esMD_Consolidated_View.Form.RecordSource = "select * from v_esMD_Consolidated_View where TransactionID = '" & Me.cboTransID.Value & "'"
            End If
        Case "All"
            Forms!frm_esMD_main!.v_esMD_Consolidated_View.Form.RecordSource = "select * from v_esMD_Consolidated_View"
   End Select
   
End Sub


Private Sub cboTransID_AfterUpdate()
    Call getTransID("ID")
End Sub

Private Sub cmdSearch_Click()
    Call getTransID("ID")
End Sub

Private Sub cmdViewAll_Click()
    Call getTransID("All")
End Sub
