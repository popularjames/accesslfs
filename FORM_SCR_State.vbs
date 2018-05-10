Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Event RestoreAll()
Public Event RecordSourceOnly()
Public Event CollapseAll()

Private Const vbSpecialEffect_Raised = 1
Private Const vbSpecialEffect_Sunken = 2


Public Sub SetState_Collapsed(Optional raiseFormEvent As Boolean = False)
    SetState 2, raiseFormEvent
End Sub
Public Sub SetState_SelectionOnly(Optional raiseFormEvent As Boolean = False)
    SetState 1, raiseFormEvent
End Sub
Public Sub SetState_Restored(Optional raiseFormEvent As Boolean = False)
    SetState 0, raiseFormEvent
End Sub

Private Sub SetState(theFormState As Integer, Optional raiseFormEvent As Boolean = False)
    ' Always reset all toggle and label states first
    
    ' Restore
    Me.btnRestore.SpecialEffect = vbSpecialEffect_Raised
    ' Selection Only
    Me.btnSelection.SpecialEffect = vbSpecialEffect_Raised
    ' Collapse
    Me.btnCollapse.SpecialEffect = vbSpecialEffect_Raised


    Select Case theFormState
        Case 0 ' Restored
            Me.btnRestore.SpecialEffect = vbSpecialEffect_Sunken
        
            If raiseFormEvent Then
                RaiseEvent RestoreAll
            End If
        Case 1 ' SelectionOnly
            Me.btnSelection.SpecialEffect = vbSpecialEffect_Sunken
        
            If raiseFormEvent Then
                RaiseEvent RecordSourceOnly
            End If
        Case 2 ' Collapsed
            Me.btnCollapse.SpecialEffect = vbSpecialEffect_Sunken
        
            If raiseFormEvent Then
                RaiseEvent CollapseAll
            End If
    End Select
    
    Me.txtNull.SetFocus
End Sub

Private Sub btnRestore_Click()
    SetState 0, True
End Sub

Private Sub btnCollapse_Click()
    SetState 2, True
End Sub

Private Sub btnSelection_Click()
    SetState 1, True
End Sub

Private Sub Form_Load()
    Me.txtNull.SetFocus
End Sub
