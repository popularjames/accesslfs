Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum mode
    Grouping = 1
    Calculations = 0
End Enum

Private mvMode As mode
Private mvResults As Integer



Public Property Get RunMode() As mode
    RunMode = mvMode
End Property
Public Property Let RunMode(Value As mode)

        mvMode = Value
        
        Select Case mvMode
        Case mode.Calculations
            Me.RecordSource = "SCR_ScreensTotalsCalculations"
            Me.CmboAggr.visible = True
            LblCmboAggr.visible = True
        Case mode.Grouping
            Me.RecordSource = "SCR_ScreensTotalsFields"
            Me.CmboAggr.visible = False
            LblCmboAggr.visible = False
            
        End Select

End Property
Public Property Get Results() As Integer
    Results = mvResults
End Property


Private Sub cmdOk_Click()
    Me.Dirty = False
    mvResults = True
    Me.visible = False
    
End Sub

Private Sub Form_Open(Cancel As Integer)
mvResults = 1
End Sub
