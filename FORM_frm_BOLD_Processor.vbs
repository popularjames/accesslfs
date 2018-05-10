Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private WithEvents coProcessor As clsProcessor


Private Sub cmdTestProcessor_Click()
    Set goBOLD_Processor = New clsBOLD_Processor
    
    goBOLD_Processor.StartProcessing
    
    

End Sub
