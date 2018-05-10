Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_NOTE_Detail_Display
' Author:      Barbara Dyroff
' Create Date: 2010-05-17
' Description:
'   Display Notes.  Sort the data.
'
' Modification History:
'
' =============================================

Private Sub Form_Open(Cancel As Integer)
    Me.OrderBy = "SeqNo DESC"
    Me.OrderByOn = True
End Sub
