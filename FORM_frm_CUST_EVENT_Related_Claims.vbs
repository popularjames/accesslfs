Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mbIsLoaded As Boolean
'Identify whether all the related claims' topics were selected to make sure they are before closing form
Public Function areTopicsSet() As Boolean
Dim rs As ADODB.RecordSet

    Set rs = Me.RecordSet.Clone
    If rs.BOF And rs.EOF Then
        areTopicsSet = True
        Exit Function
    End If
        
    rs.MoveFirst
    While rs.BOF = False And rs.EOF = False
        If rs("TopicID") = 0 Or Nz(rs("TopicID"), "") = "" Then
            areTopicsSet = False
            Exit Function
        End If
        rs.MoveNext
    Wend

    areTopicsSet = True
    
End Function



Private Sub Form_Current()
    
    If Me.Parent Is Nothing Then
    Else
'        If mbIsLoaded = True And Me.Parent.IsLoaded = True And Me.Parent.IsRefreshing = False Then
        If mbIsLoaded = True And Me.Parent.IsRefreshing = False Then
            'Get the parent to refresh the screen for the new selected claim
            If Me.Parent.RelatedClaimsSetup = True Then
                Me.Parent.RefreshCurrentClaim (Me.CnlyClaimNum)
            End If
        End If
    End If
    
End Sub


Private Sub Form_Load()
    mbIsLoaded = True
End Sub
'Set the updated topic id in the current related claim record and save the record
Private Sub SelectedTopic_Change()
Dim rs As ADODB.RecordSet
Dim TopicID As Integer
Dim CnlyClaimNum As String
Dim returnCode As Integer

    Set rs = Parent.rsRelatedClaims
    
'    CnlyClaimNum = rs("CnlyClaimNum")
'    rs.MoveFirst
'    While Not (rs.BOF And rs.EOF) And CnlyClaimNum <> rs("CnlyClaimNum")
'        rs.MoveNext
'    Wend
    
'    If rs.BOF And rs.EOF Then
'        MsgBox "The related claim row could not be found to set the topic.", vbOKOnly
'        rs.MoveFirst
'        Exit Sub
'    End If
    
'    topicID = SelectedTopic
'    Parent.rsRelatedClaims("TopicID") = topicID
    TopicID = Me.SelectedTopic
    CnlyClaimNum = Me.CnlyClaimNum
    
    If Me.Parent.frmCustEvent.Controls("cmb_EventTopic").ListIndex = -1 Then
        Me.Parent.frmCustEvent.Controls("cmb_EventTopic").Value = SelectedTopic
    End If
    returnCode = Me.Parent.UpdateRelatedClaimTopic(CnlyClaimNum, TopicID)

    
'    rs.UpdateBatch
'    rs.MoveFirst
    
End Sub
