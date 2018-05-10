Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvValUpdates As Boolean

Private Sub CmboDsBs_AfterUpdate()
    ValueUpdated
End Sub

Private Sub CmboDsCe_AfterUpdate()
    ValueUpdated
    'Disable all of the remaining controls
    Dim BlEnable As Boolean
    If CmboDsCe <> 0 Then 'Only allow exteded styles when flat
        BlEnable = False
    Else
        BlEnable = True
    End If
    
    CmdColorDsBc.Enabled = BlEnable
    CmboDsBs.Enabled = BlEnable
    CmboDsGL.Enabled = BlEnable
    CmdColorDsGl.Enabled = BlEnable
End Sub

Private Sub CmboDsGL_AfterUpdate()
    ValueUpdated
End Sub

Private Sub CmboDsHdBs_AfterUpdate()
    ValueUpdated
End Sub

Private Sub cmdApply_Click()
    SaveChanges
    GetGeneral 'refresh new values
    MvValUpdates = False 'reset flag in case user wants to update again
End Sub

Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub CmdColorDsBc_Click()
    Dim Clr As Long
    Clr = clrDsBC.BackColor
    Clr = ChooseColor(Clr, Me.hwnd)
    If Clr <> clrDsBC.BackColor Then
        clrDsBC.BackColor = Clr
        ValueUpdated
    End If
End Sub

Private Sub CmdColorDsGl_Click()
    Dim Clr As Long
    Clr = clrDsGL.BackColor
    Clr = ChooseColor(Clr, Me.hwnd)
    If Clr <> clrDsGL.BackColor Then
        clrDsGL.BackColor = Clr
        ValueUpdated
    End If
End Sub

Private Sub CmdDsDefaults_Click()
    Identity.DataSheetStylesDefaults
    GetDataSheetStyles
End Sub

Private Sub CmdFont_Click()
    Dim cls As New CT_ClsFont
    
    With cls
        .PropertiesFromControl Me.TxtFont
        .ShowEffects = True
        .ShowSize = True
        If .DialogFont = True Then
            .PropertiesToControl TxtFont
            ValueUpdated
        End If
    End With
    
    Set cls = Nothing
End Sub

Private Sub cmdOkay_Click()
    SaveChanges
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Sub SaveChanges()
    Me.CmdOkay.SetFocus
    SetDataSheetStyles
    SetGeneral
    Me.CmdApply.Enabled = False
End Sub

Private Sub CmdTxtFolderOutput_Click()
    Dim StPath As String
    StPath = TxtFolderOutput
    StPath = BrowseForFolder(TxtFolderOutput, "Default Output Path", Me.hwnd)
    If StPath <> TxtFolderOutput And "" & StPath <> "" Then
        TxtFolderOutput = StPath
        ValueUpdated
    End If
End Sub

Private Sub Form_Load()
    GetGeneral
    GetDataSheetStyles
End Sub


Private Sub SetDataSheetStyles()
    Dim Dss As CnlyDataSheetStyle
    
    With Dss
        .HeaderUnderlineStyle = CmboDsHdBs
        .GridlinesBehavior = CmboDsGL
        .GridlinesColor = clrDsGL.BackColor
        .BackGroundColor = clrDsBC.BackColor
        .BorderLineStyle = CmboDsBs
        .CellsEffect = CmboDsCe
        .fontsize = TxtFont.fontsize
        .FontFamily = TxtFont.FontName
        .FontItalic = TxtFont.FontItalic
        .FontUnderline = TxtFont.FontUnderline
        .ForeColor = TxtFont.ForeColor
        .FontWeight = TxtFont.FontWeight
    End With
    Identity.DataSheetStyle = Dss
End Sub


Private Sub SetGeneral()
    Identity.FolderOutput = "" & TxtFolderOutput
    Identity.Auditor = "" & TxtAuditor
End Sub
Private Sub GetDataSheetStyles()
    With Identity.DataSheetStyle
        CmboDsHdBs = .HeaderUnderlineStyle
        CmboDsGL = .GridlinesBehavior
        clrDsGL.BackColor = .GridlinesColor
        clrDsBC.BackColor = .BackGroundColor
    
        CmboDsBs = .BorderLineStyle
        CmboDsCe = .CellsEffect
        TxtFont.fontsize = .fontsize
        TxtFont.FontName = .FontFamily
        TxtFont.FontItalic = .FontItalic
        TxtFont.FontUnderline = .FontUnderline
        TxtFont.ForeColor = .ForeColor
        TxtFont.FontWeight = .FontWeight
    End With
End Sub

Private Sub GetGeneral()
    TxtFolderOutput = Identity.FolderOutput
    TxtAuditor = Identity.Auditor
End Sub
Private Sub ValueUpdated()
    'If MvValUpdates = False Then
    '    MvValUpdates = True
    '    Me.cmdApply.Enabled = True
    'End If
    
    If MvValUpdates = False Then
        MvValUpdates = True
        Me.CmdApply.Enabled = True
    End If
End Sub

Private Sub TxtAuditor_Change()
    ValueUpdated
End Sub

Private Sub TxtFolderOutput_Change()
    ValueUpdated
End Sub
