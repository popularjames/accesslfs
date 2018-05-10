Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2014-06-18 VS: ALJ Package Details Form
 ' ALJ
Dim packagePath As String
Dim FileName As String


Private Sub cmdViewMAXIMUS_Click()

DoCmd.OpenReport "rpt_ALJ_Fax_Cover_MAXIMUS", acViewPreview, , "HearingPackageName = '" & Replace(algPackageName, "'", "''") & "'"

End Sub

Private Sub cmdSaveMAXIMUS_Click()

algPackageName = Replace(algPackageName, "'", "''")
DoCmd.OpenReport "rpt_ALJ_Fax_Cover_MAXIMUS_And_732", acViewPreview, , "HearingPackageName = '" & algPackageName & "'"
algPackageName = Replace(algPackageName, "''", "'")

FileName = GetPackageLocation & Replace(algPackageName, ":", "_") & "_MAXIMUS_Cover_And_732.pdf"
DoCmd.OutputTo acReport, "rpt_ALJ_Fax_Cover_MAXIMUS_And_732", acFormatPDF, FileName, False

Call WriteToFaxTable(algPackageName, FileName, MAX_DOC_TYPE)

End Sub

Private Sub cmdViewAppellant_Click()

DoCmd.OpenReport "rpt_ALJ_Fax_Cover_Appellant", acViewPreview, , "HearingPackageName = '" & Replace(algPackageName, "'", "''") & "'"

End Sub

Private Sub cmdSaveAppellant_Click()

algPackageName = Replace(algPackageName, "'", "''")
DoCmd.OpenReport "rpt_ALJ_Fax_Cover_Appellant_And_732", acViewPreview, , "HearingPackageName = '" & algPackageName & "'"
algPackageName = Replace(algPackageName, "''", "'")

FileName = GetPackageLocation & Replace(algPackageName, ":", "_") & "_Appellant_Cover_And_732.pdf"
DoCmd.OutputTo acReport, "rpt_ALJ_Fax_Cover_Appellant_And_732", acFormatPDF, FileName, False

Call WriteToFaxTable(algPackageName, FileName, APP_DOC_TYPE)

End Sub

Private Sub cmdViewJudge_Click()

DoCmd.OpenReport "rpt_ALJ_Fax_Cover_Judge", acViewPreview, , "HearingPackageName = '" & Replace(algPackageName, "'", "''") & "'"

End Sub

Private Sub cmdSaveJudge_Click()

algPackageName = Replace(algPackageName, "'", "''")
DoCmd.OpenReport "rpt_ALJ_Fax_Cover_Judge_And_732", acViewPreview, , "HearingPackageName = '" & algPackageName & "'"
algPackageName = Replace(algPackageName, "''", "'")

FileName = GetPackageLocation & Replace(algPackageName, ":", "_") & "_Judge_Cover_And_732.pdf"
DoCmd.OutputTo acReport, "rpt_ALJ_Fax_Cover_Judge_And_732", acFormatPDF, FileName, False

Call WriteToFaxTable(algPackageName, FileName, JUDGE_DOC_TYPE)

End Sub

Private Sub cmdView732_Click()
algPackageName = Replace(algPackageName, "'", "''")
DoCmd.OpenReport "rpt_ALJ_732_Form", acViewPreview, , "HearingPackageName = '" & algPackageName & "'"
algPackageName = Replace(algPackageName, "''", "'")

End Sub

Private Sub cmdSave732_Click()

algPackageName = Replace(algPackageName, "'", "''")
DoCmd.OpenReport "rpt_ALJ_732_Form", acViewPreview, , "HearingPackageName = '" & algPackageName & "'"
algPackageName = Replace(algPackageName, "''", "'")

FileName = GetPackageLocation & Replace(algPackageName, ":", "_") & "_732.pdf"
DoCmd.OutputTo acReport, "rpt_ALJ_732_Form", acFormatPDF, FileName, False

End Sub

Private Sub Form_Load()

Me.txtPackageName = algPackageName

End Sub

Private Sub cmdExit_Click()

    DoCmd.Close
    
End Sub
