Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'*********************************************************************************************
'Modified by Kathleen C Flanagan, Thursday 2/28/2013
'Description: To support functionality for Claim QA 2.0 (Update submitted QA records)
'*********************************************************************************************

Dim mbFormDirty As Boolean
Dim mstrUserName As String
Dim mstrPrevCnlyClaimNum As String
Dim mbLockClaim As Boolean

Private Sub cbDRGAmountCorrect_Exit(Cancel As Integer)
    DRGQAScoreCalculation
    DRG_RequiredFormat
    DRGFieldsEnable
End Sub

Private Sub cbDRGClaimReferMN_Exit(Cancel As Integer)
    DRGQAScoreCalculation
    DRG_RequiredFormat
    DRGFieldsEnable
End Sub

Private Sub cbDRGCodingChange_Exit(Cancel As Integer)
    DRGQAScoreCalculation
    DRG_RequiredFormat
    DRGFieldsEnable
End Sub

Private Sub cbDRGCorrect_Exit(Cancel As Integer)
'KCF 1/6/2015 - Add ElseIF to handle the updates for the 'Completed QA' claims
   If Me.cbDRGCorrect.Column(0) = "Y" And Me.ClmStatus = "321" And Nz(Me.SubmitDate, "") = "" Then
    Me.cbDRGCorrectDecision.Value = ""
    Me.chkReturn = True
    Call chkReturn_Click
    
    ElseIf Me.cbDRGCorrect.Column(0) = "Y" And Me.PrevClmStatus = "321" And Nz(Me.SubmitDate, "") <> "" Then
    Me.cbDRGCorrectDecision.Value = ""
    End If
    
    DRGQAScoreCalculation
    DRG_RequiredFormat
    DRGFieldsEnable
End Sub

Private Sub cbDRGCorrectDecision_Exit(Cancel As Integer)
'KCF 1/6/2015 - Add to IF statement to handle updates from DRGCorrectDecision for 'Completed QA' claims
If Me.cbDRGCorrectDecision.Column(0) = "Y" And (Me.ClmStatus = "321" Or Me.PrevClmStatus = "321") Then
    Me.cbDRGCorrect.Value = ""
    Me.chkReturn.Value = False
    Call chkReturn_Click
    
    DRGFieldsEnable
End If
    
If Me.cbDRGCorrectDecision.Column(0) <> "Y" And Me.ClmStatus <> "321" And Nz(Me.SubmitDate, "") = "" Then
    Me.chkReturn.Value = True
    Call chkReturn_Click
    
    Me.cbDRGCorrect.Value = ""
    Me.txtDRGCorrect_Comment.Value = ""
    Me.cbDRGCorrect.Enabled = False
    
    Me.cbDRGCorrectDischarge.Value = ""
    Me.txtDRGCorrectDischarge_Comment.Value = ""
    Me.cbDRGCorrectDischarge.Enabled = False
    
    Me.cbDRGCodingChange.Value = ""
    Me.txtDRGCodingChange_Comment.Value = ""
    Me.cbDRGCodingChange.Enabled = False
    
    Me.cbRationaleCorrect.Value = ""
    Me.txtRationaleCorrect_Comment.Value = ""
    Me.cbRationaleCorrect.Enabled = False
    
ElseIf Me.cbDRGCorrectDecision.Column(0) <> "Y" And Me.PrevClmStatus <> "321" And Nz(Me.SubmitDate, "") <> "" Then
    Me.chkReturn.Value = True
    Call chkReturn_Click
    
    Me.cbDRGCorrect.Value = ""
    Me.txtDRGCorrect_Comment.Value = ""
    Me.cbDRGCorrect.Enabled = False
    
    Me.cbDRGCorrectDischarge.Value = ""
    Me.txtDRGCorrectDischarge_Comment.Value = ""
    Me.cbDRGCorrectDischarge.Enabled = False
    
    Me.cbDRGCodingChange.Value = ""
    Me.txtDRGCodingChange_Comment.Value = ""
    Me.cbDRGCodingChange.Enabled = False
    
    Me.cbRationaleCorrect.Value = ""
    Me.txtRationaleCorrect_Comment.Value = ""
    Me.cbRationaleCorrect.Enabled = False
    
ElseIf Me.cbDRGCorrectDecision.Column(0) = "Y" And Me.ClmStatus <> "321" And Nz(Me.SubmitDate, "") = "" Then
    Me.chkReturn.Value = False
    Call chkReturn_Click

    Me.cbDRGCorrect.Enabled = True
    Me.cbDRGCorrectDischarge.Enabled = True
    Me.cbDRGCodingChange.Enabled = True
    Me.cbDRGClaimReferMN.Enabled = True
    Me.cbRationaleCorrect.Enabled = True
    
ElseIf Me.cbDRGCorrectDecision.Column(0) = "Y" And Me.PrevClmStatus <> "321" And Nz(Me.SubmitDate, "") <> "" Then
    Me.chkReturn.Value = False
    Call chkReturn_Click

    Me.cbDRGCorrect.Enabled = True
    Me.cbDRGCorrectDischarge.Enabled = True
    Me.cbDRGCodingChange.Enabled = True
    Me.cbDRGClaimReferMN.Enabled = True
    Me.cbRationaleCorrect.Enabled = True
    
End If
    
    DRGQAScoreCalculation
    DRG_RequiredFormat
    'DRGFieldsEnable
End Sub

Private Sub cbDRGCorrectDischarge_Exit(Cancel As Integer)
    DRGQAScoreCalculation
    DRG_RequiredFormat
    DRGFieldsEnable
End Sub

Private Sub cbMNCodingCorrect_Exit(Cancel As Integer)
'Update 4/23/2013 KCF: to call HH QA Score Calc as needed
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Update 2/5/2014 KCF: to call Bleph Score Calc as needed; exclude Blephs from ConceptQA Calc
'Update 2/17/2016 KCF: to call PHP score Calc
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        ConceptQAScoreCalculation
    'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        PHPQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        BlephQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        SNFC2019QAScoreCalculation
    ElseIf Me.txtDataType = "IP" Then
        MNQAScoreCalculation
    ElseIf Me.txtDataType = "HH" Then
        HHQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        NPWTQAScoreCalculation
    End If
    MN_RequiredFormat
End Sub

Private Sub cbMNCompleteMR_Exit(Cancel As Integer)
'Update 4/23/2013 KCF: to call HH QA Score Calc as needed
'Update 5/30/2013 KCF: to call Extrap QA Score Calc as needed
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Update 2/5/2014 KCF: to call Bleph Score Calc as needed; exclude Blephs from ConceptQA Calc
'Update 2/7/2017 KCF: to call PHP
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        ConceptQAScoreCalculation
     'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        PHPQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        BlephQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        SNFC2019QAScoreCalculation
    ElseIf Me.txtDataType = "IP" Then
        MNQAScoreCalculation
    ElseIf Me.txtDataType = "HH" Then
        HHQAScoreCalculation
    ElseIf Me.txtDataType = "CARR" Then
        ExtrapQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        NPWTQAScoreCalculation
    End If
    MN_RequiredFormat
End Sub

Private Sub cbMNCorrectDecision_Exit(Cancel As Integer)
'Update 4/23/2013 KCF: to call HH QA Score Calc as needed
'Update 5/30/2013 KCF: to call Extrap QA Score Calc as needed
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Update 2/5/2014 KCF: to call Bleph Score Calc as needed; exclude Blephs from ConceptQA Calc
'Update 2/7/2016 KCF: to call PHP
'Update 2/22/2016 KCF: to toggle the fields based upon correct decision and claim status
    
 '=============================================================================================
 'BEGIN KCF 2/22/2016 Block to handle the toggle for correct decision for NR
 
 'Toggle the Correct Decision vs Potential Recovery for No Recoveries
 If Me.cbMNCorrectDecision.Column(0) = "Y" And (Me.ClmStatus = "321" Or Me.PrevClmStatus = "321") Then
    Me.cbMNPertLab.Value = ""
    Me.chkReturn.Value = False
    Call chkReturn_Click
    
    MNFieldsEnable
End If
 
 'END KCF 2/22/2016 Block to handle the toggle for correct decision for NR
 '==============================================================================================
    
'=============================================================================================
'BEGIN KCF 2/22/2016 Block to handle the toggle for correct decision for Recoveries
If Me.cbMNCorrectDecision.Column(0) <> "Y" And Me.ClmStatus <> "321" And Nz(Me.SubmitDate, "") = "" Then
    Me.chkReturn.Value = True
    Call chkReturn_Click
    
    Me.cbMNPertLab.Value = ""
    Me.txtMNPertLab_Comment.Value = ""
    Me.cbMNPertLab.Enabled = False
    
    Me.cbMNPhysOrder.Value = ""
    Me.txtMNPhysOrder_Comment.Value = ""
    Me.cbMNPhysOrder.Enabled = False
    
    Me.cbMNCompleteMR.Value = ""
    Me.txtMNCompleteMR_Comment.Value = ""
    Me.cbMNCompleteMR.Enabled = False
    
    Me.cbMNGrammar.Value = ""
    Me.txtMNGrammar_Comment.Value = ""
    Me.cbMNGrammar.Enabled = False
    
    Me.cbMNCodingCorrect.Value = ""
    Me.txtMNCodingCorrect_Comment.Value = ""
    Me.cbMNCodingCorrect.Enabled = False
    
    Me.cbRationaleCorrect.Value = ""
    Me.txtRationaleCorrect_Comment.Value = ""
    Me.cbRationaleCorrect.Enabled = False
    
ElseIf Me.cbMNCorrectDecision.Column(0) <> "Y" And Me.PrevClmStatus <> "321" And Nz(Me.SubmitDate, "") <> "" Then
    Me.chkReturn.Value = True
    Call chkReturn_Click
    
    Me.cbMNPertLab.Value = ""
    Me.txtMNPertLab_Comment.Value = ""
    Me.cbMNPertLab.Enabled = False
    
    Me.cbMNPhysOrder.Value = ""
    Me.txtMNPhysOrder_Comment.Value = ""
    Me.cbMNPhysOrder.Enabled = False
    
    Me.cbMNCompleteMR.Value = ""
    Me.txtMNCompleteMR_Comment.Value = ""
    Me.cbMNCompleteMR.Enabled = False
    
    Me.cbMNGrammar.Value = ""
    Me.txtMNGrammar_Comment.Value = ""
    Me.cbMNGrammar.Enabled = False
    
    Me.cbMNCodingCorrect.Value = ""
    Me.txtMNCodingCorrect_Comment.Value = ""
    Me.cbMNCodingCorrect.Enabled = False
    
    Me.cbRationaleCorrect.Value = ""
    Me.txtRationaleCorrect_Comment.Value = ""
    Me.cbRationaleCorrect.Enabled = False
    
    
ElseIf Me.cbMNCorrectDecision.Column(0) = "Y" And Me.ClmStatus <> "321" And Nz(Me.SubmitDate, "") = "" Then
    Me.chkReturn.Value = False
    Call chkReturn_Click
    
    Me.cbMNPertLab.Enabled = True
    Me.cbMNPhysOrder.Enabled = True
    Me.cbMNCompleteMR.Enabled = True
    Me.cbMNGrammar.Enabled = True
    Me.cbMNCodingCorrect.Enabled = True
    Me.cbRationaleCorrect.Enabled = True
    
ElseIf Me.cbDRGCorrectDecision.Column(0) = "Y" And Me.PrevClmStatus <> "321" And Nz(Me.SubmitDate, "") <> "" Then
    Me.chkReturn.Value = False
    Call chkReturn_Click
    
    Me.cbMNPertLab.Enabled = True
    Me.cbMNPhysOrder.Enabled = True
    Me.cbMNCompleteMR.Enabled = True
    Me.cbMNGrammar.Enabled = True
    Me.cbMNCodingCorrect.Enabled = True
    Me.cbRationaleCorrect.Enabled = True
    
End If
'END KCF 2/22/2016 Block to handle the toggle for correct decision for Recoveries
'==============================================================================================
    
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" And Me.AuditTeam <> "SNF - C2019" Then
        ConceptQAScoreCalculation
    'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        PHPQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        BlephQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        SNFC2019QAScoreCalculation
    ElseIf Me.txtDataType = "IP" Then
        MNQAScoreCalculation
    ElseIf Me.txtDataType = "HH" Then
        HHQAScoreCalculation
    ElseIf Me.txtDataType = "CARR" Then
        ExtrapQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        NPWTQAScoreCalculation
    End If
    MN_RequiredFormat
End Sub

Private Sub cbMNGrammar_Exit(Cancel As Integer)
'Update 4/23/2013 KCF: to call HH QA Score Calc as needed
'Update 5/30/2013 KCF: to call Extrap QA Score Calc as needed
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Update 2/5/2014 KCF: to call Bleph Score Calc as needed; exclude Blephs from ConceptQA Calc
'Update 2/7/2016 KCF: to call PHP
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        ConceptQAScoreCalculation
     'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        PHPQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        BlephQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        SNFC2019QAScoreCalculation
    ElseIf Me.txtDataType = "IP" Then
        MNQAScoreCalculation
    ElseIf Me.txtDataType = "HH" Then
        HHQAScoreCalculation
    ElseIf Me.txtDataType = "CARR" Then
        ExtrapQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        NPWTQAScoreCalculation
    End If
    MN_RequiredFormat
End Sub

Private Sub cbMNPertLab_Exit(Cancel As Integer)
'Update 4/23/2013 KCF: to call HH QA Score Calc as needed
'Update 5/30/2013 KCF: to call Extrap QA Score Calc as needed
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Update 2/5/2014 KCF: to call Bleph Score Calc as needed; exclude Blephs from ConceptQA Calc
'UPdate 2/7/2016 KCF: to call PHP
'Update 2/22/2016 KCF: to toggle the PertLab (Recovery Identified)

'========================================================================================
'BEGIN KCF 2/22/2016: to toggle field values based upon response to PertLab (Recovery idenitifed)
 If Me.cbMNPertLab.Column(0) = "Y" And Me.ClmStatus = "321" And Nz(Me.SubmitDate, "") = "" Then
    Me.cbMNCorrectDecision.Value = ""
    Me.chkReturn = True
    Call chkReturn_Click
    
    ElseIf Me.cbMNPertLab.Column(0) = "Y" And Me.PrevClmStatus = "321" And Nz(Me.SubmitDate, "") <> "" Then
    Me.cbMNCorrectDecision.Value = ""
    End If

'END KCF 2/22/2016: to toggle field values based upon response to PertLab (Recovery idenitifed)
'========================================================================================
    
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        ConceptQAScoreCalculation
     'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        PHPQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        BlephQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        SNFC2019QAScoreCalculation
    ElseIf Me.txtDataType = "IP" Then
        MNQAScoreCalculation
    ElseIf Me.txtDataType = "HH" Then
        HHQAScoreCalculation
    ElseIf Me.txtDataType = "CARR" Then
        ExtrapQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        NPWTQAScoreCalculation
    End If
    MN_RequiredFormat
End Sub

Private Sub cbMNPhysOrder_Exit(Cancel As Integer)
'Update 4/23/2013 KCF: to call HH QA Score Calc as needed
'Update 5/30/2013 KCF: to call Extrap QA Score Calc as needed
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Update 2/5/2014 KCF: to call Bleph Score Calc as needed; exclude Blephs from ConceptQA Calc
'Update 2/7/2016 KCF: to call PHP
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        ConceptQAScoreCalculation
     'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        PHPQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        BlephQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        SNFC2019QAScoreCalculation
    ElseIf Me.txtDataType = "IP" Then
        MNQAScoreCalculation
    ElseIf Me.txtDataType = "HH" Then
        HHQAScoreCalculation
    ElseIf Me.txtDataType = "CARR" Then
        ExtrapQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        NPWTQAScoreCalculation
    End If
    MN_RequiredFormat
End Sub

Private Sub cbRationaleCorrect_Change()
'Updated 2/5/2014 KCF: to call Bleph Score Calc & formatting; exclue Blephs from ConceptQA Calc
'Update 2/7/2016 KCF: to call PHP score & formatting
    'BEGIN 2/5/2014 KCF for Concept vs Bleph Claims
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        MN_RequiredFormat
        ConceptQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        MN_RequiredFormat
        BlephQAScoreCalculation
    'END 2/5/2014 KCF for Concept vs Bleph Claims
     'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        MN_RequiredFormat
        PHPQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        MN_RequiredFormat
        SNFC2019QAScoreCalculation
    ElseIf (Me.MedicalNecessity = "A" Or Me.MedicalNecessity = "S") Then
        MN_RequiredFormat
        MNQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" Or (Me.MedicalNecessity = "N" And Me.txtDataType = "IP") Then 'add the IF criteria for DRG
        DRG_RequiredFormat
        DRGQAScoreCalculation 'Added Wednesday 4/10/2013 for DRG Implementation
'BEGIN 4/23/2013 KCF for HH claims
    ElseIf Me.MedicalNecessity = "N" And Me.txtDataType = "HH" Then
        MN_RequiredFormat
        HHQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        MN_RequiredFormat
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        MN_RequiredFormat
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        MN_RequiredFormat
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        MN_RequiredFormat
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        MN_RequiredFormat
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        MN_RequiredFormat
        NPWTQAScoreCalculation
    End If

End Sub

Private Sub cbRationaleCorrect_Exit(Cancel As Integer)
'Update 9/11/2013 KCF: to call Concept QA Score Calc as neede
'Updated 2/5/2014 KCF: to call Bleph Score Calc & formatting; exclue Blephs from ConceptQA Calc
'Update 2/7/2016 KCF: to call PHP score & format
    'BEGIN 2/5/2014 KCF for Concept vs Bleph Claims
    If Me.MedicalNecessity = "C" And Me.AuditTeam <> "Bleph" And Me.AuditTeam <> "SNF - C2019" Then
        MN_RequiredFormat
        ConceptQAScoreCalculation
    ElseIf Me.MedicalNecessity = "C" And Me.AuditTeam = "Bleph" Then
        MN_RequiredFormat
        BlephQAScoreCalculation
    'END 2/5/2014 KCF for Concept vs Bleph Claims
 'KCF 2/17/2016
    ElseIf Me.AuditTeam = "PHP" Then
        MN_RequiredFormat
        PHPQAScoreCalculation
    ElseIf Me.AuditTeam = "SNF - C2019" Then
        MN_RequiredFormat
        SNFC2019QAScoreCalculation
    ElseIf (Me.MedicalNecessity = "A" Or Me.MedicalNecessity = "S") Then
        MN_RequiredFormat
        MNQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" Or (Me.MedicalNecessity = "N" And Me.txtDataType = "IP") Then 'add the IF criteria for DRG
        DRG_RequiredFormat
        DRGQAScoreCalculation 'Added Wednesday 4/10/2013 for DRG Implementation
        DRGFieldsEnable
'BEGIN 4/23/2013 KCF for HH claims
    ElseIf Me.MedicalNecessity = "N" And Me.txtDataType = "HH" Then
        MN_RequiredFormat
        HHQAScoreCalculation
'END 4/23/2013 KCF for HH claims
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
        MN_RequiredFormat
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
        MN_RequiredFormat
        SacralNerveQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        MN_RequiredFormat
        PTAQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        MN_RequiredFormat
        OsteoStimQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        MN_RequiredFormat
        HospiceQAScoreCalculation
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        MN_RequiredFormat
        NPWTQAScoreCalculation
    End If
End Sub

Private Sub chkReturn_Click()
'Form includes a checkbox for the user to indicate whether the Claim should be sent back to the Auditor - KCF 8/7/2012
'unit test 9/10/2012 kcf
    If Me.Parent.Name = "frm_QA_Review_Main" Then '2/28/2013 KCF - handle events from the original (unsubmitted) form
        If Me.chkReturn.Value = True Then
            Me.txtQAStatus = "R"
            Me.Parent.cmdSubmitQA.Caption = "Return to Auditor" 'Visual cue for user that the Claim will be returned
            Form_Dirty (mbFormDirty) 'Unbound controls do not set dirty - using dirty to set the Reviewer & ReviewedDate fields for the records
        Else
            Me.txtQAStatus = ""
            Me.Parent.cmdSubmitQA.Caption = "Submit QA Review"
            Form_Dirty (mbFormDirty) 'Unbound controls do not set dirty - using dirty to set the Reviewer & ReviewedDate fields for the records
        End If
    'BEGIN 2/28/2013 KCF - handle events on the submitted form
    ElseIf Me.Parent.Name = "frm_QA_Review_Main_Submitted" Then
        If Me.chkReturn.Value = True Then
            'Me.txtQAStatus = "R"
            Me.Parent.cmdReturnAuditor.Enabled = True
            Me.Parent.cmdUpdateQA.Enabled = False
        Else
            'Me.txtQAStatus = ""
            Me.Parent.cmdReturnAuditor.Enabled = False
            Me.Parent.cmdUpdateQA.Enabled = False
        End If
    End If
    'END 2/28/2013 KCF - handle events on the submitted form
    
    Me.Requery
    
    If Me.LockUser & "" = "" Then
        Me.LockUser = mstrUserName
        Me.LockDt = Now()
        mbLockClaim = True
        Me.Requery
    End If
    
End Sub



Private Sub cmdCalcAuditorStats_Click()
'Created Monday 10/22/2012 to replace the calculation done on the UI during the On Load Event
'Modified Tuesday 11/6/2012 to limit calculation to selected QAStaff - will pass the username to the stored procedure
'   which will check against list of valid QA Staff.  Need to update calculation to account for 0 values being returned for non-QA
    
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    
    On Error GoTo Err_handler
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_QA_Review_AuditorStats"
    cmd.Parameters.Refresh
    cmd.Parameters("@pAuditor") = Me.Auditor
    cmd.Parameters("@pBeginDate") = Me.txtStatFromDate
    cmd.Parameters("@pEndDate") = Me.txtStatToDate
    'cmd.Parameters("@pAdj_ProjectedSavings") = Me.txtAdjProSav
    cmd.Parameters("@pmstrUserName") = mstrUserName 'Added 11/6/2012 by KCF
    cmd.Execute
    
    strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_QA_Review_AuditorStats"
        Err.Raise 50001, "usp_QA_Review_AuditorStats", strErrMsg
    End If
    
    Me.lblQAResults.Caption = "Results for: " & Me.Auditor
    
    'BEGIN 11/6/2012: If stp recognizes user as valid QA Staff - will set the DisplayStats indicator to 'Y'
    If cmd.Parameters("@pDisplayStats") <> "N" Then
    'END 11/6/2012: If stp recognizes user as valid QA Staff - will set the DisplayStats indicator to 'Y'
    
        'BEGIN: 2/13/2013 KCF:  Update stat calculations to match the new categories for Recovery, NR & Total Claims
        Me.lblTotalQAClaims.Caption = Nz(cmd.Parameters("@pTotalQAClaims"), 0)
        Me.lblTotalAvgQAScore.Caption = Nz(cmd.Parameters("@pTotalAvgQAScore"), 0)
        Me.lblTotalClaims.Caption = Nz(cmd.Parameters("@pTotalClaims"), 0)
        If (IsNull(cmd.Parameters("@pTotalClaims")) Or cmd.Parameters("@pTotalClaims") = 0) Then
            Me.lblTotalQAPercent.Caption = "0%"
        Else
            Me.lblTotalQAPercent.Caption = (Round(cmd.Parameters("@pTotalQAClaims") / cmd.Parameters("@pTotalClaims"), 2) * 100) & "%"
        End If
        
        Me.lblRecoveryQAClaims.Caption = Nz(cmd.Parameters("@pRecoveryQAClaims"), 0)
        Me.lblRecoveryAvgQAScore.Caption = Nz(cmd.Parameters("@pRecoveryAvgQAScore"), 0)
        Me.lblRecoveryClaims.Caption = Nz(cmd.Parameters("@pRecoveryClaims"), 0)
        If (IsNull(cmd.Parameters("@pRecoveryClaims")) Or cmd.Parameters("@pRecoveryClaims") = 0) Then
            Me.lblRecoveryQAPercent.Caption = "0%"
        Else
            Me.lblRecoveryQAPercent.Caption = (Round(cmd.Parameters("@pRecoveryQAClaims") / cmd.Parameters("@pRecoveryClaims"), 2) * 100) & "%"
        End If
                
        Me.lblNRQAClaims.Caption = Nz(cmd.Parameters("@pNRQAClaims"), 0)
        Me.lblNRAvgQAScore.Caption = Nz(cmd.Parameters("@pNRAvgQAScore"), 0)
        Me.lblNRClaims.Caption = Nz(cmd.Parameters("@pNRClaims"), 0)
        If (IsNull(cmd.Parameters("@pNRClaims")) Or cmd.Parameters("@pNRClaims") = 0) Then
            Me.lblNRQAPercent.Caption = "0%"
        Else
            Me.lblNRQAPercent.Caption = (Round(cmd.Parameters("@pNRQAClaims") / cmd.Parameters("@pNRClaims"), 2) * 100) & "%"
        End If
        
        Me.lblNRtoRec.Caption = Nz(cmd.Parameters("@pNRToRec"), 0)
        Me.lblNRtoRecAmt.Caption = Format(Nz(cmd.Parameters("@pNRToRecAmt"), 0), "Currency")
        Me.lblRecToNR.Caption = Nz(cmd.Parameters("@pRecToNR"), 0)
        Me.lblRecToNRAmt.Caption = Format(Nz(cmd.Parameters("@pRecToNRAmt"), 0), "Currency")
            
     'BEGIN 11/6/2012: If stp does not recognize user as valid QA Staff - will not display any Auditor Stats
     Else
        Me.lblTotalQAClaims.Caption = "N\a"
        Me.lblTotalAvgQAScore.Caption = "N\a"
        Me.lblTotalClaims.Caption = "N\a"
        Me.lblTotalQAPercent.Caption = "N\a"
        
        Me.lblRecoveryQAClaims.Caption = "N\a"
        Me.lblRecoveryAvgQAScore.Caption = "N\a"
        Me.lblRecoveryClaims.Caption = "N\a"
        Me.lblRecoveryQAPercent.Caption = "N\a"
        
        Me.lblNRQAClaims.Caption = "N\a"
        Me.lblNRAvgQAScore.Caption = "N\a"
        Me.lblNRClaims.Caption = "N\a"
        Me.lblNRQAPercent.Caption = "N\a"
        
        Me.lblNRtoRec.Caption = "N\a"
        Me.lblNRtoRecAmt.Caption = "N\a"
        Me.lblRecToNR.Caption = "N\a"
        Me.lblRecToNRAmt.Caption = "N\a"

     End If
    'END 11/6/2012: If stp does not recognize user as valid QA Staff - will not display any Auditor Stats
    'END: 2/13/2013 KCF:  Update stat calculations to match the new categories for Recovery, NR & Total Claims
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub


Sub cmdUnlockClaim_Click()
'2/11/2013 KCF: Module added

Me.LockUser = ""
Me.LockDt = Null

Form_Current



End Sub



Private Sub Form_BeforeUpdate(Cancel As Integer)
'******************************************************************
'Revised Monday 3/23/2013 KCF - Handle the write conflict cleanly
Dim msg1 As String

    If Me.Parent.Name = "frm_QA_Review_Main" Then

    ElseIf Me.Parent.Name = "frm_QA_Review_Main_Submitted" Then
        If mbFormDirty = True Then
            msg1 = MsgBox("You have made changes to this Claim QA Record.  To keep your changes, select 'OK'.  To discard your changes, select 'Cancel'.", vbOKCancel)
            If msg1 = vbOK Then
                Me.Parent.cmdUpdateQA.Enabled = True
                mbFormDirty = False
                Me.ReviewComment.SetFocus
            ElseIf msg1 = vbCancel Then
                Cancel = True
                Me.Undo
            End If

        End If
    End If

    If Me.LockUser & "" = mstrUserName And mbLockClaim = False Then
        Me.LockUser = ""
        Me.LockDt = ""
    End If

End Sub


Private Sub Form_Close()
'Clear out the user from the Lock fields
    If Me.LockUser & "" = mstrUserName Then
        Me.LockDt = ""
        Me.LockUser = ""
    End If
End Sub


Private Sub Form_Current()
'unit test 9/10/2012 kcf

    Dim MyAdo As New clsADO
    Dim rs As ADODB.RecordSet
        
    Dim strSQL As String
    Dim strErrMsg As String
    Dim QAScore As Integer

    mstrUserName = GetUserName()
    mbFormDirty = False
    mbLockClaim = False
    QAScore = 0
    strErrMsg = ""
    
    Me.txtStatFromDate = "1/1/1900"
    Me.txtStatToDate = "12/31/9999"
    
    Call cmdCalcAuditorStats_Click
    
 'BEGIN Set up the form controls for the current record
    'Set the AuditCheckbox depending upon the QAStatus when loaded
    If Me.txtQAStatus = "R" And Me.Parent.Name = "frm_QA_Review_Main" Then
        Me.chkReturn.Value = True
    Else
        Me.chkReturn.Value = False
    End If
    
 'BEGIN 3/27/2013 KCF:  Allow returns for subcontractors by commenting out the code
'    'If a sub - should not be returned to Auditor (version 1.1)
'    If Me.txtCompanyID = 1 Then
'        Me.chkReturn.visible = True
'    Else
'        Me.chkReturn.visible = False
'    End If

    Me.chkReturn.visible = True
    
'END 3/27/2013 KCF:  Allow returns for subcontractors by commenting out the code
    
    'BEGIN 2/15/2013 KCF: For completed only allow the return if the current claim status is Recover
    '3/21/2013 FIX HERE
    If Not (Me.RecordSet Is Nothing) Then
        If Me.Parent.Name = "frm_QA_Review_Main_submitted" Then
            If (Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus = "320", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus = "320.2", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus = "321", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus = "322", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus = "314", "")) Then
                Me.chkReturn.visible = True
        'BEGIN 4/17/2013 FIX: Correct the structure of the IF statement
            'End If '4/17/2013
        'ElseIf Me.Parent.Name = "frm_QA_Review_Main_submitted" Then '4/17/2013 KCF
            'If (Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "320", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "321", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "322", "")) Then
            ElseIf (Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "320", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus = "320.2", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "321", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "322", "") Or Nz(Me.Parent("frm_QA_Claims_List")!txtAHClmStatus <> "314", "")) Then
        'END 4/17/2013 FIX: Correct the structure of the IF statement
                Me.chkReturn.visible = False
            End If
        End If
    End If
    'END 2/15/2013 KCF: For completed only allow the return if the current claim status is Recovery
    
    'visual cue for user that toggles submit button if claim to be returned to auditor
    If IsSubForm(Me) And Me.Parent.Form.Name = "frm_QA_Review_Main" Then '2/22/2013 Update to separate out command button options depending on the form.
        On Error Resume Next
        If Me.txtQAStatus = "R" Then
            Me.Parent.cmdSubmitQA.Caption = "Return to Auditor"
        Else
            Me.Parent.cmdSubmitQA.Caption = "Submit to QA"
        End If
    'BEGIN 2/22/2013 KCF: set auditor checkbox on the form
    ElseIf IsSubForm(Me) And Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" Then
        If Me.Parent.cmdReturnAuditor.Enabled = True Then
            Me.chkReturn.Value = "True"
        End If
    'END 2/22/2013 KCF: set auditor checkbox on the form
    End If
        
    'Toggle the scoring questions based upon the category of review - Medical Necessity or DRG
    'BEGIN 2/7/2016 KCF: Add for PHP
    If Me.txtAuditTeam = "PHP" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        PHPQAScoreCalculation
        MN_RequiredFormat
    
    'BEGIN 5/28/2013 KCF: Add IF for Extrapolation claims
    ElseIf Me.txtAuditTeam = "Extrapolation" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        ExtrapQAScoreCalculation
        MN_RequiredFormat
    'END 5/28/2013 KCF: Add IF for Extrapolation claims
    'BEGIN 9/11/2013 KCF: Add IF for Concept claims
    ElseIf Me.txtAuditTeam = "Cnly Concept Development Team" Or Me.txtAuditTeam = "CNLY_therapy" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        ConceptQAScoreCalculation
        MN_RequiredFormat
    'END 5/28/2013 KCF: Add IF for Extrapolation claims
    'BEGIN 2/5/2014 KCF: Add IF for Bleph claims
    ElseIf Me.txtAuditTeam = "Bleph" And Me.MedicalNecessity = "C" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        ConceptQAScoreCalculation
        MN_RequiredFormat
    'END 2/5/2014 KCF: Add IF for Bleph claims
    'BEGIN 3/31/2014 KCF: Add IF for SNF C2019 claims
    ElseIf Me.txtAuditTeam = "SNF - C2019" And Me.MedicalNecessity = "C" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        SNFC2019QAScoreCalculation
        MN_RequiredFormat
    'END 3/31/2014 KCF: Add IF for SNF C2019 claims
    ElseIf Me.MedicalNecessity = "A" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        MNQAScoreCalculation
        MN_RequiredFormat
        'Me.txtAdjProSav = 2000
    ElseIf Me.MedicalNecessity = "S" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        MNQAScoreCalculation
        MN_RequiredFormat
        'Me.txtAdjProSav = 2000
    ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam Like "%DRG%" Then
        DRGFieldsVisible
        DRGFieldsEnable
        MNFieldsInvisible
        MNFieldsDisable
        DRGQAScoreCalculation
        DRG_RequiredFormat
        'Me.txtAdjProSav = 5000
    ElseIf (Me.MedicalNecessity = "N" And Me.txtDataType = "IP") Then
            DRGFieldsVisible
            DRGFieldsEnable
            MNFieldsInvisible
            MNFieldsDisable
            DRGQAScoreCalculation
            DRG_RequiredFormat
    'BEGIN 4/23/2013 KCF: HH Implementation
    'Will rely on the MN formatting since re-useing the MN data fields, except the HHQAScoreCalculation
    ElseIf Me.MedicalNecessity = "N" And Me.txtDataType = "HH" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        HHQAScoreCalculation 'Based upon the MNQAScoreCalculation
        MN_RequiredFormat
    'END 4/23/2013 KCF: HH Implementation
    
'BEGIN New Detail Concepts: KCF 4/11/2014
    ElseIf Me.AuditTeam = "Sacral Nerve CM_1983" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        SacralNerveQAScoreCalculation
        MN_RequiredFormat
     ElseIf Me.AuditTeam = "Sacral Nerve CM_1984" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        SacralNerveQAScoreCalculation
        MN_RequiredFormat
     ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        PTAQAScoreCalculation
        MN_RequiredFormat
     ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        OsteoStimQAScoreCalculation
        MN_RequiredFormat
     ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        HospiceQAScoreCalculation
        MN_RequiredFormat
     ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsVisible
        MNFieldsEnable
        NPWTQAScoreCalculation
        MN_RequiredFormat
 'END New Detail Concepts: KCF 4/11/2014
    Else
        DRGFieldsInvisible
        DRGFieldsDisable
        MNFieldsInvisible
        MNFieldsDisable
    End If

    'Visual cue for user to toggle note on UI that the claim has \ has not been previously returned to Auditor
    If Me.txtSeqNo = 1 Then
        Me.lblReturnNote.visible = False
    Else
        Me.lblReturnNote.visible = True
    End If
    
    'If no records - don't display ReturnNote (default if SeqNo <> 1)
    If Me.RecordSet.recordCount = 0 Then
        lblReturnNote.visible = False
    End If
 'END Set up the form controls for the current record

'BEGIN Setting up the Lock on the Form
    If Not (Me.RecordSet Is Nothing) Then
        If Me.RecordSet.recordCount > 0 Then
            If Me.LockUser & "" = "" Then
                Me.LockUser = mstrUserName
                Me.LockDt = Now()
                mbLockClaim = True
                Me.Requery
            End If
        End If

    'BEGIN Enable \ Disable the form controls based upon LockUser value
        If Me.LockUser & "" = mstrUserName Then
            Me.LockUser.visible = False
            Me.LockDt.visible = False
            Me.frmLock.visible = False
            Me.cmdUnlockClaim.visible = False  '2013 KCF: Unlock Claim button for Qa Team
            Me.LockDt = Now()
            IPFieldsEnable
        
            If IsSubForm(Me) Then
                On Error Resume Next
                If Me.Reviewer <> "" Then
                    Me.Parent.Form.cmdSubmitQA.Enabled = True
                    Me.Parent.Form.cmdRemoveReviews.Enabled = False ' kcf
                Else
                    If Not Me.Parent.Form.cmdSubmitQA.getfocus Then
                        Me.Parent.Form.cmdSubmitQA.Enabled = False
                        Me.Parent.Form.cmdRemoveReviews.Enabled = False
                    End If
                End If
            End If
        Else
            Me.LockUser.visible = True
            Me.LockUser.Locked = True
            Me.LockDt.visible = True
            Me.LockDt.Locked = True
            Me.frmLock.visible = True
            Me.cmdUnlockClaim.visible = True '2013 KCF: Unlock Claim button for Qa Team
            IPFieldsDisable
            DRGFieldsDisable
            MNFieldsDisable
            
        End If
    'END Enable \ Disable the form controls based upon LockUser value
End If
'END Setting up the Lock on the Form
    
End Sub

Private Sub Form_Load()
    If IsSubForm(Me) Then
        Me.Parent.DetailFormLoaded = True
        'BEGIN 2/15/2013 KCF:  do not allow to return to audit if completed claim is no longer in Recovery status
        If Me.Parent.Name = "frm_QA_Review_Main" Then
            Me.chkReturn.Enabled = True
        ElseIf Me.Parent.Name = "frm_QA_Review_Main_Submitted" And (Me.txtAHClmStatus = "320" Or Me.txtAHClmStatus = "320.2" Or Me.txtAHClmStatus = "321" Or Me.txtAHClmStatus = "322") Then
            Me.chkReturn.Enabled = True
        Else
            Me.chkReturn.Enabled = False
        End If
        'END 2/15/2013 KCF:  do not allow to return to audit if completed claim is no longer in Recovery status
    End If
    
     
End Sub



Private Sub Form_Dirty(Cancel As Integer)
'unit test 9/10/2012 by kcf
    mbFormDirty = True
    If Me.Parent.Name = "Frm_QA_REview_Main_Submitted" Then
        Me.Parent.Form.cmdUpdateQA.Enabled = True
    End If
    
    If IsSubForm(Me) And Me.Parent.Name = "frm_QA_Review_Main" Then
        Me.Parent.Form.cmdSubmitQA.Enabled = True
    End If
    
    Me.Reviewer = mstrUserName 'QA Claim 1.1 update so that all record updates to the form will update the Reviewer field
    Me.ReviewedDate = Now() 'QA Claim 1.1 update so that all record updates to the form will update the Reviewed Date
End Sub


Public Function FormDirty() As Boolean
    FormDirty = mbFormDirty
End Function


Sub SpellCheck(strSpell)
' Run spell check
' Version 1.1 KCF - 7/6/2012
'unit test 9/10/2012 by kcf

Dim strErrMsg As String

On Error GoTo Err_handler

If IsNull(Len(strSpell)) Or Len(strSpell) = 0 Then
    Exit Sub
End If

With Me.ActiveControl
.SetFocus
.SelStart = 0
.SelLength = Len(strSpell)
End With

DoCmd.SetWarnings False
DoCmd.RunCommand acCmdSpelling
DoCmd.SetWarnings True

Exit_Sub:
    Exit Sub
    
Err_handler:
    strErrMsg = Err.Description
    Resume Exit_Sub

End Sub

Private Sub txtAdjProSav_Exit(Cancel As Integer)
'unit test 9/10/2012 by kcf
Call cmdCalcAuditorStats_Click

Me.CnlyClaimNum.SetFocus

End Sub



Private Sub txtDRGAmountCorrect_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtDRGCorrectDecision_Comment
SpellCheck (strSpell)

End Sub


Private Sub ReviewComment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/2/2012

Dim strSpell
strSpell = ReviewComment
SpellCheck (strSpell)
End Sub


Private Sub txtDRGClaimReferMN_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtDRGClaimReferMN_Comment
SpellCheck (strSpell)
End Sub

Private Sub txtDRGCodingChange_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtDRGCodingChange_Comment
SpellCheck (strSpell)
End Sub

Private Sub txtDRGCorrectDischarge_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtDRGCorrectDischarge_Comment
SpellCheck (strSpell)
End Sub


Private Sub txtDRGCorrect_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtDRGCorrect_Comment
SpellCheck (strSpell)
End Sub

Private Sub txtMNCodingCorrect_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtMNCodingCorrect_Comment
SpellCheck (strSpell)
End Sub



Private Sub txtMNCompleteMR_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtMNCompleteMR_Comment
SpellCheck (strSpell)
End Sub

Private Sub txtMNCorrectDecision_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtMNCorrectDecision_Comment
SpellCheck (strSpell)
End Sub

Private Sub txtMNGrammar_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtMNGrammar_Comment
SpellCheck (strSpell)

If Me.txtDataType = "CARR" Then
    Me.CnlyClaimNum.SetFocus
End If


End Sub

Private Sub txtMNPertLab_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtMNPertLab_Comment
SpellCheck (strSpell)

End Sub

Private Sub txtMNPhysOrder_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtMNPhysOrder_Comment
SpellCheck (strSpell)
End Sub

Private Sub txtRationaleCorrect_Comment_Exit(Cancel As Integer)
' Will check the spelling in the Comments section upon exiting the field
' Version 1.1 KCF - 7/3/2012

Dim strSpell
strSpell = txtRationaleCorrect_Comment
SpellCheck (strSpell)

Me.CnlyClaimNum.SetFocus

End Sub



Private Sub MNFieldsVisible()
'Any controls for the MN will be invisible
'unit test 9/10/2012 by kcf
Dim ctrl As Control

Application.Echo False

For Each ctrl In Me.Controls
    If ctrl.Tag = "MNQuestions" Then
        ctrl.visible = True
    End If
Next ctrl

End Sub


Private Sub MNFieldsInvisible()
'Any controls for the MN will be invisible
'unit test 9/10/2012 by kcf
Dim ctrl As Control

Application.Echo False

For Each ctrl In Me.Controls
    If ctrl.Tag = "MNQuestions" Then
        ctrl.visible = False
    End If
Next ctrl

End Sub

Private Sub MNFieldsDisable()
'unit test complete 9/10/2012 - MsgBox ("Hit MN Fields Disable subroutine")

    Me.cbMNPertLab.Enabled = False
    Me.txtMNPertLab_Comment.Enabled = False
    
    Me.cbMNPhysOrder.Enabled = False
    Me.txtMNPhysOrder_Comment.Enabled = False
    
    Me.cbMNCompleteMR.Enabled = False
    Me.txtMNCompleteMR_Comment.Enabled = False
    
    Me.cbMNGrammar.Enabled = False
    Me.txtMNGrammar_Comment.Enabled = False
    
    Me.cbMNCorrectDecision.Enabled = False
    Me.txtMNCorrectDecision_Comment.Enabled = False
    
    Me.cbMNCodingCorrect.Enabled = False
    Me.txtMNCodingCorrect_Comment.Enabled = False
    
Application.Echo True
    
End Sub

Private Sub DRGFieldsInvisible()
'unit test 9/10/2012 by kcf
Dim ctrl As Control

Application.Echo False

For Each ctrl In Me.Controls
    If ctrl.Tag = "DRGQuestions" Then
        ctrl.visible = False
    End If
Next ctrl

End Sub

Private Sub DRGFieldsVisible()
'unit test 9/10/2012 by kcf
Dim ctrl As Control

Application.Echo False

For Each ctrl In Me.Controls
    If ctrl.Tag = "DRGQuestions" Then
        ctrl.visible = True
    End If
Next ctrl


End Sub

Private Sub DRGFieldsDisable()
' QA v 1.1 Begin Comments - Set all Y/N & Comment fields for MN and DRG sections to Enabled = False and Locked = True
'unit test 9/10/2012 kcf - MsgBox ("Hit DRG Fields Disable subroutine")

    Me.cbDRGCorrectDischarge.Enabled = False
    Me.txtDRGCorrectDischarge_Comment.Enabled = False
    
    Me.cbDRGCodingChange.Enabled = False
    Me.txtDRGCodingChange_Comment.Enabled = False
    
    Me.cbDRGClaimReferMN.Enabled = False
    Me.txtDRGClaimReferMN_Comment.Enabled = False
    
    Me.cbDRGCorrectDecision.Enabled = False
    Me.txtDRGCorrectDecision_Comment.Enabled = False
    
    Me.cbDRGCorrect.Enabled = False
    Me.txtDRGCorrect_Comment.Enabled = False

    
Application.Echo True
 
End Sub

Private Sub MNFieldsEnable()
'unit test 9/10/2012 kcf MsgBox ("Hit MN Fields Enable subroutine")
'Tuesday 5/28/2013 KCF: Set up for Extrapolation Implementation
'Wednesday 9/11/2013 KCF: Set up for Concept Dev reviews

    lblMNClaimCommentHdr.Caption = "Comments"
    
    Me.cbMNPertLab.Enabled = True
    Me.txtMNPertLab_Comment.Enabled = True

    Me.cbMNPhysOrder.Enabled = True
    Me.txtMNPhysOrder_Comment.Enabled = True
    
    Me.cbMNCorrectDecision.RowSource = "Y;Yes;1;N;No;0"
    Me.cbMNPertLab.RowSource = "Y;Yes;1;N;No;0"
    Me.cbMNPhysOrder.RowSource = "Y;Yes;1;N;No;0"
    
    'frm_QA_Claim_Review_Result.Form.cbMNPhysOrder.RowsSource = "Y;Yes;1;N;No;0"
    
'    "Y";"Yes";1;"N";"No";0

    Me.cbMNCompleteMR.Enabled = True
    Me.txtMNCompleteMR_Comment.Enabled = True
    
    Me.cbMNGrammar.Enabled = True
    Me.txtMNGrammar_Comment.Enabled = True
    
    Me.cbMNCorrectDecision.Enabled = True
    Me.txtMNCorrectDecision_Comment.Enabled = True
    
    Me.cbMNCodingCorrect.Enabled = True
    Me.txtMNCodingCorrect_Comment.Enabled = True

'IF for Recovery Claims Formatting
If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then

    Me.txtMNCorrectDecision_Comment.visible = True
    Me.txtMNCorrectDecision_Comment.Enabled = True
    Me.cbMNCorrectDecision.visible = True
    Me.cboCorrectDecision_Label.visible = True
    Me.cboCorrectDecision_Label.Caption = "Record/Criteria correct as billed?"
    Me.cbMNCorrectDecision.RowSource = "Y;Yes;1"
    Me.cbMNCorrectDecision.Enabled = True
    
    Me.txtMNPertLab_Comment.visible = True
    Me.txtMNPertLab_Comment.Enabled = True
    Me.cbMNPertLab.visible = True
    Me.cboPertinentLab_Label.visible = True
    Me.cboPertinentLab_Label.Caption = "Potential Recovery Identified?"
    Me.cbMNPertLab.RowSource = "Y;Yes;1"
    Me.cbMNPertLab.Enabled = True
    
    Me.txtMNPhysOrder_Comment.visible = False
    Me.cbMNPhysOrder.visible = False
    Me.cboPhysicianOrder_Label.visible = False
    
    Me.txtMNCompleteMR_Comment.visible = False
    Me.cbMNCompleteMR.visible = False
    Me.cboCompleteMR_Label.visible = False
    
    Me.txtMNGrammar_Comment.visible = False
    Me.cbMNGrammar.visible = False
    Me.cboGrammarCorrect_Label.visible = False
    
    Me.txtMNCodingCorrect_Comment.visible = False
    Me.cbMNCodingCorrect.visible = False
    Me.cboCodingCorrect_Label.visible = False
    
    

Else 'Block for Recovery Claims formatting
    'BEGIN 4/21/2013 KCF: toggle question fields for the HH vs MN questions
    'BEGIN 5/28/2013 KCF: toggle questions for Extrapolation questions
        If Me.txtAuditTeam = "Extrapolation" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Guidelines Completed Correctly?"
            cboPhysicianOrder_Label.Caption = "History Correct?"
            cboCompleteMR_Label.Caption = "Exam Correct?"
            cboGrammarCorrect_Label.Caption = "MDM Correct?"
            
            txtRationaleCorrect_Comment.visible = False
            cbRationaleCorrect.visible = False
            cbRationaleCorrect_Label.visible = False
            
            txtMNCodingCorrect_Comment.visible = False
            cbMNCodingCorrect.visible = False
            cboCodingCorrect_Label.visible = False
            
            lblMNClaimHdr.Caption = "Extrapolation Claims"
     
     'BEGIN 2/17/2016 KCF: toggle questions for PHP Reviews
        ElseIf Me.txtAuditTeam = "PHP" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Psych Dx identified?"
            cboPhysicianOrder_Label.Caption = "Support doc for benefit of PHP?"
            cboCompleteMR_Label.Caption = "Multidisciplinary met requirements?"
            cboGrammarCorrect_Label.Caption = "Doc met participation requirements?"
            
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            txtMNCodingCorrect_Comment.visible = False
            txtMNCodingCorrect_Comment.Enabled = False
            cbMNCodingCorrect.visible = False
            cbMNCodingCorrect.Enabled = False
            cboCodingCorrect_Label.visible = False
            
            lblMNClaimHdr.Caption = "PHP Claims"
            
    'END 2/17/2016 KCF: toggle questions for PHP Reviews
     
     
     
     'BEGIN 9/11/2013 KCF: toggle questions for Concept Reviews
        ElseIf (Me.txtAuditTeam = "CNLY Concept Development Team" Or Me.txtAuditTeam = "CNLY_Therapy") And Me.txtDataType <> "HH" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Guidelines Completed Correctly?"
            cboPhysicianOrder_Label.Caption = "Appropriate NCD\LCD?"
            cboCompleteMR_Label.Caption = "Grammar\Spelling\Punction Correct?"
            
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            txtMNGrammar_Comment.visible = False
            txtMNGrammar_Comment.Enabled = False
            cbMNGrammar.visible = False
            cbMNGrammar.Enabled = False
            cboGrammarCorrect_Label.visible = False
            
            txtMNCodingCorrect_Comment.visible = False
            txtMNCodingCorrect_Comment.Enabled = False
            cbMNCodingCorrect.visible = False
            cbMNCodingCorrect.Enabled = False
            cboCodingCorrect_Label.visible = False
            
            Me.cbMNPhysOrder.RowSource = "Y;Yes;1;N;No;0;A;N/a;1"
            
            lblMNClaimHdr.Caption = "Concept Development Claims"
            
    'END 9/11/2013 KCF: toggle questions for Concept Reviews
    
    'BEGIN 2/5/2014 KCF: toggle questions for Concept Reviews
        ElseIf Me.txtAuditTeam = "Bleph" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Criteria met with required documentation?"
            cboPhysicianOrder_Label.Caption = "Supporting diagnosis codes are present?"
            cboCompleteMR_Label.Caption = "LCD language is appropriate?"
            cboGrammarCorrect_Label.Caption = "Visual fields are interpreted correctly?"
            
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            txtMNCodingCorrect_Comment.visible = False
            txtMNCodingCorrect_Comment.Enabled = False
            cbMNCodingCorrect.visible = False
            cbMNCodingCorrect.Enabled = False
            cboCodingCorrect_Label.visible = False
            
            lblMNClaimHdr.Caption = "Blepharoplasty Claims"
            
    'END 2/5/2014 KCF: toggle questions for Concept Reviews
    
    'BEGIN 3/31/2014 KCF: toggle questions for Concept Reviews
        ElseIf Me.txtAuditTeam = "SNF - C2019" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "RUG adjusted appropriately?"
            cboPhysicianOrder_Label.Caption = "Claim re-priced correctly?"
            'cboCompleteMR_Label.Caption = "LCD language is appropriate?"
            'cboGrammarCorrect_Label.Caption = "Visual fields are interpreted correctly?"
            
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            txtMNCompleteMR_Comment.visible = False
            txtMNGrammar_Comment.visible = False
            txtMNCodingCorrect_Comment.visible = False
            txtMNCodingCorrect_Comment.Enabled = False
            cbMNCompleteMR.visible = False
            cbMNGrammar.visible = False
            cbMNCodingCorrect.visible = False
            cbMNCodingCorrect.Enabled = False
            cboCodingCorrect_Label.visible = False
            
            lblMNClaimHdr.Caption = "Concept C2019 - Prepayment Review: Skilled Nursing Facility and Coding Validation"
            
    'END 3/31/2014 KCF: toggle questions for Concept Reviews
    
        ElseIf Me.txtDataType = "IP" Then
    'END 5/28/2013 KCF: toggle questions for Extrapolation questions
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Pertinent Diagnositic Results?"
            cboPhysicianOrder_Label.Caption = "Physician Order Noted Correct?"
            cboCompleteMR_Label.Caption = "Review Complete w/ Complete MR?"
            cboGrammarCorrect_Label.Caption = "Grammar/Punctuation Correct"
            If Me.txtClmStatus = "321" Then
                cboCodingCorrect_Label.Caption = "Was Coding Correct?"
            Else
                cboCodingCorrect_Label.Caption = "Was Admission / Discharge Status Correct?"
            End If
            
    'BEGIN 5/28/2013 KCF: display Rationale & Coding Correct questions
            txtMNCodingCorrect_Comment.visible = True
            txtMNCodingCorrect_Comment.Enabled = True
            cbMNCodingCorrect.visible = True
            cbMNCodingCorrect.Enabled = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "MN Claims"
    'END 5/28/2013 KCF: display Rationale & Coding Correct questions
            
        ElseIf Me.txtDataType = "HH" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Homebound Status Correct?"
            cboPhysicianOrder_Label.Caption = "Conditions of Participation /Technical Criteria?"
            cboCompleteMR_Label.Caption = "Medical Necessity?"
            cboGrammarCorrect_Label.Caption = "Grammar/Spelling/ Punctuation Correct?"
            cboCodingCorrect_Label.Caption = "Claim Edit Detail Correct?"
            
    'BEGIN 5/28/2013 KCF: display Rationale & Coding Correct questions
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "Home Health Claims"
    'END 5/28/2013 KCF: display Rationale & Coding Correct questions
    
    'TO DO HERE: KCF 4/11/2014
    
        ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1983" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Does the documentation show failed therapy?"
            cboPhysicianOrder_Label.Caption = "Is the Patient an Appropriate Surgical Candidate?"
            cboCompleteMR_Label.Caption = "Are the Secondary Manifestations Documented?"
            cboGrammarCorrect_Label.Caption = "Is the Voiding Diary Data Documented?"
            cboCodingCorrect_Label.Caption = "Does the documentation support a successful Test Stimulation?"
            
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "Sacral Nerve Claims"
            
        ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Sacral Nerve CM_1984" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Does the documentation show failed therapy?"
            cboPhysicianOrder_Label.Caption = "Is the Patient an Appropriate Surgical Candidate?"
            cboCompleteMR_Label.Caption = "Are the Secondary Manifestations Documented?"
            cboGrammarCorrect_Label.Caption = "Is the Voiding Diary Data Documented?"
            cboCodingCorrect_Label.Caption = "Does the documentation support a successful Test Stimulation?"
            
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "Sacral Nerve Claims"
    
        ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "PTA" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Was the History and Physical correctly noted?"
            cboPhysicianOrder_Label.Caption = "Was Vascular Examinatino correctly noted?"
            cboCompleteMR_Label.Caption = "Were prior NDE correctly noted?"
            cboGrammarCorrect_Label.Caption = "Were add. requirements for Co/Re/Ca/Intrac PTA correctly noted?"
            cboCodingCorrect_Label.Caption = "Were angina or MRA Reports correctly noted?"
            
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "PTA Claims"
    
    
        ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Osteo Stim" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Was the Rx/ Dispensing Order correct?"
            cboPhysicianOrder_Label.Caption = "Was the Detailed Written Order correct?"
            cboCompleteMR_Label.Caption = "Was the Certificate of Medical Necessity Correct?"
            cboGrammarCorrect_Label.Caption = "Was the Proof of Delivery correct?"
            cboCodingCorrect_Label.Caption = "Was the Spelling/ Grammar/ Punctuation Correct?"
            
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "Osteo Stim Claims"
    
        ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "Hospice" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Was the Hospice Election Statement correct?"
            cboPhysicianOrder_Label.Caption = "Was the Ceritification of Terminal Illness correct?"
            cboCompleteMR_Label.Caption = "Was the Plan of Care correct?"
            cboGrammarCorrect_Label.Caption = "Was the Re-Certification of Terminal Illness correct?"
            cboCodingCorrect_Label.Caption = "Was the Face-to-Face attestation correct?"
            
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "Hospice Claims"
    
        ElseIf Me.MedicalNecessity = "D" And Me.AuditTeam = "NPWT" Then
            cboCorrectDecision_Label.Caption = "Correct Decision?"
            cboPertinentLab_Label.Caption = "Was the Rx Dispensing Order correct?"
            cboPhysicianOrder_Label.Caption = "Was the Detailed Written Order correct?"
            cboCompleteMR_Label.Caption = "Was the Proof of Delivery correct?"
            cboGrammarCorrect_Label.Caption = "Was the Refill Documentation correct?"
            cboCodingCorrect_Label.Caption = "Was the Spelling/ Grammar/ Punctuation correct?"
            
            txtMNCodingCorrect_Comment.visible = True
            cbMNCodingCorrect.visible = True
            cboCodingCorrect_Label.visible = True
    
            txtRationaleCorrect_Comment.visible = True
            cbRationaleCorrect.visible = True
            cbRationaleCorrect_Label.visible = True
            
            lblMNClaimHdr.Caption = "NPWT Claims"
    
        End If
  End If 'End block for If NO Recovery, Else Recovery
Application.Echo True
    
End Sub

Public Sub DRGFieldsEnable()
' QA v 1.1 Begin Comments - Set all Y/N & Comment fields for MN and DRG sections to Enabled = False and Locked = True
'unit test 9/10/2012 kcf - MsgBox ("Hit DRG Fields Enable subroutine")

'KCF 1/5/2015 - Update to the interface to handle formatting for claims submitted.

'If Me.ClmStatus = "321" Then  --KCF 1/5/2015
If (Me.ClmStatus = "321" Or Me.PrevClmStatus = "321") Then
    Me.txtDRGCorrectDecision_Comment.visible = True
    Me.txtDRGCorrectDecision_Comment.Enabled = True
    Me.cbDRGCorrectDecision.visible = True
    Me.cboDRGCorrectDecision_Label.visible = True
    Me.cboDRGCorrectDecision_Label.Caption = "All Dx\Px and Discharge codes correct as billed?"
    Me.cbDRGCorrectDecision.RowSource = "Y;Yes;1"
    Me.cbDRGCorrectDecision.Enabled = True
    
    Me.txtDRGCorrect_Comment.visible = True
    Me.txtDRGCorrect_Comment.Enabled = True
    Me.cbDRGCorrect.visible = True
    Me.cboDRGCorrect_label.visible = True
    Me.cboDRGCorrect_label.Caption = "Potential recovery identified?"
    Me.cbDRGCorrect.RowSource = "Y;Yes;1"
    Me.cbDRGCorrect.Enabled = True
    
    Me.txtDRGCorrectDischarge_Comment.visible = False
    Me.cbDRGCorrectDischarge.visible = False
    Me.cboCorrectDischarge_Label.visible = False
    
    Me.txtDRGCodingChange_Comment.visible = False
    Me.cbDRGCodingChange.visible = False
    Me.cboCodingChangeCorrect_Label.visible = False

    Me.txtDRGClaimReferMN_Comment.visible = False
    Me.cbDRGClaimReferMN.visible = False
    Me.cboClaimReferMN_Label.visible = False

    Me.txtRationaleCorrect_Comment.visible = False
    Me.cbRationaleCorrect.visible = False
    Me.cbRationaleCorrect_Label.visible = False

Else
    Me.txtDRGCorrectDecision_Comment.visible = True
    Me.txtDRGCorrectDecision_Comment.Enabled = True
    Me.cbDRGCorrectDecision.visible = True
    Me.cboDRGCorrectDecision_Label.visible = True
    Me.cboDRGCorrectDecision_Label.Caption = "Was the decision to recover this claim correct?"
    Me.cbDRGCorrectDecision.RowSource = "Y;Yes;1;N;No;1"
    
    Me.txtDRGCorrect_Comment.visible = True
    Me.txtDRGCorrect_Comment.Enabled = True
    Me.cbDRGCorrect.visible = True
    Me.cboDRGCorrect_label.visible = True
    Me.cboDRGCorrect_label.Caption = "Was the review of the Dx\Px correct?"
    Me.cbDRGCorrect.RowSource = "Y;Yes;1;N;No;0"

    Me.txtDRGCorrectDischarge_Comment.visible = True
    Me.txtDRGCorrectDischarge_Comment.Enabled = True
    Me.cbDRGCorrectDischarge.visible = True
    Me.cbDRGCorrectDischarge.Enabled = True
    Me.cboCorrectDischarge_Label.visible = True
    Me.cboCorrectDischarge_Label.Caption = "Was the discharge status reviewed correctly?"
    Me.cbDRGCorrectDischarge.RowSource = "Y;Yes;1;N;No;0"
    
    Me.txtDRGCodingChange_Comment.visible = True
    Me.txtDRGCodingChange_Comment.Enabled = True
    Me.cbDRGCodingChange.visible = True
    Me.cboCodingChangeCorrect_Label.visible = True
    Me.cbDRGCodingChange.Enabled = True
    Me.cboCodingChangeCorrect_Label.Caption = "Are coding clinic and coding guideline citations noted in the Rationale?"
    Me.cbDRGCodingChange.RowSource = "Y;Yes;1;N;No;0"
    
    Me.txtDRGClaimReferMN_Comment.visible = False
    Me.cbDRGClaimReferMN.visible = False
    Me.cboClaimReferMN_Label.visible = False
   
    Me.txtRationaleCorrect_Comment.visible = True
    Me.txtRationaleCorrect_Comment.Enabled = True
    Me.cbRationaleCorrect.visible = True
    Me.cbRationaleCorrect_Label.visible = True
    Me.cbRationaleCorrect_Label.Caption = "Does the rationale properly reflect the changes made and explain the recovery?"
    Me.cbRationaleCorrect.Enabled = True
    Me.cbRationaleCorrect.RowSource = "Y;Yes;1;N;No;0"

End If

Application.Echo True
    
End Sub


Public Sub MNQAScoreCalculation()
'unit test 9/10/2012 by kcf
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For MN questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub

Public Sub HHQAScoreCalculation()
'Created Tuesday 4/23/2013 for the HH Implementation
'Based upon the MN question fields - captions are set in the MNRequireFormat section
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For HH questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


Public Sub ConceptQAScoreCalculation()
'Created Wednesday 9/11/2013 for Concept Review implementation
'Based upon the MN question fields - captions are set in the MNRequireFormat section
'KCF 11/25/2013 - updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For Concept questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 15)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub

'PHP questions KCF Wed 2/17/2016 --------------------------------------
Public Sub PHPQAScoreCalculation()
'Created Wednesday 2/5/2014 for Bleph Review implementation
'Based upon the MN question fields - captions are set in the MNRequireFormat section

Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For PHP questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320.2" Or Me.PrevClmStatus = "320" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 10)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 10)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 10)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 10)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 10)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub
'PHP Questions KCF Wed. 2/17/2016 ---------------------------------------------------------------------------------------------------------------------------------







Public Sub BlephQAScoreCalculation()
'Created Wednesday 2/5/2014 for Bleph Review implementation
'Based upon the MN question fields - captions are set in the MNRequireFormat section

Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For Bleph questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320.2" Or Me.PrevClmStatus = "320" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 30)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub
 
Public Sub SNFC2019QAScoreCalculation()
'Created Wednesday 2/5/2014 for Bleph Review implementation
'Based upon the MN question fields - captions are set in the MNRequireFormat section

Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For Bleph questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320.2" Or Me.PrevClmStatus = "320" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 30)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 10)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 10)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub
 
 
 
 Public Sub ExtrapQAScoreCalculation()
'Created Tuesday 5/28/2013 for the Extrapolation Implementation
'Based upon the MN question fields - captions are set in the MNRequireFormat section
'KCF 11/25/2013 - updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For Extrap questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 60)
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 10)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 10)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 10)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 10)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'New Concept claim scoring

'Sacral Nerve
Public Sub SacralNerveQAScoreCalculation()
'unit test 9/10/2012 by kcf
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For MN questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


'PTA
Public Sub PTAQAScoreCalculation()
'unit test 9/10/2012 by kcf
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For MN questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


'Osteo Stim
Public Sub OsteoStimQAScoreCalculation()
'unit test 9/10/2012 by kcf
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For MN questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


'Hospice
Public Sub HospiceQAScoreCalculation()
'unit test 9/10/2012 by kcf
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For MN questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


'NPWT
Public Sub NPWTQAScoreCalculation()
'unit test 9/10/2012 by kcf
'KCF 11/25/2013 - Updated to include the 320.2 claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer


QAScore_Numerator = 0

'For MN questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbMNCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 100)
    If Me.cbMNCorrectDecision.Column(2) = 0 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbMNCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCorrectDecision.Column(2) * 50) 'As per Lori not to be included in Score, but MN team said weight of 50% for entire score if Recover
        If IsNull(Me.cbMNPertLab.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPertLab.Column(2) * 5)
        If IsNull(Me.cbMNCodingCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCodingCorrect.Column(2) * 5)
        If IsNull(Me.cbMNCompleteMR.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNCompleteMR.Column(2) * 5)
        If IsNull(Me.cbMNGrammar.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNGrammar.Column(2) * 5)
        If IsNull(Me.cbMNPhysOrder.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbMNPhysOrder.Column(2) * 5)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 25)
            
        QAScore_Denominator = 100
        calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
 
lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"

If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
End If
    
End Sub


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Updated KCF 10/8/2014 - updated questions for the contract extension


Sub DRGQAScoreCalculation()
'unit test 9/10/2012 by kcf
'Modified Thursday 4/4/2013 by KCF for DRG implementation
'KCF 11/25/2013 - Updated to include the new claim status
Dim calcQAScore As Integer
Dim QAScore_Denominator As Integer
Dim QAScore_Numerator As Integer

    QAScore_Numerator = 0
    
'For DRG questions - if Non-Recovery (ClmStatus = 321), the score is based only on the 'Correct Decision' Response.

If (Me.ClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main") Or (Me.PrevClmStatus = "321" And Me.Parent.Form.Name = "frm_QA_Review_Main_submitted") Then
    'If IsNull(Me.cbDRGCorrectDecision.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0
    If Me.cbDRGCorrectDecision.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + (Me.cbDRGCorrectDecision.Column(2) * 100)
    If Me.cbDRGCorrect.Column(2) = 1 Then QAScore_Numerator = QAScore_Numerator + 50
    
    QAScore_Denominator = 100
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100

Else
    If (Me.Parent.Form.Name = "frm_QA_Review_Main" And (Me.ClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.ClmStatus = "322" Or Me.ClmStatus = "314")) Or (Me.Parent.Form.Name = "frm_QA_Review_Main_Submitted" And (Me.PrevClmStatus = "320" Or Me.ClmStatus = "320.2" Or Me.PrevClmStatus = "322" Or Me.PrevClmStatus = "314")) Then
        If IsNull(Me.cbDRGCorrectDecision.Column(2)) Then QAScore_Numerator = 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbDRGCorrectDecision.Column(2) * 50)
        If IsNull(Me.cbDRGCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbDRGCorrect.Column(2) * 20)
        If IsNull(Me.cbDRGCorrectDischarge.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbDRGCorrectDischarge.Column(2) * 10)
        If IsNull(Me.cbDRGCodingChange.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbDRGCodingChange.Column(2) * 10)
        If IsNull(Me.cbRationaleCorrect.Column(2)) Then QAScore_Numerator = QAScore_Numerator + 0 Else QAScore_Numerator = QAScore_Numerator + (Me.cbRationaleCorrect.Column(2) * 10)
  
        QAScore_Denominator = 100
    
    calcQAScore = (QAScore_Numerator / QAScore_Denominator) * 100
    End If
End If
    lblQAScoreCalc.Caption = "QA Score: " & calcQAScore & "%"
    
    If Me.Reviewer <> "" Then 'Do not record default score of 0 if no work done on the claim
    Me.txtQAScore = calcQAScore
    End If

End Sub

Sub MN_RequiredFormat()
'Created Thursday 9/12/2012 - Formatting will be based upon the ClmSTatus
'If 321 (No Recovery, only the first Coding Correct field is required.
'3/1/2013 KCF: Will check the parent form to detrmine whether to base formatting on the ClmStatus or PrevClmStatus
'2/5/2104 KCF: Set tab stop for Concept & Belph claims

     Me.lnRationale.visible = False
    
    If (Me.Parent.Name = "frm_QA_Review_Main" And Me.txtClmStatus = "321") Or (Me.Parent.Name = "frm_QA_Review_Main_Submitted" And Me.txtPrevClmStatus = "321") Then
        'Only color code the Correct Decision field for validation
        Me.cbMNCodingCorrect.BorderColor = 0
        Me.cbMNCompleteMR.BorderColor = 0
        Me.cbMNCorrectDecision.BorderColor = 255
        Me.cbMNGrammar.BorderColor = 0
        Me.cbMNPertLab.BorderColor = 0
        Me.cbMNPhysOrder.BorderColor = 0
        Me.cbRationaleCorrect.BorderColor = 0
        'Only add border for the Correct Decision comment if the corresponding response is 'No'
        If Me.cbMNCorrectDecision.Column(0) = "N" Then Me.txtMNCorrectDecision_Comment.BorderColor = 255 Else Me.txtMNCorrectDecision_Comment.BorderColor = 0
    
    Else
        'All the MN questions are required
        Me.cbMNCodingCorrect.BorderColor = 255
        Me.cbMNCompleteMR.BorderColor = 255
        Me.cbMNCorrectDecision.BorderColor = 255
        Me.cbMNGrammar.BorderColor = 255
        Me.cbMNPertLab.BorderColor = 255
        Me.cbMNPhysOrder.BorderColor = 255
        Me.cbRationaleCorrect.BorderColor = 255
        
        'Add border if any of the questions have a 'No' response
        If Me.cbMNCodingCorrect.Column(0) = "N" Then Me.txtMNCodingCorrect_Comment.BorderColor = 255 Else Me.txtMNCodingCorrect_Comment.BorderColor = 0
        If Me.cbMNCompleteMR.Column(0) = "N" Then Me.txtMNCompleteMR_Comment.BorderColor = 255 Else Me.txtMNCompleteMR_Comment.BorderColor = 0
        If Me.cbMNCorrectDecision.Column(0) = "N" Then Me.txtMNCorrectDecision_Comment.BorderColor = 255 Else Me.txtMNCorrectDecision_Comment.BorderColor = 0
        If Me.cbMNGrammar.Column(0) = "N" Then Me.txtMNGrammar_Comment.BorderColor = 255 Else Me.txtMNGrammar_Comment.BorderColor = 0
        If Me.cbMNPertLab.Column(0) = "N" Then Me.txtMNPertLab_Comment.BorderColor = 255 Else Me.txtMNPertLab_Comment.BorderColor = 0
        If Me.cbMNPhysOrder.Column(0) = "N" Then Me.txtMNPhysOrder_Comment.BorderColor = 255 Else Me.txtMNPhysOrder_Comment.BorderColor = 0
        If Me.cbRationaleCorrect.Column(0) = "N" Then Me.txtRationaleCorrect_Comment.BorderColor = 255 Else Me.txtRationaleCorrect_Comment.BorderColor = 0
    
    End If

End Sub


Sub DRG_RequiredFormat()
'unit test 9/10/2012 by kcf
'Revised Wednesday 4/10/2013 by KCF for DRG Implementation

    Me.lnRationale.visible = False
    
    If (Me.Parent.Name = "frm_QA_Review_Main" And Me.txtClmStatus = "321") Or (Me.Parent.Name = "frm_QA_Review_Main_Submitted" And Me.txtPrevClmStatus = "321") Then
        'Only Color Code the Correct Decision field for validation
        Me.cbDRGCorrectDecision.BorderColor = 0
        Me.cbDRGClaimReferMN.BorderColor = 0
        Me.cbDRGCodingChange.BorderColor = 0
        Me.cbDRGCorrect.BorderColor = 0
        Me.cbDRGCorrectDischarge.BorderColor = 0
        Me.cbRationaleCorrect.BorderColor = 0

        'Only add border for comments if correct Decision is No
        If Me.cbDRGCorrectDecision.Column(0) = "N" Then Me.txtDRGCorrectDecision_Comment.BorderColor = 255 Else Me.txtDRGCorrectDecision_Comment.BorderColor = 0
    
    Else
        'All DRG fields are required
         Me.cbDRGCorrectDecision.BorderColor = 255
        Me.cbDRGClaimReferMN.BorderColor = 255
        Me.cbDRGCodingChange.BorderColor = 255
        Me.cbDRGCorrect.BorderColor = 255
        Me.cbDRGCorrectDischarge.BorderColor = 255
        Me.cbRationaleCorrect.BorderColor = 255
        
        'Add the border if any of the questions are answered No
        If Me.cbRationaleCorrect.Column(0) = "N" Then Me.txtRationaleCorrect_Comment.BorderColor = 255 Else Me.txtRationaleCorrect_Comment.BorderColor = 0
        If Me.cbDRGCorrectDecision.Column(0) = "N" Then Me.txtDRGCorrectDecision_Comment.BorderColor = 255 Else Me.txtDRGCorrectDecision_Comment.BorderColor = 0
        If Me.cbDRGClaimReferMN.Column(0) = "N" Then Me.txtDRGClaimReferMN_Comment.BorderColor = 255 Else Me.txtDRGClaimReferMN_Comment.BorderColor = 0
        If Me.cbDRGCodingChange.Column(0) = "N" Then Me.txtDRGCodingChange_Comment.BorderColor = 255 Else Me.txtDRGCodingChange_Comment.BorderColor = 0
        If Me.cbDRGCorrect.Column(0) = "N" Then Me.txtDRGCorrect_Comment.BorderColor = 255 Else Me.txtDRGCorrect_Comment.BorderColor = 0
        If Me.cbDRGCorrectDischarge.Column(0) = "N" Then Me.txtDRGCorrectDischarge_Comment.BorderColor = 255 Else Me.txtDRGCorrectDischarge_Comment.BorderColor = 0
    End If
    
    
End Sub

Private Sub IPFieldsEnable()
' Enable the IP fields (for both DRG & MN questions) - enable when the record is not locked
'unit test 9/10/2012 by kcf
    Me.chkReturn.Enabled = True
    Me.ReviewComment.Enabled = True
    Me.cbRationaleCorrect.Enabled = True
    Me.txtRationaleCorrect_Comment.Enabled = True
    'Me.Parent.Form.cmdSubmitQA.Enabled = True

End Sub

Private Sub IPFieldsDisable()
'Disable the IP fields (for both DRG & MN questions) - enable when the record is not locked
'unit test 9/10/2012 by kcf
    Me.chkReturn.Enabled = False
    Me.ReviewComment.Enabled = False
    Me.cbRationaleCorrect.Enabled = False
    Me.txtRationaleCorrect_Comment.Enabled = False
    If Me.Parent.Name = "frm_QA_Review_Main" Then
        Me.Parent.Form.cmdSubmitQA.Enabled = False
    End If

End Sub



Private Sub txtStatFromDate_Exit(Cancel As Integer)
    Call cmdCalcAuditorStats_Click
End Sub

Private Sub txtStatToDate_Exit(Cancel As Integer)
       Call cmdCalcAuditorStats_Click
End Sub
