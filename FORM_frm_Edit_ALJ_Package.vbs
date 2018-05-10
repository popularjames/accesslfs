Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2014-06-18 VS: ALJ Package Details Form
'2014-07-16 VS: Finished Edit Package functionality
'2014-09-09 VS: Need to deploy the change
Dim packagePath As String
Dim FileName As String
Dim strSQL As String

Private Sub Form_Load()
Dim Phone As String
Dim phoneSource As String
Dim judgeFullName As String
Phone = DLookup("[HearingPhone]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
Phone = left(Phone, Len(Phone) - 1)
phoneSource = DLookup("[HearingPhoneSource]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
judgeFullName = DLookup("[JudgeName]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")

Me.txtPackageName = algPackageName
Me.cmbALJJudgeName = DLookup("[JudgeName]", "APPEAL_XREF_ALJ_Judges", "FirstName + ' ' + JudgeName ='" & Replace(judgeFullName, "'", "''") & "'")
Me.cmbParticipant = DLookup("[PartTitle]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'") & " " & _
                    DLookup("[ConnollyHearingParticipantName]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'") & " " & _
                    DLookup("[PartDegree]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
Me.cmbParticipant2 = DLookup("[PartTitle2]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'") & " " & _
                    DLookup("[ConnollyHearingParticipantName2]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'") & " " & _
                    DLookup("[PartDegree2]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
Me.txtHearingDt = Format(DLookup("[HearingDate]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "mm/dd/yyyy hh:mm:ss AMPM")
Me.cmbAppellant = DLookup("[AppellantFolderName]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
Me.cmbPartyStatus = DLookup("[PartyStatus]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
Me.txtResponseDate = Format(DLookup("[FaxDate]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "mm/dd/yyyy")

Call SetPartyStatusVisible

If phoneSource <> "Connolly" Then
    Me.txtHearingPhone = Phone
    Me.txtHearingPasscode = DLookup("[HearingPasscode]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
End If

Me.txtEditNote = DLookup("[EditNote]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'")
End Sub

Private Sub cmdExit_Click()

    DoCmd.Close
    
End Sub

Private Sub cmdSave_Click()

        Dim ResponseDate As String
        
        algHearDate = Format(Me.txtHearingDt.Value, "yyyy-mm-dd hh:mm:ss")
        algJudgeName = Nz(Trim(Me.cmbALJJudgeName.Value), "")
        algAppellant = Nz(Me.cmbAppellant.Value, "")
        algParticipant = Nz(Me.cmbParticipant.Value, "")
        algParticipant2 = Nz(Me.cmbParticipant2.Value, "")
        algHearingPhone = Nz(Me.txtHearingPhone.Value, "")
        algHearingPasscode = Nz(Me.txtHearingPasscode.Value, "")
        algEditNote = Nz(Me.txtEditNote.Value, "")
        algPartyStatus = Nz(Me.cmbPartyStatus, "")
        
        If (InStr(Me.txtResponseDate.Value, "/") = 0) Then
            ResponseDate = left(Me.txtResponseDate.Value, 2) + "/" + Right(left(Me.txtResponseDate.Value, 4), 2) + "/" + Right(Me.txtResponseDate.Value, 4)
        Else
            ResponseDate = Me.txtResponseDate.Value
        End If
        
        algNOIResponseDate = Format(ResponseDate, "yyyy-mm-dd hh:mm:ss")
        
        If algPartyStatus = PARTY Then
            algPartyWitn = Nz(Me.cmbWitn, "")
            algPartyInfo = Nz(Me.cmbInfo, "")
            algPartyDisc = Nz(Me.cmbDiscover, "")
        End If
        
        Call Edit_ALJ_Package(Me.txtPackageName)
        Call SetPartyStatusVisible
    
End Sub

Public Function SetPartyStatusVisible() As String

    If Me.cmbPartyStatus = PARTY Then
       Me.cmbDiscover.visible = True
       Me.lblDisc.visible = True
       Me.cmbInfo.visible = True
       Me.lblInfo.visible = True
       Me.cmbWitn.visible = True
       Me.lblWitn.visible = True
       Me.lnDisc.visible = True
       Me.lnInfo.visible = True
       Me.lnWit.visible = True
       
       Me.cmbDiscover = Nz(DLookup("[PartyDisc]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "")
       Me.cmbWitn = Nz(DLookup("[PartyWitn]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "")
       Me.cmbInfo = Nz(DLookup("[PartyInfo]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "")
       
    Else
       Me.cmbDiscover.visible = False
       Me.cmbInfo.visible = False
       Me.cmbWitn.visible = False
       Me.lblDisc.visible = False
       Me.lblInfo.visible = False
       Me.lblWitn.visible = False
       Me.lnDisc.visible = False
       Me.lnInfo.visible = False
       Me.lnWit.visible = False
       
    End If
    
    Me.Repaint
    
End Function
