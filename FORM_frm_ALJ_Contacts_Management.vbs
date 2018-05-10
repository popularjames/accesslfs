Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private SearchType As String
Private searchFor As String
Private strSQL As String
Private Const JUDGE_LAST_NAME As String = "Judge Last Name"
Private Const LAW_FIRM_NAME As String = "Law Firm Name"
Private Const LAW_FIRM_CONTACT As String = "Law Firm Contact Name"
Private Const PROVIDER_NAME As String = "Provider Name"
Private Const PROVIDER_CONTACT_NAME As String = "Provider Contact Name"
Private Const PART_LAST_NAME = "Connolly Participant Last Name"
Private Const PART_NAME = "Connolly Participant Full Name"

Private Sub cmdSearchALJContacts_Click()

If Nz(Me.txtSearchFor, "") = "" Or Nz(Me.cboSearchBy, "") = "" Then
    MsgBox ("Please enter Search Criteria!")
    Exit Sub
End If

SearchType = Me.cboSearchBy
searchFor = Me.txtSearchFor

Select Case SearchType
    Case JUDGE_LAST_NAME
        strSQL = "select * from APPEAL_XREF_ALJ_Judges_Temp where JudgeName Like '*" & searchFor & "*'"
        Me.APPEAL_XREF_ALJ_Judges.Form.RecordSource = strSQL
    Case LAW_FIRM_CONTACT, LAW_FIRM_NAME
            If SearchType = LAW_FIRM_NAME Then
               strSQL = "select * from APPEAL_XREF_ALJ_Hearing_LawFirms_Temp where LawFirm Like '*" & searchFor & "*'"
            Else
               strSQL = "select * from APPEAL_XREF_ALJ_Hearing_LawFirms_Temp where ContactFullName Like '*" & searchFor & "*'"
            End If
            Me.APPEAL_XREF_ALJ_Hearing_LawFirms.Form.RecordSource = strSQL
   Case PROVIDER_NAME, PROVIDER_CONTACT_NAME
       If SearchType = PROVIDER_NAME Then
               strSQL = "select * from APPEAL_XREF_ALJ_Hearing_Hospitals_Temp where HospitalName Like '*" & searchFor & "*'"
            Else
               strSQL = "select * from APPEAL_XREF_ALJ_Hearing_Hospitals_Temp where ContactFullName Like '*" & searchFor & "*'"
            End If
            Me.frm_ALJ_Providers.Form.RecordSource = strSQL
   Case PART_LAST_NAME, PART_NAME
       If SearchType = PART_NAME Then
               strSQL = "select * from APPEAL_XREF_ALJ_Hearing_Connolly_Participants_Temp where FullName Like '*" & searchFor & "*'"
            Else
               strSQL = "select * from APPEAL_XREF_ALJ_Hearing_Connolly_Participants_Temp where LastName Like '*" & searchFor & "*'"
            End If
            Me.frm_ALJ_Connolly_Participants.Form.RecordSource = strSQL

   End Select

End Sub

Private Sub cmsRefresh_Click()

Call Refresh_Contacts
Me.APPEAL_XREF_ALJ_Judges.Form.RecordSource = "select * from APPEAL_XREF_ALJ_Judges_Temp"
Me.APPEAL_XREF_ALJ_Hearing_LawFirms.Form.RecordSource = "select * from APPEAL_XREF_ALJ_Hearing_LawFirms_Temp"
Me.frm_ALJ_Connolly_Participants.Form.RecordSource = "select * from APPEAL_XREF_ALJ_Hearing_Connolly_Participants_Temp"
Me.frm_ALJ_Providers.Form.RecordSource = "select * from APPEAL_XREF_ALJ_Hearing_Hospitals_Temp"
Me.cboSearchBy.Value = ""
Me.txtSearchFor.Value = ""

End Sub
