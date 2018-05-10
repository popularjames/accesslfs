Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrUserProfile As String

Private miAppPermission As Integer
Private mstrUserName As String

Const CstrFrmAppID As String = "DashManagement"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Load()
    
Dim MyAdo As clsADO

    Call Account_Check(Me)

    Dim rsPermission As ADODB.RecordSet

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")

    mstrUserName = Identity.UserName

    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    RefreshMain
    
End Sub
    

Private Sub RefreshMain()
    
    Dim intStatus As Integer
    Dim fmrStatus As Form_ScrStatus
    Dim sMsg As String
    
    intStatus = 1
    Set fmrStatus = New Form_ScrStatus
    With fmrStatus
         .ShowCancel = True
         .ShowMessage = False
         .ShowMessage = True
         .ProgVal = 0
         .ProgMax = 5
         .TimerInterval = 50
         .show
         .visible = True
     End With
    
    DoEvents
    
    
    
    
    Dim strCurrentMonth As String
    Dim strcurrentYear As String
    Dim strCurrentdate As String
    
    strCurrentdate = Trim(str(DatePart("yyyy", Now))) & "-" & Right("0" & Trim(str(DatePart("m", Now))), 2) & "-" & Right("0" & Trim(str(DatePart("d", Now))), 2)
    strcurrentYear = Trim(str(DatePart("yyyy", Now))) & "-" & "01" & "-" & "01"
    strCurrentMonth = Trim(str(DatePart("yyyy", Now))) & "-" & Right("0" & Trim(str(DatePart("m", Now))), 2) & "-" '& Right("0" & Trim(str(DatePart("d", Now))), 2)
    
    'AR Setup Chart
    Me.mschart.RowSource = "TRANSFORM Sum([ArSetUpAmt]) AS [SumOfArSetUpAmt] SELECT [ReportWk] FROM [RPT_R0050B] WHERE ReportWk > '" & Format(DateAdd("d", -21, str(DatePart("yyyy", Now)) & "-" & Right("0" & Trim(str(DatePart("m", Now))), 2) & "-" & Right("0" & Trim(str(DatePart("d", Now))), 2)), "yyyy-mm-dd") & "'   GROUP BY [ReportWk] PIVOT [PayerName];"
    
     sMsg = "AR SETUP"
     fmrStatus.ProgVal = 2
     fmrStatus.StatusMessage sMsg
     DoEvents
    
    'Productivity Chart
    Me.mschart_Productivity.RowSource = "TRANSFORM Sum([Processed]) AS [SumOfProcessed] SELECT [ProductivityWk] FROM [RPT_R0000P00N] where team not in ('ic doc', 'management','--','cnly auto pa', 'r3test') and ProductivityWk > '" & Format(DateAdd("d", -21, str(DatePart("yyyy", Now)) & "-" & Right("0" & Trim(str(DatePart("m", Now))), 2) & "-" & Right("0" & Trim(str(DatePart("d", Now))), 2)), "yyyy-mm-dd") & "'    GROUP BY [ProductivityWk] PIVOT [Team];"
    
     sMsg = "PRODUCTIVITY"
     fmrStatus.ProgVal = 3
     fmrStatus.StatusMessage sMsg
     DoEvents
     
    'Appeal Timeline Chart
    Me.mschartAppeal.RowSource = "SELECT * FROM v_REPORT_Appeal_Lag where ARSETUP >= '" & left("2012-01", 8) & "' ORDER BY ARSETUP ASC"
        
     sMsg = "APPEALS"
     fmrStatus.ProgVal = 4
     fmrStatus.StatusMessage sMsg
     DoEvents

    
    'Aging AR Chart
    Me.msChartAgingAR.RowSource = "TRANSFORM Sum([Amt]) AS [SumOfAmt] SELECT [DayRange] FROM [v_REPORT_AR_AGING]   GROUP BY [DayRange] PIVOT [PayerName]"
    
     sMsg = "AGINNG"
     fmrStatus.ProgVal = 5
     fmrStatus.StatusMessage sMsg
     DoEvents

    Me.frm_Dash_TrialInvoice.SetFocus
    DoCmd.Close acForm, fmrStatus.Name
    Set fmrStatus = Nothing
    
End Sub
Private Sub Label23_Click()
   RefreshMain
End Sub
