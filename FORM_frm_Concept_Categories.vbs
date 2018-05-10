Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 06/10/2013
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 06/10/2013 - Created...
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################




Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Private Sub cmdAddCat_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdAddCat_Click"
    
    If Nz(Me.txtAddConceptCategory, "") = "" Then
        LogMessage strProcName, "USER ERROR", "Please add the category to be added in the text box before you click 'Add'"
        GoTo Block_Exit
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
Stop
        .sqlString = "usp_Concept_Categories_Add"
        .Parameters.Refresh
        .Parameters("@pCategory") = Me.txtAddConceptCategory
        .Execute
    End With
    Me.txtAddConceptCategory = ""
    
Block_Exit:
    Set oAdo = Nothing
    Call RefreshData
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdFinish_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Command6_Click()
    Debug.Print Me.Detail.Height
Stop
End Sub

Private Sub Form_Load()
    Me.InsideHeight = 11805
    Call RefreshData
End Sub

Public Sub RefreshData()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".RefreshData"
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT ConceptCatId, CategoryName FROM CONCEPT_Xref_ConceptCategories ORDER BY CategoryName"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not get the list of concept categories for some reason!"
            GoTo Block_Exit
        End If
    End With
    
    Me.txtConceptCategory.ControlSource = "CategoryName"
    Set Me.RecordSet = oRs
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
