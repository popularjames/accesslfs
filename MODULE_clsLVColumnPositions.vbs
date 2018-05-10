Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





''' Last Modified: 10/16/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This is essentially a dictionary of dictionaries
'''     that is used to keep track of the list view column positions
'''     given a field name, it returns the List view SubItem index
'''
'''  TODO:
'''  =====================================
'''  -
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 05/19/2015 - KD: Extended - keep the code of getting a value out of a list item short, sweet and not repeated!!
'''  - 06/19/2014 - Created class
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

Private cdctMainDict As Scripting.Dictionary


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property



Public Function SetDetails(ByVal sListViewName As String, oListView As Object) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim dctLVCols As Scripting.Dictionary
Dim oLI As ListItem
Dim iSubItem As Integer

    strProcName = ClassName & ".SetDetails"
    
    sListViewName = UCase(sListViewName)
    
    If oListView.ListItems.Count = 0 Then
        GoTo Block_Exit
    End If
    If cdctMainDict.Exists(sListViewName) = False Then
        Set dctLVCols = New Scripting.Dictionary
        
        Set oLI = oListView.ListItems(1)
        
        For iSubItem = 1 To oListView.ColumnHeaders.Count
            dctLVCols.Add UCase(oListView.ColumnHeaders(iSubItem)), iSubItem - 1
        Next
        cdctMainDict.Add sListViewName, dctLVCols
        
    Else
        ' I guess we should see if it changed..
        
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function GetDetails(ByVal sListViewName As String, ByVal sFieldName As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim dctLVCols As Scripting.Dictionary


    strProcName = ClassName & ".GetDetails"
    
    sListViewName = UCase(sListViewName)
    sFieldName = UCase(sFieldName)
    
    If cdctMainDict.Exists(sListViewName) = True Then
        
        Set dctLVCols = cdctMainDict.Item(sListViewName)
        
        If dctLVCols.Exists(sFieldName) = True Then
            GetDetails = CInt(dctLVCols.Item(sFieldName))
        Else
            ' return - 1??
            GetDetails = -1
        End If
    Else
        ' I guess we return - 1 ??
        GetDetails = -1
        Stop
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function GetLiValue(oLI As Object, ByVal sListViewName As String, ByVal sValueDesiredName As String) As Variant
On Error GoTo Block_Err
Dim strProcName As String
Dim lCol As Long

    strProcName = ClassName & ".GetLiValue"
    
    lCol = GetDetails(sListViewName, sValueDesiredName)
    If lCol = 0 Then
        GetLiValue = oLI.Text
    ElseIf lCol > 0 Then
        GetLiValue = oLI.SubItems(lCol)
    Else
        GetLiValue = "- Column Value not found, check recordsource -"
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub Class_Initialize()
    Set cdctMainDict = New Scripting.Dictionary
End Sub