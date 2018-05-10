Option Compare Database
Option Explicit



''' Last Modified: 03/20/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Just a bunch of utilities for various form / report controls
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 03/20/2012 - Created class
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



Private Const ClassName As String = "mod_ControlFunctions"


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function SelectListBoxItemFromText(ByRef oListBox As listBox, ByVal strTextToFind As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim iIndex As Integer

    strProcName = ClassName & ".SelectComboBoxItemFromText"

    For iIndex = 0 To oListBox.ListCount - 1

        If UCase(oListBox.ItemData(iIndex)) = UCase(strTextToFind) Then
            SelectListBoxItemFromText = iIndex
            oListBox.Selected(iIndex) = True
            GoTo Block_Exit
        End If
    Next
    ' if we get here then we didn't find it
    iIndex = -1

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, "ERROR!"
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Use this function to:
''' Find text in a multi column combo box and select it
''' If the combo box is invisible then this has to quickly make it visible so
''' it can setfocus which is required in order to set the ListIndex
''' But then it attempts to set it Back invisible of course Access isn't
''' as friendly as it could be so if our Screen.PreviousControl gets interrupted or
''' something weird happens, well, it's going to error (which I've "eaten" with
''' resume next..)
'''
Public Function SelectComboBoxItemFromText(ByRef oComboBox As ComboBox, ByVal strTextToFind As String, _
        Optional bSearchAllCols As Boolean, Optional iColumn As Integer = -1) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim iIndex As Integer
Dim iColIdx As Integer
Dim bWasInvisible As Boolean
Dim oCtrl As Control
Dim iHiBounds As Integer

    strProcName = ClassName & ".SelectComboBoxItemFromText"
    
    If oComboBox.visible = False Then bWasInvisible = True
            
            ' Use default of 0
    If iColumn < 0 Then iColumn = 0
    
    If bSearchAllCols = True Then
        iHiBounds = oComboBox.ListCount - 1
        iColumn = 0
    Else
        iHiBounds = iColumn
    End If

    For iIndex = iColumn To oComboBox.ListCount - 1
        If bSearchAllCols = True Then
            For iColIdx = 0 To iHiBounds
                If UCase(Nz(oComboBox.Column(iColIdx, iIndex), "")) = UCase(strTextToFind) Then
                    SelectComboBoxItemFromText = iIndex
                        ' In order to set the selection, it needs to have focus
                        ' in order to have focus it needs to be visible
                    oComboBox.visible = True
                    oComboBox.SetFocus
                    
                        ' Grab our previous control so we can reset the focus
                    Set oCtrl = screen.PreviousControl
                    
                    oComboBox.ListIndex = iIndex - 1
                    
                        ' I can't stand doing this but...
                    On Error Resume Next
                        If Not oCtrl Is Nothing Then oCtrl.SetFocus
                        If bWasInvisible = True Then oComboBox.visible = False
                    On Error GoTo Block_Err
                        ' All done, terminate
                    GoTo Block_Exit
                End If
            Next
        End If
    Next
    
    ' if we get here then we didn't find it
    iIndex = -1

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, "ERROR!"
    GoTo Block_Exit
End Function