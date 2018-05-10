'---------------------------------------------------------------------------------------
' Module    : SCR_BuildScreensProfilesControls
' Author    : SA
' Date      : 11/7/2012
' Purpose   : Moved from SCR_MainScreens. Should be deleted when top tab section is reworked.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub BuildScreensProfilesControls()
'Rebuild table SCR_ScreensProfilesControls
On Error GoTo ErrorHappened
    Dim genUtils As New CT_ClsGeneralUtilities
    Dim frm As New Form_SCR_MainScreens
    genUtils.ToggleAccessMenus (True)
    
    If MsgBox("Do you want to recalculate control positions?", vbInformation Or vbYesNo Or vbDefaultButton2, "Recalculate Control Positions?") = vbNo Then
        Exit Sub
    End If
    
    Dim ctrl As Control
    Dim ctr As Integer
    Dim pidRestore As Integer
    Dim pidPrimary As Integer
    Dim pidCollapsed As Integer
    
    Dim isPrimary As Boolean
    
    pidRestore = 4
    pidPrimary = 2
    pidCollapsed = 3
    
    CurrentDb.Execute ("DELETE FROM SCR_ScreensProfilesControls WHERE ProfileID = 2 OR ProfileID = 3 OR ProfileID = 4")
        
    For ctr = 0 To frm.FormHeader.Controls.Count - 1
        Set ctrl = frm.FormHeader.Controls(ctr)
        
        If Not ControlToSkip(ctrl.Name) Then
            isPrimary = ((ctrl.top + ctrl.Height) < ((frm.CmdScreenSave.top + frm.CmdScreenSave.Height) + 25))
            
            If isPrimary Then
                CurrentDb.Execute ( _
                    "INSERT INTO SCR_ScreensProfilesControls([ProfileID], [ControlName], [Top], [Left], [Width], [Height], [Visible], [UpdateVisibility]) " & _
                    "SELECT " & _
                    "   " & pidPrimary & ", " & _
                    "   '" & ctrl.Name & "', " & _
                    "   " & ctrl.top & ", " & _
                    "   " & ctrl.left & ", " & _
                    "   " & ctrl.Width & ", " & _
                    "   " & ctrl.Height & ", " & _
                    "   " & "1, " & _
                    "   " & "0")
            Else
                CurrentDb.Execute ( _
                    "INSERT INTO SCR_ScreensProfilesControls([ProfileID], [ControlName], [Top], [Left], [Width], [Height], [Visible], [UpdateVisibility]) " & _
                    "SELECT " & _
                    "   " & pidPrimary & ", " & _
                    "   '" & ctrl.Name & "', " & _
                    "   " & "0" & ", " & _
                    "   " & "0" & ", " & _
                    "   " & "1" & ", " & _
                    "   " & "1" & ", " & _
                    "   " & "1, " & _
                    "   " & "0")
            End If
            
                CurrentDb.Execute ( _
                    "INSERT INTO SCR_ScreensProfilesControls([ProfileID], [ControlName], [Top], [Left], [Width], [Height], [Visible], [UpdateVisibility]) " & _
                    "SELECT " & _
                    "   " & pidRestore & ", " & _
                    "   '" & ctrl.Name & "', " & _
                    "   " & ctrl.top & ", " & _
                    "   " & ctrl.left & ", " & _
                    "   " & ctrl.Width & ", " & _
                    "   " & ctrl.Height & ", " & _
                    "   " & "1, " & _
                    "   " & "0")
                    
                CurrentDb.Execute ( _
                    "INSERT INTO SCR_ScreensProfilesControls([ProfileID], [ControlName], [Top], [Left], [Width], [Height], [Visible], [UpdateVisibility]) " & _
                    "SELECT " & _
                    "   " & pidCollapsed & ", " & _
                    "   '" & ctrl.Name & "', " & _
                    "   " & "0" & ", " & _
                    "   " & "0" & ", " & _
                    "   " & "1" & ", " & _
                    "   " & "1" & ", " & _
                    "   " & "1, " & _
                    "   " & "0")
        End If
    Next
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "SCR_BuildScreensProfilesControls:BuildScreensProfilesControls"
    Resume ExitNow
    Resume
End Sub

Private Function ControlToSkip(ByVal CtrlName As String) As Boolean
    Dim Result As Boolean
    
    Result = False
    
    If CtrlName = "TabsHead" Then
        Result = True
    End If
    
    If CtrlName = "PageCondFormats" Then
        Result = True
    End If
    
    If CtrlName = "Child164" Then
        Result = True
    End If
    
    If UCase(left(CtrlName, 2)) = "PG" Then
        Result = True
    End If
    
    If Len(CtrlName) > 10 Then
        If UCase(Right(CtrlName, 10)) = "_COLLAPSED" Then
            Result = True
        End If
        
        If UCase(Right(CtrlName, 5)) = "SMALL" Then
            Result = True
        End If
        If UCase(Right(CtrlName, 5)) = "LABEL" Then
            Result = True
        End If
    End If
    
    ControlToSkip = Result

End Function