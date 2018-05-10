Option Compare Database
Option Explicit

Public ColObjectInstances As New Collection
Public ColWindows As New Collection


Public Sub RemoveObjectInstance(ctl As Object)
Dim iCurrentSetting As Integer
    iCurrentSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    On Error Resume Next
    ColObjectInstances.Remove ctl.hwnd & " "
    Err.Clear
    Application.SetOption "Error Trapping", iCurrentSetting
End Sub


Public Sub RemoveWindow(ctl As Object)
Dim iCurrentSetting As Integer
    iCurrentSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    On Error Resume Next
    ColWindows.Remove ctl.hwnd & " "
    Err.Clear
    Application.SetOption "Error Trapping", iCurrentSetting
End Sub


Public Function CheckForNull(CtlCollection As Controls, DisplayErrorMsg As Boolean) As Boolean
    Dim ctl As Control
    Dim bIsNull As Boolean
    Dim strErrMsg As String
    bIsNull = False
    
    strErrMsg = " **** ERROR **** " & vbCrLf & vbCrLf
    
    For Each ctl In CtlCollection
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            If IsNull(ctl.Value) And ctl.Tag <> "SKIP" Then
                bIsNull = True
                If DisplayErrorMsg Then
                    If ctl.Controls Is Nothing Then
                        strErrMsg = strErrMsg & "Field '" & ctl.Name & "' can not be null." & vbCrLf
                    Else
                        strErrMsg = strErrMsg & "Field '" & ctl.Controls(0).Caption & "' can not be null." & vbCrLf
                    End If
                End If
            End If
        End If
    Next
    
    If bIsNull And DisplayErrorMsg Then
        MsgBox strErrMsg, vbCritical
    End If
    
    CheckForNull = bIsNull
End Function