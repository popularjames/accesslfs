Option Compare Database
Option Explicit

Public frm_scrFilterHist As Form_SCR_ScreensFilterHistory

'This function is called from the Screen Events. this fucntion is registered in the CnlyScreensEvents table
'for the user to bind the Filter History Form from the Screen
Public Function BindFilterHisory(FormID As Byte)
    On Error GoTo BindFilterHisoryError
    If IsFormLoaded("SCR_ScreensFilterHistory") = False Then
        Set frm_scrFilterHist = New Form_SCR_ScreensFilterHistory
    End If
    
     frm_scrFilterHist.SetActiveScreen FormID

BindFilterHisoryExit:
    Exit Function

BindFilterHisoryError:
    If Err.Number = 2467 Then
        Set frm_scrFilterHist = Nothing
        DoCmd.Close acForm, "SCR_ScreensFilterHistory"
        Set frm_scrFilterHist = New Form_SCR_ScreensFilterHistory
        Resume BindFilterHisoryExit
    End If
    
End Function

'This function is called from the Screen Events. this fucntion is registered in the CnlyScreensEvents table
'for the user to unbind the Filter History Form from the Screen
Public Function UnbindFilterHisory(FormID As Byte)
    On Error Resume Next
    If IsFormLoaded("SCR_ScreensFilterHistory") Then
        frm_scrFilterHist.UnregisterScreen FormID
        
        If frm_scrFilterHist.screenCount = 0 Then
            Set frm_scrFilterHist = Nothing
        End If
    End If
End Function

'it displays the Filter History Form. if not created it creates one and show the form to the user
'it will not display anything (set form visible to false) if no screen is loaded.
Public Function FilterHistoryShow() As Boolean
    On Error GoTo FilterHistoryShowError
    Dim FormID As Byte
    
    If frm_scrFilterHist Is Nothing Then
        Set frm_scrFilterHist = New Form_SCR_ScreensFilterHistory
    End If
    
    FormID = frm_scrFilterHist.GetActiveFormID()
    
    'If a matching screen was found, bind to it.
    If FormID > 0 Then
        frm_scrFilterHist.SetActiveScreen FormID
        frm_scrFilterHist.visible = True
    Else
        'No screens is loaded. Do not show the form.
        MsgBox "Please bring a Decipher Screen to the foreground first.", vbOKOnly
        
    End If
    
FilterHistoryShowExit:
    Exit Function
FilterHistoryShowError:
    If Err.Number = 2467 Then
        Set frm_scrFilterHist = Nothing
        DoCmd.Close acForm, "SCR_ScreensFilterHistory"
        Set frm_scrFilterHist = New Form_SCR_ScreensFilterHistory
        Resume FilterHistoryShowExit
    End If
       
End Function

'Determines if the form is currently loaded or not
Public Function IsFormLoaded(strFormName As String, _
        Optional LookupFormName As Form, _
        Optional LookupControl As Control) As Boolean
    Dim frm As Form
    Dim bFound As Boolean

    If Not (IsMissing(LookupFormName) Or IsNull(LookupFormName) _
        Or LookupFormName Is Nothing) Then
        
        If Not (IsMissing(LookupControl) Or IsNull(LookupControl) _
        Or LookupControl Is Nothing) Then
            On Error GoTo IsFormLoadedError:
            If LookupControl.ControlType = acSubform Then
                If LookupControl.Form.Name = strFormName Then
                    bFound = True
                End If
            End If
        Else
            Call SearchInForm(LookupFormName, strFormName, bFound)
        End If
        
    Else
        For Each frm In Forms
            If frm.Name = strFormName Then
                bFound = True
                Exit For
            Else
                Call SearchInForm(frm, strFormName, bFound)
                If bFound Then
                    Exit For
                End If
            End If
        Next
    End If
    
IsFormLoadedError:
    IsFormLoaded = bFound
End Function

'searches for a Form that is part of a Form as a sub form
Public Function SearchInForm(frm As Form, strDoc As String, bFound As Boolean)
On Error GoTo SearchInFormError
    Dim ctl As Control

    For Each ctl In frm.Controls
        If ctl.ControlType = acSubform Then
            If ctl.Form.Name = strDoc Then
            'If ctl.SourceObject = strDoc Then
                bFound = True
            Else
                Call SearchInForm(ctl.Form, strDoc, bFound)
            End If
SkipElement:
            If bFound Then
                Exit For
            End If
        End If
    Next
    
Exit Function

SearchInFormError:
    Resume SkipElement
End Function