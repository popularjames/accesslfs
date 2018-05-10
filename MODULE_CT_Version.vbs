Option Compare Database
Option Explicit

'SA 10/23/2012 - Added CT_GetAppVersion, CT_GetAppVersionMaj, CT_GetAppVersionMin, CT_GetAppVersionRev, CT_GetAppVersionPart

Private strVersionTemplate As String

Public Function VersionTemplate() As String
On Error GoTo ErrorHappened
    'Return version of template as string
    If LenB(strVersionTemplate) = 0 Then
        strVersionTemplate = CurrentDb.TableDefs("CT_AppStartupSeq").Properties("Description")
    End If
    VersionTemplate = strVersionTemplate
ExitNow:

Exit Function
ErrorHappened:
    VersionTemplate = "ERROR"
    Resume ExitNow
End Function

Public Function CT_GetAppVersion(ByVal AppName As String) As String
'Get the version of an app
'SA 10/23/2012 - Added to module
On Error GoTo ErrorHappened
    Dim Result As String
    Result = Nz(DLookup("LocalVersion", "CT_InstalledApps", "ProductName='" & AppName & "'"), vbNullString)
ExitNow:
On Error Resume Next
    CT_GetAppVersion = Result
Exit Function
ErrorHappened:
    Result = vbNullString
    Resume ExitNow
End Function

Public Function CT_GetAppVersionMaj(ByVal AppName As String) As Integer
'Get the major version of an app by name
    CT_GetAppVersionMaj = CT_GetAppVersionPart(AppName, 0)
End Function

Public Function CT_GetAppVersionMin(ByVal AppName As String) As Integer
'Get the minor version of an app by name
    CT_GetAppVersionMin = CT_GetAppVersionPart(AppName, 1)
End Function

Public Function CT_GetAppVersionRev(ByVal AppName As String) As Integer
'Get the revision version of an app by name
    CT_GetAppVersionRev = CT_GetAppVersionPart(AppName, 2)
End Function

Private Function CT_GetAppVersionPart(ByVal AppName As String, ByVal position As Integer) As Integer
'Get the version of an app based on section
'SA 10/23/2012 - Added to module
On Error GoTo ErrorHappened
    Dim Result As Integer
    Dim Version() As String
    
    Version = Split(CT_GetAppVersion(AppName), ".")
    Result = Version(position)
    
ExitNow:
On Error Resume Next
    CT_GetAppVersionPart = Result
Exit Function
ErrorHappened:
    Result = -1
    Resume ExitNow
End Function