Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDialogs
' DateTime  : 1/20/2006 15:54
' Author    : joseph.casella
' Purpose   : Provides an interface to the File Dialog Routines
' Dependencies : Uses the existing api call in screens
' Last Mod     : 5/23/06 Added CleanFilename

'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private fso As Object

Public Enum DialogFilterType

    Mdbf = 1    ' "AccessFile (*.mdb)|*.mdb|All Files (*.*)|*.*"  'As String
    xlsf = 2    ' "ExcelFile (*.xls)|*.xls|All Files (*.*)|*.*"
    txtf = 3    ' "TextFile (*.txt)|*.txt|All Files (*.*)|*.*"
    RacXcl = 4  ' Rac Exclusion file
    RacSup = 5  ' Rac Suppression file
    RspRac = 6  ' Rac Input Return file
    RspSta = 7  ' Rac Status Return File
    docf = 8    ' "WordDocument (*.doc)
   
    
End Enum

Public Enum CleanType

    CleanName = 1
    CleanPath = 2

End Enum

Private mvBaseFileName As String
Private mvOpenFileName As String



Private Function GetFilter(FilterType) As String

    Select Case FilterType
    Case Is = 1
        GetFilter = "Access Database (*.mdb)" & Chr(0) & "*.mdb" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 2
        GetFilter = "Excel File (*.xlsx)" & Chr(0) & "*.xlsx" & Chr(0) & "Excel (97 - 2003) (*.xls)" & Chr(0) & "*.xls" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 3
        GetFilter = "Text File (*.txt)" & Chr(0) & "*.txt" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 4
        GetFilter = "Exclusion File (*.xcl)" & Chr(0) & "*.xcl" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 5
        GetFilter = "Suppression File (*.sup)" & Chr(0) & "*.sup" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 6
        GetFilter = "Input Return File (*.rac)" & Chr(0) & "*.rac" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 7
        GetFilter = "Status Return File (*.rsp.sta)" & Chr(0) & "*.rsp.Sta" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Case Is = 8
        GetFilter = "Word Document (*.docx)" & Chr(0) & "*.docx" & Chr(0) & "Word Document (97 - 2003) (*.doc)" & Chr(0) & "*.doc" & Chr(0)
    End Select

End Function

Public Function DeleteFile(ByVal FilePath As String)

    fso.DeleteFile FilePath

exitHere:
    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function

Public Function FileExists(ByVal FilePath As String) As Boolean

    FileExists = fso.FileExists(FilePath)

exitHere:
    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function


Public Function FolderExists(FilePath As String) As Boolean

    FolderExists = fso.FolderExists(FilePath)

End Function



Public Function SavePath(Folder As String, filter As DialogFilterType, Name As String) As String
 
 SavePath = FileDialog(1, "Save As", Forms(0).hwnd, Folder, GetFilter(filter), Name)
  
  
exitHere:


    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function

Public Function OpenPath(Folder As String, filter As DialogFilterType, Optional FileName As String, Optional Title As String = "Open") As String


    OpenPath = FileDialog(0, Title, Forms(0).hwnd, Folder, GetFilter(filter), FileName)

    mvBaseFileName = fso.GetBaseName(OpenPath)

exitHere:

    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function


Private Sub Class_Initialize()

    Set fso = CreateObject("Scripting.FilesystemObject")

exitHere:
    Exit Sub

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Sub


Private Sub Class_Terminate()

    Set fso = Nothing

exitHere:
    Exit Sub

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Sub

Public Function RenameFile()


End Function


Public Property Get BaseName() As String

    BaseName = mvBaseFileName

End Property

Public Function CleanFileName(Name As String, CleanWhat As CleanType) As String
'* 5/23/06 From Dave Brady

'Takes Name and "cleans" it up to be a legal win32 file name.
'Trim leading/trailing spaces, leading/trailing ".", strip "forbidden chars:
'LongForbiddenChars : set of Char = ['<', '>', '|', ''', '', '/',':','*','?'];

    Dim RegExp As Object
    Dim StrTmp As String
    Dim strDirectory As String
    Dim strName As String

    Set RegExp = CreateObject("VBScript.RegExp")

    RegExp.Pattern = "^[. ]+|[""\\<>|',/:*?\r\n\[\]\t]|[. ]+$"
    RegExp.IgnoreCase = True
    RegExp.Global = True


    If CleanWhat = CleanName Then

        StrTmp = RegExp.Replace(Name, "")

    ElseIf CleanWhat = CleanPath Then

        strName = Right$(Name, Len(Name) - InStrRev(Name, "\"))
        strDirectory = left$(Name, InStrRev(Name, "\"))

        StrTmp = strDirectory & RegExp.Replace(strName, "")

    Else
        MsgBox "Invalid CleanFileName type"

    End If

    CleanFileName = StrTmp

    Set RegExp = Nothing

End Function