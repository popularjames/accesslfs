Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrDestinationDir As String = "M:\My Documents\Data\ITR\Out"  'Out folder for ITR files.
'=============================================
' ID:          Form_frm_ADHOC_ITR
' Author:      Barbara Dyroff
' Create Date: 2010-12-01
' Description:
'      Copy the Demand Letters for an Intent to Refer Request.
'
' Modification History:
'
'
' =============================================

Private Sub cmdITRCopy_Click()

    On Error GoTo ErrorHandler
    
    Dim strSourceDirFile As String
    Dim strSourceFile As String
    Dim strDestinationDirFile As String
    Dim db As DAO.Database
    Dim rsITRFiles As DAO.RecordSet
    Dim strSQL As String
    Dim strRetVal As String
    Dim intRetVal As Integer
    Dim strErrMsg As String
    Dim ErrorText As String
    Dim strDestinationDirFilePDF As String

    
    MsgBox "Starting Intent To Refer (ITR) File Copy", vbInformation, "Intent To Refer (ITR) File Copy"
    
    'Delete current files in ITR folder.
    strRetVal = Dir$(CstrDestinationDir & "\*.*")
    If strRetVal <> "" Then
        intRetVal = MsgBox("Delete existing files?  ", vbYesNo, "Intent To Refer (ITR) File Copy")
        If intRetVal = vbYes Then    ' User chose Yes.
            Kill CstrDestinationDir & "\*.*"
        End If
    End If

    strSQL = "Select * From  ITR_CIGNA_Load_Files"
    Set db = CurrentDb
    Set rsITRFiles = db.OpenRecordSet(strSQL, dbOpenSnapshot, dbForwardOnly)
    With rsITRFiles
        If .EOF And .BOF Then
            .Close
            GoTo ExitNow
        End If
        Do Until .EOF
            strSourceDirFile = .Fields("DemandLtrDirFileNm")
            strSourceFile = Right(strSourceDirFile, 30)
            strDestinationDirFile = CstrDestinationDir + "\" + strSourceFile  ' Define target file name.
            'FileCopy strSourceDirFile, strDestinationDirFile   ' Copy source to target.
            'Copy to pdf
            strDestinationDirFilePDF = CstrDestinationDir + "\" + left(strSourceFile, 27) + "pdf"
            If ClmPkg_Doc2Pdf(strSourceDirFile, strDestinationDirFilePDF, ErrorText) = False Then
                MsgBox ErrorText
            End If
            strSourceDirFile = ""
            strSourceFile = ""
            .MoveNext
        Loop
    End With
    
    MsgBox "Completed copying Intent to Refer (ITR) files.  ", vbInformation, "Intent To Refer (ITR) File Copy"

ExitNow:
    On Error Resume Next
    Set rsITRFiles = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    strErrMsg = "Error copying Intent to Refer (ITR) files:  " & Err.Description & vbCrLf & vbCrLf
    MsgBox strErrMsg, vbCritical, "Intent To Refer (ITR) File Copy Error"
    Resume ExitNow

End Sub
