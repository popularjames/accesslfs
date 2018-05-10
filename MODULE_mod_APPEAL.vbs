Option Compare Database
Option Explicit

Public Sub CreateAppeal(ByVal strICN As String, ByVal myCnlyClaimNum As String, ByVal strLocalPath As String, ByVal strPasswd As String, ByVal strMemo As String, ByRef ReturnPath As String, ByVal LetterOnly As String, ByVal strPayerName As String)
Dim strFileName, strSQL, strPackageID, strLocalFile, strPackageName, strNotePath, strNotePathLocal, shellCommand, MsgLeft, TargetPDF, tempLocalFile As String
Dim rst As DAO.RecordSet
Dim DocCount, RtnValue As Integer
Dim SourceF, TargetF As file
Dim fs As FileSystemObject
Dim Tif2PDF, Doc2PDF As Boolean
Dim ErrorText As String

Tif2PDF = False
Doc2PDF = False

If IsMissing(LetterOnly) Then LetterOnly = "DOCONLY"

Set fs = CreateObject("Scripting.FileSystemObject")
strPackageName = strICN & ".zip"
DoCmd.Hourglass True

'Dim MYADO As clsADO
'Dim rs As ADODB.Recordset
'Set MYADO = New clsADO
'
'MYADO.ConnectionString = GetConnectString("v_DATA_Database")
'MYADO.SQLstring = "SELECT DISTINCT Prov_Xref_Payer.Contractor AS PayerName FROM AUDITCLM_Hdr INNER JOIN Prov_Xref_Payer ON AUDITCLM_Hdr.PayerNum = Prov_Xref_Payer.PayerNumber WHERE AUDITCLM_Hdr.cnlyClaimNum = '" & myCnlyClaimNum & "'"
'Set rs = MYADO.OpenRecordSet("SELECT DISTINCT Prov_Xref_Payer.Contractor AS PayerName FROM AUDITCLM_Hdr INNER JOIN Prov_Xref_Payer ON AUDITCLM_Hdr.PayerNum = Prov_Xref_Payer.PayerNumber WHERE AUDITCLM_Hdr.cnlyClaimNum = '" & myCnlyClaimNum & "'")
'If rs.RecordCount <> 1 Then
'    MsgBox "Error fetching Payer", vbCritical
'    Exit Sub
'End If
'
'strPayername = rs.Fields(0).Value
If strPayerName = "PINNACLE/TRISPAN" Then strPayerName = "PINNACLE"
If left(strPayerName, 3) = "WPS" Or strPayerName = "PINNACLE" Then Tif2PDF = True
If left(strPayerName, 3) = "WPS" Then Doc2PDF = True
'strLocalPath = Replace(strLocalPath, "<payername>", strPayername)
strNotePathLocal = strLocalPath

If LetterOnly = "APPEAL" Then strNotePath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\APPEALS\Send Log\"
If LetterOnly = "LETTER" Then strNotePath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_SENT\Send Log\"
If LetterOnly = "DOCONLY" Then strNotePath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\CMS MAIL\Send Log\"

If Not (fs.FolderExists(strLocalPath)) Then fs.CreateFolder (strLocalPath)
strLocalPath = strLocalPath & Date$ & "\"
If Not (fs.FolderExists(strLocalPath)) Then
'Send email for a dated folder only once. Any more appeals coming/generated the same day get added to the same folder
    fs.CreateFolder (strLocalPath)
    MsgLeft = "Packages in encrypted zip format"
    LeaveNote strNotePath, "Readme_" & Date$ & ".txt", MsgLeft, strMemo
    LeaveNote strNotePathLocal, "Readme_" & Date$ & ".txt", "Appeal packages in encrypted zip format", strMemo
    Shell "explorer.exe " & strLocalPath, vbNormalFocus
    ReturnPath = strLocalPath
Else
    ReturnPath = ""
End If
'strLocalPath = strLocalPath & Replace(Replace(Time, ":", "-"), " ", "") & "\"
'If Not (fs.FolderExists(strLocalPath)) Then
'    fs.CreateFolder (strLocalPath)
'End If


'Set MYADO = Nothing
'Set rs = Nothing

'If Not (LetterOnly) Then
    'Get the max strPackageID for the given CnlyClaimNum
    strSQL = "Select Nz(MAX(PackageID),0)+1 As NextPID from APPEAL_Dtl where CnlyClaimNum = '" & myCnlyClaimNum & "'"
    Set rst = CurrentDb.OpenRecordSet(strSQL)
    rst.MoveFirst
    strPackageID = rst!NextPID
    rst.Close


    'Insert PackageName and ID in Appeal_Dtl
    strSQL = "insert into APPEAL_Dtl (CnlyClaimNum,AppealStatus,PackageName,PackageID,PackageDate,GeneratedBy) Values('" & _
        myCnlyClaimNum _
        & "','" & _
        "Package Generated" _
        & "','" & _
        strPackageName _
        & "'," & strPackageID & ",'" _
        & Now() & "','" & _
        Identity.UserName _
        & "')"
    CurrentDb.Execute (strSQL)
'End If
    
'Get the RefLinks for this claim

If LetterOnly = "LETTER" Then
    strSQL = "select * from v_AUDITCLM_References where CnlyClaimNum = '" & myCnlyClaimNum & "' and (refsubtype Like 'DEM%' OR refsubtype Like 'U6000%')"
Else
    strSQL = "select * from v_AUDITCLM_References where CnlyClaimNum = '" & myCnlyClaimNum & "' and refsubtype not in ('PROVMRINV','PROVPOST')"
End If

Set rst = CurrentDb.OpenRecordSet(strSQL)
If rst.recordCount > 0 Then
    rst.MoveFirst
    DocCount = 0
    While Not (rst.EOF)
        DocCount = DocCount + 1
        strFileName = rst.Fields("RefLink")
        strLocalFile = strLocalPath & strICN & Format(DocCount, "_000") & Right(strFileName, 4) 'Generate localfile name as ICN_000.xxx
        If fs.FileExists(strFileName) Then
            SetFileReadOnly (strFileName) 'Safe!
            Set SourceF = fs.GetFile(strFileName)
            SourceF.Copy strLocalFile
            Set TargetF = fs.GetFile(strLocalFile)
            TargetF.Attributes = TargetF.Attributes - ReadOnly
            
            If Tif2PDF And Right(strFileName, 4) = ".TIF" Then
            'Convert TIF to PDF
                tempLocalFile = strLocalFile
                strLocalFile = strLocalPath & strICN & Format(DocCount, "_000") & ".pdf"
                If ClmPkg_Tif2Pdf(tempLocalFile, strLocalFile, ErrorText) = False Then
                    MsgBox ErrorText
                End If
                TargetF.Delete
            End If
            
            If Doc2PDF And (Right(strFileName, 4) = ".DOC" Or Right(strFileName, 5) = ".DOCX") Then
            'Convert DOC to PDF
                tempLocalFile = strLocalFile
                strLocalFile = strLocalPath & strICN & Format(DocCount, "_000") & ".pdf"
                If ClmPkg_Doc2Pdf(tempLocalFile, strLocalFile, ErrorText) = False Then
                    MsgBox ErrorText
                End If
                TargetF.Delete
            End If
            
            
            strSQL = "insert into APPEAL_Package (PackageName,PackageID,SerialNo,SourceDoc,TargetDoc) VALUES('" _
                    & strPackageName & "'," & strPackageID & "," & DocCount & ",'" _
                    & strFileName _
                    & "','" & strLocalFile & "')"
            CurrentDb.Execute (strSQL)
        Else
            MsgBox "Cannot Find " & strFileName, vbInformation
        End If
        rst.MoveNext
    Wend
    'Winzip not registered, throws warnings. ShellWait from mod_Shell to make the call synchronous
    'shellCommand = "c:\program files (x86)\winzip\winzip32.exe -m -s" & strPasswd & " " & Chr(34) & strLocalPath & strPackageName & Chr(34) & " " & Chr(34) & strLocalPath & "*.*" & Chr(34)
    'RtnValue = ShellWait(shellCommand, vbMaximizedFocus)
    If LetterOnly <> "LETTER" Then
        RtnValue = ShellWait("c:\program files (x86)\winzip\wzzip.exe -m -s" & strPasswd & " -x*.zip " & """" & strLocalPath & strPackageName & """" & " " & """" & strLocalPath & "*.*" & """", vbMinimizedNoFocus)
    End If
End If

End Sub
Public Sub LeaveNote(ByVal strPath As String, ByVal strFileName, ByVal strnote As String, ByVal strMemo As String)
    Dim fs As FileSystemObject
    Dim a As TextStream
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPath & strFileName) Then
        Set a = fs.OpenTextFile(strPath & strFileName, ForAppending, TristateFalse)
    Else
        Set a = fs.CreateTextFile(strPath & strFileName, True)
    End If
    a.WriteLine (strnote)
    a.WriteLine
    a.WriteLine ("***Other Info (Shipping/Receiver/Original Request)****")
    a.WriteLine (strMemo)
    a.Close
End Sub
Public Function RandomString(ByVal mask As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim options As String
    Dim char As String
    
    ' initialize result with proper lenght
    RandomString = mask
    
    For i = 1 To Len(mask)
        ' get the character
        char = Mid$(mask, i, 1)
        Select Case char
            Case "?"
                char = Chr$(1 + Rnd * 127)
                options = ""
            Case "#"
                options = "0123456789"
            Case "A"
                options = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
            Case "N"
                options = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0" _
                    & "123456789"
            Case "H"
                options = "0123456789ABCDEF"
            Case Else
                ' don't modify the character
                options = ""
        End Select
    
        ' select a random char in the option string
        If Len(options) Then
            ' select a random char
            ' note that we add an extra char, in case RND returns 1
            char = Mid$(options & Right$(options, 1), 1 + Int(Rnd * Len(options) _
                ), 1)
        End If
        
        ' insert the character in result string
        Mid(RandomString, i, 1) = char
    Next

End Function






Public Sub SendsqlMail(ByVal Subject As String, ByVal ToList As String, ByVal CCList As String, ByVal BCCList As String, ByVal Body As String)
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")

    MyAdo.sqlString = "EXEC cnly.mail.sendSqlMail '" & Subject & "','" & ToList & "','" & CCList & "','" & BCCList & "','" & Body & "',NULL"
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
End Sub