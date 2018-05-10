Option Compare Database
Option Explicit

Public algCnlyClaimNum As String
Public algHearDate As String
Public algAppealNum As String
Public algJudgeName As String
Public algPackageName As String
Public algAppellant As String
Public algParticipant As String
Public algParticipant2 As String
Public algPartyStatus As String
Public algPartyWitn As String
Public algPartyInfo As String
Public algPartyDisc As String
Public algTestimonyType As String
Public algNOIResponseDate As String

Public algHearingPhone As String
Public algHearingPasscode As String
Public algEditNote As String
Public testme As String
Public fso As Object

Public AmendedFlag As Boolean

Public PackageLocation As String
Public appeal As String

Public Const algID As String = "ALJ"
Public Const PARTY As String = "PARTY"
Public Const NON_PARTY As String = "NON-PARTY"

Public Const Participant As String = "PARTICIPANT"
Public Const Judge As String = "JUDGE"
Public Const Law_Firm As String = "LAWFIRM"
Public Const Provider As String = "PROVIDER"

Public Const NOI_STATUS_SENT = "NOI Sent"
Public Const NOI_STATUS_AMENDED = "Amended NOI Sent"
Public Const NOI_STATUS_WITHDRAW_CLAIM = "Suppressed/excluded claim"

Public Const MAX_DOC_TYPE = 3
Public Const APP_DOC_TYPE = 2
Public Const JUDGE_DOC_TYPE = 1

Public Const YES As String = "Yes"
Public Const NO As String = "No"
Public folderName As String

Private db As DAO.Database
Private rst As DAO.RecordSet
Private SQL As String

Private myCode_ADO As New clsADO
Private rs As ADODB.RecordSet
'2014-06-18 VS: The module contains functions used by ALJ forms and modules
'2014-07-03 VS: Added Found ALJ Claims function
'2014-07-16 VS: Added Edit ALJ Package function

Public Function Find_ALJ_Package()

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Find_ALJ_Package"
                myCode_ADO.Parameters("@pJudgeName") = algJudgeName
                myCode_ADO.Parameters("@pHearingDate") = algHearDate
                
                Set rs = myCode_ADO.ExecuteRS
                'Set Find_ALJ_Package = rs
                Set Forms("frm_Existing_ALJ_Packages").lstExistPkgs.RecordSet = rs

Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Function

Err_handler:
        MsgBox Err.Description
           
End Function

Public Function Update_Contact(Id As Integer, ContactType As String)

    Dim ErrorReturned As String
    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Update_ALJ_Contact"
                myCode_ADO.Parameters("@pContactType") = ContactType
                myCode_ADO.Parameters("@pId") = Id
                
                Set rs = myCode_ADO.ExecuteRS
                 
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned = "" Then
                    MsgBox ("Contact had been updated!")
                Else
                    MsgBox ErrorReturned, vbExclamation
                End If
           
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Function
    
Err_handler:
        MsgBox Err.Description
        GoTo Exit_Sub
           
End Function

Public Function Refresh_Contacts()

    Dim ErrorReturned As String
    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Refresh_ALJ_Contacts"
                
                Set rs = myCode_ADO.ExecuteRS
           
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Function
    
Err_handler:
        MsgBox Err.Description
        GoTo Exit_Sub
           
End Function

Public Function Found_ALJ_Claims(pkgName As String) As Boolean

    Dim FoundClaims As String
    Dim ErrorReturned As String
    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_ALJ_Claim_List"
                myCode_ADO.Parameters("@pALJPackageId") = pkgName
                Set rs = myCode_ADO.ExecuteRS
                
                myCode_ADO.sqlString = "usp_Get_ALJ_Claim_List"
                myCode_ADO.Parameters("@pALJPackageId") = pkgName
                Set rs = myCode_ADO.ExecuteRS
                            
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Function
                Else
                        FoundClaims = Nz(myCode_ADO.Parameters("@pFoundClaims").Value, "")
                        If FoundClaims = "No" Then
                                
                                 MsgBox ("The system can not find any claims to be added to this package!")
                                 Found_ALJ_Claims = False
                            
                        Else
                            
                                    Found_ALJ_Claims = True
                            
                                    DoCmd.OpenForm "frm_Existing_ALJ_ClaimsToAdd", , , , , acWindowNormal, Null
                                    Set Forms("frm_Existing_ALJ_ClaimsToAdd").lstExistClaimsToAdd.RecordSet = rs
                                    
                        End If
                              
                End If
            
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Function

Err_handler:
        MsgBox Err.Description
           
End Function


Public Sub Create_New_ALJ_Package()
          
    Dim ErrorReturned As String
    Dim PackageID As String

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Create_New_ALJ_Package"
                myCode_ADO.Parameters("@pJudgeName") = Trim(algJudgeName)
                myCode_ADO.Parameters("@pHearingDate") = algHearDate
                myCode_ADO.Parameters("@pAppellantName") = algAppellant
                myCode_ADO.Parameters("@pConnollyParticipant") = algParticipant
                myCode_ADO.Parameters("@pALJAppealNumber") = Trim(algAppealNum)
                myCode_ADO.Parameters("@pCnlyClaimNum") = algCnlyClaimNum
                myCode_ADO.Parameters("@pTestimonyType") = algTestimonyType
                
                If Nz(algPartyStatus, "") <> "" Then
                    myCode_ADO.Parameters("@pPartyStatus") = algPartyStatus
                End If
                
                If Nz(algHearingPhone, "") <> "" Then
                    myCode_ADO.Parameters("@pHearingPhone") = algHearingPhone
                    myCode_ADO.Parameters("@pHearingPasscode") = algHearingPasscode
                End If
                
                Set rs = myCode_ADO.ExecuteRS
                
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Sub
                Else
                    PackageID = Nz(myCode_ADO.Parameters("@pALJPackageId").Value, "")
                    If PackageID <> "" Then
                        MsgBox ("ALJ Hearing Package  " + PackageID + "  had been created.")
                        algPackageName = PackageID
                        
                        If AmendedFlag = True Then
                           Call Delete_Claim(PackageID, True)
                        End If
                        
                        CreatePackageFolder
                        CopyALJNoticeFile
                        
                    End If
                End If
                
                Find_ALJ_Package
                
                               
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        MsgBox Err.Description
           
End Sub


Public Sub Add_To_Existing_ALJ_Package(pkgName As String)
          
    Dim ErrorReturned As String
    Dim PackageID As String
    Dim FaxDate As Date
    Dim daysPassed As Integer

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Add_Details_To_ALJ_Package"
                myCode_ADO.Parameters("@pALJPackageId") = pkgName
                myCode_ADO.Parameters("@pALJAppealNumber") = algAppealNum
                myCode_ADO.Parameters("@pCnlyClaimNum") = algCnlyClaimNum
                myCode_ADO.Parameters("@pHearingDateTime") = algHearDate
                Set rs = myCode_ADO.ExecuteRS
                
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Sub
                Else
                    PackageID = Nz(myCode_ADO.Parameters("@pALJPackageId").Value, "")
                    If PackageID <> "" Then
                        MsgBox ("ALJ Hearing Package  " + PackageID + "  had been updated.")
                        algPackageName = PackageID
                        FaxDate = Nz(DLookup("[FaxDate]", "v_ALJ_732_Form", "HearingPackageName= '" & Replace(algPackageName, "'", "''") & "'"), #1/1/1900#)
                        daysPassed = DateDiff("d", FaxDate, Now())
      
                        CopyALJNoticeFile
                        
                        If AmendedFlag = True Then
                        Call Delete_Claim(pkgName, False)
                        End If
                        
                    End If
                End If
                                    
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        
        If Err.Number <> 58 Then
        MsgBox Err.Description
        End If
           
End Sub

Public Sub Delete_Claim(notPkgName As String, renameForms As Boolean)
          
    Dim ErrorReturned As String
    Dim oldPackageID As String
    Dim claimsLeft As Integer
    Dim FaxDate As Date
    Dim daysPassed As Integer

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Delete_Claim_from_ALJ_Package"
                myCode_ADO.Parameters("@pNotALJPackageName") = notPkgName
                myCode_ADO.Parameters("@pCnlyClaimNum") = algCnlyClaimNum

                Set rs = myCode_ADO.ExecuteRS
                
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Sub
                Else
                    claimsLeft = Nz(myCode_ADO.Parameters("@pClaimsLeft").Value, 0)
                    claimsLeft = claimsLeft + 1

                    oldPackageID = Nz(myCode_ADO.Parameters("@pDeletedALJPackageId").Value, "")
                    If oldPackageID <> "" Then
                        MsgBox ("Claim with Appeal Number " + algAppealNum + ", CnlyClaimNum " + algCnlyClaimNum + " had been deleted from " + oldPackageID + " package.")
                    End If
                End If
                
                If renameForms = True Then
                    Call Rename_Amended_Forms(oldPackageID, claimsLeft)
                End If
                                    
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        
        If Err.Number <> 58 Then
        MsgBox Err.Description
        End If
           
End Sub


Public Sub Edit_ALJ_Package(PkgToEd As String)
        
    Dim ErrorReturned As String
    Dim PackageID As String

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Edit_ALJ_Package"
                myCode_ADO.Parameters("@pALJPackageName") = PkgToEd
                myCode_ADO.Parameters("@pJudgeName") = algJudgeName
                myCode_ADO.Parameters("@pHearingDate") = Nz(algHearDate, "")
                myCode_ADO.Parameters("@pAppellantName") = algAppellant
                myCode_ADO.Parameters("@pConnollyParticipant") = algParticipant
                myCode_ADO.Parameters("@pConnollyParticipant2") = algParticipant2
                myCode_ADO.Parameters("@pPartyStatus") = algPartyStatus
                myCode_ADO.Parameters("@pEditNote") = algEditNote
                
                If Nz(algHearingPhone, "") <> "" Then
                    myCode_ADO.Parameters("@pHearingPhone") = algHearingPhone
                    myCode_ADO.Parameters("@pHearingPasscode") = algHearingPasscode
                End If
                
                If algPartyStatus = PARTY Then
                    myCode_ADO.Parameters("@pPartyWitn") = algPartyWitn
                    myCode_ADO.Parameters("@pPartyInfo") = algPartyInfo
                    myCode_ADO.Parameters("@pPartyDisc") = algPartyDisc
                End If
                
                If Nz(algNOIResponseDate, "") <> "" Then
                    myCode_ADO.Parameters("@pNOIResponseDate732") = Nz(algNOIResponseDate, "")
                End If
                
                Set rs = myCode_ADO.ExecuteRS
                
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Sub
                Else
                    PackageID = Nz(myCode_ADO.Parameters("@pALJPackageNameNew").Value, "")
                    If PackageID <> "" Then
                        MsgBox ("This ALJ Hearing Package had been renamed to  " + PackageID + ". Please save the documents in the new folder.")
                        algPackageName = PackageID
                        
                        CreatePackageFolder
                        CopyALJNoticeFile
                
                        Find_ALJ_Package
                                      
                    End If
                End If
                               
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        MsgBox Err.Description
           
End Sub


Public Sub Delete_ALJ_Package(pkgName As String)
          
    Dim ErrorReturned As String
    Dim PackageLocation As String
    Dim FaxDate As Date
    Dim daysPassed As Integer
    Dim objFSO As New FileSystemObject

    On Error GoTo Err_handler

    PackageLocation = GetPackageLocation()
    PackageLocation = left(PackageLocation, Len(PackageLocation) - 1)
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Delete_ALJ_Package"
                myCode_ADO.Parameters("@pALJPackageId") = pkgName
                Set rs = myCode_ADO.ExecuteRS
                
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Sub
                Else
                    MsgBox (pkgName + " had been deleted.")
                    Call Find_ALJ_Package
                    
                    'DeleteFolder here
                    Set objFSO = CreateObject("Scripting.FileSystemObject")
                    objFSO.DeleteFolder PackageLocation

                End If
                                    
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        MsgBox Err.Description
           
End Sub


Public Sub CreatePackageFolder(Optional packageNumber As Integer = 1)

Dim currentCnlyClaimNum As String
Dim folderLocation As String
folderLocation = GetPackageLocation()
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderLocation) Then
        fso.CreateFolder (folderLocation)
        
    Else
           
       'Need to create package 2 and copy all ALJ Notices
       If packageNumber = 2 Then
             
          folderLocation = left(folderLocation, Len(folderLocation) - 1) & "_Package2\"
          
          If Not fso.FolderExists(folderLocation) Then
                fso.CreateFolder (folderLocation)

                SQL = "Select CnlyClaimNum from v_ALJ_732_Form_Attach where HearingPackageName = '" & Replace(algPackageName, "'", "''") & "'"

                Set db = CurrentDb
                Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbForwardOnly)
    
                With rst
                    If .EOF And .BOF Then
                    .Close
                    GoTo ExitNow
                    End If
                
                Do Until .EOF
                currentCnlyClaimNum = .Fields("CnlyClaimNum")
                CopyALJNoticeFile currentCnlyClaimNum, folderLocation
                .MoveNext
                Loop
                End With
                
                MsgBox (folderLocation & "folder had been created. Please save fax cover pages and 732 form under this new folder!")
                
            End If
            
        End If
        
        End If
        
        SQL = "UPDATE cms_auditors_claims.dbo.APPEAL_Hearing_ALJ_732 SET FolderName = '" & Replace(folderLocation, "'", "''") & "' WHERE HearingPackageName = '" & Replace(algPackageName, "'", "''") & "'"
        
        Dim myPortalADO As clsADO
        Set myPortalADO = New clsADO
        myPortalADO.ConnectionString = GetConnectString("v_DATA_Database")
        myPortalADO.sqlString = SQL
        myPortalADO.SQLTextType = sqltext
        myPortalADO.Execute
        

ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    Set fso = Nothing
    Exit Sub
            
End Sub

Public Function GetPackageLocation() As String

    Dim appellant As String
    Dim Judge As String
    Dim HearingDate As String
    
    PackageLocation = Nz(DLookup("[folderName]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "")
    
    If PackageLocation = "" Then
    
        PackageLocation = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & algID & "'")
        'appeal = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & algID & "'")
        'Judge = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & algID & "'")
        
        Judge = left(algPackageName, InStr(algPackageName, "20") - 2)
        Judge = Replace(Judge, "''", "'")
        HearingDate = Replace(Right(algPackageName, Len(algPackageName) - InStr(algPackageName, "20") + 1), ":", "_")
     
        appellant = Nz(DLookup("[AppellantFolderName]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(algPackageName, "'", "''") & "'"), "")
        appellant = (Replace(appellant, " ", "_"))
        appellant = (Replace(appellant, ",", ""))
        appellant = (Replace(appellant, ".", ""))
        appellant = (Replace(appellant, "/", "_"))
    
        folderName = HearingDate + "_" + Judge + "_" + appellant
        
        PackageLocation = PackageLocation & folderName & "\"
    
    End If
    
    GetPackageLocation = PackageLocation
    

End Function

Public Sub CopyALJNoticeFile(Optional ClaimNum As String = "", Optional Folder As String = "")

Dim noticeFileSource As String
Dim destinationFile As String
Dim FileName As String

If fso Is Nothing Then
Set fso = CreateObject("Scripting.FileSystemObject")
End If

If ClaimNum = "" Then
  ClaimNum = algCnlyClaimNum
End If

noticeFileSource = Nz(DLookup("[ALJNoticeFileLink]", "v_ALJ_732_Form_Attach", "HearingPackageName= '" & Replace(algPackageName, "'", "''") & "' AND CnlyClaimNum='" & ClaimNum & "'"))

If noticeFileSource <> "" Then

    FileName = fso.GetFileName(noticeFileSource)
    
    If Folder = "" Then
        destinationFile = GetPackageLocation + FileName
    Else
        destinationFile = Folder + FileName
    End If
   
    If Not fso.FileExists(destinationFile) Then
        fso.CopyFile noticeFileSource, destinationFile, False
    End If

End If

End Sub

Public Function WriteToFaxTable(algPackageName As String, FileName As String, DocType As String)

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Write_To_ALJ_Fax_Table"
                myCode_ADO.Parameters("@pPackageName") = algPackageName
                myCode_ADO.Parameters("@pDocName") = FileName
                myCode_ADO.Parameters("@pDocType") = DocType
                
                Set rs = myCode_ADO.ExecuteRS

Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Function

Err_handler:
        MsgBox Err.Description
           
End Function


Public Sub SavePDFAsOtherFormat(PDFPath As String, FileExtension As String)
   
    'Saves a PDF file as another format using Adobe Professional.
   
    'By Christos Samaras
    'http://www.myengineeringworld.net
   
    'In order to use the macro you must enable the Acrobat library from VBA editor:
    'Go to Tools -> References -> Adobe Acrobat xx.0 Type Library, where xx depends
    'on your Acrobat Professional version (i.e. 9.0 or 10.0) you have installed to your PC.
   
    'Alternatively you can find it Tools -> References -> Browse and check for the path
    'C:\Program Files\Adobe\Acrobat xx.0\Acrobat\acrobat.tlb
    'where xx is your Acrobat version (i.e. 9.0 or 10.0 etc.).
   
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
   
    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        MsgBox "Cannot find the PDF file!" & vbCrLf & "Check the PDF path and retry.", _
                vbCritical, "File Path Error"
        Exit Sub
    End If
   
    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        MsgBox "The input file is not a PDF file!", vbCritical, "File Type Error"
        Exit Sub
    End If
   
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
   
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
       
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
   
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
   
    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select
    
    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        
        'Format is correct and no errors.
        
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(FileExtension) <> "xls" Then
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", "." & LCase(FileExtension))
        Else
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", ".xml")
        End If
        
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
        
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
        
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
        
        'Inform the user that conversion was successfully.
        MsgBox "The PDf file:" & vbNewLine & PDFPath & vbNewLine & vbNewLine & _
        "Was saved as: " & vbNewLine & NewFilePath, vbInformation, "Conversion finished successfully"
         
    Else
       
        'Something went wrong, so close the PDF file and the application.
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
       
        'Inform the user that something went wrong.
        MsgBox "Something went wrong!" & vbNewLine & "The conversion of the following PDF file FAILED:" & _
        vbNewLine & PDFPath, vbInformation, "Conversion failed"

    End If
       
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
       
End Sub

Public Sub Clear_Exception_ALJ(ClaimNum As String)

    Dim expType As String
    Dim screen As String
    screen = "Appeal"
    expType = "ALJ"

    On Error GoTo Err_handler
    If myCode_ADO Is Nothing Then Set myCode_ADO = New clsADO
    
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
      
      myCode_ADO.sqlString = "select * from v_QUEUE_Exception_Info where CnlyClaimNum = '" & ClaimNum & "'" & " and ExceptionType Like '" & expType & "%'"
      Set rs = myCode_ADO.OpenRecordSet

      If rs.BOF And rs.EOF Then
            'exception does not exists
            Exit Sub
      Else
            rs.MoveFirst
            expType = rs.Fields("ExceptionType").Value
            
            If (expType = "ALJ" Or expType = "ALJ2") Then
            
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_QUEUE_Exception_Clear"
                myCode_ADO.Parameters("@pCnlyClaimNum") = ClaimNum
                myCode_ADO.Parameters("@pExceptionType") = expType
                myCode_ADO.Parameters("@pScreen") = screen
                Set rs = myCode_ADO.ExecuteRS
                    
                MsgBox "ALJ Exception had been cleared for claim " & ClaimNum & "!", vbInformation
                
            End If
    
      End If

Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        MsgBox Err.Description
End Sub

Public Sub Rename_Amended_Forms(AmendPackagename As String, OrigClmCnt As Integer)

    Dim MaximForm As String
    Dim JudgeForm As String
    Dim AppelForm As String
    Dim MaximFormRen As String
    Dim JudgeFormRen As String
    Dim AppelFormRen As String
    
    Dim OldPackageLocation As String
    
    On Error GoTo Err_handler
    
    OldPackageLocation = Nz(DLookup("[folderName]", "v_ALJ_732_Form", "[HearingPackageName] ='" & Replace(AmendPackagename, "'", "''") & "'"), "")
    
    MaximForm = OldPackageLocation & Replace(AmendPackagename, ":", "_") & "_MAXIMUS_Cover_And_732.pdf"
    JudgeForm = OldPackageLocation & Replace(AmendPackagename, ":", "_") & "_Judge_Cover_And_732.pdf"
    AppelForm = OldPackageLocation & Replace(AmendPackagename, ":", "_") & "_Appellant_Cover_And_732.pdf"
    
    MaximFormRen = OldPackageLocation & Replace(AmendPackagename, ":", "_") & "_MAXIMUS_Cover_And_732" & "_Original" & OrigClmCnt & ".pdf"
    JudgeFormRen = OldPackageLocation & Replace(AmendPackagename, ":", "_") & "_Judge_Cover_And_732" & "_Original" & OrigClmCnt & ".pdf"
    AppelFormRen = OldPackageLocation & Replace(AmendPackagename, ":", "_") & "_Appellant_Cover_And_732" & "_Original" & OrigClmCnt & ".pdf"
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    If fso.FileExists(MaximForm) Then
        fso.MoveFile MaximForm, MaximFormRen
    End If
    
    If fso.FileExists(JudgeForm) Then
        fso.MoveFile JudgeForm, JudgeFormRen
    End If
        
    If fso.FileExists(AppelForm) Then
        fso.MoveFile AppelForm, AppelFormRen
    End If

Exit_Sub:
    Set fso = Nothing
    Exit Sub

Err_handler:
        MsgBox Err.Description
        GoTo Exit_Sub
End Sub

Public Sub sendALJFax(Path As String, FaxNum As String, PackageName As String, DocType As String)

Dim sFaxNum As String
Dim sEmail As String
Dim outTifFile As String
Dim sID As String
Dim docIdType As Integer
Dim Status As String

Dim ofax As New ClsCnlyFax
Dim ofaxImage As New ClsCnlyFaxImage

Status = "In Progress"
ofaxImage.Sender.Host = "cmsfax.smtp.ccaintranet.net"
ofaxImage.Sender.Suffix = "@cmsconnollyfax.com"

outTifFile = PdfToTiff(Path)
'Call SavePDFAsOtherFormat(Path, "tiff")
        
docIdType = Nz(DLookup("DocTypeId", "ALJ_Fax_Doc_Type", "DocTypeName = " & "'" & DocType & "'"), "")

sEmail = "cms@cms.fax.ccaintranet.net"
'sFaxNum = "1-203-202-6745" 'Make sure to replace this with parameter once you are done testing!
sFaxNum = "1" & FaxNum
sID = "ALJ_" & PackageName & "_" & DocType & "_" & Now & "_" & sFaxNum
               sID = Replace(sID, "/", "_")
               sID = Replace(sID, ":", "_")
               sID = Replace(sID, " ", "_")
               sID = Replace(sID, "-", "_")
               sID = Replace(sID, "'", "")
               
        Call ofaxImage.Sender.SendFax(outTifFile, sFaxNum, sEmail, sID)
        Call UpdateALJFaxStatus(PackageName, Status, Status, sID, docIdType)
        MsgBox ("The document had been faxed.")

End Sub


Sub UpdateALJFaxStatus(sPackageName As String, sStatus As String, sDescr As String, sEmlFileId As String, DocType As Integer)

  Dim oDb As DAO.Database
  Dim sqlUpdate As String
  Set oDb = CurrentDb
  Dim currDate As String
  Dim UserID As Integer

    'myCodeADO.sqlString = "Update [ALJ_Fax_Queue] set [ALJ_Fax_Queue].Status='" & sStatus & "', [ALJ_Fax_Queue].Description='" & sDescription & "'"
     '          myCodeADO.sqlString = myCodeADO.sqlString & ",[ALJ_Fax_Queue].EmlFileId='" & "1" & "' Where [ALJ_Fax_Queue].PackageName= '" & sPackageName & "' and [ALJ_Fax_Queue].DocType= " & iDocumentID
    
    'Works:
    'sqlUpdate = "Update cms_auditors_claims.dbo.ALJ_Fax_Queue set Status='" & sStatus & "', Description='" & sDescription & "' Where PackageName= '" & sPackageName & "'"
    
         'FaxDate = Nz(DLookup("[FaxDate]", "v_ALJ_732_Form", "HearingPackageName= '" & Replace(algPackageName, "'", "''") & "'"), #1/1/1900#)
          '              daysPassed = DateDiff("d", FaxDate, Now())
    currDate = Format(Now(), "YYYY-MM-DD HH:MM:SS")
    
    UserID = Nz(DLookup("Id", "APPEAL_XREF_ALJ_Hearing_Connolly_CS_Contacts", "FullName = " & "'" & GetUserName() & "'"), "")
    sqlUpdate = "Update cms_auditors_claims.dbo.ALJ_Fax_Queue set Status='" & sStatus & "', Description='" & sDescr & "', UpdateDate = '" & currDate & "', UpdateUserId = '" & UserID
    sqlUpdate = sqlUpdate & "', EmlFileId = '" & sEmlFileId
    sqlUpdate = sqlUpdate & "' Where PackageName= '" & Replace(sPackageName, "'", "''") & "' and DocType = " & DocType

    Dim myPortalADO As clsADO
        Set myPortalADO = New clsADO
        myPortalADO.ConnectionString = GetConnectString("v_DATA_Database")
        myPortalADO.sqlString = sqlUpdate
        myPortalADO.SQLTextType = sqltext
        myPortalADO.Execute
    
     Set db = Nothing
    
    End Sub
    