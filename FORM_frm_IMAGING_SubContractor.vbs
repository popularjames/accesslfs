Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbContinue As Boolean
Dim miError As Long
Dim miFileCnt As Long

Dim mstrLocalHoldPath As String
Dim mstrLocalPath As String
Dim mstrRemotePath As String
Dim mstrHoldImageName As String
Dim mstrLocalImageName As String
Dim mstrRemoteImageName As String
Dim mstrLocalImagePath As String
Dim mstrRemoteImagePath As String

Dim mstrCalledFrom As String

Const CstrFrmAppID As String = "SubContractor"
Private Sub Command0_Click()
TransferFiles
End Sub


Private Sub TransferFiles()


    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim rsScanImages As ADODB.RecordSet
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    Dim strSQL As String
    Dim strFile As String
    
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'On Error GoTo Err_handler
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    
    strSQL = " select RR.RefLink ,H.LastUpDt"
    strSQL = strSQL & " from dbo.AUDITCLM_Hdr H JOIN AUDITCLM_References RR ON H.CnlyClaimNum = RR.CnlyClaimNum "
    strSQL = strSQL & " where LOB = 'VIANT'"
    'strSQL = strSQL & " and ClmStatus = '303'"
    
    ' TK replacing the RefSubType with SubType filter
    'strSQL = strSQL & " and RR.RefSubType  IN (  'MR', 'UB92', 'CORR' )  "
    strSQL = strSQL & " and RR.RefType  = 'IMAGE' "
    
    
    
    Set rs = MyAdo.OpenRecordSet(strSQL)
    
    mstrRemotePath = "\\cca-audit\dfs-dc-01\Imaging\Misc\Hold\In\CMS\MedicalRecords\VIANT\"
      
    While Not rs.EOF
        'lstFiles.AddItem "Image does not exists;" & mstrHoldImageName
        If Not fso.FileExists(rs!RefLink) Then
            'lstFiles.AddItem "Image Not Ready;" & strFile
            GoTo NextImage
        End If
        
        strFile = fso.GetFileName(rs!RefLink)
        
        
        
        If fso.FileExists(mstrRemotePath & strFile) Then
         '   lstFiles.AddItem "Image Already Moved;" & strFile
        Else
            Call fso.CopyFile(rs!RefLink, mstrRemotePath & strFile, False)
          '  lstFiles.AddItem "Copied;" & left(strFile, 15)
        End If
       
    

        
NextImage:         rs.MoveNext
    Wend


End Sub
