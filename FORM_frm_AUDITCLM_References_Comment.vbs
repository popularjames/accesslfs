Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private mstrCommentFile As String
Private mintSaveComment As Boolean

Public Property Get FileName() As String
    FileName = mstrCommentFile
End Property
Property Let FileName(data As String)
     mstrCommentFile = data
End Property

Public Property Get SaveComment() As Boolean
    SaveComment = mintSaveComment
End Property
Property Let SaveComment(data As Boolean)
     mintSaveComment = data
End Property


Private Sub btn_Cancel_Click()
    FileName = ""
    DoCmd.Close acForm, Me.Name
    SaveComment = False
End Sub

Private Sub btn_OK_Click()
    ' Get a free file number
    FileNum = FreeFile
    
    CreateFolder "M:\tmp\"
    
    ' Create Test.txt
    FileName = "M:\tmp\Comment_" & Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "") & ".txt"
    Open FileName For Output As FileNum
    
    ' Write the contents of TextBox1 to Test.txt
    Print #FileNum, Comment
    
    ' Close the file
    Close FileNum
    SaveComment = True
    DoCmd.Close acForm, Me.Name
    
End Sub
