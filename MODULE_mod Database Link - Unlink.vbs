Option Compare Database
Option Explicit

'Private Const LINK_SRC_ACCESS = "Provider=Microsoft.Jet.OLEDB.4.0;"
 Private Const LINK_SRC_ACCESS = "Provider=Microsoft.ACE.OLEDB.12.0;"

Private Const LINK_SRC_SQL = "Provider=SQLOLEDB;Integrated Security='SSPI';"

Private Enum LinkType
    LNK_ACCESS = 0
    LNK_SQL = 1
End Enum

Private Type LinkCnf
    LinkType As LinkType
    Server As String
    Schema As String
    Database As String
    Connect As String
    Prefix As String
    Suffix As String
    UseViews As Boolean
    TableName As String
    DropIfExists As Boolean
End Type

Public Sub UnLinkTables(Optional TableName As String)
    Dim AdCat    'As ADOX.Catalog
    Dim X As Integer

    On Error GoTo CATCH

    Set AdCat = CreateObject("ADOX.Catalog")
    AdCat.ActiveConnection = LINK_SRC_ACCESS & "Data Source=" & CurrentDb.Name & ";"

    screen.MousePointer = 11
    If TableName <> "" Then
        '20130226 KD: Tired of fixing this
        If IsTable(TableName) = True Then
            AdCat.Tables.Delete TableName
        End If
    Else
        For X = AdCat.Tables.Count - 1 To 0 Step -1
            If AdCat.Tables(X).Properties("Jet OLEDB:Create Link") = True And AdCat.Tables(X).Type <> "VIEW" Then
                AdCat.Tables.Delete X
                'If Cancel = True Then Exit For
                DoEvents
            End If
        Next X
    End If
    
Done:
    On Error Resume Next
    screen.MousePointer = 0
    Set AdCat = Nothing
    Exit Sub

CATCH:
    If Application.UserControl = True Then
        MsgBox Err.Description, vbCritical, CodeContextObject.Name
    Else
        LogMessage "mod Database Link - Unlink.UnLinkTables", "ERROR", Err.Description, CodeContextObject.Name
    End If
    Resume Next
End Sub

Public Sub LinkTable(LinkType As String, ServerName As String, DatabaseName As String, Optional TableName As String, Optional Schema As String = "dbo")
    
    Dim AdCatSrc    'As ADOX.Catalog
    
    Dim strProvider As String
    Dim Link As LinkCnf
    
    On Error GoTo CATCH

    Set AdCatSrc = CreateObject("ADOX.Catalog")
    Select Case LinkType
        Case "ACCESS"
            With Link
                .LinkType = LNK_ACCESS
                .Server = ""
                .Database = "" & DatabaseName
                .Connect = ""
                .Prefix = ""
                .Suffix = ""
                .DropIfExists = True
            End With
            strProvider = LINK_SRC_ACCESS & "Data Source=" & DatabaseName & ";"
        Case "SQL"
            With Link
                .LinkType = LNK_SQL
                .Server = ServerName
                .Database = DatabaseName
                .Connect = "ODBC;DRIVER=SQL Server;SERVER=" & .Server & ";APP=ClaimsAdmin;Trusted_Connection=Yes;DATABASE=" & .Database & ";"
                .Schema = Schema
                .Prefix = ""
                .Suffix = ""
                .DropIfExists = True
            End With
            strProvider = LINK_SRC_SQL & "Server=" & ServerName & ";" & "Database=" & DatabaseName & ";"
        Case Else
            GoTo Done
    End Select
    AdCatSrc.ActiveConnection = strProvider
    If TableName <> "" Then
        LinkTable_Execute AdCatSrc, Link, TableName, Schema
        Debug.Print TableName
    Else
        LinkTable_Execute AdCatSrc, Link
    End If


Done:
    On Error Resume Next
    Set AdCatSrc = Nothing
    Exit Sub

CATCH:
    If Application.UserControl = True Then
        MsgBox Err.Description & " |Table:" & TableName, vbCritical, CodeContextObject.Name
    Else
        LogMessage "mod Database Link - Unlink.LinkTable", "ERROR", Err.Description, CodeContextObject.Name
    End If
    Resume Done
End Sub


Private Function LinkTable_Execute(AdCatSrc, Link As LinkCnf, _
                                   Optional TableName As String, _
                                   Optional Schema As String = "dbo", _
                                   Optional IncludeHidden As Boolean = False) As Boolean
    On Error GoTo CATCH
    Dim AdCat    'As ADOX.Catalog for Access DB
    Dim AdTbl    'As ADOX.Table for Access table
    Dim AdSrc    'As ADOX.Table for remote DB
    Dim strChkType As String



    If AdCatSrc Is Nothing Then GoTo Done
    
    Set AdCat = CreateObject("ADOX.Catalog")
    AdCat.ActiveConnection = CurrentProject.Connection
    
    If TableName <> "" Then
        strChkType = "CheckTable"
        Set AdSrc = AdCatSrc(TableName)
        Link.Schema = Schema
        Link.TableName = TableName
        Select Case Link.LinkType
            Case LinkType.LNK_ACCESS
                If (AdSrc.Properties("Jet OLEDB:Table Hidden In Access").Value = False Or IncludeHidden = True) And UCase(left(Link.TableName, 4)) <> "MSYS" And AdSrc.Properties("Jet OLEDB:Create Link") = False Then
                    Set AdTbl = CreateObject("ADOX.Table")
                    With AdTbl
                        .Name = Link.Prefix & Link.TableName & Link.Suffix
                        Set .ParentCatalog = AdCat
                        .Properties("Jet OLEDB:Create Link") = True
                        .Properties("Jet OLEDB:Link Datasource") = Link.Database
                        .Properties("Jet OLEDB:Remote Table Name") = Link.TableName
                    End With
                    AdCat.Tables.Append AdTbl
                End If
            Case LinkType.LNK_SQL
                If left(AdSrc.Type, 6) <> "SYSTEM" Then
                    Set AdTbl = CreateObject("ADOX.Table")
                    With AdTbl
                        .Name = Link.Prefix & Link.TableName & Link.Suffix
                        Set .ParentCatalog = AdCat
                        .Properties("Jet OLEDB:Create Link") = True
                        .Properties("Jet OLEDB:Link Provider String") = Link.Connect
                        .Properties("Jet OLEDB:Remote Table Name") = Link.Schema & "." & Link.TableName
                    End With
                    If Link.DropIfExists = True Then
                        strChkType = "Delete"
                        If IsTable(Link.TableName) = True Then  ' 20130111 KD: Got tired of the errors!
                            AdCat.Tables.Delete Link.TableName
                        End If
                    End If
                    AdCat.Tables.Append AdTbl
                End If
        End Select
    Else
        For Each AdSrc In AdCatSrc.Tables
            Link.TableName = AdSrc.Name
            Select Case Link.LinkType
                Case LinkType.LNK_ACCESS
                    If (AdSrc.Properties("Jet OLEDB:Table Hidden In Access").Value = False Or IncludeHidden = True) And UCase(left(Link.TableName, 4)) <> "MSYS" And AdSrc.Properties("Jet OLEDB:Create Link") = False Then
                        Set AdTbl = CreateObject("ADOX.Table")
                        With AdTbl
                            .Name = Link.Prefix & Link.TableName & Link.Suffix
                            Set .ParentCatalog = AdCat
                            .Properties("Jet OLEDB:Create Link") = True
                            .Properties("Jet OLEDB:Link Datasource") = Link.Database
                            .Properties("Jet OLEDB:Remote Table Name") = Link.TableName
                        End With
                        AdCat.Tables.Append AdTbl
                    End If
                Case LinkType.LNK_SQL
                    If left(AdSrc.Type, 6) <> "SYSTEM" Then
                        Set AdTbl = CreateObject("ADOX.Table")
                        With AdTbl
                            .Name = Link.Prefix & Link.TableName & Link.Suffix
                            Set .ParentCatalog = AdCat
                            .Properties("Jet OLEDB:Create Link") = True
                            .Properties("Jet OLEDB:Link Provider String") = Link.Connect
                            .Properties("Jet OLEDB:Remote Table Name") = Link.TableName
                        End With
                        If Link.DropIfExists = True Then
                            strChkType = "Delete"
                            AdCat.Tables.Delete Link.TableName
                        End If
                        AdCat.Tables.Append AdTbl
                    End If
            End Select
NextTable:
        Next AdSrc
    End If

        '' 20120703: KD added this to make sure that the tables are refreshed!
    AdCat.Tables.Refresh

Done:
    On Error Resume Next
    Set AdCat = Nothing
    Set AdTbl = Nothing
    Set AdSrc = Nothing
    Exit Function

CATCH:
    Select Case Err.Number
        Case -2147217857    'Table Already Exists
            If TableName = "" Then
                Resume NextTable
            End If
        Case 3265    'Tried to delete a table that does not exist
            If strChkType = "Delete" Then
                LogMessage "mod_Database_Link-Unlink.LinkTable_Execute", "WARNING", "Tried to delete a table that does not exist yet", Link.TableName
                Resume Next
            Else
                If strChkType = "CheckTable" Then
                    If Application.UserControl = True Then
                        MsgBox Err.Number & " -- Can not find table: " & TableName
                    Else
                        LogMessage "mod Database Link - Unlink.LinkTable_Execute", "ERROR", "Cannot find table: " & TableName, TableName
                    End If
                End If
            End If
        Case Else
            If Application.UserControl = True Then
                MsgBox Err.Description, vbCritical, CodeContextObject.Name
            Else
                LogMessage "mod Database Link - Unlink.LinkTable_Execute", "ERROR", Err.Description, CodeContextObject.Name
            End If
            Resume Done
            Resume
    End Select
End Function