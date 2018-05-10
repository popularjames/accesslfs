Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




''' Last Modified: 09/12/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 09/12/2012 - added 'TableName' as property and made sure we can use with SQL Server via ADO
'''  - 03/09/2012 - Created...
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################

Private Const cs_SETTING_TABLE_NAME As String = "[_Settings]"

Private cdctSettings As Scripting.Dictionary
Private cdctMultSettings As Scripting.Dictionary
Private cblnInitialized As Boolean
Private cdctLinkedTables As Scripting.Dictionary


Private csCurTableName As String
Private csConnString As String

Public Property Get CurrentTableName() As String
    If csCurTableName = "" Then
        CurrentTableName = cs_SETTING_TABLE_NAME
    Else
        CurrentTableName = csCurTableName
    End If
End Property
Public Property Let CurrentTableName(sTableNameToUse As String)
    csCurTableName = sTableNameToUse
End Property



Public Property Get ConnString() As String
    ConnString = csConnString
End Property
Public Property Let ConnString(sConnString As String)
    csConnString = sConnString
    ' now that we have a connection string, refresh it..
    Call Me.Refresh
End Property





Public Property Get TablePathPattern(strTableName As String) As String
    strTableName = UCase(strTableName)

    If cdctLinkedTables.Exists(strTableName) Then
        TablePathPattern = cdctLinkedTables.Item(strTableName)
    Else
        TablePathPattern = ""
    End If
End Property



Public Property Get MultipleSettings(strValueName As String) As Variant
    strValueName = UCase(strValueName)
    
    If cblnInitialized = False Then
        Call Class_Initialize
    End If
    
    If cdctMultSettings.Exists(strValueName) Then
        MultipleSettings = cdctMultSettings.Item(strValueName)
    Else
        ' Query db or return a null array.
        ' how about a single setting?
        If cdctSettings.Exists(strValueName) = True Then
            MultipleSettings = Array(cdctSettings.Item(strValueName))
        Else
            MultipleSettings = Array()
        End If
    End If

End Property


Public Property Get Setting(strValueName As String) As String
    strValueName = UCase(strValueName)
    
    If cblnInitialized = False Then
        Call Class_Initialize
    End If
    
    
    If cdctSettings.Exists(strValueName) Then
        Setting = IIf(IsNull(cdctSettings.Item(strValueName)), "", cdctSettings.Item(strValueName))
    Else
        Setting = ""
    End If

End Property
    Public Property Get GetSetting(strValueName As String) As String
       GetSetting = Me.Setting(strValueName)
    End Property



Public Function SetSetting(strValueName As String, strValue As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim objRS As DAO.RecordSet
Dim objDB As DAO.Database
Dim dtPrevMktDate As Date

    strProcName = ClassName & ".Property.Let.Setting"

    strValueName = UCase(strValueName)

        ' update the settings table:
    Set objDB = CurrentDb()
    Set objRS = objDB.OpenRecordSet("SELECT [NAME], [VALUE] FROM " & cs_SETTING_TABLE_NAME & " WHERE [Name] = '" & strValueName & "'", dbOpenDynaset)
    
    If objRS.EOF Then
        objRS.AddNew
        objRS("NAME").Value = strValueName
        objRS("VALUE").Value = strValue
        objRS.Update
    Else
        objRS.Edit
        objRS("VALUE").Value = strValue
        objRS.Update
    End If
    objRS.Close
    
        ' Now update the object
    If cdctSettings.Exists(strValueName) = True Then
        ' Change it
        cdctSettings.Item(strValueName) = strValue
    Else
        ' Add it
        cdctSettings.Add UCase(strValueName), strValue
    End If
    
'    ' If we're doing the market date then we need to change the tbl_MarketDate table:
'    If UCase(strValueName) = "MARKETDATE" Or UCase(strValueName) = "PREVIOUSMARKETDATE" Then
'        dtPrevMktDate = GetPreviousTradeDate(CDate(strValue))
'
'        Set objRS = objDb.OpenRecordset("SELECT MarketDate, PreviousMarketDate FROM tbl_MarketDate") ', dbOpenDynaset, , dbOptimistic)
'
'        If IsNull(objRS("MarketDate").Value) Then
'            RefreshMarketDate
'            objRS.Requery
'        End If
'        objRS.Edit
'        objRS("MarketDate").Value = CDate(cdctSettings.Item("MARKETDATE"))
'        objRS("PreviousMarketDate").Value = dtPrevMktDate
'        objRS.Update
'        objRS.Close
'        ' Now update the object
'        If cdctSettings.Exists("PREVIOUSMARKETDATE") = True Then
'            ' Change it
'            cdctSettings.Item("PREVIOUSMARKETDATE") = CStr(dtPrevMktDate)
'        Else
'            ' Add it
'            cdctSettings.Add UCase("PREVIOUSMARKETDATE"), CStr(dtPrevMktDate)
'        End If
'
'
'    End If
    
Block_Exit:
    Set objRS = Nothing
    Set objDB = Nothing
    Exit Function

Block_Err:
    ReportError Err, strProcName, "Attempting to set name: " & strValueName & " = '" & strValue & "'", False
    GoTo Block_Exit
End Function


Public Sub Refresh()
    Set cdctSettings = Nothing
    Call Class_Initialize
End Sub

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Private Sub Class_Initialize()
On Error GoTo Block_Err
Dim strProcName As String
Dim objRS As DAO.RecordSet
Dim objDB As DAO.Database
Dim sSql As String
Dim arrNames()
Dim arrValues()
Dim strLastName As String
Dim strThisName As String
Dim iNameCount As Integer
Dim iValCount As Integer

    strProcName = ClassName & ".Class_Initialize"

    Set cdctSettings = New Scripting.Dictionary
    
    Set objDB = CurrentDb()
    Set objRS = objDB.OpenRecordSet("SELECT [NAME], [VALUE] FROM " & cs_SETTING_TABLE_NAME) ', dbOpenDynaset, , dbOptimistic)
    
    While Not objRS.EOF
        If cdctSettings.Exists(UCase(objRS("NAME").Value)) = True Then
                ' Change it
            cdctSettings.Item(UCase(objRS("NAME").Value)) = objRS("VALUE").Value
        Else
                ' Add it
            cdctSettings.Add UCase(objRS("NAME").Value), objRS("VALUE").Value
        End If
        objRS.MoveNext
    Wend
    objRS.Close
    
    Set cdctMultSettings = New Scripting.Dictionary
    
    sSql = "SELECT " & cs_SETTING_TABLE_NAME & ".[Name], " & cs_SETTING_TABLE_NAME & ".[Value], " & cs_SETTING_TABLE_NAME & ".[Active] " & _
        " From " & cs_SETTING_TABLE_NAME & _
        " WHERE (((" & cs_SETTING_TABLE_NAME & ".[Name]) In (SELECT [Name] FROM " & cs_SETTING_TABLE_NAME & " As Tmp " & _
        " GROUP BY [Name] HAVING Count(*)>1 ))) AND " & cs_SETTING_TABLE_NAME & ".[Active] = True " & _
        " ORDER BY " & cs_SETTING_TABLE_NAME & ".[Name] "
    
    iNameCount = 1
    Set objRS = objDB.OpenRecordSet(sSql)
    While Not objRS.EOF
        strThisName = UCase(objRS("NAME").Value)
        
        If strThisName = strLastName Or strLastName = "" Then
            iValCount = iValCount + 1
            ReDim Preserve arrValues(iValCount - 1)
            arrValues(iValCount - 1) = objRS("VALUE").Value
        Else
            iNameCount = iNameCount + 1
            ReDim Preserve arrNames(iNameCount - 1)
            arrNames(iNameCount - 1) = UCase(strLastName)
            cdctMultSettings.Add strLastName, arrValues
        
            ReDim arrValues(0)
            arrValues(0) = objRS("VALUE").Value
            iValCount = 1
        End If
        
        strLastName = strThisName
        objRS.MoveNext
    Wend
    

    cdctMultSettings.Add strThisName, arrValues
    
'   Set objRS = objDb.OpenRecordset("SELECT MarketDate FROM tbl_MarketDate") ', dbOpenDynaset, , dbOptimistic)
'
'    If IsNull(objRS("MarketDate").Value) Then
'        RefreshMarketDate
'        objRS.Requery
'    End If
'    SetSetting "MarketDate", CStr(objRS("MarketDate").Value)
'    objRS.Close
    
    
    cblnInitialized = True
Block_Exit:
    Set objRS = Nothing
    Set objDB = Nothing
    Exit Sub

Block_Err:
    ReportError Err, strProcName, True
    GoTo Block_Exit
End Sub



Private Function PopulateObjectDAO()
On Error GoTo Block_Err
Dim strProcName As String
Dim objRS As DAO.RecordSet
Dim objDB As DAO.Database
Dim sSql As String
Dim arrNames()
Dim arrValues()
Dim strLastName As String
Dim strThisName As String
Dim iNameCount As Integer
Dim iValCount As Integer

    strProcName = ClassName & ".PopulateObjectDAO"

    Set cdctSettings = New Scripting.Dictionary
    
    Set objDB = CurrentDb()
    Set objRS = objDB.OpenRecordSet("SELECT [NAME], [VALUE] FROM " & cs_SETTING_TABLE_NAME) ', dbOpenDynaset, , dbOptimistic)
    
    While Not objRS.EOF
        If cdctSettings.Exists(UCase(objRS("NAME").Value)) = True Then
                ' Change it
            cdctSettings.Item(UCase(objRS("NAME").Value)) = objRS("VALUE").Value
        Else
                ' Add it
            cdctSettings.Add UCase(objRS("NAME").Value), objRS("VALUE").Value
        End If
        objRS.MoveNext
    Wend
    objRS.Close
    
    Set cdctMultSettings = New Scripting.Dictionary
    
    sSql = "SELECT " & cs_SETTING_TABLE_NAME & ".[Name], " & cs_SETTING_TABLE_NAME & ".[Value], " & cs_SETTING_TABLE_NAME & ".[Active] " & _
        " From " & cs_SETTING_TABLE_NAME & _
        " WHERE (((" & cs_SETTING_TABLE_NAME & ".[Name]) In (SELECT [Name] FROM " & cs_SETTING_TABLE_NAME & " As Tmp " & _
        " GROUP BY [Name] HAVING Count(*)>1 ))) AND " & cs_SETTING_TABLE_NAME & ".[Active] = True " & _
        " ORDER BY " & cs_SETTING_TABLE_NAME & ".[Name] "
    
    iNameCount = 1
    Set objRS = objDB.OpenRecordSet(sSql)
    While Not objRS.EOF
        strThisName = UCase(objRS("NAME").Value)
        
        If strThisName = strLastName Or strLastName = "" Then
            iValCount = iValCount + 1
            ReDim Preserve arrValues(iValCount - 1)
            arrValues(iValCount - 1) = objRS("VALUE").Value
        Else
            iNameCount = iNameCount + 1
            ReDim Preserve arrNames(iNameCount - 1)
            arrNames(iNameCount - 1) = UCase(strLastName)
            cdctMultSettings.Add strLastName, arrValues
        
            ReDim arrValues(0)
            arrValues(0) = objRS("VALUE").Value
            iValCount = 1
        End If
        
        strLastName = strThisName
        objRS.MoveNext
    Wend
    

    cdctMultSettings.Add strThisName, arrValues
    
'   Set objRS = objDb.OpenRecordset("SELECT MarketDate FROM tbl_MarketDate") ', dbOpenDynaset, , dbOptimistic)
'
'    If IsNull(objRS("MarketDate").Value) Then
'        RefreshMarketDate
'        objRS.Requery
'    End If
'    SetSetting "MarketDate", CStr(objRS("MarketDate").Value)
'    objRS.Close
    
    
    cblnInitialized = True
Block_Exit:
    Set objRS = Nothing
    Set objDB = Nothing
    Exit Function

Block_Err:
    ReportError Err, strProcName, True
    GoTo Block_Exit
End Function


Private Function PopulateObjectADO()
On Error GoTo Block_Err
Dim strProcName As String
Dim objRS As ADODB.RecordSet
Dim oAdo As clsADO

Dim sSql As String
Dim arrNames()
Dim arrValues()
Dim strLastName As String
Dim strThisName As String
Dim iNameCount As Integer
Dim iValCount As Integer

    strProcName = ClassName & ".PopulateObjectADO"

    Set cdctSettings = New Scripting.Dictionary
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.ConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT [NAME], [VALUE] FROM " & Me.CurrentTableName
        Set objRS = .ExecuteRS
            
    End With

    
    While Not objRS.EOF
        If cdctSettings.Exists(UCase(objRS("NAME").Value)) = True Then
                ' Change it
            cdctSettings.Item(UCase(objRS("NAME").Value)) = objRS("VALUE").Value
        Else
                ' Add it
            cdctSettings.Add UCase(objRS("NAME").Value), objRS("VALUE").Value
        End If
        objRS.MoveNext
    Wend
    If objRS.State = adStateOpen Then objRS.Close
    
    Set cdctMultSettings = New Scripting.Dictionary
    
    sSql = "SELECT " & Me.CurrentTableName & ".[Name], " & Me.CurrentTableName & ".[Value], " & Me.CurrentTableName & ".[Active] " & _
        " From " & Me.CurrentTableName & _
        " WHERE (((" & Me.CurrentTableName & ".[Name]) In (SELECT [Name] FROM " & Me.CurrentTableName & " As Tmp " & _
        " GROUP BY [Name] HAVING Count(*)>1 ))) AND " & Me.CurrentTableName & ".[Active] = True " & _
        " ORDER BY " & Me.CurrentTableName & ".[Name] "
    
    iNameCount = 1
    oAdo.sqlString = sSql
    Set objRS = oAdo.ExecuteRS
    
    
'    Set objRS = objDb.OpenRecordSet(sSql)
    While Not objRS.EOF
        strThisName = UCase(objRS("NAME").Value)
        
        If strThisName = strLastName Or strLastName = "" Then
            iValCount = iValCount + 1
            ReDim Preserve arrValues(iValCount - 1)
            arrValues(iValCount - 1) = objRS("VALUE").Value
        Else
            iNameCount = iNameCount + 1
            ReDim Preserve arrNames(iNameCount - 1)
            arrNames(iNameCount - 1) = UCase(strLastName)
            cdctMultSettings.Add strLastName, arrValues
        
            ReDim arrValues(0)
            arrValues(0) = objRS("VALUE").Value
            iValCount = 1
        End If
        
        strLastName = strThisName
        objRS.MoveNext
    Wend
    

    cdctMultSettings.Add strThisName, arrValues
    
'   Set objRS = objDb.OpenRecordset("SELECT MarketDate FROM tbl_MarketDate") ', dbOpenDynaset, , dbOptimistic)
'
'    If IsNull(objRS("MarketDate").Value) Then
'        RefreshMarketDate
'        objRS.Requery
'    End If
'    SetSetting "MarketDate", CStr(objRS("MarketDate").Value)
'    objRS.Close
    
    
    cblnInitialized = True
Block_Exit:
    If objRS.State = adStateOpen Then objRS.Close
    
    Set objRS = Nothing
    Set oAdo = Nothing
    Exit Function

Block_Err:
    ReportError Err, strProcName, True
    GoTo Block_Exit
End Function



'Public Function RefreshMarketDate(Optional dtNewDate As Date) As Date
'Dim dtDummy As Date
'    If dtNewDate = dtDummy Then
''        cdctSettings("MARKETDATE") = dtDummy
'        cdctSettings("MARKETDATE") = Format(Now(), "m/d/yyyy")
'    Else
'        cdctSettings("MARKETDATE") = Format(dtNewDate, "m/d/yyyy")
'    End If
'    RefreshMarketDate = cdctSettings("MARKETDATE")
'End Function


Private Sub Class_Terminate()
    Set cdctSettings = Nothing
    Set cdctMultSettings = Nothing
End Sub