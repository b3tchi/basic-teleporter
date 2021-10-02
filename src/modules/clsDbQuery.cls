VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_Query As AccessObject
Private m_AllItems As Collection
Private m_blnModifiedOnly As Boolean

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()
    
    Dim strFile As String
    Dim dbs As DAO.Database

    ' Save and sanitize file
    SaveComponentAsText acQuery, m_Query.Name, IDbComponent_SourceFile
    VCSIndex.Update Me, eatExport
    
    ' Export as SQL (if using that option)
    If Options.SaveQuerySQL Then
        Perf.OperationStart "Save Query SQL"
        Set dbs = CurrentDb
        strFile = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name) & ".sql"
        WriteFile b3AddLineBreaks(dbs.QueryDefs(m_Query.Name).SQL), strFile
        Perf.OperationEnd
        Log.Add "  " & m_Query.Name & " (SQL)", Options.ShowDebug
    End If
    
    Call QueryJsonParametersExport(dbs.QueryDefs(m_Query.Name))
   
End Sub

Private Function QueryJsonParametersExport(ByRef qdf As QueryDef) As Boolean

 ' Linked table - Save as JSON
    'Stop
    
    Dim cnt As String
    
    cnt = ""
    
    On Error Resume Next
    cnt = SanitizeConnectionString(qdf.Connect)
    On Error GoTo 0
    
    Dim dItem As Dictionary
    Set dItem = New Dictionary
    
    On Error Resume Next
    
    With dItem
        
        .Add "Name", qdf.Name
        .Add "Connect", cnt
        .Add "Type", qdf.Type
        .Add "ReturnsRecords", qdf.Properties("ReturnsRecords").Value
        .Add "Connect", qdf.Properties("Connect").Value
        .Add "Updatable", qdf.Properties("Updatable").Value
        .Add "ODBCTimeout", qdf.Properties("ODBCTimeout").Value
        .Add "RecordLocks", qdf.Properties("RecordLocks").Value
        .Add "UseTransaction", qdf.Properties("UseTransaction").Value
        .Add "TotalsRow", qdf.Properties("TotalsRow").Value
        .Add "LogMessages", qdf.Properties("LogMessages").Value
        .Add "FailOnError", qdf.Properties("FailOnError").Value
        .Add "ReadOnlyWhenDisconnected", qdf.Properties("ReadOnlyWhenDisconnected").Value
 
    End With
    
    On Error GoTo 0
    

    
    ' Write export file.
    Dim strFile As String
    strFile = IDbComponent_BaseFolder & GetSafeFileName(qdf.Name) & ".json"
    WriteJsonFile TypeName(Me), dItem, strFile, "Query Access Properties"
    
End Function




'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller / Indigo
' Date      : 10/24/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dbs As DAO.Database
    Dim strQueryName As String
    Dim strFileSql As String
    Dim strSql As String
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.sql" Then Exit Sub '"*.bas"
    
    ' Import query from file
    strQueryName = GetObjectNameFromFileName(strFile)
    'LoadComponentFromText acQuery, strQueryName, strFile
    'Set m_Query = CurrentData.AllQueries(strQueryName)
    'VCSIndex.Update Me, eatImport
    
    ' In some cases, such as when a query contains a subquery, AND has been modified in the
    ' visual query designer, it may be imported incorrectly and unable to run. For these
    ' cases we have added an option to overwrite the .SQL property with the SQL that we
    ' saved separately during the export. See the following link for further details:
    ' https://github.com/joyfullservice/msaccess-vcs-integration/issues/76
    
    ' Check option to import exact query from SQL
    If Options.ForceImportOriginalQuerySQL Then
    
        'CurrentDb.CreateQueryDef(""
    
        ' Replace .bas extension with .sql to get file content
        strFileSql = strFile 'Left$(strFIle, Len(strFIle) - 4) & ".sql"
        
        ' Tries to get SQL content from the SQL file previously exported
        strSql = ReadFile(strFileSql)

        'create new query
        Call CurrentDb.CreateQueryDef(strQueryName, strSql)
        
        Set m_Query = CurrentData.AllQueries(strQueryName)
        VCSIndex.Update Me, eatImport

        ' Update query def with saved SQL
        If strSql <> vbNullString Then
            Set dbs = CurrentDb
            dbs.QueryDefs(strQueryName).SQL = strSql
            Log.Add "  Restored original SQL for " & strQueryName, Options.ShowDebug
        Else
            Log.Add "  Couldn't get original SQL query for " & strQueryName
        End If
        
        'Add additional Access Query Parameters
        Dim strJsonFile As String
        strJsonFile = Left$(strFile, Len(strFile) - 4) & ".json"
        
        Call QueryJsonParametersImport(dbs.QueryDefs(strQueryName), strJsonFile)
        
        dbs.QueryDefs.Refresh
        
    End If
    
End Sub


Private Function QueryJsonParametersImport(ByRef qdf As QueryDef, strFile As String) As Boolean

    Dim dTable As Dictionary
    Dim dItem As Dictionary
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strSql As String
    Dim strConnect As String
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Read json file
    Set dTable = ReadJsonFile(strFile)
    If Not dTable Is Nothing Then
    
        ' Link the table
        Set dItem = dTable("Items")
        'Set dbs = CurrentDb
        'Set tdf = dbs.QueryDefs(dItem("Name"))
        strConnect = GetFullConnect(dItem("Connect"))

        If strConnect = "" Then Exit Function
        
        'With qdf
        '    .Connect = strConnect
        '    .SourceTableName = dItem("SourceTableName")
        '    .Attributes = SafeAttributes(dItem("Attributes"))
        'End With
        
        On Error Resume Next
         
        Dim varKey As Variant
        For Each varKey In dItem.Keys()
            'Range("A" & varKey).Value = oDic(varKey)
        
            Select Case varKey
            
            Case _
                  "Name" _
                , "Type" _
                , "Updatable" _
        
        
            Case Else
        
                With qdf
                    .Properties(varKey).Value = dItem(varKey)
                
                    If Err.Number = 3270 Then
                        Err.Clear
                        .Properties.Append .CreateProperty(CStr(varKey), typeOfProperty(CStr(varKey)), dItem(varKey))
                        
                    End If
                
                End With
            
                Call CatchAny(eelWarning, vbNullString, m_Query.Name & ":" & varKey & ".AddParamaters")
                'Debug.Print varKey, Err.Number, Err.Description
'                On Error GoTo 0
        
            End Select
        
        
        Next

        On Error GoTo 0
        'dbs.TableDefs.Append tdf
'        If Catch(3011) Then
'            Log.Error eelError, "Could not link table '" & dItem("SourceTableName") & "'", _
'            ModuleName & ".ImportLinkedTable"
'            Log.Add "Linked table object not found in " & strFile, False
'            Log.Add "Connection String: " & strConnect, False
'        Else
        If CatchAny(eelError, vbNullString, m_Query.Name & ".AddConnection") Then
            ' May have encountered other issue like a missing link specification.
        Else
            ' Verify that the connection matches the source file. (Issue #192)
'            If qdf.Connect <> strConnect Then
'                qdf.Connect = strConnect
'                df.RefreshLink
'            End If
'            dbs.TableDefs.Refresh
            
            ' Set index on linked table.
'            If InStr(1, tdf.Connect, ";DATABASE=", vbTextCompare) = 1 Then
'                ' Can't create a key on a linked Access database table.
'                ' Presumably this would use the Access index instead of needing the pseudo index
'            Else
'                ' Check for a primary key index (Linked SQL tables may bring over the index, but linked views won't.)
'                If dItem.Exists("PrimaryKey") And Not HasUniqueIndex(tdf) Then
'                    ' Create a pseudo index on the linked table
'                    strSql = "CREATE UNIQUE INDEX __uniqueindex ON [" & tdf.Name & "] (" & dItem("PrimaryKey") & ") WITH PRIMARY"
'                    dbs.Execute strSql, dbFailOnError
'                    dbs.TableDefs.Refresh
'                End If
'            End If
        End If
    End If
    
    ' Report any unhandled errors
    CatchAny eelError, "Error importing " & strFile, ".AddConnection"
    
End Function

Function typeOfProperty(propertyName) As Long
    
    Select Case propertyName
    Case _
          "Updatable" _
        , "ReturnsRecords" _
        , "UseTransaction" _
        , "DatasheetFontItalic" _
        , "DatasheetFontUnderline" _
        , "OrderByOn" _
        , "FilterOnLoad" _
        , "OrderByOnLoad" _
        , "TotalsRow" _
        , "HideNewField" _
        , "ReadOnlyWhenDisconnected" _
        , "FailOnError" _
        , "LogMessages" _
        , "SubdatasheetExpanded" _
    
        typeOfProperty = 1
    
    Case _
          "DatasheetFontName" _
        , "SubdatasheetName" _
    
        typeOfProperty = 10
    Case _
          "DOL" _
        , "NameMap" _
        
        typeOfProperty = 11
    Case _
          "Name" _
        , "SQL" _
        , "Connect" _
        
        typeOfProperty = 12
    Case _
          "RecordLocks" _
        , "Orientation" _
        , "PublishToWeb" _
        , "TabularCharSet" _
        , "TabularFamily" _
        , "RecordsetType" _
        , "DefaultView" _
        , "DisplayViewsOnSharePointSite" _
        
        typeOfProperty = 2
    Case _
          "Type" _
        , "ODBCTimeout" _
        , "DatasheetFontHeight" _
        , "DatasheetFontWeight" _
        , "SubdatasheetHeight" _
        
        typeOfProperty = 3
    Case _
          "RecordsAffected" _
        , "MaxRecords" _
        , "DatasheetBackColor" _
        , "DatasheetBackColor12" _
        , "DatasheetAlternateBackColor" _
        , "DatasheetForeColor" _
        , "DatasheetForeColor12" _
        , "BackThemeColorIndex" _
        , "BorderThemeColorIndex" _
        , "ThemeFontIndex" _
        , "ForeThemeColorIndex" _
        , "AlternateBackThemeColorIndex" _
        , "DatasheetGridlinesThemeColorIndex" _
        , "DatasheetForeThemeColorIndex" _
        
        typeOfProperty = 4
    Case _
          "BackTint" _
        , "BackShade" _
        , "BorderTint" _
        , "BorderShade" _
        , "ForeTint" _
        , "ForeShade" _
        , "AlternateBackTint" _
        , "AlternateBackShade" _
        
        typeOfProperty = 6
    Case _
          "DateCreated" _
        , "LastUpdated" _
        
        typeOfProperty = 8
    Case _
        "GUID" _
    
        typeOfProperty = 9
    Case Else
        typeOfProperty = -1
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    DeleteObjectIfExists acQuery, GetObjectNameFromFileName(strFile)
    IDbComponent_Import strFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection
    
    Dim qry As AccessObject
    Dim cQuery As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Or (blnModifiedOnly <> m_blnModifiedOnly) Then
        Set m_AllItems = New Collection
        m_blnModifiedOnly = blnModifiedOnly
        For Each qry In CurrentData.AllQueries
            Set cQuery = New clsDbQuery
            Set cQuery.DbObject = qry
            If blnModifiedOnly Then
                If cQuery.IsModified Then m_AllItems.Add cQuery, qry.Name
            Else
                m_AllItems.Add cQuery, qry.Name
            End If
        Next qry
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
            
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList(Optional blnModifiedOnly As Boolean = False) As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.sql")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearOrphanedSourceFiles Me, "bas", "sql", "json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = (m_Query.DateModified > VCSIndex.GetExportDate(Me))
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = m_Query.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModified() As Date
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = GetLastModifiedDate(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Queries"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "queries" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Query.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name) & ".sql"
    'IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name) & ".bas"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbQuery
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Query
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Query = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property
