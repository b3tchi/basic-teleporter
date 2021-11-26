Attribute VB_Name = "modDatabase"
'---------------------------------------------------------------------------------------
' Module    : modDatabase
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General functions for interacting with the current database.
'           : (See modVCSUtility for other functions more specific to this add-in.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : ProjectPath
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Path/Directory of the current database file.
'---------------------------------------------------------------------------------------
'
Public Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right$(ProjectPath, 1) <> PathSep Then ProjectPath = ProjectPath & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Get a database property (Default to MDB version)
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String, Optional dbs As DAO.Database) As Variant

    Dim prp As Object ' DAO.Property
    Dim oParent As Object
    
    ' Check for database reference
    If Not dbs Is Nothing Then
        Set oParent = dbs.Properties
    Else
        If DatabaseOpen Then
            ' Get parent container for properties
            If CurrentProject.ProjectType = acADP Then
                Set oParent = CurrentProject.Properties
            Else
                If dbs Is Nothing Then Set dbs = CurrentDb
                Set oParent = dbs.Properties
            End If
        Else
            ' No database open
            GetDBProperty = vbNullString
            Exit Function
        End If
    End If
    
    ' Look for property by name
    For Each prp In oParent
        If prp.Name = strName Then
            GetDBProperty = prp.Value
            Exit For
        End If
    Next prp
    Set prp = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Set a database property
'---------------------------------------------------------------------------------------
'
Public Sub SetDBProperty(ByVal strName As String, ByVal varValue As Variant, Optional ByVal prpType As Long = dbText, Optional dbs As DAO.Database)

    Dim prp As Object ' DAO.Property
    Dim blnFound As Boolean
    Dim oParent As Object
    
    ' Properties set differently for databases and ADP projects
    If CurrentProject.ProjectType = acADP Then
        Set oParent = CurrentProject.Properties
    Else
        If dbs Is Nothing Then Set dbs = CurrentDb
        Set oParent = dbs.Properties
    End If
    
    ' Look for property in collection
    For Each prp In oParent
        If prp.Name = strName Then
            ' Check for matching type
            If Not dbs Is Nothing Then
                If prp.Type <> prpType Then
                    ' Remove so we can add it back in with the correct type.
                    dbs.Properties.Delete strName
                    Exit For
                End If
            End If
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then
                Set dbs = Nothing
            Else
                ' Update value
                prp.Value = varValue
            End If
            Exit Sub
        End If
    Next prp
    
    ' Add new property
    If Not blnFound Then
        If CurrentProject.ProjectType = acADP Then
            CurrentProject.Properties.Add strName, varValue
        Else
            Set prp = dbs.CreateProperty(strName, prpType, varValue)
            dbs.Properties.Append prp
            Set dbs = Nothing
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDAOProperty
' Author    : Adam Waller
' Date      : 5/8/2020
' Purpose   : Updates a DAO property, adding if it does not exist or is the wrong type.
'---------------------------------------------------------------------------------------
'
Public Sub SetDAOProperty(objParent As Object, intType As Integer, strName As String, varValue As Variant)

    Dim prp As DAO.Property
    Dim blnFound As Boolean
    
    ' Look through existing properties.
    For Each prp In objParent.Properties
        If prp.Name = strName Then
            blnFound = True
            Exit For
        End If
    Next prp

    ' Verify type, and update value if found.
    If blnFound Then
        If prp.Type <> intType Then
            objParent.Properties.Delete strName
            blnFound = False
        Else
            If objParent.Properties(strName).Value <> varValue Then
                objParent.Properties(strName).Value = varValue
            End If
        End If
    End If
        
    ' Add new property if needed
    If Not blnFound Then
        ' Create property, then append to collection
        Set prp = objParent.CreateProperty(strName, intType, varValue)
        objParent.Properties.Append prp
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Sets a custom access object property.
'---------------------------------------------------------------------------------------
'
Public Sub SetAccessObjectProperty(objItem As AccessObject, strProperty As String, strValue As String)
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            ' Update value of property.
            prp.Value = strValue
            Exit Sub
        End If
    Next prp
    ' Property not found. Create it.
    objItem.Properties.Add strProperty, strValue
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Get the value of a custom access property
'---------------------------------------------------------------------------------------
'
Public Function GetAccessObjectProperty(objItem As AccessObject, strProperty As String, Optional strDefault As String) As Variant
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            GetAccessObjectProperty = prp.Value
            Exit Function
        End If
    Next prp
    ' Nothing found. Return default
    GetAccessObjectProperty = strDefault
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : Adam Waller
' Date      : 9/22/2017
' Purpose   : Returns true if the object is loaded and not in design view.
'---------------------------------------------------------------------------------------
'
Public Function IsLoaded(intType As AcObjectType, strName As String, Optional blnAllowDesignView As Boolean = False) As Boolean

    Dim frm As Form
    Dim ctl As Control
    
    If SysCmd(acSysCmdGetObjectState, intType, strName) <> adStateClosed Then
        If blnAllowDesignView Then
            IsLoaded = True
        Else
            Select Case intType
                Case acReport
                    IsLoaded = Reports(strName).CurrentView <> acCurViewDesign
                Case acForm
                    IsLoaded = Forms(strName).CurrentView <> acCurViewDesign
                Case acServerView
                    IsLoaded = CurrentData.AllViews(strName).CurrentView <> acCurViewDesign
                Case acStoredProcedure
                    IsLoaded = CurrentData.AllStoredProcedures(strName).CurrentView <> acCurViewDesign
                Case Else
                    ' Other unsupported object
                    IsLoaded = True
            End Select
        End If
    Else
        ' Could be loaded as subform
        If intType = acForm Then
            For Each frm In Forms
                For Each ctl In frm.Controls
                    If TypeOf ctl Is SubForm Then
                        If ctl.SourceObject = strName Then
                            IsLoaded = True
                            Exit For
                        End If
                    End If
                Next ctl
                If IsLoaded Then Exit For
            Next frm
        End If
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseAllFormsReports
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Close all open forms and reports. Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Function CloseAllFormsReports() As Boolean

    Dim strName As String
    Dim intOpened As Integer
    Dim intItem As Integer
    
    ' Get count of opened objects
    intOpened = Forms.Count + Reports.Count
    If intOpened > 0 Then
        On Error GoTo ErrorHandler
        ' Loop through forms
        For intItem = Forms.Count - 1 To 0 Step -1
            If Forms(intItem).Caption <> "MSAccessVCS" Then
                DoCmd.Close acForm, Forms(intItem).Name
                DoEvents
            End If
            intOpened = intOpened - 1
        Next intItem
        ' Loop through reports
        Do While Reports.Count > 0
            strName = Reports(0).Name
            DoCmd.Close acReport, strName
            DoEvents
            intOpened = intOpened - 1
        Loop
        If intOpened = 0 Then CloseAllFormsReports = True
    Else
        ' No forms or reports currently open.
        CloseAllFormsReports = True
    End If
    
    Exit Function

ErrorHandler:
    Debug.Print "Error closing " & strName & ": " & Err.Number & vbCrLf & Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : ProjectIsSelected
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the base project is selected in the VBE
'---------------------------------------------------------------------------------------
'
Public Function ProjectIsSelected() As Boolean
    ProjectIsSelected = (Application.VBE.SelectedVBComponent Is Nothing)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SelectionInActiveProject
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the current selection is in the active project
'---------------------------------------------------------------------------------------
'
Public Function SelectionInActiveProject() As Boolean
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.filename = GetUncPath(CurrentProject.FullName))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBProjectForCurrentDB
' Author    : Adam Waller
' Date      : 7/25/2017
' Purpose   : Get the actual VBE project for the current top-level database.
'           : (This is harder than you would think!)
'---------------------------------------------------------------------------------------
'
Public Function GetVBProjectForCurrentDB() As VBProject
    Set GetVBProjectForCurrentDB = GetProjectByName(CurrentProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeVBProject
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get a reference to the VB Project for the running code.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeVBProject() As VBProject
    Set GetCodeVBProject = GetProjectByName(CodeProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetProjectByName
' Author    : Adam Waller
' Date      : 5/26/2020
' Purpose   : Return the VBProject by file path. (Also supports network drives)
'---------------------------------------------------------------------------------------
'
Private Function GetProjectByName(ByVal strPath As String) As VBProject

    Dim objProj As VBIDE.VBProject
    Dim strUncPath As String
    
    ' Use currently active project by default
    Set GetProjectByName = VBE.ActiveVBProject
    
    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(strPath)
    
    If VBE.ActiveVBProject.filename <> strUncPath Then
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.filename = strUncPath Then
                Set GetProjectByName = objProj
                Exit For
            End If
        Next objProj
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunInCurrentProject
' Author    : Adam Waller
' Date      : 4/22/2020
' Purpose   : Use the Run command but make sure it is running in the context of the
'           : current project, not the add-in file.
'---------------------------------------------------------------------------------------
'
Public Sub RunSubInCurrentProject(strSubName As String)

    Dim strCmd As String
    
    ' Don't need the parentheses after the sub name
    strCmd = Replace(strSubName, "()", vbNullString)
    
    ' Make sure we are not trying to run a function with arguments
    If InStr(strCmd, "(") > 0 Then
        MsgBox2 "Unable to Run Command", _
            "Parameters are not supported for this command.", _
            "If you need to use parameters, please create a wrapper sub or function with" & vbCrLf & _
            "no parameters that you can call instead of " & strSubName & ".", vbExclamation
        Exit Sub
    End If
    
    ' Add project name so we can run it from the current datbase
    strCmd = "[" & GetVBProjectForCurrentDB.Name & "]." & strCmd
    
    ' Run the sub
    Application.Run strCmd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DatabaseOpen
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns true if a database (or ADP project) is currently open.
'---------------------------------------------------------------------------------------
'
Public Function DatabaseOpen() As Boolean
    DatabaseOpen = Not (CurrentDb Is Nothing And CurrentProject.Connection Is Nothing)
    'DatabaseOpen = Workspaces(0).Databases.Count > 0   ' Another approach
End Function


'---------------------------------------------------------------------------------------
' Procedure : TableExists
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Returns true if the table object is found in the dabase.
'---------------------------------------------------------------------------------------
'
Public Function TableExists(strName As String) As Boolean
    TableExists = Not (DCount("*", "MSysObjects", "Name=""" & strName & """ AND Type=1") = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteObject
' Author    : Adam Waller
' Date      : 11/23/2020
' Purpose   : Deletes the object if it exists. (Surpresses error)
'---------------------------------------------------------------------------------------
'
Public Sub DeleteObjectIfExists(intType As AcObjectType, strName As String)
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    DoCmd.DeleteObject intType, strName
    Catch 7874 ' Object not found
    CatchAny eelError, "Deleting object " & strName
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbVersion
' Author    : Adam Waller
' Date      : 5/4/2021
' Purpose   : Return the database version as an integer. Works in non-English locales
'           : where CInt(CurrentDb.Version) doesn't work correctly.
'---------------------------------------------------------------------------------------
'
Public Function DbVersion() As Integer
    DbVersion = CInt(Split(CurrentDb.Version, ".")(0))
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormLoaded
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Helps identify if a form has been closed, but is still running code
'           : after the close event.
'---------------------------------------------------------------------------------------
'
Public Function FormLoaded(frmMe As Form) As Boolean
    Dim strName As String
    ' If no forms are open, we already have our answer.  :-)
    If Forms.Count > 0 Then
        ' We will throw an error accessing the name property if the form is closed
        If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
        strName = frmMe.Name
        ' Return true if we were able to read the name property
        FormLoaded = strName <> vbNullString
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyFocus
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Verify that a control currently has the focus. (Is the active control)
'---------------------------------------------------------------------------------------
'
Public Function VerifyFocus(ctlWithFocus As Control) As Boolean

    Dim frmParent As Form
    Dim objParent As Object
    Dim ctlCurrentFocus As Control
    
    ' Determine parent form for control
    Set objParent = ctlWithFocus
    Do While Not TypeOf objParent Is Form
        Set objParent = objParent.Parent
    Loop
    Set frmParent = objParent
    
    ' Ignore any errors with Screen.* functions
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    
    ' Verify focus of parent form
    Set frmParent = Screen.ActiveForm
    If Not frmParent Is objParent Then
        Set frmParent = objParent
        frmParent.SetFocus
        DoEvents
    End If
    
    ' Verify focus of control on form
    Set ctlCurrentFocus = frmParent.ActiveControl
    If Not ctlCurrentFocus Is ctlWithFocus Then
        ctlWithFocus.SetFocus
        DoEvents
    End If
    
    ' Return true if the control currently has the focus
    VerifyFocus = frmParent.ActiveControl Is ctlWithFocus
    
    ' Discard any errors
    CatchAny eelNoError, vbNullString, , False
    
End Function

Public Function ListAllPropertiesJsonInObject( _
    ByRef aobj As Variant _
    , Optional ByRef proplist As Dictionary _
    , Optional ByRef onlylisted As Boolean = False _
    ) As Dictionary

    Dim haveproplist As Boolean

    'initiate dictionery if not set
    haveproplist = Not proplist Is Nothing
    
    Dim oFields()
    
    If haveproplist Then
        ReDim oFields(1 To proplist.Count)
    Else
        Set proplist = New Dictionary
    End If

    'Table Fields
    Dim uFields()
    ReDim uFields(1, 1 To aobj.Properties.Count)
    
    Dim ucounter
    ucounter = 0

    Dim propposition As Long
    Dim var_Value As Variant

    'two lookup system by proplist and by obj propreties
    If onlylisted Then
        Dim key
        
        For Each key In proplist.Keys
    
           propposition = proplist(key)
           
           If propposition > 0 Then
               
                On Error Resume Next
                
                var_Value = aobj.Properties(key).Value
                
                If Err.Number = 0 Then
                    oFields(propposition) = var_Value
                Else
                    Err.Clear
                End If
               
                On Error GoTo 0
    
            End If
    
    
        Next

    Else

       Dim prp As Property

       For Each prp In aobj.Properties
           
          ' Debug.Assert prp.Name <> "Format"
           
           If proplist.Exists(prp.Name) Then
               propposition = proplist(prp.Name)
           Else
               propposition = -2
           End If
           
           If propposition <> -1 Then
               
               On Error Resume Next
               
               var_Value = prp.Value
               
               If Err.Number = 0 Then
               
                   If propposition = -2 Then
                       ucounter = ucounter + 1
                       uFields(0, ucounter) = prp.Name
                       uFields(1, ucounter) = var_Value
                   Else
                       oFields(propposition) = var_Value
                   End If
               
               Else
                   Err.Clear
               End If
               
               On Error GoTo 0
           
           End If
    
       Next
    
    End If
    
   
    Dim rdct As Dictionary
    
    Set rdct = New Dictionary
    
    'sOfields = ""
    If haveproplist Then
       
        For Each key In proplist.Keys
        
            On Error Resume Next
            rdct.Add key, oFields(proplist(key))
            On Error GoTo 0
        
        Next
       
       ' sOfields = Join(oFields, "")
    End If
    
    If Not onlylisted Then
    
        If ucounter > 0 Then
        
            ReDim Preserve uFields(1, 1 To ucounter)
        
            Dim i As Long
        
        
            For i = 1 To UBound(uFields, 2)
            
                rdct.Add uFields(0, i), uFields(1, i)
            
            Next
            
        End If
    
        'sUfields = Join(uFields, "")
    End If
    
    
    Set ListAllPropertiesJsonInObject = rdct
    
End Function

Public Function FieldPropList( _
    Optional ByVal fldType As DataTypeEnum _
    , Optional ByVal linked As Integer = 0 _
    ) As Dictionary
    
    Dim proplist As Dictionary
    Set proplist = New Dictionary

    With proplist

        'Skip properties
        .Add "GUID", -1
        .Add "IMEMode", -1
        .Add "IMESentenceMode", -1
        .Add "ColumnWidth", -1
        .Add "ColumnOrder", -1
        .Add "ColumnHidden", -1
        .Add "SourceField", -1
        .Add "SourceTable", -1
        .Add "CollatingOrder", -1
        .Add "OrdinalPosition", -1
       
        .Add "DataUpdatable", -1 'not relevant for table data
        .Add "Name", -1
        .Add "Type", -1
        
        .Add "Attributes", -1 'commented not needed when XML all info there
        
        'Listed Properties
        .Add "Caption", .Count + 1
        .Add "Description", .Count + 1
        .Add "Size", .Count + 1
        .Add "DefaultValue", .Count + 1
        .Add "Required", .Count + 1
        .Add "AllowZeroLength", .Count + 1
        .Add "ValidationRule", .Count + 1
        .Add "ValidationText", .Count + 1
        .Add "UnicodeCompression", .Count + 1
        .Add "AppendOnly", .Count + 1
        .Add "Format", .Count + 1
        .Add "DecimalPlaces", .Count + 1
        
     
    End With
    
    'specific to file
    Select Case fldType
    Case dbAttachment, 11, 15, 109 '11 - OleObject, 15 - GUID, 109 - complex field
        proplist("ValidationRule") = -1
        proplist("ValidationText") = -1
        proplist("DefaultValue") = -1
        proplist("Required") = -1
    End Select
    
    'not relevant for linked tables
    If linked Then
        proplist("DefaultValue") = -1
        proplist("Required") = -1
        proplist("AllowZeroLength") = -1
        proplist("ValidationRule") = -1
        proplist("ValidationText") = -1
        proplist("AppendOnly") = -1
        proplist("Expression") = -1
    End If
    
    Set FieldPropList = proplist

End Function

Public Function TablePropList( _
    Optional ByVal linked As Integer = 0 _
    ) As Dictionary

    Dim proplist As Dictionary
    Set proplist = New Dictionary

    With proplist
    
        .Add "Name", -1
        .Add "Connect", -1
        .Add "SourceTableName", -1
        .Add "NameMap", -1
        .Add "GUID", -1
        .Add "DateCreated", -1
        .Add "LastUpdated", -1
        .Add "BackTint", -1
        .Add "BackShade", -1
        .Add "ThemeFontIndex", -1
        .Add "AlternateBackThemeColorIndex", -1
        .Add "AlternateBackTint", -1
        .Add "AlternateBackShade", -1
        .Add "DatasheetGridlinesThemeColorIndex", -1
        .Add "DatasheetForeThemeColorIndex", -1
        .Add "Updatable", -1
        .Add "RecordCount", -1
        .Add "LinkChildFields", -1
        .Add "LinkMasterFields", -1
        
        'Currently commented
        .Add "FCMinReadVer", -1
        .Add "FCMinWriteVer", -1
        .Add "FCMinDesignVer", -1
        
        .Add "Attributes", .Count + 1
        .Add "Caption", .Count + 1
        .Add "Description", .Count + 1
        .Add "ValidationRule", .Count + 1
        .Add "ValidationText", .Count + 1
        .Add "PublishToWeb", .Count + 1
        .Add "Orientation", .Count + 1
        .Add "OrderByOn", .Count + 1
        .Add "DefaultView", .Count + 1
        .Add "DisplayViewsOnSharePointSite", .Count + 1
        .Add "TotalsRow", .Count + 1
        .Add "FilterOnLoad", .Count + 1
        .Add "OrderByOnLoad", .Count + 1
        .Add "HideNewField", .Count + 1
        .Add "ReadOnlyWhenDisconnected", .Count + 1

    End With
    
    If linked Then
        proplist("ValidationRule") = -1
        proplist("ValidationText") = -1
    End If
    
    Set TablePropList = proplist

End Function
