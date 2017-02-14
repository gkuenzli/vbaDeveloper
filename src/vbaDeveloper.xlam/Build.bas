Attribute VB_Name = "Build"
'''
' Build instructions:
' 1. Open a new workbook in excel, then open the VB editor (Alt+F11)  and from the menu File->Import, import this file:
'     * src/vbaDeveloper.xlam/Build.bas
' 2. From tools references... add
'     * Microsoft Visual Basic for Applications Extensibility 5.3
'     * Microsoft Scripting Runtime
'     * Microsoft Forms 2.0 Object Library
' 3. Rename the project to 'vbaDeveloper'
' 5. Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'  (In excel 2010: 'Trust access to the vba project object model')
' 6. If using a non-English version of Excel, rename your current workbook into ThisWorkbook (in VB Editor, press F4,
'    then under the local name for Microsoft Excel Objects, select the workbook. Set the property '(Name)' to ThisWorkbook)
' 7. In VB Editor, press F4, then under Microsoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
' 8. In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'
' 9. Close excel. Open excel with a new workbook, then open the just saved vbaDeveloper.xlam
' 10.Let vbaDeveloper import its own code. Put the cursor in the function 'testImport' and press F5
' 11.If necessary rename module 'Build1' to Build. Menu File-->Save vbaDeveloper.xlam
'''

Option Explicit

Private Const SHEET_EXPORT_EXTENSION As String = ".sheet.cls"
Private Const MODULE_EXPORT_EXTENSION As String = ".bas"
Private Const CLASS_EXPORT_EXTENSION As String = ".cls"
Private Const FORM_EXPORT_EXTENSION As String = ".frm"

Private Const PROJECT_NAME_PARAM As String = "%ProjectName%"
Private Const FILE_NAME_PARAM As String = "%FileName%"

Private Const DEFAULT_RELATIVE_SOURCE_PATH As String = "src\" & FILE_NAME_PARAM
Private Const CONFIG_MODULE As String = "VbaDeveloperConfig"

Public Type Configuration
    VBProject As VBProject
    ProjectFolder As String
    ' Attributes found in VbaDeveloperConfiguration
    RelativeSourcePath As String
    ImportAfterOpen As Boolean
    FormatBeforeSave As Boolean
    ExportAfterSave As Boolean
    IgnoredComponents As Collection     ' Value = ComponentNama, Key = ComponentName
    
    ' Data for import job
    Components As Dictionary            ' Key = componentName, Value = componentFilePath
    Sheets As Dictionary                ' Key = componentName, Value = File object
End Type

Private Const IMPORT_DELAY As String = "00:00:03"

'We need to make this variable public such that they can be given as arguments to application.ontime()
Public ImportJob As Configuration


Public Sub testImport()
    Dim proj_name As String
    proj_name = "VbaDeveloper"

    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode vbaProject, True
End Sub


Public Sub testExport()
    Dim proj_name As String
    proj_name = "VbaDeveloper"

    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.exportVbaCode vbaProject
End Sub


' Returns the directory where code is exported to or imported from.
' When createIfNotExists:=True, the directory will be created if it does not exist yet.
' This is desired when we get the directory for exporting.
' When createIfNotExists:=False and the directory does not exist, an empty String is returned.
' This is desired when we get the directory for importing.
'
' Directory names always end with a '\', unless an empty string is returned.
' Usually called with: fullWorkbookPath = wb.FullName or fullWorkbookPath = vbProject.fileName
' if the workbook is new and has never been saved,
' vbProject.fileName will throw an error while wb.FullName will return a name without slashes.
Public Function getSourceDir(Config As Configuration, createIfNotExists As Boolean) As String
    ' First check if the fullWorkbookPath contains a \.
    If InStr(Config.ProjectFolder, "\") = 0 Then
        'In this case it is a new workbook, we skip it
        Exit Function
    End If

    Dim iFso As New Scripting.FileSystemObject
    Dim exportDir As String
    exportDir = JoinPath(Config.ProjectFolder, Config.RelativeSourcePath)

    If createIfNotExists Then
        exportDir = CreateSubFolderPath(Config.ProjectFolder, Config.RelativeSourcePath)
    Else
        If Not iFso.FolderExists(exportDir) Then
            Debug.Print "Folder does not exist: " & exportDir
            exportDir = ""
        End If
    End If
    getSourceDir = exportDir
End Function

''
' If not already exist, create a complete folder sub path, starting at a root path
'
' @param {String} RootPath - must exist
' @param {String} SubPath - created if necessary
' @return {String} resulting path (if successfull)
''
Public Function CreateSubFolderPath(ByVal RootPath As String, ByVal SubPath As String) As String
    Dim iFso As New Scripting.FileSystemObject
    Dim iParts() As String
    Dim i As Long
    Dim s As String
    iParts = Split(SubPath, "\")
    s = RootPath
    For i = LBound(iParts) To UBound(iParts)
        s = JoinPath(s, iParts(i))
        If Not iFso.FolderExists(s) Then
            Debug.Print "Build.CreateSubFolderPath : creating folder " + s
            On Error Resume Next
            Call iFso.CreateFolder(s)
            On Error GoTo 0
        End If
    Next i
    If iFso.FolderExists(s) Then
        CreateSubFolderPath = s
    End If
End Function

''
' Join Path with \
'
' @example
' ```VB.net
' Debug.Print JoinPath("a\", "\b")
' Debug.Print JoinPath("a", "b")
' Debug.Print JoinPath("a\", "b")
' Debug.Print JoinPath("a", "\b")
' -> a\b
' ```
'
' @param {String} LeftSide
' @param {String} RightSide
' @return {String} Joined path
''
Public Function JoinPath(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "\" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "\" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If

    If LeftSide <> "" And RightSide <> "" Then
        JoinPath = LeftSide & "\" & RightSide
    Else
        JoinPath = LeftSide & RightSide
    End If
End Function




' Usually called after the given workbook is saved
Public Sub exportVbaCode(vbaProject As VBProject)
    Dim iCnf As Configuration
    GetConfiguration vbaProject, iCnf
    If iCnf.ProjectFolder = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(iCnf, createIfNotExists:=True)

    Debug.Print "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In vbaProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        If hasCodeToExport(component) And Not IgnoreComponent(iCnf, component.name) Then
            'Debug.Print "exporting type is " & component.Type
            Select Case component.Type
                Case vbext_ct_ClassModule
                    exportComponent export_path, component
                Case vbext_ct_StdModule
                    exportComponent export_path, component, MODULE_EXPORT_EXTENSION
                Case vbext_ct_MSForm
                    BuildForm.exportMSForm export_path, component
                Case vbext_ct_Document
                    exportLines export_path, component
                Case Else
                    'Raise "Unkown component type"
            End Select
        End If
    Next component
End Sub


Private Function hasCodeToExport(component As VBComponent) As Boolean
    hasCodeToExport = True
    If component.codeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.codeModule.lines(1, 1))
        'Debug.Print firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function


'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = CLASS_EXPORT_EXTENSION)
    Debug.Print "exporting " & component.name & extension
    component.Export JoinPath(exportPath, component.name & extension)
End Sub


'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = SHEET_EXPORT_EXTENSION
    Dim fileName As String
    fileName = JoinPath(exportPath, component.name & extension)
    Debug.Print "exporting " & component.name & extension
    'component.Export exportPath & "\" & component.name & extension
    Dim FSO As New Scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = FSO.CreateTextFile(fileName, True, False)
    outStream.Write (component.codeModule.lines(1, component.codeModule.CountOfLines))
    outStream.Close
End Sub


' Usually called after the given workbook is opened.
' The option includeClassFiles is True by default providing that git repo is correctly handling line endings as crlf (Windows-style) instead of lf (Unix-style)
Public Sub importVbaCode(vbaProject As VBProject, Optional includeClassFiles As Boolean = True)
    Dim iCnf As Configuration
    GetConfiguration vbaProject, iCnf
    Dim vbProjectFileName As String
    If iCnf.ProjectFolder = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim import_path As String
    import_path = getSourceDir(iCnf, createIfNotExists:=False)
    If import_path = "" Then
        'The source directory does not exist, code has never been exported for this vbaProject.
        Debug.Print "No import directory for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    'initialize globals for Application.OnTime
    ImportJob = iCnf

    Dim FSO As New Scripting.FileSystemObject
    Dim projContents As Folder
    Set projContents = FSO.GetFolder(import_path)
    Dim file As Scripting.file
    For Each file In projContents.Files()
        'check if and how to import the file
        checkHowToImport file, includeClassFiles
    Next

    Dim ComponentName As String
    Dim vComponentName As Variant
    'Remove all the modules and class modules
    For Each vComponentName In ImportJob.Components.Keys
        ComponentName = vComponentName
        removeComponent vbaProject, ComponentName
    Next
    'Then import them
    Debug.Print "Invoking 'Build.importComponents'with Application.Ontime with delay " & IMPORT_DELAY
    ' to prevent duplicate modules, like MyClass1 etc.
    Application.OnTime Now() + TimeValue(IMPORT_DELAY), "'Build.importComponents'"
    Debug.Print "almost finished importing code for " & vbaProject.name
End Sub


Private Sub checkHowToImport(file As Scripting.file, includeClassFiles As Boolean)
    Dim fileName As String
    fileName = file.name
    Dim ComponentName As String
    ComponentName = Left(fileName, InStr(fileName, ".") - 1)
    If ComponentName = "Build" And GetFileName(ImportJob) = ThisWorkbook.name Then
        '"don't remove or import ourself
        Exit Sub
    End If
    If IgnoreComponent(ImportJob, ComponentName) Then
        ' don't import ignored component
        Exit Sub
    End If

    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = Right(fileName, 4)
        Select Case lastPart
            Case CLASS_EXPORT_EXTENSION ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And LCase(Right(fileName, 10)) = SHEET_EXPORT_EXTENSION Then
                    'import lines into sheet: importLines ImportJob.VBProject, file
                    ImportJob.Sheets.Add ComponentName, file
                Else
                    ' .cls files don't import correctly because of a bug in excel, therefore we can exclude them.
                    ' In that case they'll have to be imported manually.
                    If includeClassFiles Then
                        'importComponent vbaProject, file
                        ImportJob.Components.Add ComponentName, file.Path
                    End If
                End If
            Case MODULE_EXPORT_EXTENSION, FORM_EXPORT_EXTENSION
                'importComponent vbaProject, file
                ImportJob.Components.Add ComponentName, file.Path
            Case Else
                'do nothing
                Debug.Print "Skipping file " & fileName
        End Select
    End If
End Sub


' Only removes the vba component if it exists
Private Sub removeComponent(vbaProject As VBProject, ComponentName As String)
    If componentExists(vbaProject, ComponentName) Then
        Dim c As VBComponent
        Set c = vbaProject.VBComponents(ComponentName)
        Debug.Print "removing " & c.name
        vbaProject.VBComponents.Remove c
    End If
End Sub


Public Sub importComponents()
    If ImportJob.Components Is Nothing Then
        Debug.Print "Failed to import! Dictionary 'ImportJob.Components' was not initialized."
        Exit Sub
    End If
    Dim ComponentName As String
    Dim vComponentName As Variant
    For Each vComponentName In ImportJob.Components.Keys
        ComponentName = vComponentName
        importComponent ImportJob.VBProject, ImportJob.Components(ComponentName)
    Next

    'Import the sheets
    For Each vComponentName In ImportJob.Sheets.Keys
        ComponentName = vComponentName
        importLines ImportJob.VBProject, ImportJob.Sheets(ComponentName)
    Next

    Debug.Print "Finished importing code for " & ImportJob.VBProject.name
    'We're done, clear globals explicitly to free memory.
    Dim iCnf As Configuration
    ImportJob = iCnf
'    Set ImportJob.Components = Nothing
'    Set ImportJob.VBProject = Nothing
'    Set ImportJob.Sheets = Nothing
End Sub


' Assumes any component with same name has already been removed.
Private Sub importComponent(vbaProject As VBProject, filePath As String)
    Debug.Print "Importing component from  " & filePath
    Dim newComp As VBComponent
    Set newComp = vbaProject.VBComponents.Import(filePath)
    Do While Trim(newComp.codeModule.lines(1, 1)) = "" And newComp.codeModule.CountOfLines > 1
        newComp.codeModule.DeleteLines 1
    Loop
End Sub


Private Sub importLines(vbaProject As VBProject, file As Object)
    Dim ComponentName As String
    ComponentName = Left(file.name, InStr(file.name, ".") - 1)
    Dim c As VBComponent
    If Not componentExists(vbaProject, ComponentName) Then
        ' Create a sheet to import this code into. We cannot set the ws.codeName property which is read-only,
        ' instead we set its vbComponent.name which leads to the same result.
        Dim addedSheetCodeName As String
        addedSheetCodeName = addSheetToWorkbook(ComponentName, vbaProject.fileName)
        Set c = vbaProject.VBComponents(addedSheetCodeName)
        c.name = ComponentName
    End If
    Set c = vbaProject.VBComponents(ComponentName)
    Debug.Print "Importing lines from " & ComponentName & " into component " & c.name

    ' At this point compilation errors may cause a crash, so we ignore those.
    On Error Resume Next
    c.codeModule.DeleteLines 1, c.codeModule.CountOfLines
    c.codeModule.AddFromFile file.Path
    On Error GoTo 0
End Sub


Public Function componentExists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    componentExists = True
    Exit Function
doesnt:
    componentExists = False
End Function


' Returns a reference to the workbook. Opens it if it is not already opened.
' Raises error if the file cannot be found.
Public Function openWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath) 'can raise error
    End If
    Set openWorkbook = wb
End Function


' Returns the CodeName of the added sheet or an empty String if the workbook could not be opened.
Public Function addSheetToWorkbook(sheetName As String, workbookFilePath As String) As String
    Dim wb As Workbook
    On Error Resume Next 'can throw if given path does not exist
    Set wb = openWorkbook(workbookFilePath)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Dim ws As Worksheet
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = sheetName
        'ws.CodeName = sheetName: cannot assign to read only property
        Debug.Print "Sheet added " & sheetName
        addSheetToWorkbook = ws.CodeName
    Else
        Debug.Print "Skipping file " & sheetName & ". Could not open workbook " & workbookFilePath
        addSheetToWorkbook = ""
    End If
End Function



' ============================================= '
' Configuration Management
' ============================================= '

' Implemtation details
' --------------------
' Normally we would go for a Configuration Class, not a Type
' However, the very nature of this project is to export/import components
' and with a particularity to rebuild itself from this Build.bas module alone.
'
' So we have no other option than the good old procedural way. So be it !



''
' Prepare a Configuration for a given VBProject
'
' @method GetConfiguration
' @param {VBProject} Project
' @param out {Configuration} Config
''
Public Sub GetConfiguration(ByVal project As VBProject, ByRef Config As Configuration)
    Set Config.VBProject = project
    Dim iPath As String
    'this can throw if the workbook has never been saved.
    On Error Resume Next
    iPath = project.fileName
    On Error GoTo 0
    GetConfigurationForPath iPath, Config
End Sub

Public Sub GetConfigurationWB(ByVal Workbook As Workbook, ByRef Config As Configuration)
    Set Config.VBProject = Workbook.VBProject
    GetConfigurationForPath Workbook.FullName, Config
End Sub

Private Sub GetConfigurationForPath(ByVal ProjectFileName, ByRef Config As Configuration)
    With Config
        Dim iFso As New Scripting.FileSystemObject
        .ProjectFolder = iFso.GetParentFolderName(ProjectFileName)
        Dim iPath As String
        iPath = ReplaceParams(Config, DEFAULT_RELATIVE_SOURCE_PATH)
        .RelativeSourcePath = iPath
        .ImportAfterOpen = True
        .FormatBeforeSave = True
        .ExportAfterSave = True
        Set .IgnoredComponents = New Collection
        Set .Components = New Dictionary
        Set .Sheets = New Dictionary
    End With
    ExtractAttributes Config
End Sub


' --------------------------------------------- '
' Private Functions
' --------------------------------------------- '

Private Function GetConfigModule(Config As Configuration) As codeModule
    On Error Resume Next
    Set GetConfigModule = Config.VBProject.VBComponents(CONFIG_MODULE).codeModule
End Function

Private Function GetProjectName(Config As Configuration) As String
    If Not Config.VBProject Is Nothing Then
        GetProjectName = Config.VBProject.name
    End If
End Function

Private Function GetFileName(Config As Configuration) As String
    If Not Config.VBProject Is Nothing Then
        Dim iFso As New FileSystemObject
        GetFileName = iFso.GetFileName(Config.VBProject.fileName)
    End If
End Function

Private Sub ExtractAttributes(ByRef Config As Configuration)
    Dim src As codeModule
    Set src = GetConfigModule(Config)
    If Not src Is Nothing Then
        ParseModule Config, src
    End If
End Sub

Private Function IgnoreComponent(ByRef Config As Configuration, ByVal ComponentName As String) As Boolean
    Dim s As String
    On Error Resume Next
    s = Config.IgnoredComponents(LCase(ComponentName))
    IgnoreComponent = (s <> "")
End Function

Private Sub ParseModule(ByRef Config As Configuration, ByVal SourceCode As codeModule)
    Dim i As Long
    For i = 1 To SourceCode.CountOfLines
        ParseLine Config, i, SourceCode.lines(i, 1)
    Next i
End Sub

Private Sub ParseLine(ByRef Config As Configuration, ByVal Index As Long, ByVal line As String)
    line = Trim(line)
    
    ' Code line must start with as comment mark
    If Left(line, 1) <> "'" Then
        Exit Sub
    End If
    line = Trim(Mid(line, 2))
    
    ' Line must contain a "=" after first char
    Dim iPos As Long
    iPos = InStr(line, "=")
    If iPos < 2 Then
        Exit Sub
    End If
    
    Dim iName As String
    Dim iVal As String
    iName = Trim(Left(line, iPos - 1))
    iVal = Trim(Mid(line, iPos + 1))
    
    ' Value may be surrounded by quotation marks
    If Left(iVal, 1) = """" Then
        If Right(iVal, 1) <> """" Or Len(iVal) = 1 Then
            Debug.Print "VbaDeveloperConfig @ line #" & Index & " :  has invalid value format, closing quotation mark expected"
            Exit Sub
        End If
        iVal = Mid(iVal, 2, Len(iVal) - 2)
    End If
    
    ' Resolve params
    iVal = ReplaceParams(Config, iVal)
    
    ' Assign value to attribute
    With Config
        Select Case LCase(iName)
            Case "exportaftersave"
                .ExportAfterSave = GetBool(Index, iVal)
            Case "formatbeforesave"
                .FormatBeforeSave = GetBool(Index, iVal)
            Case "importafteropen"
                .ImportAfterOpen = GetBool(Index, iVal)
            Case "relativesourcepath"
                .RelativeSourcePath = iVal
            Case "ignore"
                AddIgnoredComponent Config, iVal
            Case Else
                Debug.Print "VbaDeveloperConfig @ line #" & Index & " : unknown attribute name """ & iName & """"
        End Select
    End With
End Sub

Private Function ReplaceParams(ByRef Config As Configuration, ByVal text As String) As String
    ReplaceParams = Replace(text, PROJECT_NAME_PARAM, GetProjectName(Config), , , vbTextCompare)
    ReplaceParams = Replace(ReplaceParams, FILE_NAME_PARAM, GetFileName(Config), , , vbTextCompare)
End Function

Private Sub AddIgnoredComponent(ByRef Config As Configuration, ByVal ComponentName As String)
    On Error Resume Next
    Config.IgnoredComponents.Add ComponentName, LCase(ComponentName)
End Sub

Private Function GetBool(ByVal Index As Long, ByVal text As String) As Boolean
    On Error GoTo GetBool_Err
    GetBool = CBool(text)
    Exit Function
GetBool_Err:
    Debug.Print "VbaDeveloperConfig @ line #" & Index & " : invalid value, boolean expected """ & text & """"
End Function
