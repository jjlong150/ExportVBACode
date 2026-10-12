Attribute VB_Name = "modExport"
' MIT License

' Copyright (c) 2025-2026 Jeffrey Long

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit

' ================================================================
' Settings
' ================================================================

' Skip sheet/ThisWorkbook modules that contain nothing but blank lines and
' "Option ..." statements, so readers aren't shown a stack of empty Sheet files.
Private Const SKIP_EMPTY_DOCUMENT_MODULES As Boolean = True

' Unzip the workbook package into src\<ext>. Set to False to skip it entirely.
Private Const UNZIP_PACKAGE As Boolean = True

' Files and folders to remove from the unzipped package, separated by ";".
'   - An entry without a backslash matches that name at any depth ("_rels").
'   - An entry with a backslash is a path from the package root ("xl\media").
'   - Wildcards * and ? are allowed. Matching is case-insensitive.
'   - A literal "[" must be written as "[[]", e.g. "[[]Content_Types].xml".
Private Const PACKAGE_EXCLUDES As String = "xl;_rels;docProps;[[]Content_Types].xml"

' Pretty-printing is switched on and off by the cell named in PRETTY_PRINT_TOGGLE_NAME
' on ExportSheet (TRUE/FALSE). If that name doesn't exist, pretty-printing is off.
Private Const PRETTY_PRINT_TOGGLE_NAME As String = "PRETTY_PRINT_XML"

' XML files in the unzipped package to write a pretty-printed copy of, separated
' by ";". Same matching rules as PACKAGE_EXCLUDES. The original is left untouched
' and the copy is written beside it with PRETTY_PRINT_SUFFIX before the extension,
' e.g. customUI14.xml -> customUI14.pretty.xml. Use "" to turn this off.
Private Const PRETTY_PRINT_FILES As String = "customUI\*.xml"
Private Const PRETTY_PRINT_SUFFIX As String = ".pretty"

' ================================================================
' Export folder names (no spaces, so wiki/doc generators link them cleanly)
' ================================================================

Private Const DIR_MODULES As String = "Modules"
Private Const DIR_CLASS_MODULES As String = "ClassModules"
Private Const DIR_FORMS As String = "Forms"
Private Const DIR_ACTIVEX As String = "ActiveX"
Private Const DIR_EXCEL_OBJECTS As String = "MicrosoftExcelObjects"

' How long the Shell.Application fallback unzip may take before giving up.
Private Const UNZIP_TIMEOUT_SECONDS As Long = 60

' ================================================================
' Constants
' ================================================================

' Component types as returned by VBComponent.Type.
' Named so they don't collide with VBIDE.VBComponent / vbext_ct_* when the
' "Microsoft Visual Basic for Applications Extensibility" reference is set.
' https://docs.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#type
Private Enum ComponentKind
    ckStdModule = 1
    ckClassModule = 2
    ckMSForm = 3
    ckActiveXDesigner = 11
    ckDocument = 100
End Enum

Private Const VBPROJECT_LOCKED As Long = 1     ' vbext_pp_locked
Private Const ERR_EXPORT As Long = vbObjectError + 513

' ================================================================
' Entry points
' ================================================================

Public Sub ExportVBA()
    #If Mac Then
        MsgBox "This exporter relies on Windows-only components (FileSystemObject, WScript.Shell) and cannot run on a Mac.", vbExclamation, "Export VBA"
        Exit Sub
    #End If

    Dim filePath As String
    Dim exportFolder As String
    filePath = ReadSetting("FILENAME")
    exportFolder = ReadSetting("EXPORT_FOLDER")

    If filePath = vbNullString Or exportFolder = vbNullString Then
        MsgBox "Choose a workbook and an export folder first.", vbExclamation, "Export VBA"
        Exit Sub
    End If

    If Not FileExists(filePath) Then
        MsgBox "The workbook could not be found:" & vbNewLine & filePath, vbExclamation, "Export VBA"
        Exit Sub
    End If

    If Not DirectoryExists(exportFolder) Then
        MsgBox "The export folder does not exist:" & vbNewLine & exportFolder, vbExclamation, "Export VBA"
        Exit Sub
    End If

    If Not IsVBProjectAccessTrusted() Then
        MsgBox "Excel is blocking programmatic access to VBA projects." & vbNewLine & vbNewLine & _
               "Turn on File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
               """Trust access to the VBA project object model"", then run the export again.", _
               vbExclamation, "Export VBA"
        Exit Sub
    End If

    Dim summary As String
    On Error GoTo ExportFailed
    summary = ExportComponents(filePath, exportFolder)
    Application.StatusBar = False
    MsgBox summary, vbInformation, "Export complete"
    Exit Sub

ExportFailed:
    Application.StatusBar = False
    MsgBox "Export failed:" & vbNewLine & vbNewLine & Err.Description & _
           vbNewLine & vbNewLine & "(Error " & Err.Number & IIf(Err.Source <> vbNullString, " in " & Err.Source, vbNullString) & ")", _
           vbCritical, "Export VBA"
End Sub

Public Sub GetFileName()
    Dim picked As String
    picked = ChooseExcelFile(ReadSetting("FILENAME"))
    ' Keep the current value if the user cancelled
    If picked <> vbNullString Then ExportSheet.Range("FILENAME").value = picked
End Sub

Public Sub ClearFileName()
    ExportSheet.Range("FILENAME").ClearContents
End Sub

Public Sub GetExportFolder()
    ExportSheet.Range("EXPORT_FOLDER").value = ChooseDirectory(ReadSetting("EXPORT_FOLDER"))
End Sub

Public Sub ClearExportFolder()
    ExportSheet.Range("EXPORT_FOLDER").ClearContents
End Sub

' ================================================================
' Export pipeline
' ================================================================

Private Function ExportComponents(ByVal filePathIn As String, ByVal exportFolder As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    filePathIn = fso.GetAbsolutePathName(filePathIn)

    Dim fileName As String
    Dim baseName As String
    Dim fileExt As String
    fileName = fso.GetFileName(filePathIn)
    baseName = fso.GetBaseName(filePathIn)
    fileExt = LCase$(fso.GetExtensionName(filePathIn))

    ' ----------------------------------------------------------------
    ' Directory layout
    '   <exportFolder>\<baseName>\src\excel\<workbook copy>
    '   <exportFolder>\<baseName>\src\vba\<component folders>
    '   <exportFolder>\<baseName>\src\<ext>\<unzipped package>
    ' Only the generated "src" tree is wiped and rebuilt. Anything else in
    ' <baseName> (a .git folder, a README, etc.) is left untouched.
    ' ----------------------------------------------------------------
    Dim exportDir As String
    Dim srcDir As String
    Dim workbookDir As String
    Dim vbaDir As String
    Dim packageDir As String

    exportDir = JoinPath(exportFolder, baseName)
    srcDir = JoinPath(exportDir, "src")
    workbookDir = JoinPath(srcDir, "excel")
    vbaDir = JoinPath(srcDir, "vba")
    packageDir = JoinPath(srcDir, fileExt)

    SetStatus "Clearing " & srcDir
    DeleteFolderIfExists srcDir

    EnsureFolder workbookDir
    EnsureFolder JoinPath(vbaDir, DIR_MODULES)
    EnsureFolder JoinPath(vbaDir, DIR_CLASS_MODULES)
    EnsureFolder JoinPath(vbaDir, DIR_FORMS)
    EnsureFolder JoinPath(vbaDir, DIR_ACTIVEX)
    EnsureFolder JoinPath(vbaDir, DIR_EXCEL_OBJECTS)

    ' ----------------------------------------------------------------
    ' Step 1 - Copy the workbook
    ' ----------------------------------------------------------------
    Dim workbookCopy As String
    workbookCopy = JoinPath(workbookDir, fileName)
    SetStatus "Copying " & filePathIn
    fso.CopyFile filePathIn, workbookCopy, True

    ' ----------------------------------------------------------------
    ' Step 2 - Export the VBA components
    ' ----------------------------------------------------------------
    Dim exportedCount As Long
    Dim skippedCount As Long
    Dim warning As String
    exportedCount = ExportVbaProject(filePathIn, vbaDir, skippedCount, warning)

    ' ----------------------------------------------------------------
    ' Step 3 - Unzip the package (Open XML formats only)
    ' ----------------------------------------------------------------
    Dim unzipNote As String
    Dim removedCount As Long
    If Not UNZIP_PACKAGE Then
        unzipNote = "Unzip skipped (UNZIP_PACKAGE is off)."
    ElseIf Not IsZipPackage(fileExt) Then
        unzipNote = "Unzip skipped: ." & fileExt & " is not a zip-based file format."
    Else
        EnsureFolder packageDir
        SetStatus "Unzipping package contents to " & packageDir
        UnzipPackage workbookCopy, packageDir

        SetStatus "Removing excluded package files"
        removedCount = PrunePackage(packageDir)

        Dim prettyCount As Long
        Dim prettyFailures As String
        If PRETTY_PRINT_FILES <> vbNullString And ReadToggle(PRETTY_PRINT_TOGGLE_NAME, False) Then
            SetStatus "Pretty-printing package XML"
            prettyCount = PrettyPrintPackageXml(packageDir, prettyFailures)
        End If

        If CreateObject("Scripting.FileSystemObject").GetFolder(packageDir).Files.Count = 0 And _
           CreateObject("Scripting.FileSystemObject").GetFolder(packageDir).SubFolders.Count = 0 Then
            DeleteFolderIfExists packageDir
            unzipNote = "Package unzipped, but every item matched PACKAGE_EXCLUDES, so src\" & fileExt & " was removed."
        Else
            unzipNote = "Package contents unzipped to src\" & fileExt & _
                        IIf(removedCount > 0, " (" & removedCount & " excluded item(s) removed).", ".")
            If prettyCount > 0 Then
                unzipNote = unzipNote & vbNewLine & prettyCount & " XML file(s) pretty-printed to *" & PRETTY_PRINT_SUFFIX & ".xml copies."
            End If
        End If

        If prettyFailures <> vbNullString Then
            If warning <> vbNullString Then warning = warning & vbNewLine & vbNewLine
            warning = warning & "These XML files could not be pretty-printed (originals are unchanged):" & prettyFailures
        End If
    End If

    ' ----------------------------------------------------------------
    ' Summary
    ' ----------------------------------------------------------------
    ' ----------------------------------------------------------------
    ' Step 4 - Remove empty folders (e.g. ActiveX when there are no designers)
    ' ----------------------------------------------------------------
    Dim emptyFolderCount As Long
    SetStatus "Removing empty folders"
    emptyFolderCount = RemoveEmptyFolders(srcDir)

    Dim summary As String
    summary = "Exported " & exportedCount & " component(s) from" & vbNewLine & _
              filePathIn & vbNewLine & "to" & vbNewLine & exportDir
    If skippedCount > 0 Then
        summary = summary & vbNewLine & vbNewLine & "Skipped " & skippedCount & " empty or unrecognized component(s)."
    End If
    summary = summary & vbNewLine & vbNewLine & unzipNote
    If emptyFolderCount > 0 Then
        summary = summary & vbNewLine & "Removed " & emptyFolderCount & " empty folder(s)."
    End If
    If warning <> vbNullString Then
        summary = summary & vbNewLine & vbNewLine & "Warning: " & warning
    End If

    ExportComponents = summary
End Function

' Exports every component of the workbook's VBA project into vbaDir.
' Returns the number exported. Restores Excel state and closes the workbook
' (if this routine opened it) whether or not an error occurs.
Private Function ExportVbaProject(ByVal filePathIn As String, ByVal vbaDir As String, _
                                  ByRef skippedCount As Long, ByRef warning As String) As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents

    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim exportedCount As Long

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    Dim fileName As String
    fileName = CreateObject("Scripting.FileSystemObject").GetFileName(filePathIn)

    ' Excel can't hold two workbooks with the same name, and Workbooks.Open on a
    ' file that is already open just returns that instance - closing it afterwards
    ' would close the user's workbook (or this tool itself). Reuse it instead.
    Set wb = GetOpenWorkbookByName(fileName)
    If Not wb Is Nothing Then
        If Not IsSameWorkbookPath(wb.FullName, filePathIn) Then
            Err.Raise ERR_EXPORT, "ExportVbaProject", _
                      "A different workbook named """ & fileName & """ is already open:" & vbNewLine & _
                      wb.FullName & vbNewLine & "Close it and run the export again."
        End If
        If Not wb.Saved Then
            warning = fileName & " is open with unsaved changes. The VBA was exported from the open copy, " & _
                      "but the workbook copy and unzipped package came from the last saved version on disk."
        End If
    Else
        Application.ScreenUpdating = False
        ' Prevent the target's Workbook_Open / Workbook_BeforeClose from running
        Application.EnableEvents = False
        SetStatus "Opening " & filePathIn
        Set wb = Workbooks.Open(fileName:=filePathIn, UpdateLinks:=0, ReadOnly:=True, _
                                IgnoreReadOnlyRecommended:=True, AddToMru:=False)
        openedHere = True
        HideWindows wb
        ThisWorkbook.Activate
    End If

    Dim project As Object
    Set project = wb.VBProject
    If project.Protection = VBPROJECT_LOCKED Then
        Err.Raise ERR_EXPORT, "ExportVbaProject", _
                  "The VBA project in " & fileName & " is password-protected." & vbNewLine & _
                  "Open the workbook, unlock the project in the VBA editor, and run the export again while it is open."
    End If

    Dim component As Object
    Dim subdir As String
    Dim ext As String
    Dim filePathOut As String

    For Each component In project.VBComponents
        If Not TryGetExportTarget(component.Type, subdir, ext) Then
            skippedCount = skippedCount + 1
        ElseIf SKIP_EMPTY_DOCUMENT_MODULES And component.Type = ckDocument And IsEffectivelyEmpty(component.codeModule) Then
            skippedCount = skippedCount + 1
        Else
            filePathOut = JoinPath(JoinPath(vbaDir, subdir), component.Name & "." & ext)
            SetStatus "Exporting " & component.Name
            component.Export filePathOut
            exportedCount = exportedCount + 1
        End If
    Next component

    ExportVbaProject = exportedCount

SafeExit:
    On Error Resume Next
    Set component = Nothing
    Set project = Nothing
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Function

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume SafeExit
End Function

Private Function TryGetExportTarget(ByVal componentType As Long, ByRef subdir As String, ByRef ext As String) As Boolean
    TryGetExportTarget = True
    Select Case componentType
        Case ckStdModule:       subdir = DIR_MODULES:       ext = "bas"
        Case ckClassModule:     subdir = DIR_CLASS_MODULES: ext = "cls"
        Case ckMSForm:          subdir = DIR_FORMS:         ext = "frm"
        Case ckActiveXDesigner: subdir = DIR_ACTIVEX:       ext = "dsr"
        Case ckDocument:        subdir = DIR_EXCEL_OBJECTS: ext = "cls"
        Case Else
            subdir = vbNullString
            ext = vbNullString
            TryGetExportTarget = False
    End Select
End Function

' True when a code module holds nothing but blank lines and Option statements.
Private Function IsEffectivelyEmpty(ByVal codeModule As Object) As Boolean
    Dim i As Long
    Dim lineText As String
    For i = 1 To codeModule.CountOfLines
        lineText = Trim$(codeModule.Lines(i, 1))
        If lineText <> vbNullString And Not (LCase$(lineText) Like "option *") Then Exit Function
    Next i
    IsEffectivelyEmpty = True
End Function

' ================================================================
' Unzipping
' ================================================================

Private Function IsZipPackage(ByVal fileExt As String) As Boolean
    Select Case LCase$(fileExt)
        Case "xlsm", "xlsx", "xlsb", "xlam", "xltm", "xltx"
            IsZipPackage = True
    End Select
End Function

Private Sub UnzipPackage(ByVal packagePath As String, ByVal destDir As String)
    ' tar.exe (built into Windows 10 1803+) extracts synchronously and reads the
    ' package directly, so no renaming is needed. Fall back to the shell otherwise.
    If TryUnzipWithTar(packagePath, destDir) Then Exit Sub
    UnzipWithShell packagePath, destDir
End Sub

Private Function TryUnzipWithTar(ByVal packagePath As String, ByVal destDir As String) As Boolean
    Dim tarPath As String
    tarPath = Environ$("SystemRoot") & "\System32\tar.exe"
    If Not FileExists(tarPath) Then Exit Function

    Dim command As String
    command = """" & tarPath & """ -xf """ & packagePath & """ -C """ & destDir & """"

    On Error GoTo Failed
    Dim exitCode As Long
    exitCode = CreateObject("WScript.Shell").Run(command, 0, True)   ' hidden window, wait for exit
    TryUnzipWithTar = (exitCode = 0)
    Exit Function

Failed:
    TryUnzipWithTar = False
End Function

Private Sub UnzipWithShell(ByVal packagePath As String, ByVal destDir As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Shell.Application only treats files with a .zip extension as archives, so
    ' extract from a temporary copy rather than renaming the exported workbook.
    Dim tempZip As String
    tempZip = JoinPath(fso.GetSpecialFolder(2).Path, fso.GetTempName() & ".zip")
    fso.CopyFile packagePath, tempZip, True

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    Dim shellApp As Object
    Dim sourceNs As Object
    Dim destNs As Object
    Set shellApp = CreateObject("Shell.Application")
    ' Namespace() silently returns Nothing when passed a String variable; it must be a Variant.
    Set sourceNs = shellApp.Namespace(CVar(tempZip))
    Set destNs = shellApp.Namespace(CVar(destDir))
    If sourceNs Is Nothing Or destNs Is Nothing Then
        Err.Raise ERR_EXPORT, "UnzipWithShell", "Windows could not open " & packagePath & " as a zip archive."
    End If

    Dim expectedCount As Long
    expectedCount = sourceNs.Items.Count

    ' 4 = no progress dialog, 16 = yes to all, 1024 = no error UI
    destNs.CopyHere sourceNs.Items, 4 + 16 + 1024

    ' CopyHere returns immediately and copies in the background. Wait for all
    ' top-level items to appear, then for the folder size to stop changing.
    Dim deadline As Date
    deadline = Now + TimeSerial(0, 0, UNZIP_TIMEOUT_SECONDS)

    Do While destNs.Items.Count < expectedCount
        WaitOneSecond deadline
    Loop

    Dim lastSize As Double
    Dim currentSize As Double
    lastSize = -1
    Do
        currentSize = fso.GetFolder(destDir).Size
        If currentSize = lastSize Then Exit Do
        lastSize = currentSize
        WaitOneSecond deadline
    Loop

SafeExit:
    On Error Resume Next
    Set sourceNs = Nothing
    Set destNs = Nothing
    Set shellApp = Nothing
    If fso.FileExists(tempZip) Then fso.DeleteFile tempZip, True
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume SafeExit
End Sub

' Removes everything in the unzipped package that matches PACKAGE_EXCLUDES.
' Returns the number of files and folders removed (a removed folder counts once).
Private Function PrunePackage(ByVal packageDir As String) As Long
    Dim patterns() As String
    patterns = Split(PACKAGE_EXCLUDES, ";")
    If UBound(patterns) < 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim rootPath As String
    rootPath = fso.GetFolder(packageDir).Path
    PrunePackage = PruneFolder(rootPath, rootPath, patterns, fso)
End Function

Private Function PruneFolder(ByVal folderPath As String, ByVal rootPath As String, _
                             ByRef patterns() As String, ByVal fso As Object) As Long
    Dim removed As Long
    Dim doomed As Collection
    Dim item As Object
    Dim itemPath As Variant

    ' Collect matches first and delete afterwards: changing a Files or
    ' SubFolders collection while enumerating it is unreliable.
    Set doomed = New Collection
    For Each item In fso.GetFolder(folderPath).Files
        If MatchesAnyPattern(item.Path, item.Name, rootPath, patterns) Then doomed.Add item.Path
    Next item
    For Each itemPath In doomed
        fso.DeleteFile itemPath, True
        removed = removed + 1
    Next itemPath

    Set doomed = New Collection
    For Each item In fso.GetFolder(folderPath).SubFolders
        If MatchesAnyPattern(item.Path, item.Name, rootPath, patterns) Then doomed.Add item.Path
    Next item
    For Each itemPath In doomed
        fso.DeleteFolder itemPath, True
        removed = removed + 1
    Next itemPath

    ' Recurse into whatever survived
    For Each item In fso.GetFolder(folderPath).SubFolders
        removed = removed + PruneFolder(item.Path, rootPath, patterns, fso)
    Next item

    PruneFolder = removed
End Function

' ================================================================
' Pretty-printing
' ================================================================

' Writes a formatted copy of every package file matching PRETTY_PRINT_FILES.
' Returns the number written; failures are listed in `failures` and don't stop the export.
Private Function PrettyPrintPackageXml(ByVal packageDir As String, ByRef failures As String) As Long
    Dim patterns() As String
    patterns = Split(PRETTY_PRINT_FILES, ";")
    If UBound(patterns) < 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim rootPath As String
    rootPath = fso.GetFolder(packageDir).Path

    ' Collect first so the copies we write aren't picked up as sources
    Dim sources As Collection
    Set sources = New Collection
    CollectMatchingFiles rootPath, rootPath, patterns, fso, sources

    Dim sourcePath As Variant
    Dim destPath As String
    Dim written As Long
    For Each sourcePath In sources
        destPath = JoinPath(fso.GetParentFolderName(sourcePath), _
                            fso.GetBaseName(sourcePath) & PRETTY_PRINT_SUFFIX & "." & fso.GetExtensionName(sourcePath))
        On Error Resume Next
        PrettyPrintXmlFile CStr(sourcePath), destPath
        If Err.Number <> 0 Then
            failures = failures & vbNewLine & "  " & Mid$(sourcePath, Len(rootPath) + 2) & ": " & Err.Description
            Err.Clear
        Else
            written = written + 1
        End If
        On Error GoTo 0
    Next sourcePath

    PrettyPrintPackageXml = written
End Function

Private Sub CollectMatchingFiles(ByVal folderPath As String, ByVal rootPath As String, _
                                 ByRef patterns() As String, ByVal fso As Object, ByVal results As Collection)
    Dim item As Object
    For Each item In fso.GetFolder(folderPath).Files
        If MatchesAnyPattern(item.Path, item.Name, rootPath, patterns) Then results.Add item.Path
    Next item
    For Each item In fso.GetFolder(folderPath).SubFolders
        CollectMatchingFiles item.Path, rootPath, patterns, fso, results
    Next item
End Sub

' Re-indents an XML file into destPath (UTF-8, no BOM). sourcePath is only read.
Private Sub PrettyPrintXmlFile(ByVal sourcePath As String, ByVal destPath As String)
    ' Load into a DOM with whitespace stripped, so any existing indentation
    ' isn't doubled up by the writer's own indentation.
    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    dom.preserveWhiteSpace = False
    If Not dom.Load(sourcePath) Then
        Err.Raise ERR_EXPORT, "PrettyPrintXmlFile", Trim$(Replace(dom.parseError.reason, vbCrLf, " "))
    End If

    Dim writer As Object
    Set writer = CreateObject("MSXML2.MXXMLWriter.6.0")
    writer.indent = True
    writer.omitXMLDeclaration = True   ' written below with the correct encoding

    Dim reader As Object
    Set reader = CreateObject("MSXML2.SAXXMLReader.6.0")
    ' Report xmlns declarations as attributes, otherwise the writer drops them
    reader.putFeature "http://xml.org/sax/features/namespace-prefixes", True
    Set reader.contentHandler = writer
    reader.putProperty "http://xml.org/sax/properties/lexical-handler", writer     ' keeps comments and CDATA
    reader.putProperty "http://xml.org/sax/properties/declaration-handler", writer
    reader.parse dom

    WriteUtf8NoBom destPath, "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & writer.output
End Sub

Private Sub WriteUtf8NoBom(ByVal filePath As String, ByVal text As String)
    Dim textStream As Object
    Dim binaryStream As Object

    Set textStream = CreateObject("ADODB.Stream")
    textStream.Type = 2          ' adTypeText
    textStream.Charset = "utf-8"
    textStream.Open
    textStream.WriteText text
    textStream.Position = 3      ' skip the 3-byte BOM ADODB always writes

    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Type = 1        ' adTypeBinary
    binaryStream.Open
    textStream.CopyTo binaryStream
    binaryStream.SaveToFile filePath, 2   ' adSaveCreateOverWrite

    binaryStream.Close
    textStream.Close
End Sub

Private Function MatchesAnyPattern(ByVal fullPath As String, ByVal itemName As String, _
                                   ByVal rootPath As String, ByRef patterns() As String) As Boolean
    Dim relativePath As String
    relativePath = LCase$(Mid$(fullPath, Len(rootPath) + 2))   ' strip root and separator

    Dim entry As Variant
    Dim pattern As String
    For Each entry In patterns
        pattern = LCase$(Trim$(Replace(CStr(entry), "/", "\")))
        If pattern <> vbNullString Then
            If InStr(pattern, "\") > 0 Then
                If relativePath Like pattern Then MatchesAnyPattern = True: Exit Function
            Else
                If LCase$(itemName) Like pattern Then MatchesAnyPattern = True: Exit Function
            End If
        End If
    Next entry
End Function

Private Sub WaitOneSecond(ByVal deadline As Date)
    If Now > deadline Then
        Err.Raise ERR_EXPORT, "UnzipWithShell", "Timed out after " & UNZIP_TIMEOUT_SECONDS & " seconds waiting for the unzip to finish."
    End If
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
End Sub

' ================================================================
' Dialogs
' ================================================================

Private Function ChooseExcelFile(ByVal startPath As String) As String
    Dim picker As Office.FileDialog

    On Error GoTo ChooseExcelFile_error_handler

    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "Select a macro-enabled Excel file"
        .ButtonName = "Select"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Macro-enabled Excel files", "*.xlsm;*.xlsb;*.xlam;*.xltm;*.xls;*.xla"
        .Filters.Add "All files", "*.*"
        If FileExists(startPath) Then .InitialFileName = startPath
        If .Show = -1 Then ChooseExcelFile = .SelectedItems.item(1)
    End With

Exit_Point:
    Set picker = Nothing
    Exit Function

ChooseExcelFile_error_handler:
    MsgBox "ChooseExcelFile - " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Exit_Point
End Function

' Returns the chosen folder, or startDir unchanged if the user cancels.
Private Function ChooseDirectory(ByVal startDir As String) As String
    ChooseDirectory = startDir

    Dim picker As Office.FileDialog
    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    With picker
        .Title = "Select the export folder"
        .AllowMultiSelect = False
        If DirectoryExists(startDir) Then .InitialFileName = JoinPath(startDir, vbNullString)
        If .Show = -1 Then ChooseDirectory = .SelectedItems.item(1)   ' -1 = user clicked OK
    End With
    Set picker = Nothing
End Function

' ================================================================
' Helpers
' ================================================================

Private Function ReadSetting(ByVal rangeName As String) As String
    Dim value As Variant
    value = ExportSheet.Range(rangeName).value
    If Not IsError(value) Then ReadSetting = Trim$(CStr(value))
End Function

' Reads an on/off switch from a named cell on ExportSheet. Accepts TRUE/FALSE
' (including a checkbox's linked cell), Yes/No, On/Off and 1/0. Returns
' defaultValue if the name isn't defined or the cell is blank or unrecognized.
Private Function ReadToggle(ByVal rangeName As String, ByVal defaultValue As Boolean) As Boolean
    ReadToggle = defaultValue

    Dim value As Variant
    On Error Resume Next
    value = ExportSheet.Range(rangeName).value
    If Err.Number <> 0 Then Exit Function   ' name not defined
    On Error GoTo 0

    If IsError(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        ReadToggle = value
        Exit Function
    End If

    Select Case LCase$(Trim$(CStr(value)))
        Case "true", "yes", "y", "on", "1":  ReadToggle = True
        Case "false", "no", "n", "off", "0": ReadToggle = False
    End Select
End Function

' Deletes every empty folder below rootPath, working bottom-up so a folder that
' only contained empty folders is removed too. rootPath itself is kept.
' Returns the number of folders removed.
Private Function RemoveEmptyFolders(ByVal rootPath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootPath) Then Exit Function

    ' Snapshot the children first; deleting while enumerating SubFolders is unreliable
    Dim children As Collection
    Set children = New Collection
    Dim item As Object
    For Each item In fso.GetFolder(rootPath).SubFolders
        children.Add item.Path
    Next item

    Dim removed As Long
    Dim childPath As Variant
    Dim isEmptyFolder As Boolean
    For Each childPath In children
        removed = removed + RemoveEmptyFolders(CStr(childPath))
        With fso.GetFolder(childPath)
            isEmptyFolder = (.Files.Count = 0 And .SubFolders.Count = 0)
        End With
        If isEmptyFolder Then
            fso.DeleteFolder CStr(childPath), True
            removed = removed + 1
        End If
    Next childPath

    RemoveEmptyFolders = removed
End Function

Private Function IsVBProjectAccessTrusted() As Boolean
    ' Reading .VBProject raises error 1004 when "Trust access to the VBA project object model" is off.
    Dim probe As Object
    On Error Resume Next
    Set probe = ThisWorkbook.VBProject
    IsVBProjectAccessTrusted = (Err.Number = 0)
    On Error GoTo 0
    Set probe = Nothing
End Function

Private Function GetOpenWorkbookByName(ByVal fileName As String) As Workbook
    ' Workbooks(name) also finds installed add-ins, which For Each over Workbooks does not.
    On Error Resume Next
    Set GetOpenWorkbookByName = Workbooks(fileName)
    On Error GoTo 0
End Function

Private Function IsSameWorkbookPath(ByVal openFullName As String, ByVal filePath As String) As Boolean
    ' OneDrive/SharePoint-synced workbooks report a URL as FullName, so the local
    ' path can't be compared; treat a same-named synced workbook as the same file.
    If LCase$(Left$(openFullName, 4)) = "http" Then
        IsSameWorkbookPath = True
    Else
        IsSameWorkbookPath = (StrComp(openFullName, filePath, vbTextCompare) = 0)
    End If
End Function

Private Sub HideWindows(ByVal wb As Workbook)
    Dim wbWindow As Window
    For Each wbWindow In wb.Windows
        wbWindow.Visible = False
    Next wbWindow
End Sub

Private Function JoinPath(ByVal basePath As String, ByVal child As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Do While Len(basePath) > 0 And (Right$(basePath, 1) = sep Or Right$(basePath, 1) = "/")
        basePath = Left$(basePath, Len(basePath) - 1)
    Loop
    JoinPath = basePath & sep & child
End Function

Private Sub SetStatus(ByVal message As String)
    Application.StatusBar = message
    DoEvents
End Sub

Private Function DirectoryExists(ByVal dirPath As String) As Boolean
    If Len(dirPath) = 0 Then Exit Function
    DirectoryExists = CreateObject("Scripting.FileSystemObject").FolderExists(dirPath)
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then Exit Function
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
End Function

' Creates a folder and any missing parent folders. Raises on failure.
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> vbNullString Then
        If Not fso.FolderExists(parentPath) Then EnsureFolder parentPath
    End If

    SetStatus "Creating directory " & folderPath
    fso.CreateFolder folderPath
End Sub

Private Sub DeleteFolderIfExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub

    On Error GoTo DeleteFailed
    fso.DeleteFolder folderPath, True   ' True = also delete read-only files
    Exit Sub

DeleteFailed:
    Err.Raise Err.Number, "DeleteFolderIfExists", _
              "Could not clear " & folderPath & " (" & Err.Description & ")." & vbNewLine & _
              "Close any program or Explorer window using files in it and try again."
End Sub
