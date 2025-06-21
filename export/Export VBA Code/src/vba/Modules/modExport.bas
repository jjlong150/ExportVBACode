Attribute VB_Name = "modExport"
' MIT License

' Copyright (c) 2025 Jeffrey Long

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

'@IgnoreModule ModuleWithoutFolder
Option Explicit

' https://docs.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#type
Private Enum VBComponent
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
    vbext_ct_ActiveXDesigner = 11
    vbext_ct_Document = 100
End Enum

'@EntryPoint
Public Sub ExportVBA()
    Dim fileName As String
    fileName = ExportSheet.Range("FILENAME").Value
    
    Dim exportFolder As String
    exportFolder = ExportSheet.Range("EXPORT_FOLDER").Value
    
    If fileName <> vbNullString And exportFolder <> vbNullString Then
        ExportComponents fileName, exportFolder
        Application.StatusBar = "Done!"
    End If
End Sub

Private Sub ExportComponents(ByVal filePathIn As String, ByVal exportFolder As String)

    ' ================================================================
    ' Create the directory tree structure
    ' ================================================================
    
    Dim baseName As String
    Dim fileName As String
    
    ' Break the input file path down into pieces
    With CreateObject("Scripting.FileSystemObject")
        fileName = .GetFileName(filePathIn)
        baseName = .GetBaseName(filePathIn)
    End With

    ' Create the base export directory
    Dim exportDir As String
    exportDir = exportFolder & Application.PathSeparator & baseName
    DeleteDirAndContents exportDir
    CreateDirectory exportDir

    ' Create a directory to hold all the exported source code
    Dim exportSrcDir As String
    exportSrcDir = exportDir & Application.PathSeparator & "src"
    CreateDirectory exportSrcDir
    
    ' Create a directory to hold a copy of the workbook
    Dim exportWorkbookDir As String
    exportWorkbookDir = exportSrcDir & Application.PathSeparator & "excel"
    CreateDirectory exportWorkbookDir

    ' Create a directory tree to organize the exported VBA components
    Dim exportSrcVbaDir As String
    exportSrcVbaDir = exportSrcDir & Application.PathSeparator & "vba"
    CreateDirectory exportSrcVbaDir
    
    CreateDirectory exportSrcVbaDir & Application.PathSeparator & "Modules"
    CreateDirectory exportSrcVbaDir & Application.PathSeparator & "Class Modules"
    CreateDirectory exportSrcVbaDir & Application.PathSeparator & "Forms"
    CreateDirectory exportSrcVbaDir & Application.PathSeparator & "ActiveX"
    CreateDirectory exportSrcVbaDir & Application.PathSeparator & "Microsoft Excel Objects"

    ' Create a directory to hold the contents of the unzipped spreadsheet
    Dim exportSrcZipDir As String
    exportSrcZipDir = exportSrcDir & Application.PathSeparator & "xlsm"
    CreateDirectory exportSrcZipDir
    
    ' ================================================================
    ' Step 1 - Create a copy of the workbook
    ' ================================================================
    
    ' Remove any file with the same name from the destination directory
    Dim exportWorkbookFile As String
    exportWorkbookFile = exportWorkbookDir & Application.PathSeparator & fileName
    If FileExists(exportWorkbookFile) Then
        Kill exportWorkbookFile
    End If
    
    ' Copy the workbook to the export directory
    Application.StatusBar = "Copying " & filePathIn & " to " & exportWorkbookDir
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile (filePathIn), exportWorkbookDir & Application.PathSeparator, True
    Set FSO = Nothing

    ' ================================================================
    ' Step 2 - Export the VBA Code
    ' ================================================================
    
    ' Open the workbook, but don't display it
    Dim wrkbook As Excel.Workbook
    Application.ScreenUpdating = False
    Set wrkbook = Workbooks.Open(filePathIn)
    ActiveWindow.Visible = False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    
    Dim visualBasicProject As Object
    Set visualBasicProject = wrkbook.VBProject
    
    Dim NumComponents As Long
    NumComponents = visualBasicProject.VBComponents.Count
    
    Dim ext As String
    Dim subdir As String
    
    Dim filePathOut As String
    Dim i As Long
    For i = 1 To NumComponents
        Select Case visualBasicProject.VBComponents(i).Type
        Case VBComponent.vbext_ct_StdModule         ' = 1
            ext = "bas"
            subdir = "Modules"
        Case VBComponent.vbext_ct_ClassModule       ' = 2
            ext = "cls"
            subdir = "Class Modules"
        Case VBComponent.vbext_ct_MSForm            ' = 3
            ext = "frm"
            subdir = "Forms"
        Case VBComponent.vbext_ct_ActiveXDesigner   ' = 11
            ext = "ocx"
            subdir = "ActiveX"
        Case VBComponent.vbext_ct_Document          ' = 100
            ext = "cls"
            subdir = "Microsoft Excel Objects"
        End Select
        
        filePathOut = exportSrcVbaDir & Application.PathSeparator & subdir & Application.PathSeparator & visualBasicProject.VBComponents(i).Name & "." & ext
        
        Application.StatusBar = "Exporting " & filePathOut
        visualBasicProject.VBComponents(i).Export (filePathOut)
        DoEvents
    Next i
    
    ' Cleanup
    Set visualBasicProject = Nothing
    
    wrkbook.Close SaveChanges:=False
    Set wrkbook = Nothing
    
    ' ================================================================
    ' Step 3 - Unzip the workbook
    ' ================================================================
    
    ' Rename the copy of the workbook to have a "zip" extension
    Dim filePathZip As String
    filePathZip = exportWorkbookDir & Application.PathSeparator & baseName & ".zip"
    If FileExists(filePathZip) Then
        Kill filePathZip
    End If
    
    If FileExists(exportWorkbookFile) Then
        Name exportWorkbookFile As filePathZip
    End If
    
    Dim shellApplication As Shell
    Set shellApplication = CreateObject("Shell.Application")
    
    Dim zipFileItems As FolderItems
    Set zipFileItems = shellApplication.Namespace(CVar(filePathZip)).Items

    ' Unzip all files to the "source" folder
    Application.StatusBar = "Unzipping spreadsheet contents to " & exportSrcZipDir
    shellApplication.Namespace(exportSrcZipDir).CopyHere zipFileItems
    
    ' Rename the zip file back to the original filename
    Name filePathZip As exportWorkbookFile
    
    ' Cleanup
    Set zipFileItems = Nothing
    Set shellApplication = Nothing

End Sub

'@EntryPoint
Public Sub GetFileName()
    Dim fileName As String
    fileName = ChooseExcelFile()
    ExportSheet.Range("FILENAME").Value = fileName
End Sub

'@EntryPoint
Public Sub ClearFileName()
    ExportSheet.Range("FILENAME").ClearContents
End Sub

'@EntryPoint
Public Sub GetExportFolder()
    Dim folderName As String
    folderName = ExportSheet.Range("EXPORT_FOLDER").Value
    folderName = ChooseDirectory(folderName)
    ExportSheet.Range("EXPORT_FOLDER").Value = folderName
End Sub

'@EntryPoint
Public Sub ClearExportFolder()
    ExportSheet.Range("EXPORT_FOLDER").ClearContents
End Sub

Private Function ChooseExcelFile() As String
    Dim fdExcelFilePicker As Office.FileDialog
    
    On Error GoTo ChooseExcelFile_error_handler
    
    Set fdExcelFilePicker = Application.FileDialog(msoFileDialogOpen)
    With fdExcelFilePicker
       .Title = "Select Macro-enabled Excel File"
       .ButtonName = "Open"
       .Filters.Clear
       .Filters.Add "Excel Files", "*.xlsm"
       .Show
       If .SelectedItems.Count > 0 Then
          ChooseExcelFile = .SelectedItems.Item(1)
       End If
    End With
    
Exit_Point:
    Exit Function
    
ChooseExcelFile_error_handler:
    MsgBox "ChooseExcelFile - " & Err.Number & " - " & Err.Description
    On Error GoTo 0: Resume Exit_Point
End Function

Private Function ChooseDirectory(ByVal startDir As String) As String
    ChooseDirectory = startDir
    Dim fileDialogHandle As FileDialog
    Set fileDialogHandle = Application.FileDialog(msoFileDialogFolderPicker)
    
    If fileDialogHandle Is Nothing Then
        MsgBox "msgboxNoFolderPicker"
    Else
        If Trim$(startDir) <> vbNullString Then
            If DirectoryExists(startDir) Then
                fileDialogHandle.InitialFileName = startDir
            End If
        End If
    
        '  Get the number of the button chosen
        Dim selected As Long
        selected = fileDialogHandle.Show
        If selected = -1 Then
            ' User clicked on CANCEL
        
            ' Set path of directory chosen
            ChooseDirectory = fileDialogHandle.SelectedItems.Item(1)
        End If
    End If
    Set fileDialogHandle = Nothing
End Function

Private Function DirectoryExists(ByVal dirPath As String) As Boolean

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    DirectoryExists = False
    
    If Len(dirPath) > 0 Then
        If FSO.FolderExists(dirPath) = True Then
            DirectoryExists = True
        End If
    End If

    Set FSO = Nothing
End Function

Private Sub CreateDirectory(ByVal directoryName As String)
    On Error Resume Next
    Application.StatusBar = "Creating directory " & directoryName
    MkDir directoryName
    On Error GoTo 0
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = False

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Len(filePath) > 0 Then
        If FSO.FileExists(filePath) = True Then
            FileExists = True
        End If
    End If

    Set FSO = Nothing
End Function

Private Sub DeleteDirAndContents(ByVal folderPath As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Len(folderPath) > 0 Then
        If FSO.FolderExists(folderPath) = True Then
            FSO.deletefolder folderPath
        End If
    End If

    Set FSO = Nothing
End Sub

