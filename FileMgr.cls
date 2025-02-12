VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FileMgr Class Module
Option Explicit

Private fileDialogFilters As Collection
Private fileDialogTitle As String
Private allowMultiSelect As Boolean
Private fdType As MsoFileDialogType

Private Sub Class_Initialize()
    Set fileDialogFilters = New Collection
    fileDialogTitle = "Select a file"
    allowMultiSelect = False
    fdType = msoFileDialogFilePicker ' Default to file picker, can be changed if needed
End Sub

Public Sub AddFilter(DESCRIPTION As String, extensions As String)
    fileDialogFilters.Add Array(DESCRIPTION, extensions)
End Sub

Public Sub ClearFilters()
    Set fileDialogFilters = New Collection
End Sub

Public Function SelectFile(Optional customTitle As String, Optional multiSelect As Boolean) As String
    If customTitle = "" Then customTitle = fileDialogTitle
    If multiSelect Then allowMultiSelect = multiSelect
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(fdType)
    
    With fd
        .allowMultiSelect = allowMultiSelect
        .Title = customTitle
        .InitialFileName = getWBPath & "\"
        
        Dim filterItem As Variant
            .Filters.Clear
        For Each filterItem In fileDialogFilters
            .Filters.Add filterItem(0), filterItem(1)
        Next filterItem
        
        If .Show = True Then
            If allowMultiSelect Then
                Dim selectedFiles As String
                Dim i As Integer
                For i = 1 To .SelectedItems.Count
                    selectedFiles = selectedFiles & .SelectedItems(i) & ";"
                Next i
                SelectFile = selectedFiles
            Else
                SelectFile = .SelectedItems(1)
            End If
        End If
    End With
    
    'SelectFile = localPath(SelectFile)
    
End Function

Public Function SelectFolder(Optional customTitle As String) As String
    fileDialogTitle = IIf(customTitle = "", "Select a folder", customTitle)
    fdType = msoFileDialogFolderPicker
    SelectFile customTitle, False ' Call SelectFile with folder picker setup
End Function

Public Function FolderExists(sFolderPath As String) As Boolean
    FolderExists = (dir(sFolderPath, vbDirectory) <> vbNullString And sFolderPath <> vbNullString)
End Function

Public Function FileExists(filePath As String) As Boolean
    Dim testStr As String
    If Not filePath = "" Then
        testStr = dir(filePath)
        FileExists = (testStr <> vbNullString)
    End If
End Function

Public Function GetFilenameFromPath(ByVal strPath As String) As String
    GetFilenameFromPath = Mid(strPath, InStrRev(strPath, "\") + 1)
End Function

' Gets the directory of a given filename
Public Function GetDir(fileName As String) As String
    GetDir = Left(fileName, InStrRev(fileName, "\"))
End Function

' Reduces the folder path level by one
Public Function GetRootDir(folderName As String) As String
    If Right(folderName, 1) = "\" Then folderName = Left(folderName, Len(folderName) - 1)
    GetRootDir = Left(folderName, InStrRev(folderName, "\"))
End Function

' Adds a subfolder to a root folder and returns True if successful
Public Function addSubFolder(rootFolder As String, subFolder As String) As Boolean
    Dim newPath As String
    newPath = rootFolder & subFolder & "\"
    If Not FolderExists(newPath) Then
        MkDir newPath
        addSubFolder = FolderExists(newPath)
    Else
        addSubFolder = True ' Folder already exists
    End If
End Function

' Copies a file to a new location with an optional new name
Public Function copyFile(orgFileName As String, copyFolder As String, Optional copyName As String) As Boolean
    Dim destFile As String
    If copyName = "" Then
        destFile = copyFolder & "\" & GetFilenameFromPath(orgFileName)
    Else
        destFile = copyFolder & "\" & copyName & Mid(orgFileName, InStrRev(orgFileName, "."))
    End If

    If FileExists(destFile) Then Kill destFile ' Remove existing file if it exists

    On Error Resume Next
    FileCopy orgFileName, destFile
    copyFile = (Err.Number = 0)
    On Error GoTo 0
End Function

' Opens a file using the default system application and returns True if successful
Public Function openFile(fileName As String) As Boolean
    Dim result As Long
    Dim targetFile As String
    Dim fileSpec As String
    Dim fileFound As String

    Dim fileNameOnly As String
    
    ' Check if fileName includes an extension
    If InStrRev(fileName, ".") > InStrRev(fileName, "\") Then
        targetFile = fileName
    Else
        fileNameOnly = GetFilenameFromPath(fileName)
        ' If no extension, search for any file with the given name regardless of extension
        fileSpec = GetDir(fileName) & fileNameOnly & ".*"
        Debug.Print fileSpec
        fileFound = dir(fileSpec)
        If fileFound = "" Then
            openFile = False
            Exit Function
        Else
            targetFile = GetDir(fileName) & fileFound
        End If
    End If

    ' Attempt to open the file
    On Error Resume Next
    result = Shell("explorer.exe """ & targetFile & """", vbNormalFocus)
    If result > 0 Then
        openFile = True
    Else
        openFile = False
    End If
    On Error GoTo 0
End Function



