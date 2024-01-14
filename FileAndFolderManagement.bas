Attribute VB_Name = "FileAndFolderManagement"
Option Explicit

'Set variables
Dim MyPath As String
Dim NewPath As String
Dim DirName As String, NextFile As String, ChangeTo As String
Dim ErrorMsg As String
Dim Path As String, Msg As String
Dim RowCounter As Integer
Dim Unchanged As Integer
Dim fso As Object

Const RenamedColour As Integer = 36
Const ProblemColour As Integer = 40
Const UnchangedColour As Integer = 35

Type FileInfo
    DateModified As Date
    FileName As String
End Type



Sub GetSourcePath()

    'Get the path using shell32.dll routine
    Msg = "Select a directory for the file list"
    Path = GetDirectory(Msg)
    If Path <> "" Then
        Range("Path").Value = Path
    End If
    
End Sub


Sub ListFiles()
'Create a list of files

    'Check for errors
    'Create error message
    ErrorMsg = "Problem creating list - check path."
    'On Error GoTo ErrorHandler
    If Range("Path").Value = "" Then GoTo ErrorHandler
    
    'If no error then main code
    Application.ScreenUpdating = False
    
    DirName = Range("Path").Value
    If Right(DirName, 1) <> "\" Then DirName = DirName & "\"
    NextFile = Dir(DirName)
    'Check if there are no files
    If NextFile = "" Then MsgBox "Incorrect path specified or no files detected", vbInformation, "List files": Exit Sub

    'Clear area for list
    Range("Filelist").Offset(1, 0).Select
    RowCounter = 0
    Range("B" & ActiveCell.Row & ":G65536").ClearContents
    Range("B" & ActiveCell.Row & ":G65536").Interior.ColorIndex = 2
    
    'Loop to insert file name and details
    Do While NextFile <> ""
        ActiveCell.Offset(RowCounter, 0).Value = NextFile
        ActiveCell.Offset(RowCounter, 5).Value = NextFile
        ActiveCell.Offset(RowCounter, 1).Value = FileLen(DirName & NextFile)
        ActiveCell.Offset(RowCounter, 2).Value = FileDateTime(DirName & NextFile)
        NextFile = Dir()
        RowCounter = RowCounter + 1
    Loop
    If ActiveCell.Offset(1, 0).Value = "" Then [A1].Select: Exit Sub
    
    'Sort alphabetically
    Selection.CurrentRegion.Select
    Selection.Sort key1:=Range(ActiveCell.Address), order1:=xlAscending, Header:=xlYes
    [A1].Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
    Exit Sub
ErrorHandler:
    MsgBox ErrorMsg, vbInformation, "List files"
    [A1].Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
    Exit Sub
End Sub
Sub ListFolders()
    'Create a list of folders

    'Check for errors
    'Create error message
    ErrorMsg = "Problem creating list - check path."
    'On Error GoTo ErrorHandler
    If Range("Path").Value = "" Then GoTo ErrorHandler
    
    'If no error then main code
    Application.ScreenUpdating = False
    
    DirName = Range("Path").Value
    If Right(DirName, 1) <> "\" Then DirName = DirName & "\"
    
    'Clear area for list
    Range("Filelist").Offset(1, 0).Select
    RowCounter = 0
    Range("B" & ActiveCell.Row & ":G65536").ClearContents
    Range("B" & ActiveCell.Row & ":G65536").Interior.ColorIndex = 2
    
    ' Loop to insert folder names and details
    NextFile = Dir(DirName & "*.*", vbDirectory)
    Do While NextFile <> ""
        If (GetAttr(DirName & NextFile) And vbDirectory) = vbDirectory Then
            If NextFile <> "." And NextFile <> ".." Then
                ' Populate the list with folder details
                ActiveCell.Offset(RowCounter, 0).Value = NextFile
                ActiveCell.Offset(RowCounter, 5).Value = NextFile ' Copy folder name to column G
                ' Add other folder details as needed
                RowCounter = RowCounter + 1
            End If
        End If
        NextFile = Dir() ' Get next folder
    Loop
    If ActiveCell.Offset(1, 0).Value = "" Then [A1].Select: Exit Sub
    
    'Sort alphabetically
    Selection.CurrentRegion.Select
    Selection.Sort key1:=Range(ActiveCell.Address), order1:=xlAscending, Header:=xlYes
    [A1].Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
    Exit Sub
ErrorHandler:
    MsgBox ErrorMsg, vbInformation, "List folders"
    [A1].Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
End Sub

Sub AuditFolders()
    'Create a list of folders and find the newest file's date in each

    'Check for errors
    ErrorMsg = "Problem creating list - check path."
    If Range("Path").Value = "" Then GoTo ErrorHandler
    
    'If no error then main code
    Application.ScreenUpdating = False
    
    DirName = Range("Path").Value
    If Right(DirName, 1) <> "\" Then DirName = DirName & "\"
    
    'Clear area for list
    Range("Filelist").Offset(1, 0).Select
    RowCounter = 0
    Range("B" & ActiveCell.Row & ":G65536").ClearContents
    Range("B" & ActiveCell.Row & ":G65536").Interior.ColorIndex = 2
    
    ' Loop to insert folder names and details
    NextFile = Dir(DirName & "*.*", vbDirectory)
    Do While NextFile <> ""
        If (GetAttr(DirName & NextFile) And vbDirectory) = vbDirectory Then
            If NextFile <> "." And NextFile <> ".." Then
                ' Populate the list with folder details
                ActiveCell.Offset(RowCounter, 0).Value = NextFile
                'ActiveCell.Offset(RowCounter, 5).Value = NextFile ' Copy folder name to column G

                ' Find the newest file in the folder
                Dim NewestFileInfo As FileInfo
                NewestFileInfo = GetNewestFileInfo(DirName & NextFile)

                ' Add the newest file's date to column D and name to the next column
                If NewestFileInfo.DateModified <> 0 Then
                    ActiveCell.Offset(RowCounter, 2).Value = NewestFileInfo.DateModified
                    ActiveCell.Offset(RowCounter, 3).Value = NewestFileInfo.FileName

                    ' Calculate the age in months and add to the next column
                    Dim FileAgeMonths As Integer
                    FileAgeMonths = DateDiff("m", NewestFileInfo.DateModified, Now)
                    ActiveCell.Offset(RowCounter, 4).Value = FileAgeMonths
                End If

                RowCounter = RowCounter + 1
            End If
        End If
        NextFile = Dir() ' Get next folder
    Loop
    If ActiveCell.Offset(1, 0).Value = "" Then [A1].Select: Exit Sub
    
    'Sort alphabetically
    Selection.CurrentRegion.Select
    Selection.Sort key1:=Range(ActiveCell.Address), order1:=xlAscending, Header:=xlYes
    [A1].Select
    Exit Sub

ErrorHandler:
    MsgBox ErrorMsg, vbInformation, "List folders"
    [A1].Select
End Sub

Sub AuditFiles()
    'Create a list of all files in folders

    'Check for errors
    ErrorMsg = "Problem creating list - check path."
    If Range("Path").Value = "" Then GoTo ErrorHandler
    
    'If no error then main code
    Application.ScreenUpdating = False
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    DirName = Range("Path").Value
    If Right(DirName, 1) <> "\" Then DirName = DirName & "\"

    'Clear area for list
    Range("Filelist").Offset(1, 0).Select
    RowCounter = 0
    Range("B" & ActiveCell.Row & ":F65536").ClearContents
    Range("B" & ActiveCell.Row & ":F65536").Interior.ColorIndex = 2

    ' Get the root folder
    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(DirName)
    
    ' Loop through each subfolder
    Dim subFolder As Object
    For Each subFolder In rootFolder.SubFolders
        ' Loop through each file in the subfolder
        Dim file As Object
        For Each file In subFolder.Files
            ' Populate the list with file details
            ActiveCell.Offset(RowCounter, 0).Value = subFolder.Name
            ActiveCell.Offset(RowCounter, 3).Value = file.Name
            ActiveCell.Offset(RowCounter, 2).Value = file.DateLastModified

            ' Calculate the age in months and add to the next column
            Dim FileAgeMonths As Integer
            FileAgeMonths = DateDiff("m", file.DateLastModified, Now)
            ActiveCell.Offset(RowCounter, 4).Value = FileAgeMonths

            RowCounter = RowCounter + 1
        Next file
    Next subFolder

    If ActiveCell.Offset(1, 0).Value = "" Then [A1].Select: Exit Sub
    
    'Sort alphabetically
    Selection.CurrentRegion.Select
    Selection.Sort key1:=Range(ActiveCell.Address), order1:=xlAscending, Header:=xlYes
    [A1].Select
    Exit Sub

ErrorHandler:
    MsgBox ErrorMsg, vbInformation, "List files"
    [A1].Select
End Sub


Function GetNewestFileDate(FolderPath As String) As Date
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(FolderPath)
    
    Dim file As Object
    Dim mostRecentDate As Date
    mostRecentDate = 0 ' Initialize to a zero date

    For Each file In folder.Files
        If file.DateLastModified > mostRecentDate Then
            mostRecentDate = file.DateLastModified
        End If
    Next file

    GetNewestFileDate = mostRecentDate
End Function

Function GetNewestFileInfo(FolderPath As String) As FileInfo
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(FolderPath)
    
    Dim file As Object
    Dim mostRecentInfo As FileInfo
    mostRecentInfo.DateModified = 0 ' Initialize to a zero date

    For Each file In folder.Files
        If file.DateLastModified > mostRecentInfo.DateModified Then
            mostRecentInfo.DateModified = file.DateLastModified
            mostRecentInfo.FileName = file.Name
        End If
    Next file

    GetNewestFileInfo = mostRecentInfo
End Function




Sub FindAndReplace()
    'ActiveSheet.Unprotect
    Range("FileList").Offset(2, 4).Select
    If ActiveCell.Value <> "" Then
        ActiveCell.Offset(-1, 0).Select
        Range(ActiveCell, ActiveCell.End(xlDown)).Select
        Application.Dialogs(xlDialogFormulaReplace).Show
    End If
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
End Sub

Sub RemoveAllFlags()
    'ActiveSheet.Unprotect
    Application.ScreenUpdating = False
    Range("Filelist").Offset(1, 0).Select
    Range("E" & ActiveCell.Row & ":E65536").ClearContents
    Range("B" & ActiveCell.Row & ":G65536").Interior.ColorIndex = 2
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
End Sub

Sub Removeformulas()
 
    Application.ScreenUpdating = False
    Range("Filelist").Offset(1, 4).Select
    Range(ActiveCell, "F65536").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues
    
    Range("Filelist").Offset(1, 5).Select
    Range(ActiveCell, "G65536").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues
    Range("Filelist").Offset(1, 5).Select
    Application.CutCopyMode = False
    
End Sub

Sub RenameFiles()
    'ActiveSheet.Unprotect
    Application.ScreenUpdating = False
    Range("Filelist").Offset(1, 0).Select
    RowCounter = 0
    Unchanged = 0
    If ActiveCell.Value = "" Then
        MsgBox "No files detected", vbInformation, "Rename files"
        Exit Sub
    End If
    
    MyPath = Range("Path").Value
    NewPath = MyPath
    
    If MyPath = "" Then
        Application.ScreenUpdating = True
        MsgBox "No Path specified", vbInformation, "Rename files"
        Exit Sub
    End If
    
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    On Error GoTo BadFile
    
    Do
        If ActiveCell.Offset(RowCounter, 0).Interior.ColorIndex <> RenamedColour Then
            NextFile = MyPath & ActiveCell.Offset(RowCounter, 0)
            
            NewPath = ActiveCell.Offset(RowCounter, 4).Value
            
            If NewPath = "" Then
                NewPath = MyPath
            ElseIf Right(NewPath, 1) <> "\" Then
                NewPath = NewPath & "\"
            End If
                        
            ChangeTo = NewPath & ActiveCell.Offset(RowCounter, 5)
            RowCounter = RowCounter + 1
            If NextFile = ChangeTo Then
                Range("B" & RowCounter + Range("Filelist").Row & ":G" & RowCounter + Range("Filelist").Row).Interior.ColorIndex = UnchangedColour
                Range("E" & RowCounter + Range("Filelist").Row).Value = "U"
                Unchanged = Unchanged + 1
            Else
                Name NextFile As ChangeTo
                Range("B" & RowCounter + Range("Filelist").Row & ":G" & RowCounter + Range("Filelist").Row).Interior.ColorIndex = RenamedColour
                Range("E" & RowCounter + Range("Filelist").Row).Value = "R"
            End If
        Else
            RowCounter = RowCounter + 1
        End If
        Application.StatusBar = "Renaming File " & (RowCounter + 1)
        
    Loop Until ActiveCell.Offset(RowCounter, 0).Value = ""
    
    Application.StatusBar = RowCounter - Unchanged & " Files Renamed"
    
    Application.ScreenUpdating = True
    MsgBox RowCounter - Unchanged & " files renamed" & Chr(13) & Unchanged & " files unchanged", vbInformation, "Rename files"
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
    Exit Sub
BadFile:
    Range("B" & RowCounter + Range("Filelist").Row & ":G" & RowCounter + Range("Filelist").Row).Interior.ColorIndex = ProblemColour
    Range("E" & RowCounter + Range("Filelist").Row).Value = "P"
    Range("Filelist").Offset(RowCounter, 0).Select
    Application.ScreenUpdating = True
    MsgBox "Problem with file..." & Chr(13) & Chr(13) & NextFile & Chr(13) & Chr(13) & "Error=" & Err.Description, vbCritical, "Rename files"
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
End Sub

Sub RenameFilesAndFolders()
    Application.ScreenUpdating = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Range("Filelist").Offset(1, 0).Select
    Dim RowCounter As Integer
    RowCounter = 0
    Dim Unchanged As Integer
    Unchanged = 0
    If ActiveCell.Value = "" Then
        MsgBox "No files detected", vbInformation, "Rename files"
        Exit Sub
    End If

    Dim MyPath As String
    MyPath = Range("Path").Value
    Dim NewPath As String
    NewPath = MyPath

    If MyPath = "" Then
        Application.ScreenUpdating = True
        MsgBox "No Path specified", vbInformation, "Rename files"
        Exit Sub
    End If

    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"

    On Error GoTo BadFile
    Do
        If ActiveCell.Offset(RowCounter, 0).Interior.ColorIndex <> RenamedColour Then
            Dim NextFile As String
            NextFile = MyPath & ActiveCell.Offset(RowCounter, 0)

            NewPath = ActiveCell.Offset(RowCounter, 4).Value
            If NewPath = "" Then
                NewPath = MyPath
            ElseIf Right(NewPath, 1) <> "\" Then
                NewPath = NewPath & "\"
            End If

            Dim ChangeTo As String
            ChangeTo = NewPath & ActiveCell.Offset(RowCounter, 5)

            If fso.FileExists(NextFile) Then
                ' Rename file
                Name NextFile As ChangeTo
            ElseIf fso.FolderExists(NextFile) Then
                ' Rename folder
                fso.MoveFolder Source:=NextFile, Destination:=ChangeTo
            End If

            ' Check if the renaming was successful
            RowCounter = RowCounter + 1
            If NextFile <> ChangeTo Then
                Range("B" & RowCounter + Range("Filelist").Row & ":G" & RowCounter + Range("Filelist").Row).Interior.ColorIndex = RenamedColour
                Range("E" & RowCounter + Range("Filelist").Row).Value = "R"
                Unchanged = Unchanged + 1
            Else
                Range("B" & RowCounter + Range("Filelist").Row & ":G" & RowCounter + Range("Filelist").Row).Interior.ColorIndex = UnchangedColour
                Range("E" & RowCounter + Range("Filelist").Row).Value = "U"
                Unchanged = Unchanged + 1
            End If
            Else
            RowCounter = RowCounter + 1
        End If

        
        Application.StatusBar = "Processing " & RowCounter
    Loop Until ActiveCell.Offset(RowCounter, 0).Value = ""

    Application.StatusBar = RowCounter - Unchanged & " Items Renamed"
    Application.ScreenUpdating = True
    MsgBox RowCounter - Unchanged & " items renamed" & Chr(13) & Unchanged & " items unchanged", vbInformation, "Rename files"
    Exit Sub

BadFile:
    Range("B" & RowCounter + Range("Filelist").Row & ":G" & RowCounter + Range("Filelist").Row).Interior.ColorIndex = ProblemColour
    Range("E" & RowCounter + Range("Filelist").Row).Value = "P"
    Range("Filelist").Offset(RowCounter, 0).Select
    Application.ScreenUpdating = True
    MsgBox "Problem with file..." & Chr(13) & Chr(13) & NextFile & Chr(13) & Chr(13) & "Error=" & Err.Description, vbCritical, "Rename files"
End Sub





Sub ClearList()
'Clear file list
    'ActiveSheet.Unprotect
    Range("Filelist").Offset(1, 0).Select
    RowCounter = 0
    Range("B" & ActiveCell.Row & ":G65536").ClearContents
    Range("B" & ActiveCell.Row & ":G65536").Interior.ColorIndex = 2
    [A1].Select
    Application.StatusBar = False
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True
End Sub

