Attribute VB_Name = "GetDirectoryFunction"
#If VBA7 Then
    Public Type BROWSEINFO
        hOwner As LongPtr
        pidlRoot As LongPtr
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As LongPtr
        lParam As Long
        iImage As Long
    End Type

    Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As Long
    Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (ByRef lpBrowseInfo As BROWSEINFO) As LongPtr
#Else
    Public Type BROWSEINFO
        hOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
    End Type

    Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (ByRef lpBrowseInfo As BROWSEINFO) As Long
#End If

Function OldGetDirectory(Optional Msg) As String
    Dim bInfo As BROWSEINFO
    Dim Path As String
    Dim r As Long
    #If VBA7 Then
        Dim x As LongPtr
    #Else
        Dim x As Long
    #End If
    Dim pos As Integer
 
'   Root folder = Desktop
    bInfo.pidlRoot = 0&

'   Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
    Else
        bInfo.lpszTitle = Msg
    End If
    
'   Type of directory to return
    bInfo.ulFlags = &H1

'   Display the dialog
    x = SHBrowseForFolder(bInfo)
    
'   Parse the result
    Path = Space$(512)
    r = SHGetPathFromIDList(ByVal x, ByVal Path)
    If r Then
        pos = InStr(Path, Chr$(0))
        'GetDirectory = Left(Path, pos - 1)
    Else
        'GetDirectory = ""
    End If
End Function


Function GetDirectory(Optional Msg As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        If Msg <> "" Then .Title = Msg ' Set the title of the dialog box if provided
        If .Show = -1 Then ' If the user makes a selection
            GetDirectory = .SelectedItems(1) ' Return the selected path
        Else
            GetDirectory = "" ' If the user cancels, return an empty string
        End If
    End With
End Function


