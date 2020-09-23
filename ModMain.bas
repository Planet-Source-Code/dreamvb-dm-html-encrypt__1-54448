Attribute VB_Name = "ModMain"
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const BIF_RETURNONLYFSDIRS = &H1

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Global Const mFile = "C:\tlib.dat"
Public mCounter As Integer

Function OpenFile(lzFileName As String) As String
Dim nFile As Long
Dim StrBuff As String
    nFile = FreeFile
    
    Open lzFileName For Binary As #nFile
        StrBuff = Space(LOF(1))
        Get #nFile, , StrBuff
    Close #nFile
    
    OpenFile = StrBuff
    
End Function

Function SaveFile(lzFile As String, sData As String)
Dim nFile As Long
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        Put #nFile, , sData
    Close #nFile
    
End Function

Public Function FindFile(lzFile As String) As Boolean
    ' This function will retun a result of a file of exsitence file found will return with a true value
    If Dir$(lzFile) = "" Then FindFile = False Else FindFile = True
End Function

Public Function FixPath(lzPath As String) As String
    ' Fixes a path by adding a back slash if required
    If Right$(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function GetFileExt(lzFile As String) As String
Dim I As Long, iPart As Long, StrA As String
   For I = Len(lzFile) To 1 Step -1
        StrA = Mid(lzFile, I, 1)
        If StrA = "." Then
            iPart = I
            Exit For
        End If
   Next
   
   If iPart = 0 Then
        GetFileExt = ""
    Else
        GetFileExt = UCase$(Mid$(lzFile, iPart + 1, Len(lzFile)))
   End If
   iPart = 0: I = 0
   StrA = ""
   
End Function

Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim OffSet As Integer

    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        OffSet = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, OffSet - 1)
    End If

End Function
