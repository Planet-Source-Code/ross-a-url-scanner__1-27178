Attribute VB_Name = "INetOpt"
Option Explicit

Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function FindFirstUrlCacheGroup Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwFilter As Long, ByRef lpSearchCondition As Long, ByVal dwSearchCondition As Long, ByRef lpGroupId As Date, ByRef lpReserved As Long) As Long
Declare Function FindNextUrlCacheGroup Lib "wininet.dll" (ByVal hFind As Long, ByRef lpGroupId As Date, ByRef lpReserved As Long) As Long
Declare Function DeleteUrlCacheGroup Lib "wininet.dll" (ByVal sGroupID As Date, ByVal dwFlags As Long, ByRef lpReserved As Long) As Long
Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, ByRef lpFirstCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As Long) As Long
Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByRef lpNextCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long

Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    szRestOfData(1024) As Long
End Type

Const CACHGROUP_SEARCH_ALL = &H0
Const ERROR_NO_MORE_FILES = 18
Const ERROR_NO_MORE_ITEMS = 259
Const CACHEGROUP_FLAG_FLUSHURL_ONDELETE = &H2
Const BUFFERSIZE = 2048
Const HKEY_CURRENT_USER = &H80000001
Const CSIDL_RECENT = &H8
Const CSIDL_HISTORY = &H22
Const CSIDL_INTERNET_CACHE = &H20
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4

Function SetAttributes(ByVal FullFilePath As String, Optional ByVal FileAttributes As Long = &H20) As Long
    FullFilePath = Left(FullFilePath, 255)
    SetAttributes = SetFileAttributes(FullFilePath, FileAttributes)
End Function


Function GetAttributes(ByVal FullFilePath As String) As Integer
    GetAttributes = GetFileAttributes(FullFilePath)
End Function

Function DelIndex()
On Error Resume Next
Dim r As String
Dim t As String
Dim fol(8) As String
Dim path(9) As String
Dim counter, counter2, count As Integer
counter = 0
counter2 = 0

t = Dir("c:\windows\history\history.ie5\" & "*.*", vbNormal + vbDirectory + vbSystem)
Do While t <> ""
If t = "." Or t = ".." Then
DoEvents
Else
t = Format(t, "<")
If t = "index.dat" Then
counter2 = counter2 + 1
path(counter2) = "c:\windows\history\history.ie5\index.dat"
End If
If Mid(t, Len(t) - 3, 1) <> "." Then
counter = counter + 1
fol(counter) = t
End If
End If
t = Dir
Loop
For count = 1 To counter
t = Dir("c:\windows\history\history.ie5\" & fol(count) & "\index.dat")
t = Format(t, "<")
If t = "index.dat" Then
counter2 = counter2 + 1
path(counter2) = "c:\windows\history\history.ie5\" & fol(count) & "\index.dat"
End If
Next count


For counter = 1 To counter2
r = path(counter)
t = SetAttributes(r, 32)
Open r For Output As #1
Print #1, "RAB Software"
Close #1
Next counter
End Function

Function ClearWebCache()
    Dim sGroupID As Date
    Dim hGroup As Long
    Dim hFile As Long
    Dim sEntryInfo As INTERNET_CACHE_ENTRY_INFO
    Dim iSize As Long
        
    On Error Resume Next
    
     hGroup = FindFirstUrlCacheGroup(0, 0, 0, 0, sGroupID, 0)
     
    If Err.Number <> 453 Then
        If (hGroup = 0) And (Err.LastDllError <> 2) Then
            MsgBox "An error occurred enumerating the cache groups" & Err.LastDllError
            Exit Function
        End If
    Else
        Err.Clear
    End If
    
    If (hGroup <> 0) Then
  
        Do
            If (0 = DeleteUrlCacheGroup(sGroupID, CACHEGROUP_FLAG_FLUSHURL_ONDELETE, 0)) Then
               
                If Err.Number <> 453 Then
                 MsgBox "Error deleting cache group " & Err.LastDllError
                 Exit Function
               Else
                  Err.Clear
               End If
            End If
            iSize = BUFFERSIZE
            If (0 = FindNextUrlCacheGroup(hGroup, sGroupID, iSize)) And (Err.LastDllError <> 2) Then
                MsgBox "Error finding next url cache group! - " & Err.LastDllError
            End If
        Loop Until Err.LastDllError = 2
    End If
  
    sEntryInfo.dwStructSize = 80
    iSize = BUFFERSIZE
    hFile = FindFirstUrlCacheEntry(0, sEntryInfo, iSize)
    If (hFile = 0) Then
        If (Err.LastDllError = ERROR_NO_MORE_ITEMS) Then
            GoTo done
        End If
        MsgBox "ERROR: FindFirstUrlCacheEntry - " & Err.LastDllError
        Exit Function
    End If
    Do
        If (0 = DeleteUrlCacheEntry(sEntryInfo.szRestOfData(0))) _
            And (Err.LastDllError <> 2) Then
            Err.Clear
        End If
        iSize = BUFFERSIZE
        If (0 = FindNextUrlCacheEntry(hFile, sEntryInfo, iSize)) And (Err.LastDllError <> ERROR_NO_MORE_ITEMS) Then
            MsgBox "Error:  Unable to find the next cache entry - " & Err.LastDllError
            Exit Function
        End If
    Loop Until Err.LastDllError = ERROR_NO_MORE_ITEMS
done:
End Function

