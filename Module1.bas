Attribute VB_Name = "UrlMod"
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Dim Data() As Byte, Length As Long, MainUrl As String, DownUrl As String
Const CompStr = "href="""
Const SpecChr = """"

Private Function DownloadFile(URL As String, LocalFilename As String) As Boolean
Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Function GetUrls(sListBox As ListBox, URL As String, FName As Integer) As Boolean
On Error Resume Next
Dim fNo As Integer, SetA As Boolean, Filename As String
GetURL URL
Filename = App.path & "\HTML" & FName & ".tmp"
SetA = DownloadFile(MainUrl, Filename)
If SetA = True Then
    GetUrls = True
    fNo = FreeFile
    Open Filename For Binary As #fNo
        Length = LOF(fNo)
        ReDim Data(1 To Length)
        Get #fNo, , Data
    Close #fNo
    CollectData sListBox
Else
    GetUrls = False
    MsgBox "Url not found", vbExclamation
    Exit Function
End If
Kill Filename
End Function

Private Function GetURL(URL As String)
On Error Resume Next
Dim counter As Long, SetUrl As String

SetUrl = URL
If LCase(Mid(URL, 1, 4)) <> "http" Then SetUrl = "http://" & URL
If Mid(SetUrl, Len(SetUrl), 1) <> "/" Then SetUrl = SetUrl & "/"

MainUrl = SetUrl

For counter = Len(MainUrl) To 1 Step -1
    If Mid(MainUrl, counter, 1) = "/" Then
        DownUrl = Mid(MainUrl, 1, counter - 1)
        Exit Function
    End If
Next counter
End Function

Private Function CollectData(sListBox As ListBox)
On Error Resume Next
Dim counter As Long, Temp As Long, StrData(1 To 6) As String, count As Integer
sListBox.AddItem MainUrl
For counter = 1 To Length
DoEvents
    For count = 0 To 5
        StrData(count + 1) = LCase(Chr(Data(counter + count)))
    Next count
        If StrData(1) = "h" And StrData(2) = "r" And StrData(3) = "e" And StrData(4) = "f" And StrData(5) = "=" And StrData(6) = """" Then
        AddUrl counter + 6, sListBox
    End If
Next counter
End Function

Private Function AddUrl(No As Long, sListBox As ListBox)
On Error Resume Next
Dim UrlName As String
UrlName = GetName(No)
If Mid(UrlName, 1, 1) = "/" Then UrlName = Mid(UrlName, 2, Len(UrlName) - 1)
If LCase(Mid(UrlName, 1, 4)) <> "http" Then UrlName = DownUrl & "/" & UrlName
If Mid(UrlName, Len(UrlName), 1) <> "/" Then UrlName = UrlName & "/"
sListBox.AddItem UrlName
End Function

Private Function GetName(No As Long) As String
On Error Resume Next
Dim No2 As Long, Found As Boolean, counter As Long, Temp As String
Found = False
counter = No
Do While Found = False
    DoEvents
    counter = counter + 1
    Temp = Chr(Data(counter))
    If Temp = SpecChr Then
        No2 = counter - 1
        Found = True
    End If
Loop
For counter = No To No2
    GetName = GetName & Chr(Data(counter))
Next counter
End Function
