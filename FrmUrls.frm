VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm FrmUrls 
   BackColor       =   &H8000000C&
   Caption         =   "URL Search"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "FrmUrls.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Savefile 
      Left            =   2640
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu save 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu options 
      Caption         =   "O&ptions"
      Begin VB.Menu history 
         Caption         =   "Delete &History"
         Shortcut        =   ^H
      End
      Begin VB.Menu deltemp 
         Caption         =   "Dele&te Tempory Files"
         Shortcut        =   ^T
      End
      Begin VB.Menu delind 
         Caption         =   "Delete &Index.dat Files"
         Shortcut        =   ^I
      End
      Begin VB.Menu delurl 
         Caption         =   "Delete Typed &URL's"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu tilev 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu tileh 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu arrange 
         Caption         =   "&Arrange Icons"
      End
   End
End
Attribute VB_Name = "FrmUrls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lSearchCount As Long, frm As FrmUrlSearch, DocCount As Integer

Function LoadNew(Optional URL As String)
DocCount = DocCount + 1
Set frm = New FrmUrlSearch
frm.Caption = "URL Search"
frm.Show
frm.URL.Text = URL
frm.SName.Caption = DocCount
End Function

Private Sub arrange_Click()
Me.arrange vbArrangeIcons
End Sub

Private Sub cascade_Click()
Me.arrange vbCascade
End Sub

Private Sub delind_Click()
DelIndex
Beep
End Sub

Private Sub deltemp_Click()
ClearWebCache
Beep
End Sub

Private Sub delurl_Click()
RegDeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs"
Beep
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub MDIForm_Load()
DocCount = 0
LoadNew
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub New_Click()
LoadNew
End Sub

Private Sub save_Click()
Dim sFile As String
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.URLList.ListCount = 0 Then Exit Sub
    With Savefile
        .DialogTitle = "Save As"
        .CancelError = False
        .Flags = cdlOFNHideReadOnly
        .Filter = "Text Files (*.txt)|*.txt"
        .ShowSave
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
    End With
SaveDetails sFile
End Sub

Function SaveDetails(Filename As String)
Dim counter As Long, fNo As Integer, Data As String
fNo = FreeFile
Open Filename For Output As #fNo
For counter = 0 To (ActiveForm.URLList.ListCount - 1)
    Data = ActiveForm.URLList.List(counter)
    Print #fNo, Data
Next counter
Close #fNo
End Function

Private Sub tileh_Click()
Me.arrange vbTileHorizontal
End Sub

Private Sub tilev_Click()
Me.arrange vbTileVertical
End Sub
