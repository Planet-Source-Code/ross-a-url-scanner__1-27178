VERSION 5.00
Begin VB.Form FrmUrlSearch 
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   6570
   Begin VB.CommandButton ScanUrl 
      Caption         =   "Scan URL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   4560
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton OpenUrl 
      Caption         =   "Open URL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2280
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.ListBox URLList 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton Trace 
      Caption         =   "Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox URL 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label SName 
      Height          =   855
      Left            =   1800
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   160
      Width           =   855
   End
End
Attribute VB_Name = "FrmUrlSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = 7000
Me.Height = 5500
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim SetWidth As Long
If Me.WindowState <> vbMinimized Then
    If Me.Width < 7000 Then Me.Width = 7000
    If Me.Height < 5500 Then Me.Height = 5500
    URL.Width = Me.Width - 800
    URLList.Width = Me.Width - 320
    URLList.Height = Me.Height - 1800
    Trace.Top = Me.Height - 1200
    OpenUrl.Top = Me.Height - 1200
    ScanUrl.Top = Me.Height - 1200
    
    SetWidth = (Me.Width / 3) - 195
    Trace.Width = SetWidth
    OpenUrl.Width = SetWidth
    ScanUrl.Width = SetWidth
    Trace.Left = 120
    OpenUrl.Left = (SetWidth * 2) + 360
    ScanUrl.Left = SetWidth + 240
End If
End Sub

Private Sub OpenUrl_Click()
On Error Resume Next
If URLList.ListCount > 0 Then
    If URLList.Text = "" Then Exit Sub
    r = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & URLList.Text, vbMaximizedFocus)
Else
    r = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & URL.Text, vbMaximizedFocus)
End If
End Sub

Private Sub ScanUrl_Click()
Dim UrlTxt As String
UrlTxt = URLList.Text
If Len(UrlTxt) > 0 Then FrmUrls.LoadNew UrlTxt
End Sub

Private Sub Trace_Click()
Dim Answ As Boolean
URLList.Clear
Me.Caption = UCase(URL.Text)
Answ = GetUrls(URLList, URL.Text, SName.Caption)
If Answ = True Then
    Me.Caption = UCase(URL.Text) & " - Complete"
Else
    Me.Caption = UCase(URL.Text) & " - Error"
End If
End Sub

Private Sub URLList_DblClick()
On Error Resume Next
If URLList.Text = "" Then Exit Sub
r = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & URLList.Text, vbMaximizedFocus)
End Sub

Private Sub URLList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu popup, , X + URLList.Left, Y + URLList.Top

End If
End Sub
