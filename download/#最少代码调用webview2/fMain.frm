VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "webview2内核的网页浏览器"
   ClientHeight    =   8470
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   847
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1352
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   370
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   5290
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "打开百网址"
      Height          =   400
      Left            =   5760
      TabIndex        =   1
      Top             =   90
      Width           =   2130
   End
   Begin VB.PictureBox PicWeb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4460
      ScaleWidth      =   8060
      TabIndex        =   0
      Top             =   540
      Width           =   8055
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   45
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Web1 As cWebView2 'declare a WebView-variable WithEvents
Attribute Web1.VB_VarHelpID = -1
Private Sub Form_Load()
    Text1.Text = "https://www.baidu.com"
    Me.Visible = True
    Set Web1 = New_c.WebView2
    Dim LoadWebOk As Boolean, SdkPath As String
    SdkPath = "Z:\Wgx\webview2Win7绿色优化版109.0.1518.140"
    
    If SdkPath <> "" Then
        LoadWebOk = Web1.BindTo(PicWeb.hWnd, , SdkPath)
    Else
        LoadWebOk = Web1.BindTo(PicWeb.hWnd)
    End If
    If LoadWebOk = 0 Then
    'If Web1.BindTo(PicWeb.hWnd) = 0 Then '不用SDK目录直接加载
        MsgBox "初始化webview2失败入": Exit Sub
    End If
End Sub
    
    
Private Sub Command1_Click()
    Web1.Navigate "https://www.baidu.com", 0
End Sub

Private Sub cmdNavigate_Click()
 Web1.Navigate Text1.Text, 0
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        PicWeb.Move 10, PicWeb.Top, ScaleWidth - 20, ScaleHeight - PicWeb.Top - 10
        If Not Web1 Is Nothing Then Web1.SyncSizeToHostWindow
     End If
End Sub
 
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdNavigate_Click
End If
End Sub

Private Sub Web1_InitComplete()
 Web1.Navigate "https://www.baidu.com", 0
End Sub
