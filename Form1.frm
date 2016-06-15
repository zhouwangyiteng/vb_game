VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打气球"
   ClientHeight    =   7770
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   11145
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11145
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   10680
      Top             =   5280
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   10200
      Top             =   5280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9720
      Top             =   5280
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form1.frx":0000
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000015&
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      Picture         =   "Form1.frx":0006
      ScaleHeight     =   7815
      ScaleWidth      =   9615
      TabIndex        =   5
      Top             =   0
      Width           =   9615
      Begin VB.Image Image2 
         Height          =   975
         Index           =   0
         Left            =   9240
         Top             =   6840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   960
         Picture         =   "Form1.frx":87A2
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   2
      Left            =   9720
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   1
      Left            =   10560
      TabIndex        =   3
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   0
      Left            =   9720
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "time"
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Menu St 
      Caption         =   "开始"
      Begin VB.Menu run 
         Caption         =   "run"
      End
      Begin VB.Menu pa 
         Caption         =   "pause"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "工具"
      Begin VB.Menu calc 
         Caption         =   "计算器"
      End
      Begin VB.Menu note 
         Caption         =   "记事本"
      End
   End
   Begin VB.Menu Hp 
      Caption         =   "帮助"
      Begin VB.Menu help 
         Caption         =   "Help"
      End
      Begin VB.Menu ab 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time%, score%, head%, num%, i%
Dim start As Boolean

Private Sub ab_Click()
    frmAbout.Show
End Sub
Private Sub reset()
    For i = head To num - 1
        Unload Image2(i)
    Next i
    head = 1
    num = 1
    score = 0
    Command2.Enabled = False
    Command1.Enabled = True
    start = False
    Me.KeyPreview = True
    time = 60
    Label2(0).Caption = time
    Text1.Text = score
End Sub

Private Sub calc_Click()
    Shell "calc", vbNormalFocus
End Sub




Private Sub Command1_Click()
    Command1.Enabled = False
    start = True
    Command2.Enabled = True
End Sub

Private Sub Command2_Click()
    If start Then
        Command2.Caption = "继续"
        start = False
    Else
        Command2.Caption = "暂停"
        start = True
    End If
End Sub


Private Sub Command3_Click()
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If start Then
        If KeyAscii = Asc("a") Then
            Image1.Picture = LoadPicture("plane_l.gif")
            Image1.Left = Image1.Left - 100
        End If
        If KeyAscii = Asc("d") Then
            Image1.Picture = LoadPicture("plane_r.gif")
            Image1.Left = Image1.Left + 100
        End If
        If KeyAscii = Asc("j") Then
            Image1.Picture = LoadPicture("plane_l.gif")
            Image1.Left = Image1.Left - 1000
        End If
        If KeyAscii = Asc("k") Then
            Image1.Picture = LoadPicture("plane_r.gif")
            Image1.Left = Image1.Left + 1000
        End If
        If Image1.Left < Picture1.Left Then Image1.Left = Picture1.Left
        If Image1.Left + Image1.Width > Picture1.Left + Picture1.Width Then Image1.Left = Picture1.Left + Picture1.Width - Image1.Width
    End If
End Sub

Private Sub Form_Load()
    Form1.Top = Screen.Height / 8
    Form1.Left = Screen.Width / 4
    head = 1
    num = 1
    Image1.Picture = LoadPicture("plane_r.gif")
    Command2.Enabled = False
    start = False
    Me.KeyPreview = True
    Text1.FontSize = 17
    Text1.FontBold = True
    Label1.FontSize = 19
    Label1.ForeColor = vbRed
    Label1.FontBold = True
    time = 6
    Label2(0).Caption = time
    Label2(1).Caption = "秒"
    Label2(2).Caption = "SCORE"
    For i = 0 To 2
        Label2(i).Alignment = 2
        Label2(i).FontSize = 17
        Label2(i).ForeColor = vbRed
        Label2(i).FontBold = True
    Next i
    Text2.FontSize = 14
    Text2.FontBold = True
    Text2.Text = "a: 左移d: 右移j:左加速k:右加速"
    Text1.Text = score
End Sub

Private Sub note_Click()
    Shell "notepad", vbNormalFocus
End Sub

Private Sub pa_Click()
    Command2_Click
End Sub

Private Sub run_Click()
    Command1_Click
End Sub

Private Sub Timer1_Timer()
    If start Then
        Label2(0).Caption = time
        time = time - 1
        If time < 0 Then
            MsgBox "您的成绩为" & score, vbOKOnly, "游戏结果"
            reset
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If start Then
        Load Image2(num)
        Image2(num).Picture = LoadPicture("bouble.gif")
        Image2(num).Top = Picture1.Top + Picture1.Height - Image2(num).Height
        Image2(num).Left = Picture1.Width * Rnd + Picture1.Left
        If Image2(num).Left > Picture1.Width - Image2(num).Width Then Image2(num).Left = Picture1.Width - Image2(num).Width
        Image2(num).Visible = True
        num = num + 1
    End If
End Sub

Private Sub Timer3_Timer()
    If start Then
        For i = head To num - 1
            If Image1.Top < Image2(i).Top Then Image2(i).Top = Image2(i).Top - 30
            If Image2(i).Top >= Image1.Top And Image2(i).Top <= Image1.Height Then
                If Image2(i).Left >= Image1.Left And Image2(i).Left <= Image1.Left + Image1.Width Then
                    Image2(i).Top = Picture1.Height
                    score = score + 1
                    Text1.Text = score
                End If
                Unload Image2(i)
                head = head + 1
            End If
        Next i
    End If
End Sub
