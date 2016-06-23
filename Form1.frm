VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打气球"
   ClientHeight    =   7770
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9720
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9720
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   9240
      Top             =   -120
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   8760
      Top             =   -120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8280
      Top             =   -120
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   8280
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form1.frx":0000
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000015&
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      Picture         =   "Form1.frx":0006
      ScaleHeight     =   7815
      ScaleWidth      =   8175
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      Begin VB.Image Image2 
         Height          =   975
         Index           =   0
         Left            =   8640
         Top             =   6600
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
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   9240
      TabIndex        =   11
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   8280
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   2
      Left            =   8280
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   1
      Left            =   9120
      TabIndex        =   3
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   0
      Left            =   8280
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "time"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Menu St 
      Caption         =   "游戏(&R)"
      Begin VB.Menu run 
         Caption         =   "run"
      End
      Begin VB.Menu pa 
         Caption         =   "pause"
      End
      Begin VB.Menu rec 
         Caption         =   "排行榜"
      End
   End
   Begin VB.Menu level 
      Caption         =   "难度(&L)"
      Begin VB.Menu L1 
         Caption         =   "Level 1"
      End
      Begin VB.Menu L2 
         Caption         =   "Level 2"
      End
      Begin VB.Menu L3 
         Caption         =   "Level 3"
      End
      Begin VB.Menu L4 
         Caption         =   "Level 4"
      End
      Begin VB.Menu L5 
         Caption         =   "Level 5"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "工具(&T)"
      Begin VB.Menu calc 
         Caption         =   "计算器"
      End
      Begin VB.Menu note 
         Caption         =   "记事本"
      End
   End
   Begin VB.Menu Hp 
      Caption         =   "帮助(&H)"
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
Dim time%, score%, head%, num%, i%, rate!, randnum!, l%, step%
Dim start As Boolean
Dim record(1 To 5) As String

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
    Label4.Caption = l
    step = 1000
    Label5.Caption = "高速"
    Command1.SetFocus
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
        If KeyAscii = Asc("d") Then
            Image1.Picture = LoadPicture("plane_l.gif")
            Image1.Left = Image1.Left - step
        End If
        If KeyAscii = Asc("f") Then
            Image1.Picture = LoadPicture("plane_r.gif")
            Image1.Left = Image1.Left + step
        End If
        If KeyAscii = Asc("j") Then
            Image1.Picture = LoadPicture("plane_l.gif")
            Image1.Left = Image1.Left - step
        End If
        If KeyAscii = Asc("k") Then
            Image1.Picture = LoadPicture("plane_r.gif")
            Image1.Left = Image1.Left + step
        End If
        If Image1.Left < Picture1.Left Then Image1.Left = Picture1.Left
        If Image1.Left + Image1.Width > Picture1.Left + Picture1.Width Then Image1.Left = Picture1.Left + Picture1.Width - Image1.Width
    End If
End Sub

Private Sub Form_Load()
    Open "record.txt" For Input As #1
    For i = 1 To 5
        Line Input #1, record(i)
    Next i
    Close 1
    l = 1
    setLevel (l)
    rate = 0.5
    Form1.Top = (Screen.Height - Form1.Height) \ 2
    Form1.Left = (Screen.Width - Form1.Width) \ 2
    head = 1
    num = 1
    Image1.Picture = LoadPicture("plane_r.gif")
    Command2.Enabled = False
    start = False
    Me.KeyPreview = True
    Text1.FontSize = 17
    Text1.FontBold = True
    Text1.Enabled = False
    Label1.FontSize = 19
    Label1.ForeColor = vbRed
    Label1.FontBold = True
    time = 60
    Label2(0).Caption = time
    Label2(1).Caption = "秒"
    Label2(2).Caption = "SCORE"
    For i = 0 To 2
        Label2(i).Alignment = 2
        Label2(i).FontSize = 17
        Label2(i).ForeColor = vbRed
        Label2(i).FontBold = True
    Next i
    Label3.Alignment = 2
    Label3.FontSize = 17
    Label3.ForeColor = vbRed
    Label3.FontBold = True
    Label3.Caption = "Level"
    Label4.Alignment = 2
    Label4.FontSize = 17
    Label4.ForeColor = vbRed
    Label4.FontBold = True
    Label4.Caption = l
    Text2.FontSize = 14
    Text2.FontBold = True
    Text2.Text = "d: 左移f: 右移j: 左移k: 右移"
    Label5.Alignment = 2
    Label5.FontSize = 17
    Label5.FontBold = True
    Label5.Caption = "高速"
    step = 1000
    Text1.Text = score
End Sub
Private Sub setLevel(l%)
    Select Case l
    Case 1
        Timer2.Interval = 700
        rate = 0.1
    Case 2
        Timer2.Interval = 600
        rate = 0.3
    Case 3
        Timer2.Interval = 500
        rate = 0.5
    Case 4
        Timer2.Interval = 300
        rate = 0.7
    Case 5
        Timer2.Interval = 200
        rate = 0.8
    End Select
End Sub

Private Sub help_Click()
    frmHelp.Show
End Sub

Private Sub L1_Click()
    setLevel (1)
    l = 1
    reset
End Sub

Private Sub L2_Click()
    setLevel (2)
    l = 2
    reset
End Sub
Private Sub L3_Click()
    setLevel (3)
    l = 3
    reset
End Sub
Private Sub L4_Click()
    setLevel (4)
    l = 4
    reset
End Sub
Private Sub L5_Click()
    setLevel (5)
    l = 5
    reset
End Sub

Private Sub note_Click()
    Shell "notepad", vbNormalFocus
End Sub

Private Sub pa_Click()
    If start Then
        Command2_Click
    End If
End Sub


Private Sub rec_Click()
    MsgBox "              最高分" & vbCrLf & "Level 1:    " & record(1) & vbCrLf & "Level 2:    " & record(2) & vbCrLf & "Level 3:    " & record(3) & vbCrLf & "Level 4:    " & record(4) & vbCrLf & "Level 5:    " & record(5), vbOKOnly, "排行榜"
End Sub

Private Sub run_Click()
    Command1_Click
End Sub
Private Sub gameover()
    MsgBox "您的成绩为" & score, vbOKOnly, "游戏结果"
    If (score > record(l)) Then
        MsgBox "恭喜你打破了Level" & l & "的记录", vbOKOnly, "新记录"
        record(l) = score
        Open "record.txt" For Output As #1
        For i = 1 To 5
            Print #1, record(i)
        Next i
        Close 1
    End If
    reset
End Sub

Private Sub Timer1_Timer()
    If start Then
        Label2(0).Caption = time
        time = time - 1
        If time < 0 Then
            gameover
            start = False
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If start Then
        Load Image2(num)
        randnum = Rnd
        If randnum < rate * 0.4 Then
            Image2(num).Picture = LoadPicture("balloon_r.gif")
            Image2(num).Tag = 1
        ElseIf randnum < rate Then
            Image2(num).Picture = LoadPicture("balloon_y.gif")
            Image2(num).Tag = 2
        ElseIf randnum < rate + (1 - rate) * 0.1 Then
            Image2(num).Picture = LoadPicture("balloon_b.gif")
            Image2(num).Tag = 3
        Else
            Image2(num).Picture = LoadPicture("balloon_g.gif")
            Image2(num).Tag = 4
        End If
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
                If Image2(i).Left >= Image1.Left - Image2(i).Width \ 2 And Image2(i).Left <= Image1.Left + Image1.Width - Image2(i).Width \ 2 Then
                    Select Case Image2(i).Tag
                    Case 1
                        gameover
                        Exit For
                    Case 2
                        If score > 0 Then score = score \ 2
                        step = 300
                        Label5.Caption = "慢速"
                    Case 3
                        If score > 0 Then
                            score = score * 2
                        Else
                            score = 1
                        End If
                        step = 1000
                        Label5.Caption = "高速"
                    Case 4
                        score = score + 1
                    End Select
                    Text1.Text = score
                End If
                Unload Image2(i)
                head = head + 1
            End If
        Next i
    End If
End Sub
