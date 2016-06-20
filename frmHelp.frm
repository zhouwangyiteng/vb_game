VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   5385
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.Image Image2 
         Height          =   495
         Left            =   3840
         Top             =   4440
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   960
         Top             =   4440
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n%

Private Sub Form_Load()
    frmHelp.Top = Form1.Top + (Form1.Height - frmHelp.Height) \ 2
    frmHelp.Left = Form1.Left + (Form1.Width - frmHelp.Width) \ 2
    Picture1.Top = 0
    Picture1.Left = 0
    Image1.Picture = LoadPicture("arrow_l.gif")
    Image2.Picture = LoadPicture("arrow_r.gif")
    Image1.Left = Picture1.Left + 400
    Image2.Left = Picture1.Width - 400 - Image2.Width
    n = 0
    selectPic (n)
End Sub

Private Sub selectPic(i%)
    Select Case i
    Case 0
        Picture1.Picture = LoadPicture("help0.gif")
    Case 1
        Picture1.Picture = LoadPicture("help1.gif")
    Case 2
        Picture1.Picture = LoadPicture("help2.gif")
    End Select
End Sub

Private Sub Image1_Click()
    n = (n + 2) Mod 3
    selectPic (n)
End Sub

Private Sub Image2_Click()
    n = (n + 1) Mod 3
    selectPic (n)
End Sub

