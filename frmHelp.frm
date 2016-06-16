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
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmHelp.Top = Form1.Top + (Form1.Height - frmHelp.Height) \ 2
    frmHelp.Left = Form1.Left + (Form1.Width - frmHelp.Width) \ 2
End Sub
