VERSION 5.00
Begin VB.Form frmAnimation 
   BackColor       =   &H00000000&
   Caption         =   "Animation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimation 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.PictureBox picAnimation 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------Learning How To Animate-----------

Private Sub Form_Load()
'The code written below will load picture into picture box picAnimation
picAnimation.Picture = LoadPicture(App.Path & "\circle.bmp")

End Sub

Private Sub tmrAnimation_Timer()
'This is the code which does the real job


If frmAnimation.ScaleHeight > picAnimation.Top Then
'The code entered below moves the picture downwards
'To increase the speed just increase the number from
'10 to 20 to anything you want
picAnimation.Top = picAnimation.Top + 10


'First of all it checks that whether the is
'inside the boundary of the form or not and
'if not this code does so.

Else

picAnimation.Top = 1

End If

End Sub
'Happy Programming
'Shashwat Srivastava
