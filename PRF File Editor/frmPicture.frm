VERSION 5.00
Begin VB.Form frmPicture 
   AutoRedraw      =   -1  'True
   Caption         =   "Picture - [no picture]"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "frmPicture.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   1
      Top             =   0
      Width           =   1800
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
  picPicture.Width = Me.ScaleWidth
  picPicture.Height = Me.ScaleHeight
  picPicture.Cls
  StretchBlt picPicture.hdc, 0, 0, picPicture.ScaleWidth, picPicture.ScaleHeight, picOriginal.hdc, 0, 0, picOriginal.ScaleWidth, picOriginal.ScaleHeight, vbSrcCopy
  picPicture.Refresh
End Sub
