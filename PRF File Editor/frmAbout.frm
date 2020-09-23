VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About - PRFile Editor"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4740
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PrVbTool@hotmail.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1020
      MouseIcon       =   "frmAbout.frx":038A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2160
      Width           =   2715
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0694
      ForeColor       =   &H000000B9&
      Height          =   855
      Left            =   60
      TabIndex        =   5
      Top             =   2580
      Width           =   4635
   End
   Begin VB.Label lblBug 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If you find any bug in this software please contact with us at:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE3FE&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   4755
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4740
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblBy 
      BackStyle       =   0  'Transparent
      Caption         =   "By: PinkRabbit Soft"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EDC4FF&
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   3135
   End
   Begin VB.Label lblPRFileVer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Use PRFile version 1.00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EDC4FF&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "v 1.01 (beta)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FD6FFD&
      Height          =   255
      Left            =   2100
      TabIndex        =   1
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "PRFile Editor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EDC4FF&
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4755
   End
   Begin VB.Image imgBack 
      Height          =   3120
      Left            =   0
      Picture         =   "frmAbout.frx":0762
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4740
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblEMail_Click()
  ShellExecute Me.hwnd, "Open", "mailto:PrVbTool@Hotmail.com", "", "", 0
End Sub
