VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug window"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4995
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDebug 
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4995
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
