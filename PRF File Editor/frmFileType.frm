VERSION 5.00
Begin VB.Form frmFileType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File type"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2655
   Icon            =   "frmFileType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1380
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1275
   End
   Begin VB.ComboBox cboKind 
      Height          =   315
      ItemData        =   "frmFileType.frx":038A
      Left            =   60
      List            =   "frmFileType.frx":039D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
   Begin VB.ListBox lstVideo 
      Height          =   2595
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   420
      Width           =   2535
   End
   Begin VB.ListBox lstVarious 
      Height          =   2595
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   420
      Width           =   2535
   End
   Begin VB.ListBox lstText 
      Height          =   2595
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   420
      Width           =   2535
   End
   Begin VB.ListBox lstSound 
      Height          =   2595
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   420
      Width           =   2535
   End
   Begin VB.ListBox lstPicture 
      Height          =   2595
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   2535
   End
End
Attribute VB_Name = "frmFileType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboKind_Click()
  lstPicture.ListIndex = -1
  lstSound.ListIndex = -1
  lstText.ListIndex = -1
  lstVarious.ListIndex = -1
  lstVideo.ListIndex = -1
  cmdRemove.Enabled = False
  Select Case cboKind.ListIndex
    Case 0
      lstPicture.ZOrder
    Case 1
      lstSound.ZOrder
    Case 2
      lstText.ZOrder
    Case 3
      lstVarious.ZOrder
    Case 4
      lstVideo.ZOrder
  End Select
End Sub

Private Sub cmdAdd_Click()
  Dim Ext As String
  
  Select Case cboKind.ListIndex
    Case 0
      Ext = InputBox("New extention for picture", "Picture")
      If Ext <> "" And CheckExtExist(Ext) Then lstPicture.AddItem Ext
    Case 1
      Ext = InputBox("New extention for sound", "Sound")
      If Ext <> "" And CheckExtExist(Ext) Then lstSound.AddItem Ext
    Case 2
      Ext = InputBox("New extention for text", "Text")
      If Ext <> "" And CheckExtExist(Ext) Then lstText.AddItem Ext
    Case 3
      Ext = InputBox("New extention for various", "Various")
      If Ext <> "" And CheckExtExist(Ext) Then lstVarious.AddItem Ext
    Case 4
      Ext = InputBox("New extention for Video", "Video")
      If Ext <> "" And CheckExtExist(Ext) Then lstVideo.AddItem Ext
  End Select
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  ReDim PictureExt(lstPicture.ListCount - 1)
  For i = 0 To lstPicture.ListCount - 1
    PictureExt(i) = lstPicture.List(i)
  Next i
  
  ReDim SoundExt(lstSound.ListCount - 1)
  For i = 0 To lstSound.ListCount - 1
    SoundExt(i) = lstSound.List(i)
  Next i

  ReDim TextExt(lstText.ListCount - 1)
  For i = 0 To lstText.ListCount - 1
    TextExt(i) = lstText.List(i)
  Next i

  ReDim VariousExt(lstVarious.ListCount - 1)
  For i = 0 To lstVarious.ListCount - 1
    VariousExt(i) = lstVarious.List(i)
  Next i

  ReDim VideoExt(lstVideo.ListCount - 1)
  For i = 0 To lstVideo.ListCount - 1
    VideoExt(i) = lstVideo.List(i)
  Next i

  SaveBindExt
  
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  Select Case cboKind.ListIndex
    Case 0
      lstPicture.RemoveItem lstPicture.ListIndex
    Case 1
      lstSound.RemoveItem lstSound.ListIndex
    Case 2
      lstText.RemoveItem lstText.ListIndex
    Case 3
      lstVarious.RemoveItem lstVarious.ListIndex
    Case 4
      lstVideo.RemoveItem lstVideo.ListIndex
  End Select
  cmdRemove.Enabled = False
End Sub

Private Sub Form_Load()
  For i = 0 To UBound(PictureExt)
    If PictureExt(i) <> "" Then lstPicture.AddItem PictureExt(i)
  Next i

  For i = 0 To UBound(SoundExt)
    If SoundExt(i) <> "" Then lstSound.AddItem SoundExt(i)
  Next i

  For i = 0 To UBound(TextExt)
    If TextExt(i) <> "" Then lstText.AddItem TextExt(i)
  Next i

  For i = 0 To UBound(VariousExt)
    If VariousExt(i) <> "" Then lstVarious.AddItem VariousExt(i)
  Next i

  For i = 0 To UBound(VideoExt)
    If VideoExt(i) <> "" Then lstVideo.AddItem VideoExt(i)
  Next i
  
  cboKind.ListIndex = 0
End Sub

Private Sub lstPicture_Click()
  cmdRemove.Enabled = True
End Sub

Private Sub lstSound_Click()
  cmdRemove.Enabled = True
End Sub

Private Sub lstText_Click()
  cmdRemove.Enabled = True
End Sub

Private Sub lstVarious_Click()
  cmdRemove.Enabled = True
End Sub

Private Sub lstVideo_Click()
  cmdRemove.Enabled = True
End Sub

Function CheckExtExist(Ext As String) As Boolean
  Dim Exist As Boolean
  
  For i = 0 To lstPicture.ListCount - 1
    If lstPicture.List(i) = Ext Then Exist = True
  Next i
  
  For i = 0 To lstSound.ListCount - 1
    If lstSound.List(i) = Ext Then Exist = True
  Next i

  For i = 0 To lstText.ListCount - 1
    If lstText.List(i) = Ext Then Exist = True
  Next i

  For i = 0 To lstVarious.ListCount - 1
    If lstVarious.List(i) = Ext Then Exist = True
  Next i

  For i = 0 To lstVideo.ListCount - 1
    If lstVideo.List(i) = Ext Then Exist = True
  Next i

  If Exist Then MsgBox "Extension already exist in one of these list"
  
  CheckExtExist = Not Exist
End Function
