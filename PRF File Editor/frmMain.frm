VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRFile Editor"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   8865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList iltFolder 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":178C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":225A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":298E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":345C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   6255
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   240
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picEmpty 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   180
      ScaleHeight     =   225
      ScaleWidth      =   1785
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label lblEmpty 
         Alignment       =   2  'Center
         Caption         =   "File list is empty"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   15
         Width           =   1080
      End
   End
   Begin VB.FileListBox File 
      Height          =   285
      Left            =   30
      TabIndex        =   22
      Top             =   15
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSComctlLib.TreeView trvKind 
      Height          =   6255
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   11033
      _Version        =   393217
      Indentation     =   176
      Style           =   7
      ImageList       =   "iltFolder"
      Appearance      =   1
   End
   Begin VB.Frame fraMain 
      Height          =   6255
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      Begin VB.Frame fraPreview 
         Height          =   4155
         Left            =   60
         TabIndex        =   12
         Top             =   2040
         Width           =   6555
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find app"
            Height          =   240
            Left            =   1380
            TabIndex        =   25
            Top             =   0
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CommandButton cmdPreview 
            Caption         =   "Show preview"
            Height          =   240
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   1275
         End
         Begin VB.PictureBox picText 
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   60
            ScaleHeight     =   3855
            ScaleWidth      =   6435
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   6435
            Begin RichTextLib.RichTextBox rtfText 
               Height          =   3555
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   6271
               _Version        =   393217
               Enabled         =   -1  'True
               TextRTF         =   $"frmMain.frx":465E
            End
            Begin VB.Label lblNbChar 
               Caption         =   "Nb Char: 0"
               Height          =   255
               Left            =   0
               TabIndex        =   21
               Top             =   3600
               Width           =   5955
            End
         End
         Begin VB.PictureBox picVideo 
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   60
            ScaleHeight     =   3855
            ScaleWidth      =   6435
            TabIndex        =   17
            Top             =   240
            Visible         =   0   'False
            Width           =   6435
         End
         Begin VB.PictureBox picSound 
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   60
            ScaleHeight     =   3855
            ScaleWidth      =   6435
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   6435
         End
         Begin VB.PictureBox picPicture 
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   60
            ScaleHeight     =   3855
            ScaleWidth      =   6435
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   6435
            Begin VB.PictureBox picOriginal 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   375
               Left            =   900
               ScaleHeight     =   21
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   19
               Top             =   600
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.PictureBox picPic 
               AutoRedraw      =   -1  'True
               Height          =   3615
               Left            =   0
               ScaleHeight     =   237
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   425
               TabIndex        =   18
               Top             =   0
               Width           =   6435
            End
            Begin VB.Label lblSize 
               Caption         =   "Size: 0x0 (pixel)"
               Height          =   255
               Left            =   0
               TabIndex        =   20
               Top             =   3600
               Width           =   5595
            End
         End
      End
      Begin VB.Frame fraHeader 
         Caption         =   "Header     "
         Height          =   1875
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   6555
         Begin VB.Frame fraInfo 
            Height          =   1215
            Left            =   60
            TabIndex        =   7
            Top             =   600
            Width           =   6435
            Begin VB.Label lblFileNo 
               Caption         =   "File no: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   420
               Width           =   6195
            End
            Begin VB.Label lblEndAt 
               Caption         =   "End at: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   900
               Width           =   6255
            End
            Begin VB.Label lblStartAt 
               Caption         =   "Start at: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   660
               Width           =   6255
            End
            Begin VB.Label lblFileSize 
               Caption         =   "File size: 0 (0Kb)"
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Top             =   180
               Width           =   6195
            End
         End
         Begin VB.ComboBox cboKind 
            Height          =   315
            ItemData        =   "frmMain.frx":46E0
            Left            =   4440
            List            =   "frmMain.frx":46F6
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Height          =   240
            Left            =   3330
            TabIndex        =   4
            Top             =   270
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtFileName 
            Height          =   285
            Left            =   840
            TabIndex        =   3
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblKind 
            Caption         =   "File type:"
            Height          =   255
            Left            =   3780
            TabIndex        =   5
            Top             =   300
            Width           =   675
         End
         Begin VB.Label lblFileName 
            Caption         =   "File name:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   300
            Width           =   1635
         End
      End
   End
   Begin MediaPlayerCtl.MediaPlayer mprVideo 
      Height          =   315
      Left            =   540
      TabIndex        =   29
      Top             =   5460
      Visible         =   0   'False
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer mprSound 
      Height          =   375
      Left            =   1500
      TabIndex        =   28
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBindExt 
         Caption         =   "Bind extention"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Library data"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add file"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove file"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContent 
         Caption         =   "Content"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "                           Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuShowTFH 
         Caption         =   "Show header file"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempRcFile As String
Dim MainHeader As PRF_Header
Dim FileHeader() As TempFileHeader
Dim CurIndex As Integer
Dim RcFileName As String

'Object procedure

Private Sub cboKind_Click()
  If CurIndex = -1 Then Exit Sub
  
  FileHeader(CurIndex).FH.FileType = cboKind.ListIndex
  
  RefreshFileList
End Sub

Private Sub cmdFind_Click()
  Dim TheFile As String
  
  TheFile = OriExtention(FileHeader(CurIndex).TempFile)
  
  FileCopy FileHeader(CurIndex).TempFile, TheFile
  
  retval = ShellExecute(Me.hwnd, "Open", TheFile, "", 0, 1)
  
  If retval = 0 Then MsgBox "Error in Windows API"
  If retval = 31 Then MsgBox "Sorry! It's seem that nothing on your computer can open this" & vbNewLine & vbNewLine & "(uncertain function: check if no app has open your file" & vbNewLine & "so if you see this for nothing please repport)"
End Sub

Private Sub cmdPreview_Click()
  OnWork "Please wait while preparing preview"
  cmdPreview.Enabled = False
  Select Case FileHeader(CurIndex).FH.FileType
    Case 0  'Text
      OpenTextFile (FileHeader(CurIndex).TempFile)
    Case 1 'Picture
      OpenPictureFile (FileHeader(CurIndex).TempFile)
    Case 2 'Sound
      OpenSoundFile (FileHeader(CurIndex).TempFile)
    Case 3 'Text
      OpenTextFile (FileHeader(CurIndex).TempFile)
    Case 4 'Text
      OpenTextFile (FileHeader(CurIndex).TempFile)
    Case 5 'Video
      OpenVideoFile (FileHeader(CurIndex).TempFile)
  End Select
  StopWork
End Sub

Private Sub Form_Load()
  App.Title = "PRFile Editor"
  
  OnWork "Search for temp ressource directory"
  If Dir(App.Path & "\TempRc", vbDirectory) = "" Then MkDir App.Path & "\TempRc"
  StopWork
  
  OnWork "Search for bind extention setting and read it"
  If Dir(App.Path & "\BindExt.pfe") = "" Then CreateDefault
  LoadBindExt
  StopWork
  
  Me.Show
  Me.Refresh
  If Dir(App.Path & "\TempRc\*.*") <> "" Then
    If MsgBox("An abnormal close has been detected. Try to recover last PRF project?", vbQuestion + vbYesNo, "Abnormal close") = vbYes Then
      If TryLoadFromAC Then
        MsgBox "Recovering last project successful"
        RefreshFileList
      Else
        Kill App.Path & "\TempRc\*.*"
        mnuNew_Click
      End If
    Else
      Kill App.Path & "\TempRc\*.*"
      mnuNew_Click
    End If
  Else
    mnuNew_Click
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Terminate
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub cmdOK_Click()
  Dim StrLen As Long
  
  FileHeader(CurIndex).FH.FileName = txtFileName.Text
  
  SizeRc
  
  OnWork "Write new file header"
  WriteTempRc
  StopWork
  
  RefreshFileList
  
  cmdOK.Visible = False
End Sub

Private Sub mnuAdd_Click()
  Dim StrLen As Integer
  Dim TempFileName As String
  Dim Index As Integer
  Dim Fno As Integer
  Dim TheFile As String
  Dim FoundType As Integer
  
  On Error Resume Next
  cdlBrowse.Filter = "All file (*.*)"
  cdlBrowse.DialogTitle = "Choose a file to add in ressource file"
  cdlBrowse.FileName = ""
  cdlBrowse.ShowOpen
  If Err.Number <> 0 Or Dir(cdlBrowse.FileName) = "" Then Exit Sub
  On Error GoTo 0
  
  OnWork "Check if file already exist"
  For i = 0 To MainHeader.NbFile - 1
    If FileHeader(i).FH.FileName = GetFileName(cdlBrowse.FileName) Then
      MsgBox "File already exist in this ressource file"
      StopWork
      Exit Sub
    End If
  Next i
  StopWork
  
  OnWork "Creating unexisting temp file name and copy original file"
  Do
    TempFileName = App.Path & "\TempRc\TempFile" & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & ".tmp"
  Loop Until Dir(TempFileName) = ""
  FileCopy cdlBrowse.FileName, TempFileName
  StopWork
  
  Index = MainHeader.NbFile
  
  OnWork "Find the best file type"
  TheFile = LCase(cdlBrowse.FileName)
  
  FoundType = -1
  
  For i = 0 To UBound(PictureExt)
    If Mid(TheFile, Len(TheFile) + 1 - Len(PictureExt(i))) = PictureExt(i) And PictureExt(i) <> "" Then FoundType = 1
  Next i
  
  For i = 0 To UBound(SoundExt)
    If Mid(TheFile, Len(TheFile) + 1 - Len(SoundExt(i))) = SoundExt(i) And SoundExt(i) <> "" Then FoundType = 2
  Next i

  For i = 0 To UBound(TextExt)
    If Mid(TheFile, Len(TheFile) + 1 - Len(TextExt(i))) = TextExt(i) And TextExt(i) <> "" Then FoundType = 3
  Next i
    
  For i = 0 To UBound(VariousExt)
    If Mid(TheFile, Len(TheFile) + 1 - Len(VariousExt(i))) = VariousExt(i) And VariousExt(i) <> "" Then FoundType = 4
  Next i
    
  For i = 0 To UBound(VideoExt)
    If Mid(TheFile, Len(TheFile) + 1 - Len(VideoExt(i))) = VideoExt(i) And VideoExt(i) <> "" Then FoundType = 5
  Next i
    
  If FoundType = -1 Then FoundType = 0
  StopWork
  
  OnWork "Create file header"
  ReDim Preserve FileHeader(Index)
  FileHeader(Index).TempFile = TempFileName
  FileHeader(Index).FH.FileName = GetFileName(cdlBrowse.FileName)
  FileHeader(Index).FH.StartAt = MainHeader.FileLenght - FileLen(cdlBrowse.FileName)
  FileHeader(Index).FH.EndAt = MainHeader.FileLenght
  FileHeader(Index).FH.FileType = FoundType
  FileHeader(Index).FH.FileLenght = FileLen(cdlBrowse.FileName)
  StopWork
  
  OnWork "Deleting temp ressource file and rewrite new main file header"
  Kill TempRcFile
  StopWork
  
  MainHeader.NbFile = MainHeader.NbFile + 1
  
  WriteTempRc
  
  SizeRc
  
  RefreshFileList
  
  SelFile (FileHeader(Index).FH.FileName)
End Sub

Private Sub mnuBindExt_Click()
  frmFileType.Show
End Sub

Private Sub mnuContent_Click()
  MsgBox "Sorry!!! Not done yet"
End Sub

Private Sub mnuData_Click()
  On Error Resume Next
  retval = trvKind.SelectedItem.Index
  If Err.Number <> 0 Or trvKind.SelectedItem.Key <> "" Then
    mnuRemove.Enabled = False
    mnuExport.Enabled = False
  Else
    mnuRemove.Enabled = True
    mnuExport.Enabled = True
  End If
End Sub

Private Sub mnuExport_Click()
  Dim ExportTo As String
  
  On Error Resume Next
  cdlBrowse.DialogTitle = "Choose a file to export this file"
  cdlBrowse.Filter = "All file (*.*)|*.*"
  cdlBrowse.FileName = ""
  cdlBrowse.ShowSave
  If Err.Number <> 0 Then Exit Sub
  If Dir(cdlBrowse.FileName) <> "" Then If MsgBox("File already exist. Overwrite?", vbQuestion + vbYesNo, "File exist") = vbNo Then Exit Sub Else Kill cdlBrowse.FileName
  
  If Right(cdlBrowse.FileName, Len(OriExtention(FileHeader(CurIndex).TempFile, True))) <> OriExtention(FileHeader(CurIndex).TempFile, True) Then cdlBrowse.FileName = cdlBrowse.FileName & "." & OriExtention(FileHeader(CurIndex).TempFile, True)
  
  FileCopy FileHeader(CurIndex).TempFile, cdlBrowse.FileName
End Sub

Private Sub mnuFile_Click()
  mnuSave.Enabled = MainHeader.NbFile
  mnuSaveAs.Enabled = MainHeader.NbFile
End Sub

Private Sub mnuNew_Click()
  Dim Fno As Integer
  
  ResetInterface
  
  RcFileName = ""
  
  OnWork "Delete old temp ressource file"
  On Error Resume Next
  Kill App.Path & "\TempRc\*.*"
  On Error GoTo 0
  StopWork
  
  ReDim FileHeader(0)
  
  OnWork "Creating random temp ressource file name"
  TempRcFile = App.Path & "\TempRc\TempPRFile" & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & ".tmp"
  StopWork
  
  OnWork "Create temp ressource file and main file header"
  MainHeader.PRFversion = PRF_Version
  MainHeader.NbFile = 0
  MainHeader.FileLenght = 20
  Fno = FreeFile
  Open TempRcFile For Binary Access Write Lock Read Write As Fno
    Put Fno, , MainHeader
  Close Fno
  StopWork
  
  OnWork "Empting previous file list"
  trvKind.Nodes.Clear
  trvKind.Nodes.Add , , "Other", "Other"
  trvKind.Nodes.Add , , "Picture", "Picture"
  trvKind.Nodes.Add , , "Sound", "Sound"
  trvKind.Nodes.Add , , "Text", "Text"
  trvKind.Nodes.Add , , "Various", "Various"
  trvKind.Nodes.Add , , "Video", "Video"
  StopWork
  
  RefreshFileList
End Sub

Private Sub mnuOpen_Click()
  On Error Resume Next
  cdlBrowse.DialogTitle = "Choose PRFile to open"
  cdlBrowse.Filter = "PRF ressource file (*.RAB)|*.RAB"
  cdlBrowse.FileName = ""
  cdlBrowse.ShowOpen
  If Err.Number <> 0 Or Dir(cdlBrowse.FileName) = "" Then Exit Sub
  On Error GoTo 0
  
  mnuNew_Click
  
  RcFileName = cdlBrowse.FileName
  
  Dim Fno As Integer
  Dim FnoF As Integer
  Dim CharArray() As Byte
  
  Fno = FreeFile
  
  OnWork "Opening PRFile and reading mainheader and fileheader"
  Open RcFileName For Binary Access Read Lock Read Write As Fno
    Get Fno, 1, MainHeader
    
    If MainHeader.PRFversion <> PRF_Version Then MsgBox "Wrong PRFile version. Can't open this file": Close: Exit Sub
    If MainHeader.FileLenght <> FileLen(RcFileName) Then MsgBox "PRFile size info mismatch. Can't open this file": Close: Exit Sub
    
    ReDim FileHeader(MainHeader.NbFile - 1)
    
    For i = 0 To MainHeader.NbFile - 1
      Get Fno, , FileHeader(i).FH
      Do
        FileHeader(i).TempFile = App.Path & "\TempRc\TempFile" & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & ".tmp"
      Loop Until Dir(FileHeader(i).TempFile) = ""
    Next i
    StopWork
    
    For i = 0 To MainHeader.NbFile - 1
      
      OnWork "Read file " & FileHeader(i).FH.FileName & " and write to " & DoShortPath(FileHeader(i).TempFile, i)
      ReDim CharArray(FileHeader(i).FH.FileLenght - 1)
      Get Fno, FileHeader(i).FH.StartAt, CharArray
      
      FnoF = FreeFile
      Open FileHeader(i).TempFile For Binary Access Write Lock Read Write As FnoF
        Put FnoF, 1, CharArray
      Close FnoF
      StopWork
      
    Next i
    
  Close Fno
  
  WriteTempRc
  
  RefreshFileList
  
  SelFile (trvKind.Nodes.Item(1).Child.Text)

End Sub

Private Sub mnuQuit_Click()
  Terminate
End Sub

Private Sub mnuRemove_Click()
  If MainHeader.NbFile = 1 Then MsgBox "Can't remove last file": Exit Sub
  
  Dim tmpFH() As TempFileHeader
  Dim ii As Integer
  
  Kill FileHeader(CurIndex).TempFile
  
  ReDim tmpFH(MainHeader.NbFile - 1)
  
  OnWork "Creating a list of file that will not be deleted"
  For i = CurIndex + 1 To MainHeader.NbFile - 1
    tmpFH(i).TempFile = FileHeader(i).TempFile
    tmpFH(i).FH.EndAt = FileHeader(i).FH.EndAt
    tmpFH(i).FH.FileLenght = FileHeader(i).FH.FileLenght
    tmpFH(i).FH.FileName = FileHeader(i).FH.FileName
    tmpFH(i).FH.FileType = FileHeader(i).FH.FileType
    tmpFH(i).FH.StartAt = FileHeader(i).FH.StartAt
  Next i
  StopWork
  
  MainHeader.NbFile = MainHeader.NbFile - 1
  
  ReDim Preserve FileHeader(MainHeader.NbFile - 1)
  
  OnWork "Restore file to library list"
  For i = CurIndex To MainHeader.NbFile - 1
    FileHeader(i).TempFile = tmpFH(i + 1).TempFile
    FileHeader(i).FH.EndAt = tmpFH(i + 1).FH.EndAt
    FileHeader(i).FH.FileLenght = tmpFH(i + 1).FH.FileLenght
    FileHeader(i).FH.FileName = tmpFH(i + 1).FH.FileName
    FileHeader(i).FH.FileType = tmpFH(i + 1).FH.FileType
    FileHeader(i).FH.StartAt = tmpFH(i + 1).FH.StartAt
  Next i
  StopWork
  
  If CurIndex <> 0 Then CurIndex = CurIndex - 1
  
  SizeRc
  
  WriteTempRc
  
  RefreshFileList
End Sub

Private Sub mnuSave_Click()
  If RcFileName = "" Then SaveAs
  If RcFileName = "" Then Exit Sub
  SaveOutputFile RcFileName
End Sub

Private Sub mnuSaveAs_Click()
  SaveAs
  mnuSave_Click
End Sub

Private Sub mnuShowTFH_Click()
  Dim txt As String
  frmDebug.Show
  txt = "mainheader.FileLenght = " & MainHeader.FileLenght & vbNewLine
  txt = txt & "mainheader.NbFile = " & MainHeader.NbFile & vbNewLine & "- - - - - - - - - - - -" & vbNewLine
  For i = 0 To MainHeader.NbFile - 1
    txt = txt & "fileheader(" & i & ").FH.EndAt = " & FileHeader(i).FH.EndAt & vbNewLine
    txt = txt & "fileheader(" & i & ").FH.FileLenght = " & FileHeader(i).FH.FileLenght & vbNewLine
    txt = txt & "fileheader(" & i & ").FH.FileName = " & FileHeader(i).FH.FileName & vbNewLine
    txt = txt & "fileheader(" & i & ").FH.FileType = " & FileHeader(i).FH.FileType & vbNewLine
    txt = txt & "fileheader(" & i & ").FH.StartAt = " & FileHeader(i).FH.StartAt & vbNewLine
    txt = txt & "- - - - - - - - - - - -" & vbNewLine
  Next i
  frmDebug.txtDebug.Text = txt
End Sub

Private Sub picPic_Click()
  Dim AddX As Integer, AddY As Integer
  Load frmPicture
  With frmPicture
    .picOriginal.Width = picOriginal.ScaleWidth
    .picOriginal.Height = picOriginal.ScaleHeight
    .picOriginal.Picture = picOriginal.Picture
    AddX = (.Width - .ScaleWidth) / 15
    AddY = (.Height - .ScaleHeight) / 15
    .Width = 15 * picOriginal.ScaleWidth + AddX
    .Height = 15 * picOriginal.ScaleHeight + AddY
    .Show
    .Caption = "Picture - " & txtFileName.Text
  End With
End Sub

Private Sub trvKind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then PopupMenu mnuData
End Sub

Private Sub trvKind_NodeClick(ByVal Node As MSComctlLib.Node)
  
  ResetInterface

  If Node.Key = "" Then
  
    Dim Index As Integer
    Dim TmpF As String
    Dim SelFile As String
    
    SelFile = Node.Text
    
    CurIndex = -1
    
    For i = 0 To MainHeader.NbFile - 1
      If FileHeader(i).FH.FileName = SelFile Then Index = i: Exit For
    Next i
    
    txtFileName.Text = FileHeader(Index).FH.FileName: cmdOK.Visible = False
    cboKind.ListIndex = FileHeader(Index).FH.FileType
    lblFileSize.Caption = "File size: " & SelBestUnit(FileHeader(Index).FH.FileLenght)
    lblFileNo.Caption = "File no: " & Index + 1
    lblStartAt.Caption = "Start at: " & FileHeader(Index).FH.StartAt
    lblEndAt.Caption = "End at: " & FileHeader(Index).FH.EndAt
    
    CurIndex = Index
    
    Select Case FileHeader(Index).FH.FileType
      Case 0  'Text
        picText.Visible = True
        picText.ZOrder
        cmdFind.Visible = True
      Case 1 'Picture
        picPicture.Visible = True
        picPicture.ZOrder
      Case 2 'Sound
        picSound.Visible = True
        picSound.ZOrder
      Case 3 'Text
        picText.Visible = True
        picText.ZOrder
      Case 4 'Text
        picText.Visible = True
        picText.ZOrder
        cmdFind.Visible = True
      Case 5 'Video
        picVideo.Visible = True
        picVideo.ZOrder
    End Select
    
    cmdPreview.Enabled = True
    
    stbStatus.SimpleText = SelFile & " (" & DoShortPath(FileHeader(CurIndex).TempFile, CurIndex) & ")"
  Else
    CurIndex = -1
  End If
End Sub

Private Sub txtFileName_Change()
  cmdOK.Visible = True
End Sub


'Function
Function GetFileName(FilePath As String) As String
  GetFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
End Function

Function RefreshFileList()
  Dim CFile As Boolean
  
  OnWork "Refresh file list"
  trvKind.Nodes.Clear
  trvKind.Nodes.Add , , "Other", "Other", 1, 1
  trvKind.Nodes.Add , , "Picture", "Picture", 2, 2
  trvKind.Nodes.Add , , "Sound", "Sound", 3, 3
  trvKind.Nodes.Add , , "Text", "Text", 4, 4
  trvKind.Nodes.Add , , "Various", "Various", 5, 5
  trvKind.Nodes.Add , , "Video", "Video", 6, 6
  For i = 0 To MainHeader.NbFile - 1
    Select Case FileHeader(i).FH.FileType
      Case 0 'Other
        trvKind.Nodes.Add "Other", tvwChild, , FileHeader(i).FH.FileName, 7, 13
      Case 1 'Picture
        trvKind.Nodes.Add "Picture", tvwChild, , FileHeader(i).FH.FileName, 8, 14
      Case 2 'Sound
        trvKind.Nodes.Add "Sound", tvwChild, , FileHeader(i).FH.FileName, 9, 15
      Case 3 'Text
        trvKind.Nodes.Add "Text", tvwChild, , FileHeader(i).FH.FileName, 10, 16
      Case 4 'Various
        trvKind.Nodes.Add "Various", tvwChild, , FileHeader(i).FH.FileName, 11, 17
      Case 5 'Video
        trvKind.Nodes.Add "Video", tvwChild, , FileHeader(i).FH.FileName, 12, 18
    End Select
  Next i
  
  CFile = False
  For i = trvKind.Nodes.Count To 1 Step -1
    trvKind.Nodes.Item(i).Expanded = True
    If trvKind.Nodes.Item(i).Key <> "" Then If trvKind.Nodes.Item(i).Children > 0 Then CFile = True Else trvKind.Nodes.Remove i
  Next i
  
  picEmpty.Visible = Not CFile
  
  If CurIndex = -1 Then
    For i = 1 To trvKind.Nodes.Count
      If trvKind.Nodes.Item(i).Text = FileHeader(CurIndex).FH.FileName Then
        trvKind.SelectedItem = trvKind.Nodes.Item(i)
        Exit For
      End If
    Next i
  End If
  
  StopWork
End Function

Public Function OnWork(Msg As String)
  stbStatus.SimpleText = Msg
  stbStatus.Refresh
  Me.MousePointer = 11
End Function

Public Function StopWork()
  stbStatus.SimpleText = stbStatus.SimpleText & " -> Done!"
  stbStatus.Refresh
  Sleep 100
  stbStatus.SimpleText = ""
  Me.MousePointer = 0
End Function

Function SelBestUnit(DATA As Long) As String
  '1024 Octets = 1 Ko
  If DATA < (1024 ^ 2) Then  ' Entre 1 Ko et 1023Ko
    SelBestUnit = DATA & " (" & (Round((DATA / 1024), 2)) & " Kb" & ")"
          
  ElseIf DATA >= (1024 ^ 2) And DATA < ((1024 ^ 2) ^ 2) Then 'Entre 1 Mo et 1023 Mo
    SelBestUnit = DATA & " (" & (Round((DATA / (1024 ^ 2)), 2)) & " Mb" & ")"
  
  ElseIf DATA >= ((1024 ^ 2) ^ 2) And DATA < (((1024 ^ 2) ^ 2) ^ 2) Then 'Entre 1 Go et 1023 Go
    SelBestUnit = DATA & " (" & (Round((DATA / ((1024 ^ 2) ^ 2)), 2)) & " Gb" & ")"
    
  End If
End Function

Function WriteTempRc()
  Dim Fno As Integer
  
  Fno = FreeFile
  Open TempRcFile For Binary Access Write Lock Read Write As Fno
    Put Fno, , MainHeader
    For i = 0 To MainHeader.NbFile - 1
      Put Fno, , FileHeader(i)
    Next i
  Close Fno
End Function

Function SizeRc()
  Dim Lenght As Long
  Dim Fno As Integer
  
  OnWork "Calculing file lenght"
  
  Lenght = FileLen(TempRcFile)
  
  For i = 0 To MainHeader.NbFile - 1
    FileHeader(i).FH.StartAt = Lenght + 1
    FileHeader(i).FH.EndAt = Lenght + FileHeader(i).FH.FileLenght
    Lenght = Lenght + FileHeader(i).FH.FileLenght
  Next i
  MainHeader.FileLenght = Lenght
  
  StopWork
End Function

Function OriExtention(File As String, Optional JustExt As Boolean) As String
  OriExtention = Mid(File, 1, InStrRev(File, ".") - 1) & Mid(FileHeader(CurIndex).FH.FileName, InStrRev(FileHeader(CurIndex).FH.FileName, "."))
  If JustExt Then OriExtention = Mid(OriExtention, InStrRev(OriExtention, ".") + 1)
End Function

Function DoShortPath(Path As String, ByVal Index As Integer) As String
  DoShortPath = Mid(FileHeader(Index).TempFile, 1, InStr(Path, "\")) & "..." & Mid(Path, InStrRev(Path, "\TempRc"))
End Function

Function SaveAs()
  On Error Resume Next
  cdlBrowse.DialogTitle = "Choose ressource file name"
  cdlBrowse.FileName = ""
  cdlBrowse.Filter = "PRF ressource file (*.RAB)|*.RAB"
  cdlBrowse.ShowSave
  If Err.Number <> 0 Then Exit Function
  If Dir(cdlBrowse.FileName) <> "" Then If MsgBox("File already exist. Overwrite?", vbQuestion + vbYesNo, "File exist") = vbNo Then Exit Function Else Kill cdlBrowse.FileName
  RcFileName = cdlBrowse.FileName
End Function

Function SaveOutputFile(FilePath As String)
  Dim FnoF As Integer
  Dim FnoR As Integer
  
  Dim CharArray() As Byte
  
  If Dir(FilePath) <> "" Then Kill FilePath
  
  FnoF = FreeFile
  Open FilePath For Binary Access Write Lock Read Write As FnoF
    
    OnWork "Writing headers to output file"
    Put FnoF, , MainHeader
    
    For i = 0 To MainHeader.NbFile - 1
      Put FnoF, , FileHeader(i).FH
    Next i
    StopWork
    
    For i = 0 To MainHeader.NbFile - 1
      
      OnWork "Read temp file " & DoShortPath(FileHeader(i).TempFile, i) & " and write to output file " & GetFileName(FilePath)
      FnoR = FreeFile
      Open FileHeader(i).TempFile For Binary Access Read Lock Read Write As FnoR
        ReDim CharArray(LOF(FnoR) - 1)
        Get FnoR, 1, CharArray
      Close FnoR
      Put FnoF, FileHeader(i).FH.StartAt, CharArray
      StopWork
    Next i
    
    If LOF(FnoF) = MainHeader.FileLenght Then
      MsgBox "Save successful"
    Else
      MsgBox "Error when save"
    End If
  Close FnoF
End Function

Function OpenTextFile(File As String)
  Dim Fno As Integer
  
  OnWork "Please wait while opening file"
  Fno = FreeFile
  Open File For Binary Access Read Lock Read Write As Fno
    
    On Error Resume Next
    rtfText.Text = Input(LOF(Fno), Fno)
    If Err.Number = 7 Then MsgBox "Not enough memory"
    On Error GoTo 0
 
  Close
  StopWork
  
  lblNbChar.Caption = "Nb Char: " & Len(rtfText.Text)
  
End Function

Function OpenSoundFile(File As String)
  Dim TheSound As String
  
  TheSound = OriExtention(File)
  
  If Dir(TheSound) <> "" Then Kill TheSound
  
  FileCopy File, TheSound
  
  mprSound.FileName = TheSound
End Function

Function OpenPictureFile(File As String)
  Dim ThePic As String
  
  ThePic = OriExtention(File)
  
  If Dir(ThePic) <> "" Then Kill ThePic
  
  FileCopy File, ThePic
  
  On Error Resume Next
  picOriginal.Picture = LoadPicture(ThePic)
  If Err.Number = 481 Then
    If MsgBox("Sorry! The picture viewer included in this software can't open this picture file. Try to found an application on your computer that can open this picture?", vbQuestion + vbYesNo, "Can't open") = vbYes Then cmdFind_Click
    Exit Function
  End If
  
  StretchBlt picPic.hdc, 0, 0, picPic.ScaleWidth, picPic.ScaleHeight, picOriginal.hdc, 0, 0, picOriginal.ScaleWidth, picOriginal.ScaleHeight, vbSrcCopy
  picPic.Refresh
  
  lblSize.Caption = "Size: " & picOriginal.ScaleWidth & "x" & picOriginal.ScaleHeight & " (pixel)"
End Function

Function OpenVideoFile(File As String)
  Dim TheVideo As String
  
  TheVideo = OriExtention(File)
  
  If Dir(TheVideo) <> "" Then Kill TheVideo
  
  FileCopy File, TheVideo
  
  mprVideo.FileName = TheVideo
End Function

Function TryLoadFromAC() As Boolean
  OnWork "Try to find last temp header file"
  File.Path = App.Path & "\TempRc"
  File.Refresh
  File.Pattern = "*.tmp"
  For i = 0 To File.ListCount - 1
    If InStr(File.List(i), "TempPRFile") Then
      TempRcFile = App.Path & "\TempRc\" & File.List(i)
      Exit For
    End If
  Next i
  If TempRcFile = "" Then MsgBox "Can't find last temp rc file": TryLoadFromAC = False: Exit Function
  StopWork
  
  OnWork "Try to read last rc temp file"
  Dim Fno As Integer
  
  Fno = FreeFile
  Open TempRcFile For Binary Access Read Lock Read Write As Fno
    Get Fno, 1, MainHeader
    If MainHeader.PRFversion <> PRF_Version Then MsgBox "Bad temp rc file(wrong PRF version)": TryLoadFromAC = False: Close: Exit Function
    
    If MainHeader.NbFile = 0 Then MsgBox "Previous project was empty": TryLoadFromAC = False: Close: Exit Function
    
    ReDim FileHeader(MainHeader.NbFile - 1)
    
    For i = 0 To MainHeader.NbFile - 1
      Get Fno, , FileHeader(i)
    Next i
  Close
  StopWork
  
  For i = 0 To MainHeader.NbFile - 1
    OnWork "Check if file " & GetFileName(FileHeader(i).TempFile) & " exist"
    If Dir(FileHeader(i).TempFile) = "" Then MsgBox "Can't find some temp file": TryLoadFromAC = False: Exit Function
    StopWork
    OnWork "Check if fileheader(" & i & ") is correct. Compare lenght of " & FileHeader(i).FH.FileName & " with " & GetFileName(FileHeader(i).TempFile)
    If FileLen(FileHeader(i).TempFile) <> FileHeader(i).FH.FileLenght Then MsgBox "Bad temp rc file(contain bad lenght info)": TryLoadFromAC = False: Close: Exit Function
    StopWork
  Next i
  
  OnWork "Remove last temp preview file"
  File.Pattern = "*.*"
  For i = 0 To File.ListCount - 1
    If Right(App.Path & "\TempRc" & File.List(i), 4) <> ".tmp" Then Kill App.Path & "\TempRc\" & File.List(i)
  Next i
  StopWork
  
  SizeRc
  
  WriteTempRc
  
  TryLoadFromAC = True
End Function

Function ResetInterface()
  picPicture.Visible = False
  picSound.Visible = False
  picVideo.Visible = False
  picText.Visible = False
  cmdFind.Visible = False
  
  cmdPreview.Enabled = False
  
  rtfText.Text = ""
  lblNbChar.Caption = "Nb Char: 0"
  mprSound.FileName = "*.*"
  mprVideo.FileName = "*.*"
  picPic.Cls
  
  txtFileName.Text = ""
  cmdOK.Visible = False
  cboKind.Text = ""
  lblSize.Caption = "File size: 0 (0Kb)"
  lblFileNo.Caption = "File no: 0"
  lblStartAt.Caption = "Start at: 0"
  lblEndAt.Caption = "End at: 0"
  
  File.Path = App.Path & "\TempRc"
  File.Refresh
  For i = 0 To File.ListCount - 1
    If Right(File.List(i), 4) <> ".tmp" Then
      Kill App.Path & "\TempRC\" & File.List(i)
    End If
  Next i
End Function

Function SelFile(Text As String)
  For i = 1 To trvKind.Nodes.Count
    If trvKind.Nodes.Item(i).Text = Text Then
      trvKind.SelectedItem = trvKind.Nodes.Item(i)
      trvKind_NodeClick trvKind.SelectedItem
      Exit For
    End If
  Next i
End Function

Function Terminate()
  If Dir(App.Path & "\TempRc\*.*") <> "" Then Kill App.Path & "\TempRc\*.*"
  End
End Function

