VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAll 
      Caption         =   "Extract all"
      Height          =   315
      Left            =   420
      TabIndex        =   6
      Top             =   4320
      Width           =   2715
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Show text"
      Height          =   315
      Left            =   420
      TabIndex        =   5
      Top             =   3540
      Width           =   2715
   End
   Begin VB.TextBox txtText 
      Height          =   915
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   2580
      Width           =   2715
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "Sound"
      Height          =   315
      Left            =   420
      TabIndex        =   3
      Top             =   3900
      Width           =   2715
   End
   Begin MCI.MMControl Mci 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   556
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdShowPic 
      Caption         =   "Show pic"
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   2220
      Width           =   2715
   End
   Begin VB.PictureBox picPic 
      Height          =   2055
      Left            =   420
      ScaleHeight     =   1995
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PRFile by: PinkRabbit84 (PinkRabbit84@Hotmail.com)

'First to include the PRFile Dll in your project:
'   Click on Project -> References
'     Check ~Vb PRFile Dll~ in the list. If it not in, click on browse and hunt your drive to find the Dll!

'Now you have added the Dll to the project

'You will create the main object
Dim PRF As New clsPRFile
'Note that you can create more than one for use more PRFile at once


'Error code:
'999001 = Bad folder path
'999002 = Bad file path
'999003 = PRFile version mismatch (file is not open)
'999004 = PRFile size info mismatch (file is not open)
'999005 = File does not exist in library(file is not extracted)
'999006 = File can't be writed to destination path(file is not extracted)
'999007 = Destination already exist and overwrite not activated(file is not extracted)
'999008 = Destination file already exist and in use; can't overwrite(file is not extracted)

Private Sub Form_Load()
  'Init the Dll
  '1) The application path(a temp folder will be create in application path; if it's impossible(AppPath is on a CD for example) the temp dir will be created in your Windows\temp dir)
  '2) Where the file will be extracted by default
  ret = PRF.Init(App.Path, App.Path)
  
  If ret <> 0 Then
    'There is an error
    'Possible: 999001
  End If
  
  'Open the PRFile
  '1) Path of the PRFile
  ret = PRF.SetPRFile(App.Path & "\Test.rab")
  
  If ret <> 0 Then
    'There is an error
    'Possible: 999002; 999003; 999004
  End If
End Sub

Private Sub cmdShowPic_Click()
    
  'Extract a file(PinkRabbit.bmp)
  '1) Name of the file in the PRFile
  '2) Path to extract the file(optional if nothing: default path is used)
  '3) OverWrite the file if already exist?(optional if nothing: True)
  ret = PRF.GetFile("PinkRabbit.bmp")
  
  If ret <> 0 Then
    'There is an error
    'Possible: 999005; 999006; 999007
  End If
  
  picPic.Picture = LoadPicture(App.Path & "\PinkRabbit.bmp")
End Sub

Private Sub cmdText_Click()
  'Extract a file(Text File.txt)
  '1) Name of the file in the PRFile
  '2) Path to extract the file(optional if nothing: default path is used)
  '3) OverWrite the file if already exist?(optional if nothing: True)
  ret = PRF.GetFile("Text File.txt")
  
  If ret <> 0 Then
    'There is an error
    'Possible: 999005; 999006; 999007
  End If
  
  Open App.Path & "\Text File.txt" For Input As #1
    txtText.Text = Input(LOF(1), #1)
  Close
End Sub

Private Sub cmdSound_Click()
  Mci.Command = "Close"
  
  'Extract a file(chimes.wav)
  '1) Name of the file in the PRFile
  '2) Path to extract the file(optional if nothing: default path is used)
  '3) OverWrite the file if already exist?(optional if nothing: True)
  ret = PRF.GetFile("chimes.wav")
  
  If ret <> 0 Then
    'There is an error
    'Possible: 999005; 999006; 999007
  End If
  
  Mci.FileName = App.Path & "\chimes.wav"
  Mci.Command = "Open"
  Mci.Command = "Play"
End Sub

Private Sub cmdAll_Click()
  'Files in PRFile can all extracted at once. Just call GetAllFile
  
  '1) Path to extract the files(optional if nothing: default path is used)
  '2) OverWrite the file if already exist?(optional if nothing: True)
  ret = PRF.GetAllFile
  
  If ret <> 0 Then
    'There is an error
    'Possible: 999005; 999006; 999007
    'Note that here you get only the last error that occur in GetFile (GetAllFile is calling GetFile)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Mci.Command = "Close"
  
  'Clear up memory
  Set PRF = Nothing
  'When the class(clsPRFile) is closed the temp folder and all extracted file is automaticaly deleted
End Sub



'Scuse me for my english
