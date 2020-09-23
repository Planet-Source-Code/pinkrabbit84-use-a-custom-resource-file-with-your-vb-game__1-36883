Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents iltFolder As AxMSComctlLib.AxImageList
	Public WithEvents stbStatus As AxComctlLib.AxStatusBar
	Public WithEvents cdlBrowse As AxMSComDlg.AxCommonDialog
	Public WithEvents lblEmpty As System.Windows.Forms.Label
	Public WithEvents picEmpty As System.Windows.Forms.Panel
	Public WithEvents File As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
	Public WithEvents trvKind As AxMSComctlLib.AxTreeView
	Public WithEvents cmdFind As System.Windows.Forms.Button
	Public WithEvents cmdPreview As System.Windows.Forms.Button
	Public WithEvents rtfText As AxRichTextLib.AxRichTextBox
	Public WithEvents lblNbChar As System.Windows.Forms.Label
	Public WithEvents picText As System.Windows.Forms.Panel
	Public WithEvents picVideo As System.Windows.Forms.PictureBox
	Public WithEvents picSound As System.Windows.Forms.PictureBox
	Public WithEvents picOriginal As System.Windows.Forms.PictureBox
	Public WithEvents picPic As System.Windows.Forms.PictureBox
	Public WithEvents lblSize As System.Windows.Forms.Label
	Public WithEvents picPicture As System.Windows.Forms.Panel
	Public WithEvents fraPreview As System.Windows.Forms.GroupBox
	Public WithEvents lblFileNo As System.Windows.Forms.Label
	Public WithEvents lblEndAt As System.Windows.Forms.Label
	Public WithEvents lblStartAt As System.Windows.Forms.Label
	Public WithEvents lblFileSize As System.Windows.Forms.Label
	Public WithEvents fraInfo As System.Windows.Forms.GroupBox
	Public WithEvents cboKind As System.Windows.Forms.ComboBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtFileName As System.Windows.Forms.TextBox
	Public WithEvents lblKind As System.Windows.Forms.Label
	Public WithEvents lblFileName As System.Windows.Forms.Label
	Public WithEvents fraHeader As System.Windows.Forms.GroupBox
	Public WithEvents fraMain As System.Windows.Forms.GroupBox
	Public WithEvents mprVideo As AxMediaPlayer.AxMediaPlayer
	Public WithEvents mprSound As AxMediaPlayer.AxMediaPlayer
	Public WithEvents mnuNew As System.Windows.Forms.MenuItem
	Public WithEvents mnuOpen As System.Windows.Forms.MenuItem
	Public WithEvents mnuSave As System.Windows.Forms.MenuItem
	Public WithEvents mnuSaveAs As System.Windows.Forms.MenuItem
	Public WithEvents mnuLine4 As System.Windows.Forms.MenuItem
	Public WithEvents mnuBindExt As System.Windows.Forms.MenuItem
	Public WithEvents mnuLine1 As System.Windows.Forms.MenuItem
	Public WithEvents mnuQuit As System.Windows.Forms.MenuItem
	Public WithEvents mnuFile As System.Windows.Forms.MenuItem
	Public WithEvents mnuAdd As System.Windows.Forms.MenuItem
	Public WithEvents mnuRemove As System.Windows.Forms.MenuItem
	Public WithEvents mnuLine2 As System.Windows.Forms.MenuItem
	Public WithEvents mnuExport As System.Windows.Forms.MenuItem
	Public WithEvents mnuData As System.Windows.Forms.MenuItem
	Public WithEvents mnuContent As System.Windows.Forms.MenuItem
	Public WithEvents mnuLine3 As System.Windows.Forms.MenuItem
	Public WithEvents mnuAbout As System.Windows.Forms.MenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.MenuItem
	Public WithEvents mnuShowTFH As System.Windows.Forms.MenuItem
	Public WithEvents mnuDebug As System.Windows.Forms.MenuItem
	Public MainMenu1 As System.Windows.Forms.MainMenu
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.iltFolder = New AxMSComctlLib.AxImageList()
		Me.stbStatus = New AxComctlLib.AxStatusBar()
		Me.cdlBrowse = New AxMSComDlg.AxCommonDialog()
		Me.picEmpty = New System.Windows.Forms.Panel()
		Me.lblEmpty = New System.Windows.Forms.Label()
		Me.File = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
		Me.trvKind = New AxMSComctlLib.AxTreeView()
		Me.fraMain = New System.Windows.Forms.GroupBox()
		Me.fraPreview = New System.Windows.Forms.GroupBox()
		Me.cmdFind = New System.Windows.Forms.Button()
		Me.cmdPreview = New System.Windows.Forms.Button()
		Me.picPicture = New System.Windows.Forms.Panel()
		Me.picPic = New System.Windows.Forms.PictureBox()
		Me.picOriginal = New System.Windows.Forms.PictureBox()
		Me.lblSize = New System.Windows.Forms.Label()
		Me.picText = New System.Windows.Forms.Panel()
		Me.rtfText = New AxRichTextLib.AxRichTextBox()
		Me.lblNbChar = New System.Windows.Forms.Label()
		Me.picVideo = New System.Windows.Forms.PictureBox()
		Me.picSound = New System.Windows.Forms.PictureBox()
		Me.fraHeader = New System.Windows.Forms.GroupBox()
		Me.fraInfo = New System.Windows.Forms.GroupBox()
		Me.lblFileNo = New System.Windows.Forms.Label()
		Me.lblEndAt = New System.Windows.Forms.Label()
		Me.lblStartAt = New System.Windows.Forms.Label()
		Me.lblFileSize = New System.Windows.Forms.Label()
		Me.cboKind = New System.Windows.Forms.ComboBox()
		Me.cmdOK = New System.Windows.Forms.Button()
		Me.txtFileName = New System.Windows.Forms.TextBox()
		Me.lblKind = New System.Windows.Forms.Label()
		Me.lblFileName = New System.Windows.Forms.Label()
		Me.mprVideo = New AxMediaPlayer.AxMediaPlayer()
		Me.mprSound = New AxMediaPlayer.AxMediaPlayer()
		Me.MainMenu1 = New System.Windows.Forms.MainMenu()
		Me.mnuFile = New System.Windows.Forms.MenuItem()
		Me.mnuNew = New System.Windows.Forms.MenuItem()
		Me.mnuOpen = New System.Windows.Forms.MenuItem()
		Me.mnuSave = New System.Windows.Forms.MenuItem()
		Me.mnuSaveAs = New System.Windows.Forms.MenuItem()
		Me.mnuLine4 = New System.Windows.Forms.MenuItem()
		Me.mnuBindExt = New System.Windows.Forms.MenuItem()
		Me.mnuLine1 = New System.Windows.Forms.MenuItem()
		Me.mnuQuit = New System.Windows.Forms.MenuItem()
		Me.mnuData = New System.Windows.Forms.MenuItem()
		Me.mnuAdd = New System.Windows.Forms.MenuItem()
		Me.mnuRemove = New System.Windows.Forms.MenuItem()
		Me.mnuLine2 = New System.Windows.Forms.MenuItem()
		Me.mnuExport = New System.Windows.Forms.MenuItem()
		Me.mnuHelp = New System.Windows.Forms.MenuItem()
		Me.mnuContent = New System.Windows.Forms.MenuItem()
		Me.mnuLine3 = New System.Windows.Forms.MenuItem()
		Me.mnuAbout = New System.Windows.Forms.MenuItem()
		Me.mnuDebug = New System.Windows.Forms.MenuItem()
		Me.mnuShowTFH = New System.Windows.Forms.MenuItem()
		CType(Me.iltFolder, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.stbStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cdlBrowse, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.picEmpty.SuspendLayout()
		CType(Me.trvKind, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.fraMain.SuspendLayout()
		Me.fraPreview.SuspendLayout()
		Me.picPicture.SuspendLayout()
		Me.picText.SuspendLayout()
		CType(Me.rtfText, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.fraHeader.SuspendLayout()
		Me.fraInfo.SuspendLayout()
		CType(Me.mprVideo, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mprSound, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'iltFolder
		'
		Me.iltFolder.Enabled = True
		Me.iltFolder.Location = New System.Drawing.Point(72, 80)
		Me.iltFolder.Name = "iltFolder"
		Me.iltFolder.OcxState = CType(resources.GetObject("iltFolder.OcxState"), System.Windows.Forms.AxHost.State)
		Me.iltFolder.Size = New System.Drawing.Size(38, 38)
		Me.iltFolder.TabIndex = 0
		'
		'stbStatus
		'
		Me.stbStatus.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.stbStatus.Location = New System.Drawing.Point(0, 417)
		Me.stbStatus.Name = "stbStatus"
		Me.stbStatus.OcxState = CType(resources.GetObject("stbStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.stbStatus.Size = New System.Drawing.Size(591, 21)
		Me.stbStatus.TabIndex = 27
		'
		'cdlBrowse
		'
		Me.cdlBrowse.Enabled = True
		Me.cdlBrowse.Location = New System.Drawing.Point(16, 92)
		Me.cdlBrowse.Name = "cdlBrowse"
		Me.cdlBrowse.OcxState = CType(resources.GetObject("cdlBrowse.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cdlBrowse.Size = New System.Drawing.Size(32, 32)
		Me.cdlBrowse.TabIndex = 28
		'
		'picEmpty
		'
		Me.picEmpty.BackColor = System.Drawing.SystemColors.Control
		Me.picEmpty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.picEmpty.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblEmpty})
		Me.picEmpty.Cursor = System.Windows.Forms.Cursors.Default
		Me.picEmpty.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picEmpty.Location = New System.Drawing.Point(12, 16)
		Me.picEmpty.Name = "picEmpty"
		Me.picEmpty.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picEmpty.Size = New System.Drawing.Size(121, 17)
		Me.picEmpty.TabIndex = 23
		Me.picEmpty.TabStop = True
		Me.picEmpty.Visible = False
		'
		'lblEmpty
		'
		Me.lblEmpty.BackColor = System.Drawing.SystemColors.Control
		Me.lblEmpty.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEmpty.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblEmpty.Location = New System.Drawing.Point(24, 1)
		Me.lblEmpty.Name = "lblEmpty"
		Me.lblEmpty.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEmpty.Size = New System.Drawing.Size(72, 17)
		Me.lblEmpty.TabIndex = 24
		Me.lblEmpty.Text = "File list is empty"
		Me.lblEmpty.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'File
		'
		Me.File.BackColor = System.Drawing.SystemColors.Window
		Me.File.Cursor = System.Windows.Forms.Cursors.Default
		Me.File.ForeColor = System.Drawing.SystemColors.WindowText
		Me.File.Location = New System.Drawing.Point(2, 1)
		Me.File.Name = "File"
		Me.File.Pattern = "*.*"
		Me.File.Size = New System.Drawing.Size(35, 17)
		Me.File.TabIndex = 22
		Me.File.Visible = False
		'
		'trvKind
		'
		Me.trvKind.Location = New System.Drawing.Point(4, 0)
		Me.trvKind.Name = "trvKind"
		Me.trvKind.OcxState = CType(resources.GetObject("trvKind.OcxState"), System.Windows.Forms.AxHost.State)
		Me.trvKind.Size = New System.Drawing.Size(137, 417)
		Me.trvKind.TabIndex = 26
		'
		'fraMain
		'
		Me.fraMain.BackColor = System.Drawing.SystemColors.Control
		Me.fraMain.Controls.AddRange(New System.Windows.Forms.Control() {Me.fraPreview, Me.fraHeader})
		Me.fraMain.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraMain.Location = New System.Drawing.Point(144, 0)
		Me.fraMain.Name = "fraMain"
		Me.fraMain.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraMain.Size = New System.Drawing.Size(445, 417)
		Me.fraMain.TabIndex = 0
		Me.fraMain.TabStop = False
		'
		'fraPreview
		'
		Me.fraPreview.BackColor = System.Drawing.SystemColors.Control
		Me.fraPreview.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdFind, Me.cmdPreview, Me.picPicture, Me.picText, Me.picVideo, Me.picSound})
		Me.fraPreview.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraPreview.Location = New System.Drawing.Point(4, 136)
		Me.fraPreview.Name = "fraPreview"
		Me.fraPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraPreview.Size = New System.Drawing.Size(437, 277)
		Me.fraPreview.TabIndex = 12
		Me.fraPreview.TabStop = False
		'
		'cmdFind
		'
		Me.cmdFind.BackColor = System.Drawing.SystemColors.Control
		Me.cmdFind.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdFind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdFind.Location = New System.Drawing.Point(92, 0)
		Me.cmdFind.Name = "cmdFind"
		Me.cmdFind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdFind.Size = New System.Drawing.Size(53, 16)
		Me.cmdFind.TabIndex = 25
		Me.cmdFind.Text = "Find app"
		Me.cmdFind.Visible = False
		'
		'cmdPreview
		'
		Me.cmdPreview.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPreview.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPreview.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPreview.Location = New System.Drawing.Point(8, 0)
		Me.cmdPreview.Name = "cmdPreview"
		Me.cmdPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPreview.Size = New System.Drawing.Size(85, 16)
		Me.cmdPreview.TabIndex = 13
		Me.cmdPreview.Text = "Show preview"
		'
		'picPicture
		'
		Me.picPicture.BackColor = System.Drawing.SystemColors.Control
		Me.picPicture.Controls.AddRange(New System.Windows.Forms.Control() {Me.picPic, Me.picOriginal, Me.lblSize})
		Me.picPicture.Cursor = System.Windows.Forms.Cursors.Default
		Me.picPicture.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picPicture.Location = New System.Drawing.Point(4, 16)
		Me.picPicture.Name = "picPicture"
		Me.picPicture.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picPicture.Size = New System.Drawing.Size(429, 257)
		Me.picPicture.TabIndex = 15
		Me.picPicture.TabStop = True
		Me.picPicture.Visible = False
		'
		'picPic
		'
		Me.picPic.BackColor = System.Drawing.SystemColors.Control
		Me.picPic.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.picPic.Cursor = System.Windows.Forms.Cursors.Default
		Me.picPic.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picPic.Name = "picPic"
		Me.picPic.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picPic.Size = New System.Drawing.Size(429, 241)
		Me.picPic.TabIndex = 18
		Me.picPic.TabStop = False
		'
		'picOriginal
		'
		Me.picOriginal.BackColor = System.Drawing.SystemColors.Control
		Me.picOriginal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.picOriginal.Cursor = System.Windows.Forms.Cursors.Default
		Me.picOriginal.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picOriginal.Location = New System.Drawing.Point(60, 40)
		Me.picOriginal.Name = "picOriginal"
		Me.picOriginal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picOriginal.Size = New System.Drawing.Size(29, 25)
		Me.picOriginal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picOriginal.TabIndex = 19
		Me.picOriginal.TabStop = False
		Me.picOriginal.Visible = False
		'
		'lblSize
		'
		Me.lblSize.BackColor = System.Drawing.SystemColors.Control
		Me.lblSize.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSize.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSize.Location = New System.Drawing.Point(0, 240)
		Me.lblSize.Name = "lblSize"
		Me.lblSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSize.Size = New System.Drawing.Size(373, 17)
		Me.lblSize.TabIndex = 20
		Me.lblSize.Text = "Size: 0x0 (pixel)"
		'
		'picText
		'
		Me.picText.BackColor = System.Drawing.SystemColors.Control
		Me.picText.Controls.AddRange(New System.Windows.Forms.Control() {Me.rtfText, Me.lblNbChar})
		Me.picText.Cursor = System.Windows.Forms.Cursors.Default
		Me.picText.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picText.Location = New System.Drawing.Point(4, 16)
		Me.picText.Name = "picText"
		Me.picText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picText.Size = New System.Drawing.Size(429, 257)
		Me.picText.TabIndex = 16
		Me.picText.TabStop = True
		Me.picText.Visible = False
		'
		'rtfText
		'
		Me.rtfText.ContainingControl = Me
		Me.rtfText.Name = "rtfText"
		Me.rtfText.OcxState = CType(resources.GetObject("rtfText.OcxState"), System.Windows.Forms.AxHost.State)
		Me.rtfText.Size = New System.Drawing.Size(425, 237)
		Me.rtfText.TabIndex = 30
		'
		'lblNbChar
		'
		Me.lblNbChar.BackColor = System.Drawing.SystemColors.Control
		Me.lblNbChar.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblNbChar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblNbChar.Location = New System.Drawing.Point(0, 240)
		Me.lblNbChar.Name = "lblNbChar"
		Me.lblNbChar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblNbChar.Size = New System.Drawing.Size(397, 17)
		Me.lblNbChar.TabIndex = 21
		Me.lblNbChar.Text = "Nb Char: 0"
		'
		'picVideo
		'
		Me.picVideo.BackColor = System.Drawing.SystemColors.Control
		Me.picVideo.Cursor = System.Windows.Forms.Cursors.Default
		Me.picVideo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picVideo.Location = New System.Drawing.Point(4, 16)
		Me.picVideo.Name = "picVideo"
		Me.picVideo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picVideo.Size = New System.Drawing.Size(429, 257)
		Me.picVideo.TabIndex = 17
		Me.picVideo.TabStop = False
		Me.picVideo.Visible = False
		'
		'picSound
		'
		Me.picSound.BackColor = System.Drawing.SystemColors.Control
		Me.picSound.Cursor = System.Windows.Forms.Cursors.Default
		Me.picSound.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picSound.Location = New System.Drawing.Point(4, 16)
		Me.picSound.Name = "picSound"
		Me.picSound.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picSound.Size = New System.Drawing.Size(429, 257)
		Me.picSound.TabIndex = 14
		Me.picSound.TabStop = False
		Me.picSound.Visible = False
		'
		'fraHeader
		'
		Me.fraHeader.BackColor = System.Drawing.SystemColors.Control
		Me.fraHeader.Controls.AddRange(New System.Windows.Forms.Control() {Me.fraInfo, Me.cboKind, Me.cmdOK, Me.txtFileName, Me.lblKind, Me.lblFileName})
		Me.fraHeader.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraHeader.Location = New System.Drawing.Point(4, 8)
		Me.fraHeader.Name = "fraHeader"
		Me.fraHeader.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraHeader.Size = New System.Drawing.Size(437, 125)
		Me.fraHeader.TabIndex = 1
		Me.fraHeader.TabStop = False
		Me.fraHeader.Text = "Header     "
		'
		'fraInfo
		'
		Me.fraInfo.BackColor = System.Drawing.SystemColors.Control
		Me.fraInfo.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblFileNo, Me.lblEndAt, Me.lblStartAt, Me.lblFileSize})
		Me.fraInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraInfo.Location = New System.Drawing.Point(4, 40)
		Me.fraInfo.Name = "fraInfo"
		Me.fraInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraInfo.Size = New System.Drawing.Size(429, 81)
		Me.fraInfo.TabIndex = 7
		Me.fraInfo.TabStop = False
		'
		'lblFileNo
		'
		Me.lblFileNo.BackColor = System.Drawing.SystemColors.Control
		Me.lblFileNo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFileNo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFileNo.Location = New System.Drawing.Point(8, 28)
		Me.lblFileNo.Name = "lblFileNo"
		Me.lblFileNo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFileNo.Size = New System.Drawing.Size(413, 13)
		Me.lblFileNo.TabIndex = 11
		Me.lblFileNo.Text = "File no: 0"
		'
		'lblEndAt
		'
		Me.lblEndAt.BackColor = System.Drawing.SystemColors.Control
		Me.lblEndAt.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEndAt.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblEndAt.Location = New System.Drawing.Point(8, 60)
		Me.lblEndAt.Name = "lblEndAt"
		Me.lblEndAt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEndAt.Size = New System.Drawing.Size(417, 13)
		Me.lblEndAt.TabIndex = 10
		Me.lblEndAt.Text = "End at: 0"
		'
		'lblStartAt
		'
		Me.lblStartAt.BackColor = System.Drawing.SystemColors.Control
		Me.lblStartAt.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblStartAt.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblStartAt.Location = New System.Drawing.Point(8, 44)
		Me.lblStartAt.Name = "lblStartAt"
		Me.lblStartAt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblStartAt.Size = New System.Drawing.Size(417, 13)
		Me.lblStartAt.TabIndex = 9
		Me.lblStartAt.Text = "Start at: 0"
		'
		'lblFileSize
		'
		Me.lblFileSize.BackColor = System.Drawing.SystemColors.Control
		Me.lblFileSize.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFileSize.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFileSize.Location = New System.Drawing.Point(8, 12)
		Me.lblFileSize.Name = "lblFileSize"
		Me.lblFileSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFileSize.Size = New System.Drawing.Size(413, 13)
		Me.lblFileSize.TabIndex = 8
		Me.lblFileSize.Text = "File size: 0 (0Kb)"
		'
		'cboKind
		'
		Me.cboKind.BackColor = System.Drawing.SystemColors.Window
		Me.cboKind.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboKind.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboKind.Items.AddRange(New Object() {"Other", "Picture", "Sound", "Text", "Various", "Video"})
		Me.cboKind.Location = New System.Drawing.Point(296, 16)
		Me.cboKind.Name = "cboKind"
		Me.cboKind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboKind.Size = New System.Drawing.Size(137, 21)
		Me.cboKind.TabIndex = 6
		'
		'cmdOK
		'
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Location = New System.Drawing.Point(222, 18)
		Me.cmdOK.Name = "cmdOK"
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.Size = New System.Drawing.Size(25, 16)
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Visible = False
		'
		'txtFileName
		'
		Me.txtFileName.AcceptsReturn = True
		Me.txtFileName.AutoSize = False
		Me.txtFileName.BackColor = System.Drawing.SystemColors.Window
		Me.txtFileName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFileName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFileName.Location = New System.Drawing.Point(56, 16)
		Me.txtFileName.MaxLength = 0
		Me.txtFileName.Name = "txtFileName"
		Me.txtFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFileName.Size = New System.Drawing.Size(193, 19)
		Me.txtFileName.TabIndex = 3
		Me.txtFileName.Text = ""
		'
		'lblKind
		'
		Me.lblKind.BackColor = System.Drawing.SystemColors.Control
		Me.lblKind.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblKind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblKind.Location = New System.Drawing.Point(252, 20)
		Me.lblKind.Name = "lblKind"
		Me.lblKind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblKind.Size = New System.Drawing.Size(45, 17)
		Me.lblKind.TabIndex = 5
		Me.lblKind.Text = "File type:"
		'
		'lblFileName
		'
		Me.lblFileName.BackColor = System.Drawing.SystemColors.Control
		Me.lblFileName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFileName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFileName.Location = New System.Drawing.Point(8, 20)
		Me.lblFileName.Name = "lblFileName"
		Me.lblFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFileName.Size = New System.Drawing.Size(109, 13)
		Me.lblFileName.TabIndex = 2
		Me.lblFileName.Text = "File name:"
		'
		'mprVideo
		'
		Me.mprVideo.Location = New System.Drawing.Point(36, 364)
		Me.mprVideo.Name = "mprVideo"
		Me.mprVideo.OcxState = CType(resources.GetObject("mprVideo.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mprVideo.Size = New System.Drawing.Size(286, 225)
		Me.mprVideo.TabIndex = 29
		Me.mprVideo.Visible = False
		'
		'mprSound
		'
		Me.mprSound.Location = New System.Drawing.Point(100, 272)
		Me.mprSound.Name = "mprSound"
		Me.mprSound.OcxState = CType(resources.GetObject("mprSound.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mprSound.Size = New System.Drawing.Size(286, 225)
		Me.mprSound.TabIndex = 28
		Me.mprSound.Visible = False
		'
		'MainMenu1
		'
		Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuData, Me.mnuHelp, Me.mnuDebug})
		'
		'mnuFile
		'
		Me.mnuFile.Index = 0
		Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNew, Me.mnuOpen, Me.mnuSave, Me.mnuSaveAs, Me.mnuLine4, Me.mnuBindExt, Me.mnuLine1, Me.mnuQuit})
		Me.mnuFile.Text = "&File"
		'
		'mnuNew
		'
		Me.mnuNew.Index = 0
		Me.mnuNew.Shortcut = System.Windows.Forms.Shortcut.CtrlN
		Me.mnuNew.Text = "&New"
		'
		'mnuOpen
		'
		Me.mnuOpen.Index = 1
		Me.mnuOpen.Shortcut = System.Windows.Forms.Shortcut.CtrlO
		Me.mnuOpen.Text = "&Open"
		'
		'mnuSave
		'
		Me.mnuSave.Index = 2
		Me.mnuSave.Shortcut = System.Windows.Forms.Shortcut.CtrlS
		Me.mnuSave.Text = "&Save"
		'
		'mnuSaveAs
		'
		Me.mnuSaveAs.Index = 3
		Me.mnuSaveAs.Text = "Save &as"
		'
		'mnuLine4
		'
		Me.mnuLine4.Index = 4
		Me.mnuLine4.Text = "-"
		'
		'mnuBindExt
		'
		Me.mnuBindExt.Index = 5
		Me.mnuBindExt.Shortcut = System.Windows.Forms.Shortcut.CtrlB
		Me.mnuBindExt.Text = "Bind extention"
		'
		'mnuLine1
		'
		Me.mnuLine1.Index = 6
		Me.mnuLine1.Text = "-"
		'
		'mnuQuit
		'
		Me.mnuQuit.Index = 7
		Me.mnuQuit.Text = "&Quit"
		'
		'mnuData
		'
		Me.mnuData.Index = 1
		Me.mnuData.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAdd, Me.mnuRemove, Me.mnuLine2, Me.mnuExport})
		Me.mnuData.Text = "Library data"
		'
		'mnuAdd
		'
		Me.mnuAdd.Index = 0
		Me.mnuAdd.Shortcut = System.Windows.Forms.Shortcut.CtrlA
		Me.mnuAdd.Text = "Add file"
		'
		'mnuRemove
		'
		Me.mnuRemove.Index = 1
		Me.mnuRemove.Shortcut = System.Windows.Forms.Shortcut.CtrlR
		Me.mnuRemove.Text = "Remove file"
		'
		'mnuLine2
		'
		Me.mnuLine2.Index = 2
		Me.mnuLine2.Text = "-"
		'
		'mnuExport
		'
		Me.mnuExport.Index = 3
		Me.mnuExport.Shortcut = System.Windows.Forms.Shortcut.CtrlE
		Me.mnuExport.Text = "Export"
		'
		'mnuHelp
		'
		Me.mnuHelp.Index = 2
		Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuContent, Me.mnuLine3, Me.mnuAbout})
		Me.mnuHelp.Text = "Help"
		'
		'mnuContent
		'
		Me.mnuContent.Index = 0
		Me.mnuContent.Text = "Content"
		Me.mnuContent.Visible = False
		'
		'mnuLine3
		'
		Me.mnuLine3.Index = 1
		Me.mnuLine3.Text = "-"
		Me.mnuLine3.Visible = False
		'
		'mnuAbout
		'
		Me.mnuAbout.Index = 2
		Me.mnuAbout.Text = "About"
		'
		'mnuDebug
		'
		Me.mnuDebug.Index = 3
		Me.mnuDebug.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuShowTFH})
		Me.mnuDebug.Text = "                           Debug"
		Me.mnuDebug.Visible = False
		'
		'mnuShowTFH
		'
		Me.mnuShowTFH.Index = 0
		Me.mnuShowTFH.Text = "Show header file"
		'
		'frmMain
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(591, 438)
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.iltFolder, Me.stbStatus, Me.cdlBrowse, Me.picEmpty, Me.File, Me.trvKind, Me.fraMain, Me.mprVideo, Me.mprSound})
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Location = New System.Drawing.Point(10, 31)
		Me.MaximizeBox = False
		Me.Menu = Me.MainMenu1
		Me.MinimizeBox = False
		Me.Name = "frmMain"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "PRFile Editor"
		CType(Me.iltFolder, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.stbStatus, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cdlBrowse, System.ComponentModel.ISupportInitialize).EndInit()
		Me.picEmpty.ResumeLayout(False)
		CType(Me.trvKind, System.ComponentModel.ISupportInitialize).EndInit()
		Me.fraMain.ResumeLayout(False)
		Me.fraPreview.ResumeLayout(False)
		Me.picPicture.ResumeLayout(False)
		Me.picText.ResumeLayout(False)
		CType(Me.rtfText, System.ComponentModel.ISupportInitialize).EndInit()
		Me.fraHeader.ResumeLayout(False)
		Me.fraInfo.ResumeLayout(False)
		CType(Me.mprVideo, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mprSound, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmMain
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmMain
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmMain()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim TempRcFile As String
	Dim MainHeader As PRF_Header
	Dim FileHeader() As TempFileHeader
	Dim CurIndex As Short
	Dim RcFileName As String
	
	'Object procedure
	
	'UPGRADE_WARNING: Event cboKind.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cboKind_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboKind.SelectedIndexChanged
		If CurIndex = -1 Then Exit Sub
		
		FileHeader(CurIndex).FH.FileType = cboKind.SelectedIndex
		
		RefreshFileList()
	End Sub
	
	Private Sub cmdFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFind.Click
		Dim retval As Object
		Dim TheFile As String
		
		TheFile = OriExtention(FileHeader(CurIndex).TempFile)
		
		FileCopy(FileHeader(CurIndex).TempFile, TheFile)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object retval. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
		retval = ShellExecute(Me.Handle.ToInt32, "Open", TheFile, "", CStr(0), 1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object retval. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
		If retval = 0 Then MsgBox("Error in Windows API")
		'UPGRADE_WARNING: Couldn't resolve default property of object retval. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
		If retval = 31 Then MsgBox("Sorry! It's seem that nothing on your computer can open this" & vbNewLine & vbNewLine & "(uncertain function: check if no app has open your file" & vbNewLine & "so if you see this for nothing please repport)")
	End Sub
	
	Private Sub cmdPreview_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreview.Click
		OnWork("Please wait while preparing preview")
		cmdPreview.Enabled = False
		Select Case FileHeader(CurIndex).FH.FileType
			Case 0 'Text
				OpenTextFile(FileHeader(CurIndex).TempFile)
			Case 1 'Picture
				OpenPictureFile(FileHeader(CurIndex).TempFile)
			Case 2 'Sound
				OpenSoundFile(FileHeader(CurIndex).TempFile)
			Case 3 'Text
				OpenTextFile(FileHeader(CurIndex).TempFile)
			Case 4 'Text
				OpenTextFile(FileHeader(CurIndex).TempFile)
			Case 5 'Video
				OpenVideoFile(FileHeader(CurIndex).TempFile)
		End Select
		StopWork()
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		OnWork("Search for temp ressource directory")
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(VB6.GetPath & "\TempRc", FileAttribute.Directory) = "" Then MkDir(VB6.GetPath & "\TempRc")
		StopWork()
		
		OnWork("Search for bind extention setting and read it")
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(VB6.GetPath & "\BindExt.pfe") = "" Then CreateDefault()
		LoadBindExt()
		StopWork()
		
		Me.Show()
		Me.Refresh()
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(VB6.GetPath & "\TempRc\*.*") <> "" Then
			If MsgBox("An abnormal close has been detected. Try to recover last PRF project?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Abnormal close") = MsgBoxResult.Yes Then
				If TryLoadFromAC Then
					MsgBox("Recovering last project successful")
					RefreshFileList()
				Else
					Kill(VB6.GetPath & "\TempRc\*.*")
					mnuNew_Click(mnuNew, New System.EventArgs())
				End If
			Else
				Kill(VB6.GetPath & "\TempRc\*.*")
				mnuNew_Click(mnuNew, New System.EventArgs())
			End If
		Else
			mnuNew_Click(mnuNew, New System.EventArgs())
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmMain.Unload has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		Terminate()
	End Sub
	
	Public Sub mnuAbout_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Popup
		mnuAbout_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
		frmAbout.DefInstance.ShowDialog()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim StrLen As Integer
		
		FileHeader(CurIndex).FH.FileName = txtFileName.Text
		
		SizeRc()
		
		OnWork("Write new file header")
		WriteTempRc()
		StopWork()
		
		RefreshFileList()
		
		cmdOK.Visible = False
	End Sub
	
	Public Sub mnuAdd_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAdd.Popup
		mnuAdd_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAdd.Click
		Dim i As Object
		Dim StrLen As Short
		Dim TempFileName As String
		Dim Index As Short
		Dim Fno As Short
		Dim TheFile As String
		Dim FoundType As Short
		
		On Error Resume Next
		cdlBrowse.Filter = "All file (*.*)"
		cdlBrowse.DialogTitle = "Choose a file to add in ressource file"
		cdlBrowse.FileName = ""
		cdlBrowse.ShowOpen()
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Err.Number <> 0 Or Dir(cdlBrowse.FileName) = "" Then Exit Sub
		On Error GoTo 0
		
		OnWork("Check if file already exist")
		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If FileHeader(i).FH.FileName = GetFileName((cdlBrowse.FileName)) Then
				MsgBox("File already exist in this ressource file")
				StopWork()
				Exit Sub
			End If
		Next i
		StopWork()
		
		OnWork("Creating unexisting temp file name and copy original file")
		Do 
			TempFileName = VB6.GetPath & "\TempRc\TempFile" & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & ".tmp"
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		Loop Until Dir(TempFileName) = ""
		FileCopy(cdlBrowse.FileName, TempFileName)
		StopWork()
		
		Index = MainHeader.NbFile
		
		OnWork("Find the best file type")
		TheFile = LCase(cdlBrowse.FileName)
		
		FoundType = -1
		
		For i = 0 To UBound(PictureExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Mid(TheFile, Len(TheFile) + 1 - Len(PictureExt(i))) = PictureExt(i) And PictureExt(i) <> "" Then FoundType = 1
		Next i
		
		For i = 0 To UBound(SoundExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Mid(TheFile, Len(TheFile) + 1 - Len(SoundExt(i))) = SoundExt(i) And SoundExt(i) <> "" Then FoundType = 2
		Next i
		
		For i = 0 To UBound(TextExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Mid(TheFile, Len(TheFile) + 1 - Len(TextExt(i))) = TextExt(i) And TextExt(i) <> "" Then FoundType = 3
		Next i
		
		For i = 0 To UBound(VariousExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Mid(TheFile, Len(TheFile) + 1 - Len(VariousExt(i))) = VariousExt(i) And VariousExt(i) <> "" Then FoundType = 4
		Next i
		
		For i = 0 To UBound(VideoExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Mid(TheFile, Len(TheFile) + 1 - Len(VideoExt(i))) = VideoExt(i) And VideoExt(i) <> "" Then FoundType = 5
		Next i
		
		If FoundType = -1 Then FoundType = 0
		StopWork()
		
		OnWork("Create file header")
		ReDim Preserve FileHeader(Index)
		FileHeader(Index).TempFile = TempFileName
		FileHeader(Index).FH.FileName = GetFileName((cdlBrowse.FileName))
		FileHeader(Index).FH.StartAt = MainHeader.FileLenght - FileLen(cdlBrowse.FileName)
		FileHeader(Index).FH.EndAt = MainHeader.FileLenght
		FileHeader(Index).FH.FileType = FoundType
		FileHeader(Index).FH.FileLenght = FileLen(cdlBrowse.FileName)
		StopWork()
		
		OnWork("Deleting temp ressource file and rewrite new main file header")
		Kill(TempRcFile)
		StopWork()
		
		MainHeader.NbFile = MainHeader.NbFile + 1
		
		WriteTempRc()
		
		SizeRc()
		
		RefreshFileList()
		
		SelFile(FileHeader(Index).FH.FileName)
	End Sub
	
	Public Sub mnuBindExt_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBindExt.Popup
		mnuBindExt_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuBindExt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBindExt.Click
		frmFileType.DefInstance.Show()
	End Sub
	
	Public Sub mnuContent_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuContent.Popup
		mnuContent_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuContent_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuContent.Click
		MsgBox("Sorry!!! Not done yet")
	End Sub
	
	Public Sub mnuData_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuData.Popup
		mnuData_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuData.Click
		Dim retval As Object
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object retval. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
		retval = trvKind.SelectedItem.Index
		If Err.Number <> 0 Or trvKind.SelectedItem.Key <> "" Then
			mnuRemove.Enabled = False
			mnuExport.Enabled = False
		Else
			mnuRemove.Enabled = True
			mnuExport.Enabled = True
		End If
	End Sub
	
	Public Sub mnuExport_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExport.Popup
		mnuExport_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuExport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExport.Click
		Dim ExportTo As String
		
		On Error Resume Next
		cdlBrowse.DialogTitle = "Choose a file to export this file"
		cdlBrowse.Filter = "All file (*.*)|*.*"
		cdlBrowse.FileName = ""
		cdlBrowse.ShowSave()
		If Err.Number <> 0 Then Exit Sub
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(cdlBrowse.FileName) <> "" Then If MsgBox("File already exist. Overwrite?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "File exist") = MsgBoxResult.No Then Exit Sub Else Kill(cdlBrowse.FileName)
		
		If VB.Right(cdlBrowse.FileName, Len(OriExtention(FileHeader(CurIndex).TempFile, True))) <> OriExtention(FileHeader(CurIndex).TempFile, True) Then cdlBrowse.FileName = cdlBrowse.FileName & "." & OriExtention(FileHeader(CurIndex).TempFile, True)
		
		FileCopy(FileHeader(CurIndex).TempFile, cdlBrowse.FileName)
	End Sub
	
	Public Sub mnuFile_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFile.Popup
		mnuFile_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFile.Click
		mnuSave.Enabled = MainHeader.NbFile
		mnuSaveAs.Enabled = MainHeader.NbFile
	End Sub
	
	Public Sub mnuNew_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNew.Popup
		mnuNew_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNew.Click
		Dim Fno As Short
		
		ResetInterface()
		
		RcFileName = ""
		
		OnWork("Delete old temp ressource file")
		On Error Resume Next
		Kill(VB6.GetPath & "\TempRc\*.*")
		On Error GoTo 0
		StopWork()
		
		ReDim FileHeader(0)
		
		OnWork("Creating random temp ressource file name")
		TempRcFile = VB6.GetPath & "\TempRc\TempPRFile" & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & ".tmp"
		StopWork()
		
		OnWork("Create temp ressource file and main file header")
		MainHeader.PRFversion = PRF_Version
		MainHeader.NbFile = 0
		MainHeader.FileLenght = 20
		Fno = FreeFile
		FileOpen(Fno, TempRcFile, OpenMode.Binary, OpenAccess.Write, OpenShare.LockReadWrite)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, MainHeader)
		FileClose(Fno)
		StopWork()
		
		OnWork("Empting previous file list")
		trvKind.Nodes.Clear()
		trvKind.Nodes.Add( ,  , "Other", "Other")
		trvKind.Nodes.Add( ,  , "Picture", "Picture")
		trvKind.Nodes.Add( ,  , "Sound", "Sound")
		trvKind.Nodes.Add( ,  , "Text", "Text")
		trvKind.Nodes.Add( ,  , "Various", "Various")
		trvKind.Nodes.Add( ,  , "Video", "Video")
		StopWork()
		
		RefreshFileList()
	End Sub
	
	Public Sub mnuOpen_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Popup
		mnuOpen_Click(eventSender, eventArgs)
	End Sub

	Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click
		Dim i As Object
		On Error Resume Next
		cdlBrowse.DialogTitle = "Choose PRFile to open"
		cdlBrowse.Filter = "PRF ressource file (*.RAB)|*.RAB"
		cdlBrowse.FileName = ""
		cdlBrowse.ShowOpen()
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Err.Number <> 0 Or Not IO.File.Exists(cdlBrowse.FileName) Then Exit Sub
		On Error GoTo 0

		mnuNew_Click(mnuNew, New System.EventArgs())

		RcFileName = cdlBrowse.FileName

		Dim Fno As Short
		Dim FnoF As Short
		Dim CharArray() As Byte

		Fno = FreeFile()

		OnWork("Opening PRFile and reading mainheader and fileheader")
		FileOpen(Fno, RcFileName, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, MainHeader, 1)

		If MainHeader.PRFversion <> PRF_Version Then MsgBox("Wrong PRFile version. Can't open this file") : FileClose() : Exit Sub
		If MainHeader.FileLenght <> FileLen(RcFileName) Then MsgBox("PRFile size info mismatch. Can't open this file") : FileClose() : Exit Sub

		ReDim FileHeader(MainHeader.NbFile - 1)

		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FileGet(Fno, FileHeader(i).FH)
			Do
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
				FileHeader(i).TempFile = VB6.GetPath & "\TempRc\TempFile" & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & Int(Rnd() * 9) & ".tmp"
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			Loop Until Dir(FileHeader(i).TempFile) = ""
		Next i
		StopWork()

		For i = 0 To MainHeader.NbFile - 1

			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			OnWork("Read file " & FileHeader(i).FH.FileName & " and write to " & DoShortPath(FileHeader(i).TempFile, i))
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			ReDim CharArray(FileHeader(i).FH.FileLenght - 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FileGet(Fno, CharArray, FileHeader(i).FH.StartAt)

			FnoF = FreeFile()
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileOpen(FnoF, FileHeader(i).TempFile, OpenMode.Binary, OpenAccess.Write, OpenShare.LockReadWrite)
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FilePut(FnoF, CharArray, 1)
			FileClose(FnoF)
			StopWork()

		Next i

		FileClose(Fno)

		WriteTempRc()

		RefreshFileList()

		SelFile(trvKind.Nodes.Item(1).Child.Text)

	End Sub

	Public Sub mnuQuit_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuQuit.Popup
		mnuQuit_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuQuit.Click
		Terminate()
	End Sub

	Public Sub mnuRemove_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRemove.Popup
		mnuRemove_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRemove.Click
		Dim i As Object
		If MainHeader.NbFile = 1 Then MsgBox("Can't remove last file") : Exit Sub

		Dim tmpFH() As TempFileHeader
		Dim ii As Short

		Kill(FileHeader(CurIndex).TempFile)

		ReDim tmpFH(MainHeader.NbFile - 1)

		OnWork("Creating a list of file that will not be deleted")
		For i = CurIndex + 1 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			tmpFH(i).TempFile = FileHeader(i).TempFile
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			tmpFH(i).FH.EndAt = FileHeader(i).FH.EndAt
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			tmpFH(i).FH.FileLenght = FileHeader(i).FH.FileLenght
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			tmpFH(i).FH.FileName = FileHeader(i).FH.FileName
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			tmpFH(i).FH.FileType = FileHeader(i).FH.FileType
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			tmpFH(i).FH.StartAt = FileHeader(i).FH.StartAt
		Next i
		StopWork()

		MainHeader.NbFile = MainHeader.NbFile - 1

		ReDim Preserve FileHeader(MainHeader.NbFile - 1)

		OnWork("Restore file to library list")
		For i = CurIndex To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).TempFile = tmpFH(i + 1).TempFile
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.EndAt = tmpFH(i + 1).FH.EndAt
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.FileLenght = tmpFH(i + 1).FH.FileLenght
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.FileName = tmpFH(i + 1).FH.FileName
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.FileType = tmpFH(i + 1).FH.FileType
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.StartAt = tmpFH(i + 1).FH.StartAt
		Next i
		StopWork()

		If CurIndex <> 0 Then CurIndex = CurIndex - 1

		SizeRc()

		WriteTempRc()

		RefreshFileList()
	End Sub

	Public Sub mnuSave_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Popup
		mnuSave_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click
		If RcFileName = "" Then SaveAs()
		If RcFileName = "" Then Exit Sub
		SaveOutputFile(RcFileName)
	End Sub

	Public Sub mnuSaveAs_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveAs.Popup
		mnuSaveAs_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveAs.Click
		SaveAs()
		mnuSave_Click(mnuSave, New System.EventArgs())
	End Sub

	Public Sub mnuShowTFH_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShowTFH.Popup
		mnuShowTFH_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuShowTFH_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShowTFH.Click
		Dim i As Object
		Dim txt As String
		frmDebug.DefInstance.Show()
		txt = "mainheader.FileLenght = " & MainHeader.FileLenght & vbNewLine
		txt = txt & "mainheader.NbFile = " & MainHeader.NbFile & vbNewLine & "- - - - - - - - - - - -" & vbNewLine
		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			txt = txt & "fileheader(" & i & ").FH.EndAt = " & FileHeader(i).FH.EndAt & vbNewLine
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			txt = txt & "fileheader(" & i & ").FH.FileLenght = " & FileHeader(i).FH.FileLenght & vbNewLine
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			txt = txt & "fileheader(" & i & ").FH.FileName = " & FileHeader(i).FH.FileName & vbNewLine
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			txt = txt & "fileheader(" & i & ").FH.FileType = " & FileHeader(i).FH.FileType & vbNewLine
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			txt = txt & "fileheader(" & i & ").FH.StartAt = " & FileHeader(i).FH.StartAt & vbNewLine
			txt = txt & "- - - - - - - - - - - -" & vbNewLine
		Next i
		frmDebug.DefInstance.txtDebug.Text = txt
	End Sub

	Private Sub picPic_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picPic.Click
		Dim fPicture As frmPicture
		Dim AddX, AddY As Short
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1039"'

		fPicture = New frmPicture()
		With fPicture
			.picOriginal.Width = picOriginal.ClientRectangle.Width
			.picOriginal.Height = picOriginal.ClientRectangle.Height
			.picOriginal.Image = picOriginal.Image
			AddX = .Width - .ClientRectangle.Width
			AddY = .Height - .ClientRectangle.Height
			.Width = picOriginal.ClientRectangle.Width + AddX
			.Height = picOriginal.ClientRectangle.Height + AddY
			.Show()
			.Text = "Picture - " & txtFileName.Text
		End With
	End Sub

	Private Sub trvKind_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ITreeViewEvents_MouseUpEvent) Handles trvKind.MouseUpEvent
		'UPGRADE_ISSUE: Form method frmMain.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2064"'
		If eventArgs.button = 2 Then mnuData.PerformClick()
	End Sub

	Private Sub trvKind_NodeClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ITreeViewEvents_NodeClickEvent) Handles trvKind.NodeClick
		Dim i As Object

		ResetInterface()

		Dim Index As Short
		Dim TmpF As String
		Dim SelFile As String
		If eventArgs.node.Key = "" Then


			SelFile = eventArgs.node.Text

			CurIndex = -1

			For i = 0 To MainHeader.NbFile - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
				If FileHeader(i).FH.FileName = SelFile Then
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					Index = i : Exit For
				End If
			Next i

			txtFileName.Text = FileHeader(Index).FH.FileName : cmdOK.Visible = False
			cboKind.SelectedIndex = FileHeader(Index).FH.FileType
			lblFileSize.Text = "File size: " & SelBestUnit(FileHeader(Index).FH.FileLenght)
			lblFileNo.Text = "File no: " & Index + 1
			lblStartAt.Text = "Start at: " & FileHeader(Index).FH.StartAt
			lblEndAt.Text = "End at: " & FileHeader(Index).FH.EndAt

			CurIndex = Index

			Select Case FileHeader(Index).FH.FileType
				Case 0				'Text
					picText.Visible = True
					picText.BringToFront()
					cmdFind.Visible = True
				Case 1				'Picture
					picPicture.Visible = True
					picPicture.BringToFront()
				Case 2				'Sound
					picSound.Visible = True
					picSound.BringToFront()
				Case 3				'Text
					picText.Visible = True
					picText.BringToFront()
				Case 4				'Text
					picText.Visible = True
					picText.BringToFront()
					cmdFind.Visible = True
				Case 5				'Video
					picVideo.Visible = True
					picVideo.BringToFront()
			End Select

			cmdPreview.Enabled = True

			stbStatus.SimpleText = SelFile & " (" & DoShortPath(FileHeader(CurIndex).TempFile, CurIndex) & ")"
		Else
			CurIndex = -1
		End If
	End Sub

	'UPGRADE_WARNING: Event txtFileName.TextChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub txtFileName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFileName.TextChanged
		cmdOK.Visible = True
	End Sub


	'Function
	Function GetFileName(ByRef FilePath As String) As String
		GetFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
	End Function

	Function RefreshFileList() As Object
		Dim i As Object
		Dim CFile As Boolean

		OnWork("Refresh file list")
		trvKind.Nodes.Clear()
		trvKind.Nodes.Add(, , "Other", "Other", 1, 1)
		trvKind.Nodes.Add(, , "Picture", "Picture", 2, 2)
		trvKind.Nodes.Add(, , "Sound", "Sound", 3, 3)
		trvKind.Nodes.Add(, , "Text", "Text", 4, 4)
		trvKind.Nodes.Add(, , "Various", "Various", 5, 5)
		trvKind.Nodes.Add(, , "Video", "Video", 6, 6)
		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			Select Case FileHeader(i).FH.FileType
				Case 0				'Other
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					trvKind.Nodes.Add("Other", MSComctlLib.TreeRelationshipConstants.tvwChild, , FileHeader(i).FH.FileName, 7, 13)
				Case 1				'Picture
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					trvKind.Nodes.Add("Picture", MSComctlLib.TreeRelationshipConstants.tvwChild, , FileHeader(i).FH.FileName, 8, 14)
				Case 2				'Sound
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					trvKind.Nodes.Add("Sound", MSComctlLib.TreeRelationshipConstants.tvwChild, , FileHeader(i).FH.FileName, 9, 15)
				Case 3				'Text
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					trvKind.Nodes.Add("Text", MSComctlLib.TreeRelationshipConstants.tvwChild, , FileHeader(i).FH.FileName, 10, 16)
				Case 4				'Various
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					trvKind.Nodes.Add("Various", MSComctlLib.TreeRelationshipConstants.tvwChild, , FileHeader(i).FH.FileName, 11, 17)
				Case 5				'Video
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
					trvKind.Nodes.Add("Video", MSComctlLib.TreeRelationshipConstants.tvwChild, , FileHeader(i).FH.FileName, 12, 18)
			End Select
		Next i

		CFile = False
		For i = trvKind.Nodes.Count To 1 Step -1
			trvKind.Nodes.Item(i).Expanded = True
			If trvKind.Nodes.Item(i).Key <> "" Then If trvKind.Nodes.Item(i).Children > 0 Then CFile = True Else trvKind.Nodes.Remove(i)
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

		StopWork()
	End Function

	Public Function OnWork(ByRef Msg As String) As Object
		stbStatus.SimpleText = Msg
		stbStatus.CtlRefresh()
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
	End Function

	Public Function StopWork() As Object
		stbStatus.SimpleText = stbStatus.SimpleText & " -> Done!"
		stbStatus.CtlRefresh()
		Sleep(100)
		stbStatus.SimpleText = ""
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Function

	Function SelBestUnit(ByRef DATA As Integer) As String
		'1024 Octets = 1 Ko
		If DATA < (1024 ^ 2) Then		' Entre 1 Ko et 1023Ko
			SelBestUnit = DATA & " (" & (System.Math.Round(DATA / 1024, 2)) & " Kb" & ")"

		ElseIf DATA >= (1024 ^ 2) And DATA < ((1024 ^ 2) ^ 2) Then		'Entre 1 Mo et 1023 Mo
			SelBestUnit = DATA & " (" & (System.Math.Round(DATA / (1024 ^ 2), 2)) & " Mb" & ")"

		ElseIf DATA >= ((1024 ^ 2) ^ 2) And DATA < (((1024 ^ 2) ^ 2) ^ 2) Then		'Entre 1 Go et 1023 Go
			SelBestUnit = DATA & " (" & (System.Math.Round(DATA / ((1024 ^ 2) ^ 2), 2)) & " Gb" & ")"

		End If
	End Function

	Function WriteTempRc() As Object
		Dim i As Object
		Dim Fno As Short

		Fno = FreeFile()
		FileOpen(Fno, TempRcFile, OpenMode.Binary, OpenAccess.Write, OpenShare.LockReadWrite)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, MainHeader)
		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FilePut(Fno, FileHeader(i))
		Next i
		FileClose(Fno)
	End Function

	Function SizeRc() As Object
		Dim i As Object
		Dim Lenght As Integer
		Dim Fno As Short

		OnWork("Calculing file lenght")

		Lenght = FileLen(TempRcFile)

		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.StartAt = Lenght + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileHeader(i).FH.EndAt = Lenght + FileHeader(i).FH.FileLenght
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			Lenght = Lenght + FileHeader(i).FH.FileLenght
		Next i
		MainHeader.FileLenght = Lenght

		StopWork()
	End Function

	Function OriExtention(ByRef File As String, Optional ByRef JustExt As Boolean = False) As String
		OriExtention = Mid(File, 1, InStrRev(File, ".") - 1) & Mid(FileHeader(CurIndex).FH.FileName, InStrRev(FileHeader(CurIndex).FH.FileName, "."))
		If JustExt Then OriExtention = Mid(OriExtention, InStrRev(OriExtention, ".") + 1)
	End Function

	Function DoShortPath(ByRef Path As String, ByVal Index As Short) As String
		DoShortPath = Mid(FileHeader(Index).TempFile, 1, InStr(Path, "\")) & "..." & Mid(Path, InStrRev(Path, "\TempRc"))
	End Function

	Function SaveAs() As Object
		On Error Resume Next
		cdlBrowse.DialogTitle = "Choose ressource file name"
		cdlBrowse.FileName = ""
		cdlBrowse.Filter = "PRF ressource file (*.RAB)|*.RAB"
		cdlBrowse.ShowSave()
		If Err.Number <> 0 Then Exit Function
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(cdlBrowse.FileName) <> "" Then If MsgBox("File already exist. Overwrite?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "File exist") = MsgBoxResult.No Then Exit Function Else Kill(cdlBrowse.FileName)
		RcFileName = cdlBrowse.FileName
	End Function

	Function SaveOutputFile(ByRef FilePath As String) As Object
		Dim FnoF As Short
		Dim FnoR As Short
		Dim CharArray() As Byte
		Dim i As Integer

		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(FilePath) <> "" Then Kill(FilePath)

		FnoF = FreeFile()
		FileOpen(FnoF, FilePath, OpenMode.Binary, OpenAccess.Write, OpenShare.LockReadWrite)

		OnWork("Writing headers to output file")
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(FnoF, MainHeader)

		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FilePut(FnoF, FileHeader(i).FH)
		Next i
		StopWork()

		For i = 0 To MainHeader.NbFile - 1

			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			OnWork("Read temp file " & DoShortPath(FileHeader(i).TempFile, i) & " and write to output file " & GetFileName(FilePath))
			FnoR = FreeFile()
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileOpen(FnoR, FileHeader(i).TempFile, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
			ReDim CharArray(LOF(FnoR) - 1)
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FileGet(FnoR, CharArray, 1)
			FileClose(FnoR)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FilePut(FnoF, CharArray, FileHeader(i).FH.StartAt)
			StopWork()
		Next i

		If LOF(FnoF) = MainHeader.FileLenght Then
			MsgBox("Save successful")
		Else
			MsgBox("Error when save")
		End If
		FileClose(FnoF)
	End Function

	Function OpenTextFile(ByRef File As String) As Object
		Dim Fno As Short

		OnWork("Please wait while opening file")
		Fno = FreeFile()
		FileOpen(Fno, File, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)

		On Error Resume Next
		rtfText.Text = InputString(Fno, LOF(Fno))
		If Err.Number = 7 Then MsgBox("Not enough memory")
		On Error GoTo 0

		FileClose()
		StopWork()

		lblNbChar.Text = "Nb Char: " & Len(rtfText.Text)

	End Function

	Function OpenSoundFile(ByRef File As String) As Object
		Dim TheSound As String

		TheSound = OriExtention(File)

		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(TheSound) <> "" Then Kill(TheSound)

		FileCopy(File, TheSound)

		mprSound.FileName = TheSound
	End Function

	Function OpenPictureFile(ByRef File As String)
		Dim bitImage As Bitmap
		Dim ThePic As String

		ThePic = OriExtention(File)

		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(ThePic) <> "" Then Kill(ThePic)

		FileCopy(File, ThePic)

		Try
			bitImage = New Bitmap(New Bitmap(ThePic))
		Catch
			If MsgBox("Sorry! The picture viewer included in this software can't open this picture file. Try to found an application on your computer that can open this picture?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Can't open") = MsgBoxResult.Yes Then cmdFind_Click(cmdFind, New System.EventArgs())
			Exit Function
		End Try

		picPic.CreateGraphics.DrawImage(bitImage, New Rectangle(0, 0, picPic.Width, picPic.Height))

		lblSize.Text = "Size: " & bitImage.Width & "x" & bitImage.Height & " (pixel)"

		bitImage.Dispose()
		bitImage = Nothing
	End Function

	Function OpenVideoFile(ByRef File As String) As Object
		Dim TheVideo As String

		TheVideo = OriExtention(File)

		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(TheVideo) <> "" Then Kill(TheVideo)

		FileCopy(File, TheVideo)

		mprVideo.FileName = TheVideo
	End Function

	Function TryLoadFromAC() As Boolean
		Dim i As Object
		OnWork("Try to find last temp header file")
		File.Path = VB6.GetPath & "\TempRc"
		File.Refresh()
		File.Pattern = "*.tmp"
		For i = 0 To File.Items.Count - 1
			If InStr(File.Items(i), "TempPRFile") Then
				TempRcFile = VB6.GetPath & "\TempRc\" & File.Items(i)
				Exit For
			End If
		Next i
		If TempRcFile = "" Then MsgBox("Can't find last temp rc file") : TryLoadFromAC = False : Exit Function
		StopWork()

		OnWork("Try to read last rc temp file")
		Dim Fno As Short

		Fno = FreeFile()
		FileOpen(Fno, TempRcFile, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, MainHeader, 1)
		If MainHeader.PRFversion <> PRF_Version Then MsgBox("Bad temp rc file(wrong PRF version)") : TryLoadFromAC = False : FileClose() : Exit Function

		If MainHeader.NbFile = 0 Then MsgBox("Previous project was empty") : TryLoadFromAC = False : FileClose() : Exit Function

		ReDim FileHeader(MainHeader.NbFile - 1)

		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			FileGet(Fno, FileHeader(i))
		Next i
		FileClose()
		StopWork()

		For i = 0 To MainHeader.NbFile - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			OnWork("Check if file " & GetFileName(FileHeader(i).TempFile) & " exist")
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
			If Dir(FileHeader(i).TempFile) = "" Then MsgBox("Can't find some temp file") : TryLoadFromAC = False : Exit Function
			StopWork()
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			OnWork("Check if fileheader(" & i & ") is correct. Compare lenght of " & FileHeader(i).FH.FileName & " with " & GetFileName(FileHeader(i).TempFile))
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If FileLen(FileHeader(i).TempFile) <> FileHeader(i).FH.FileLenght Then MsgBox("Bad temp rc file(contain bad lenght info)") : TryLoadFromAC = False : FileClose() : Exit Function
			StopWork()
		Next i

		OnWork("Remove last temp preview file")
		File.Pattern = "*.*"
		For i = 0 To File.Items.Count - 1
			If VB.Right(VB6.GetPath & "\TempRc" & File.Items(i), 4) <> ".tmp" Then Kill(VB6.GetPath & "\TempRc\" & File.Items(i))
		Next i
		StopWork()

		SizeRc()

		WriteTempRc()

		TryLoadFromAC = True
	End Function

	Function ResetInterface() As Object
		Dim i As Object
		picPicture.Visible = False
		picSound.Visible = False
		picVideo.Visible = False
		picText.Visible = False
		cmdFind.Visible = False

		cmdPreview.Enabled = False

		rtfText.Text = ""
		lblNbChar.Text = "Nb Char: 0"
		mprSound.FileName = "*.*"
		mprVideo.FileName = "*.*"
		'UPGRADE_ISSUE: PictureBox method picPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2064"'
		picPic.CreateGraphics.Clear(Color.Black)

		txtFileName.Text = ""
		cmdOK.Visible = False
		cboKind.Text = ""
		lblSize.Text = "File size: 0 (0Kb)"
		lblFileNo.Text = "File no: 0"
		lblStartAt.Text = "Start at: 0"
		lblEndAt.Text = "End at: 0"

		File.Path = VB6.GetPath & "\TempRc"
		File.Refresh()
		For i = 0 To File.Items.Count - 1
			If VB.Right(File.Items(i), 4) <> ".tmp" Then
				IO.File.Delete(Application.StartupPath & "\TempRC\" & File.Items(i))
			End If
		Next i
	End Function

	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1061"'
	Function SelFile(ByRef Text_Renamed As String) As Object
		Dim i As Object
		For i = 1 To trvKind.Nodes.Count
			If trvKind.Nodes.Item(i).Text = Text_Renamed Then
				trvKind.SelectedItem = trvKind.Nodes.Item(i)
				trvKind_NodeClick(trvKind, New AxMSComctlLib.ITreeViewEvents_NodeClickEvent(trvKind.SelectedItem))
				Exit For
			End If
		Next i
	End Function

	Function Terminate() As Object
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(VB6.GetPath & "\TempRc\*.*") <> "" Then Kill(VB6.GetPath & "\TempRc\*.*")
		End
	End Function
End Class