Option Strict Off
Option Explicit On
Friend Class frmAbout
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
	Public WithEvents lblEMail As System.Windows.Forms.Label
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
	Public WithEvents lblBug As System.Windows.Forms.Label
	Public WithEvents Line1 As System.Windows.Forms.Label
	Public WithEvents lblBy As System.Windows.Forms.Label
	Public WithEvents lblPRFileVer As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents imgBack As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.lblEMail = New System.Windows.Forms.Label()
		Me.lblDisclaimer = New System.Windows.Forms.Label()
		Me.lblBug = New System.Windows.Forms.Label()
		Me.Line1 = New System.Windows.Forms.Label()
		Me.lblBy = New System.Windows.Forms.Label()
		Me.lblPRFileVer = New System.Windows.Forms.Label()
		Me.lblVersion = New System.Windows.Forms.Label()
		Me.lblTitle = New System.Windows.Forms.Label()
		Me.imgBack = New System.Windows.Forms.PictureBox()
		Me.SuspendLayout()
		'
		'lblEMail
		'
		Me.lblEMail.BackColor = System.Drawing.Color.Transparent
		Me.lblEMail.Font = New System.Drawing.Font("Comic Sans MS", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEMail.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
		Me.lblEMail.Location = New System.Drawing.Point(68, 144)
		Me.lblEMail.Name = "lblEMail"
		Me.lblEMail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEMail.Size = New System.Drawing.Size(181, 21)
		Me.lblEMail.TabIndex = 6
		Me.lblEMail.Text = "PrVbTool@hotmail.com"
		Me.lblEMail.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'lblDisclaimer
		'
		Me.lblDisclaimer.BackColor = System.Drawing.Color.Transparent
		Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDisclaimer.ForeColor = System.Drawing.Color.FromArgb(CType(185, Byte), CType(0, Byte), CType(0, Byte))
		Me.lblDisclaimer.Location = New System.Drawing.Point(4, 172)
		Me.lblDisclaimer.Name = "lblDisclaimer"
		Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDisclaimer.Size = New System.Drawing.Size(309, 57)
		Me.lblDisclaimer.TabIndex = 5
		Me.lblDisclaimer.Text = "Disclaimer: We are not reponsable for any damage (data lost for example) caused  " & _
		"by this software or occur when this software is running. This software is curren" & _
		"tly in beta test so run at your own risk!"
		Me.lblDisclaimer.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'lblBug
		'
		Me.lblBug.BackColor = System.Drawing.Color.Transparent
		Me.lblBug.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBug.Font = New System.Drawing.Font("Comic Sans MS", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBug.ForeColor = System.Drawing.Color.FromArgb(CType(254, Byte), CType(227, Byte), CType(255, Byte))
		Me.lblBug.Location = New System.Drawing.Point(0, 96)
		Me.lblBug.Name = "lblBug"
		Me.lblBug.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBug.Size = New System.Drawing.Size(317, 41)
		Me.lblBug.TabIndex = 4
		Me.lblBug.Text = "If you find any bug in this software please contact with us at:"
		Me.lblBug.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'Line1
		'
		Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
		Me.Line1.Location = New System.Drawing.Point(0, 88)
		Me.Line1.Name = "Line1"
		Me.Line1.Size = New System.Drawing.Size(316, 1)
		Me.Line1.TabIndex = 7
		'
		'lblBy
		'
		Me.lblBy.BackColor = System.Drawing.Color.Transparent
		Me.lblBy.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBy.Font = New System.Drawing.Font("Comic Sans MS", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBy.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(196, Byte), CType(237, Byte))
		Me.lblBy.Location = New System.Drawing.Point(4, 44)
		Me.lblBy.Name = "lblBy"
		Me.lblBy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBy.Size = New System.Drawing.Size(209, 33)
		Me.lblBy.TabIndex = 3
		Me.lblBy.Text = "By: PinkRabbit Soft"
		'
		'lblPRFileVer
		'
		Me.lblPRFileVer.BackColor = System.Drawing.Color.Transparent
		Me.lblPRFileVer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPRFileVer.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPRFileVer.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(196, Byte), CType(237, Byte))
		Me.lblPRFileVer.Location = New System.Drawing.Point(136, 32)
		Me.lblPRFileVer.Name = "lblPRFileVer"
		Me.lblPRFileVer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPRFileVer.Size = New System.Drawing.Size(89, 13)
		Me.lblPRFileVer.TabIndex = 2
		Me.lblPRFileVer.Text = "Use PRFile version 1.00"
		Me.lblPRFileVer.TextAlign = System.Drawing.ContentAlignment.TopRight
		'
		'lblVersion
		'
		Me.lblVersion.BackColor = System.Drawing.Color.Transparent
		Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVersion.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblVersion.ForeColor = System.Drawing.Color.FromArgb(CType(253, Byte), CType(111, Byte), CType(253, Byte))
		Me.lblVersion.Location = New System.Drawing.Point(140, 16)
		Me.lblVersion.Name = "lblVersion"
		Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVersion.Size = New System.Drawing.Size(173, 17)
		Me.lblVersion.TabIndex = 1
		Me.lblVersion.Text = "v 1.01 (beta)"
		'
		'lblTitle
		'
		Me.lblTitle.BackColor = System.Drawing.Color.Transparent
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.Font = New System.Drawing.Font("Comic Sans MS", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(196, Byte), CType(237, Byte))
		Me.lblTitle.Location = New System.Drawing.Point(4, 4)
		Me.lblTitle.Name = "lblTitle"
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.Size = New System.Drawing.Size(317, 29)
		Me.lblTitle.TabIndex = 0
		Me.lblTitle.Text = "PRFile Editor"
		'
		'imgBack
		'
		Me.imgBack.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgBack.Image = CType(resources.GetObject("imgBack.Image"), System.Drawing.Bitmap)
		Me.imgBack.Name = "imgBack"
		Me.imgBack.Size = New System.Drawing.Size(316, 208)
		Me.imgBack.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me.imgBack.TabIndex = 8
		Me.imgBack.TabStop = False
		'
		'frmAbout
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.Color.Black
		Me.ClientSize = New System.Drawing.Size(316, 226)
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblEMail, Me.lblDisclaimer, Me.lblBug, Me.Line1, Me.lblBy, Me.lblPRFileVer, Me.lblVersion, Me.lblTitle, Me.imgBack})
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Location = New System.Drawing.Point(3, 24)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "frmAbout"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "About - PRFile Editor"
		Me.ResumeLayout(False)

	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmAbout
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmAbout
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmAbout()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private Sub lblEMail_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblEMail.Click
		ShellExecute(Me.Handle.ToInt32, "Open", "mailto:PrVbTool@Hotmail.com", "", "", 0)
	End Sub
End Class