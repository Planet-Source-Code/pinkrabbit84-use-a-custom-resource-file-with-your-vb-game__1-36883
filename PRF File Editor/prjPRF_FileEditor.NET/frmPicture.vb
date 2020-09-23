Option Strict Off
Option Explicit On
Friend Class frmPicture
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()

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
	Public WithEvents picPicture As System.Windows.Forms.PictureBox
	Public WithEvents picOriginal As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPicture))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.picPicture = New System.Windows.Forms.PictureBox()
		Me.picOriginal = New System.Windows.Forms.PictureBox()
		Me.Text = "Picture - [no picture]"
		Me.ClientSize = New System.Drawing.Size(312, 211)
		Me.Location = New System.Drawing.Point(4, 25)
		Me.Icon = CType(resources.GetObject("frmPicture.Icon"), System.Drawing.Icon)
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPicture"
		Me.picPicture.BackColor = System.Drawing.SystemColors.Window
		Me.picPicture.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picPicture.Size = New System.Drawing.Size(120, 53)
		Me.picPicture.Location = New System.Drawing.Point(0, 0)
		Me.picPicture.TabIndex = 1
		Me.picPicture.Dock = System.Windows.Forms.DockStyle.None
		Me.picPicture.CausesValidation = True
		Me.picPicture.Enabled = True
		Me.picPicture.Cursor = System.Windows.Forms.Cursors.Default
		Me.picPicture.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picPicture.TabStop = True
		Me.picPicture.Visible = True
		Me.picPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picPicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.picPicture.Name = "picPicture"
		Me.picOriginal.Size = New System.Drawing.Size(49, 41)
		Me.picOriginal.Location = New System.Drawing.Point(0, 0)
		Me.picOriginal.TabIndex = 0
		Me.picOriginal.Visible = False
		Me.picOriginal.Dock = System.Windows.Forms.DockStyle.None
		Me.picOriginal.BackColor = System.Drawing.SystemColors.Control
		Me.picOriginal.CausesValidation = True
		Me.picOriginal.Enabled = True
		Me.picOriginal.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picOriginal.Cursor = System.Windows.Forms.Cursors.Default
		Me.picOriginal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picOriginal.TabStop = True
		Me.picOriginal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picOriginal.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picOriginal.Name = "picOriginal"
		Me.Controls.Add(picPicture)
		Me.Controls.Add(picOriginal)
	End Sub
#End Region 
	'UPGRADE_WARNING: Event frmPicture.Resize may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub frmPicture_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		picPicture.Width = Me.ClientRectangle.Width
		picPicture.Height = Me.ClientRectangle.Height
		'UPGRADE_ISSUE: PictureBox method picPicture.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2064"'
		picPicture.CreateGraphics.Clear(Color.Black)
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2070"'
		'UPGRADE_ISSUE: PictureBox property picOriginal.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: PictureBox property picPicture.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2064"'
		StretchBlt(picPicture.CreateGraphics.GetHdc.ToInt32, 0, 0, picPicture.ClientRectangle.Width, picPicture.ClientRectangle.Height, picOriginal.CreateGraphics.GetHdc.ToInt32, 0, 0, picOriginal.ClientRectangle.Width, picOriginal.ClientRectangle.Height, &HCC0020)

		picPicture.Refresh()
	End Sub
End Class