Option Strict Off
Option Explicit On
Friend Class frmFileType
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdRemove As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents cboKind As System.Windows.Forms.ComboBox
	Public WithEvents lstVideo As System.Windows.Forms.ListBox
	Public WithEvents lstVarious As System.Windows.Forms.ListBox
	Public WithEvents lstText As System.Windows.Forms.ListBox
	Public WithEvents lstSound As System.Windows.Forms.ListBox
	Public WithEvents lstPicture As System.Windows.Forms.ListBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFileType))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdRemove = New System.Windows.Forms.Button
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.cboKind = New System.Windows.Forms.ComboBox
		Me.lstVideo = New System.Windows.Forms.ListBox
		Me.lstVarious = New System.Windows.Forms.ListBox
		Me.lstText = New System.Windows.Forms.ListBox
		Me.lstSound = New System.Windows.Forms.ListBox
		Me.lstPicture = New System.Windows.Forms.ListBox
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "File type"
		Me.ClientSize = New System.Drawing.Size(177, 267)
		Me.Location = New System.Drawing.Point(3, 24)
		Me.Icon = CType(resources.GetObject("frmFileType.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmFileType"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(92, 240)
		Me.cmdCancel.TabIndex = 5
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(4, 240)
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdRemove.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRemove.Text = "Remove"
		Me.cmdRemove.Enabled = False
		Me.cmdRemove.Size = New System.Drawing.Size(85, 25)
		Me.cmdRemove.Location = New System.Drawing.Point(88, 204)
		Me.cmdRemove.TabIndex = 3
		Me.cmdRemove.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRemove.CausesValidation = True
		Me.cmdRemove.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRemove.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRemove.TabStop = True
		Me.cmdRemove.Name = "cmdRemove"
		Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdd.Text = "Add"
		Me.cmdAdd.Size = New System.Drawing.Size(85, 25)
		Me.cmdAdd.Location = New System.Drawing.Point(4, 204)
		Me.cmdAdd.TabIndex = 2
		Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdd.CausesValidation = True
		Me.cmdAdd.Enabled = True
		Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdd.TabStop = True
		Me.cmdAdd.Name = "cmdAdd"
		Me.cboKind.Size = New System.Drawing.Size(169, 21)
		Me.cboKind.Location = New System.Drawing.Point(4, 4)
		Me.cboKind.Items.AddRange(New Object(){"Picture", "Sound", "Text", "Various", "Video"})
		Me.cboKind.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboKind.TabIndex = 0
		Me.cboKind.BackColor = System.Drawing.SystemColors.Window
		Me.cboKind.CausesValidation = True
		Me.cboKind.Enabled = True
		Me.cboKind.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboKind.IntegralHeight = True
		Me.cboKind.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboKind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboKind.Sorted = False
		Me.cboKind.TabStop = True
		Me.cboKind.Visible = True
		Me.cboKind.Name = "cboKind"
		Me.lstVideo.Size = New System.Drawing.Size(169, 176)
		Me.lstVideo.Location = New System.Drawing.Point(4, 28)
		Me.lstVideo.Sorted = True
		Me.lstVideo.TabIndex = 9
		Me.lstVideo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstVideo.BackColor = System.Drawing.SystemColors.Window
		Me.lstVideo.CausesValidation = True
		Me.lstVideo.Enabled = True
		Me.lstVideo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstVideo.IntegralHeight = True
		Me.lstVideo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstVideo.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstVideo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstVideo.TabStop = True
		Me.lstVideo.Visible = True
		Me.lstVideo.MultiColumn = False
		Me.lstVideo.Name = "lstVideo"
		Me.lstVarious.Size = New System.Drawing.Size(169, 176)
		Me.lstVarious.Location = New System.Drawing.Point(4, 28)
		Me.lstVarious.Sorted = True
		Me.lstVarious.TabIndex = 8
		Me.lstVarious.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstVarious.BackColor = System.Drawing.SystemColors.Window
		Me.lstVarious.CausesValidation = True
		Me.lstVarious.Enabled = True
		Me.lstVarious.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstVarious.IntegralHeight = True
		Me.lstVarious.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstVarious.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstVarious.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstVarious.TabStop = True
		Me.lstVarious.Visible = True
		Me.lstVarious.MultiColumn = False
		Me.lstVarious.Name = "lstVarious"
		Me.lstText.Size = New System.Drawing.Size(169, 176)
		Me.lstText.Location = New System.Drawing.Point(4, 28)
		Me.lstText.Sorted = True
		Me.lstText.TabIndex = 7
		Me.lstText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstText.BackColor = System.Drawing.SystemColors.Window
		Me.lstText.CausesValidation = True
		Me.lstText.Enabled = True
		Me.lstText.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstText.IntegralHeight = True
		Me.lstText.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstText.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstText.TabStop = True
		Me.lstText.Visible = True
		Me.lstText.MultiColumn = False
		Me.lstText.Name = "lstText"
		Me.lstSound.Size = New System.Drawing.Size(169, 176)
		Me.lstSound.Location = New System.Drawing.Point(4, 28)
		Me.lstSound.Sorted = True
		Me.lstSound.TabIndex = 6
		Me.lstSound.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstSound.BackColor = System.Drawing.SystemColors.Window
		Me.lstSound.CausesValidation = True
		Me.lstSound.Enabled = True
		Me.lstSound.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstSound.IntegralHeight = True
		Me.lstSound.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstSound.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstSound.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstSound.TabStop = True
		Me.lstSound.Visible = True
		Me.lstSound.MultiColumn = False
		Me.lstSound.Name = "lstSound"
		Me.lstPicture.Size = New System.Drawing.Size(169, 176)
		Me.lstPicture.Location = New System.Drawing.Point(4, 28)
		Me.lstPicture.Sorted = True
		Me.lstPicture.TabIndex = 1
		Me.lstPicture.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstPicture.BackColor = System.Drawing.SystemColors.Window
		Me.lstPicture.CausesValidation = True
		Me.lstPicture.Enabled = True
		Me.lstPicture.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstPicture.IntegralHeight = True
		Me.lstPicture.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstPicture.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstPicture.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstPicture.TabStop = True
		Me.lstPicture.Visible = True
		Me.lstPicture.MultiColumn = False
		Me.lstPicture.Name = "lstPicture"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdRemove)
		Me.Controls.Add(cmdAdd)
		Me.Controls.Add(cboKind)
		Me.Controls.Add(lstVideo)
		Me.Controls.Add(lstVarious)
		Me.Controls.Add(lstText)
		Me.Controls.Add(lstSound)
		Me.Controls.Add(lstPicture)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmFileType
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmFileType
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmFileType()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'UPGRADE_WARNING: Event cboKind.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cboKind_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboKind.SelectedIndexChanged
		lstPicture.SelectedIndex = -1
		lstSound.SelectedIndex = -1
		lstText.SelectedIndex = -1
		lstVarious.SelectedIndex = -1
		lstVideo.SelectedIndex = -1
		cmdRemove.Enabled = False
		Select Case cboKind.SelectedIndex
			Case 0
				lstPicture.BringToFront()
			Case 1
				lstSound.BringToFront()
			Case 2
				lstText.BringToFront()
			Case 3
				lstVarious.BringToFront()
			Case 4
				lstVideo.BringToFront()
		End Select
	End Sub
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		Dim Ext As String
		
		Select Case cboKind.SelectedIndex
			Case 0
				Ext = InputBox("New extention for picture", "Picture")
				If Ext <> "" And CheckExtExist(Ext) Then lstPicture.Items.Add(Ext)
			Case 1
				Ext = InputBox("New extention for sound", "Sound")
				If Ext <> "" And CheckExtExist(Ext) Then lstSound.Items.Add(Ext)
			Case 2
				Ext = InputBox("New extention for text", "Text")
				If Ext <> "" And CheckExtExist(Ext) Then lstText.Items.Add(Ext)
			Case 3
				Ext = InputBox("New extention for various", "Various")
				If Ext <> "" And CheckExtExist(Ext) Then lstVarious.Items.Add(Ext)
			Case 4
				Ext = InputBox("New extention for Video", "Video")
				If Ext <> "" And CheckExtExist(Ext) Then lstVideo.Items.Add(Ext)
		End Select
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim i As Object
		ReDim PictureExt(lstPicture.Items.Count - 1)
		For i = 0 To lstPicture.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			PictureExt(i) = VB6.GetItemString(lstPicture, i)
		Next i
		
		ReDim SoundExt(lstSound.Items.Count - 1)
		For i = 0 To lstSound.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			SoundExt(i) = VB6.GetItemString(lstSound, i)
		Next i
		
		ReDim TextExt(lstText.Items.Count - 1)
		For i = 0 To lstText.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			TextExt(i) = VB6.GetItemString(lstText, i)
		Next i
		
		ReDim VariousExt(lstVarious.Items.Count - 1)
		For i = 0 To lstVarious.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			VariousExt(i) = VB6.GetItemString(lstVarious, i)
		Next i
		
		ReDim VideoExt(lstVideo.Items.Count - 1)
		For i = 0 To lstVideo.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			VideoExt(i) = VB6.GetItemString(lstVideo, i)
		Next i
		
		SaveBindExt()
		
		Me.Close()
	End Sub
	
	Private Sub cmdRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRemove.Click
		Select Case cboKind.SelectedIndex
			Case 0
				lstPicture.Items.RemoveAt(lstPicture.SelectedIndex)
			Case 1
				lstSound.Items.RemoveAt(lstSound.SelectedIndex)
			Case 2
				lstText.Items.RemoveAt(lstText.SelectedIndex)
			Case 3
				lstVarious.Items.RemoveAt(lstVarious.SelectedIndex)
			Case 4
				lstVideo.Items.RemoveAt(lstVideo.SelectedIndex)
		End Select
		cmdRemove.Enabled = False
	End Sub
	
	Private Sub frmFileType_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Object
		For i = 0 To UBound(PictureExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If PictureExt(i) <> "" Then lstPicture.Items.Add(PictureExt(i))
		Next i
		
		For i = 0 To UBound(SoundExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If SoundExt(i) <> "" Then lstSound.Items.Add(SoundExt(i))
		Next i
		
		For i = 0 To UBound(TextExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If TextExt(i) <> "" Then lstText.Items.Add(TextExt(i))
		Next i
		
		For i = 0 To UBound(VariousExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If VariousExt(i) <> "" Then lstVarious.Items.Add(VariousExt(i))
		Next i
		
		For i = 0 To UBound(VideoExt)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If VideoExt(i) <> "" Then lstVideo.Items.Add(VideoExt(i))
		Next i
		
		cboKind.SelectedIndex = 0
	End Sub
	
	'UPGRADE_WARNING: Event lstPicture.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub lstPicture_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPicture.SelectedIndexChanged
		cmdRemove.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Event lstSound.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub lstSound_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstSound.SelectedIndexChanged
		cmdRemove.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Event lstText.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub lstText_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstText.SelectedIndexChanged
		cmdRemove.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Event lstVarious.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub lstVarious_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstVarious.SelectedIndexChanged
		cmdRemove.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Event lstVideo.SelectedIndexChanged may fire when form is intialized. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub lstVideo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstVideo.SelectedIndexChanged
		cmdRemove.Enabled = True
	End Sub
	
	Function CheckExtExist(ByRef Ext As String) As Boolean
		Dim i As Object
		Dim Exist As Boolean
		
		For i = 0 To lstPicture.Items.Count - 1
			If VB6.GetItemString(lstPicture, i) = Ext Then Exist = True
		Next i
		
		For i = 0 To lstSound.Items.Count - 1
			If VB6.GetItemString(lstSound, i) = Ext Then Exist = True
		Next i
		
		For i = 0 To lstText.Items.Count - 1
			If VB6.GetItemString(lstText, i) = Ext Then Exist = True
		Next i
		
		For i = 0 To lstVarious.Items.Count - 1
			If VB6.GetItemString(lstVarious, i) = Ext Then Exist = True
		Next i
		
		For i = 0 To lstVideo.Items.Count - 1
			If VB6.GetItemString(lstVideo, i) = Ext Then Exist = True
		Next i
		
		If Exist Then MsgBox("Extension already exist in one of these list")
		
		CheckExtExist = Not Exist
	End Function
End Class