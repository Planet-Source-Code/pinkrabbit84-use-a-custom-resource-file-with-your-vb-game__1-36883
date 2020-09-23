Option Strict Off
Option Explicit On
Friend Class frmDebug
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
	Public WithEvents txtDebug As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDebug))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.txtDebug = New System.Windows.Forms.TextBox
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Debug window"
		Me.ClientSize = New System.Drawing.Size(333, 247)
		Me.Location = New System.Drawing.Point(3, 24)
		Me.Icon = CType(resources.GetObject("frmDebug.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmDebug"
		Me.txtDebug.AutoSize = False
		Me.txtDebug.Size = New System.Drawing.Size(333, 249)
		Me.txtDebug.Location = New System.Drawing.Point(0, 0)
		Me.txtDebug.MultiLine = True
		Me.txtDebug.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.txtDebug.WordWrap = False
		Me.txtDebug.TabIndex = 0
		Me.txtDebug.AcceptsReturn = True
		Me.txtDebug.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDebug.BackColor = System.Drawing.SystemColors.Window
		Me.txtDebug.CausesValidation = True
		Me.txtDebug.Enabled = True
		Me.txtDebug.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDebug.HideSelection = True
		Me.txtDebug.ReadOnly = False
		Me.txtDebug.Maxlength = 0
		Me.txtDebug.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDebug.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDebug.TabStop = True
		Me.txtDebug.Visible = True
		Me.txtDebug.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDebug.Name = "txtDebug"
		Me.Controls.Add(txtDebug)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmDebug
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmDebug
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmDebug()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
End Class