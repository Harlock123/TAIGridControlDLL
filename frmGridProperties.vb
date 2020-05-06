Public Class frmGridProperties
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Private _pg As TAIGridControl

    Public Sub New(ByVal ParentGrid As TAIGridControl)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        _pg = ParentGrid

        chkShowHeaders.Checked = _pg.GridheaderVisible
        chkShowTitle.Checked = _pg.TitleVisible

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFontsSmaller As System.Windows.Forms.Button
    Friend WithEvents btnFontsLarger As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents chkShowTitle As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowHeaders As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFontsSmaller = New System.Windows.Forms.Button
        Me.btnFontsLarger = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.chkShowTitle = New System.Windows.Forms.CheckBox
        Me.chkShowHeaders = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Grids Fonts"
        '
        'btnFontsSmaller
        '
        Me.btnFontsSmaller.Location = New System.Drawing.Point(12, 28)
        Me.btnFontsSmaller.Name = "btnFontsSmaller"
        Me.btnFontsSmaller.Size = New System.Drawing.Size(16, 20)
        Me.btnFontsSmaller.TabIndex = 1
        Me.btnFontsSmaller.Text = "<"
        '
        'btnFontsLarger
        '
        Me.btnFontsLarger.Location = New System.Drawing.Point(44, 28)
        Me.btnFontsLarger.Name = "btnFontsLarger"
        Me.btnFontsLarger.Size = New System.Drawing.Size(16, 20)
        Me.btnFontsLarger.TabIndex = 2
        Me.btnFontsLarger.Text = ">"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(364, 20)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(72, 24)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Close"
        '
        'chkShowTitle
        '
        Me.chkShowTitle.Location = New System.Drawing.Point(144, 16)
        Me.chkShowTitle.Name = "chkShowTitle"
        Me.chkShowTitle.Size = New System.Drawing.Size(152, 16)
        Me.chkShowTitle.TabIndex = 4
        Me.chkShowTitle.Text = "Show Grid Title Bar"
        '
        'chkShowHeaders
        '
        Me.chkShowHeaders.Location = New System.Drawing.Point(144, 36)
        Me.chkShowHeaders.Name = "chkShowHeaders"
        Me.chkShowHeaders.Size = New System.Drawing.Size(172, 16)
        Me.chkShowHeaders.TabIndex = 5
        Me.chkShowHeaders.Text = "Show Grid Column Headers"
        '
        'frmGridProperties
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Info
        Me.ClientSize = New System.Drawing.Size(442, 68)
        Me.Controls.Add(Me.chkShowHeaders)
        Me.Controls.Add(Me.chkShowTitle)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnFontsLarger)
        Me.Controls.Add(Me.btnFontsSmaller)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGridProperties"
        Me.Text = "Properties for the Grid"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Hide()
    End Sub

    Private Sub btnFontsSmaller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFontsSmaller.Click
        _pg.AllFontsSmaller()
    End Sub

    Private Sub btnFontsLarger_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFontsLarger.Click
        _pg.AllFontsLarger()
    End Sub

    Private Sub chkShowTitle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowTitle.CheckedChanged
        _pg.TitleVisible = chkShowTitle.Checked
    End Sub

    Private Sub chkShowHeaders_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowHeaders.CheckedChanged
        _pg.GridheaderVisible = chkShowHeaders.Checked
    End Sub
End Class
