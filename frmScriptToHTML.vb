Public Class frmScriptToHTML
    Inherits System.Windows.Forms.Form

    Private _Taig As TAIGridControl

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal TAIG As TAIGridControl)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        _Taig = TAIG

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSelectFile As System.Windows.Forms.Button
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtBorder As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkMatchColors As System.Windows.Forms.CheckBox
    Friend WithEvents chkOmitNulls As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSelectFile = New System.Windows.Forms.Button
        Me.txtFileName = New System.Windows.Forms.TextBox
        Me.OFD = New System.Windows.Forms.OpenFileDialog
        Me.txtBorder = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkMatchColors = New System.Windows.Forms.CheckBox
        Me.chkOmitNulls = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(4, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(212, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Save the resulting HTML Table script to"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(316, 4)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(56, 20)
        Me.btnOk.TabIndex = 10
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(256, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 20)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(348, 44)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(24, 20)
        Me.btnSelectFile.TabIndex = 8
        Me.btnSelectFile.Text = "..."
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(4, 44)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(340, 20)
        Me.txtFileName.TabIndex = 7
        Me.txtFileName.Text = ""
        '
        'OFD
        '
        Me.OFD.CheckFileExists = False
        Me.OFD.DefaultExt = "html"
        Me.OFD.Title = "Select file name to save the script to"
        '
        'txtBorder
        '
        Me.txtBorder.Location = New System.Drawing.Point(4, 20)
        Me.txtBorder.Name = "txtBorder"
        Me.txtBorder.Size = New System.Drawing.Size(28, 20)
        Me.txtBorder.TabIndex = 12
        Me.txtBorder.Text = "1"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 12)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Border"
        '
        'chkMatchColors
        '
        Me.chkMatchColors.Checked = True
        Me.chkMatchColors.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMatchColors.Location = New System.Drawing.Point(40, 24)
        Me.chkMatchColors.Name = "chkMatchColors"
        Me.chkMatchColors.Size = New System.Drawing.Size(96, 16)
        Me.chkMatchColors.TabIndex = 14
        Me.chkMatchColors.Text = "Match Colors"
        '
        'chkOmitNulls
        '
        Me.chkOmitNulls.Checked = True
        Me.chkOmitNulls.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOmitNulls.Location = New System.Drawing.Point(144, 24)
        Me.chkOmitNulls.Name = "chkOmitNulls"
        Me.chkOmitNulls.Size = New System.Drawing.Size(96, 16)
        Me.chkOmitNulls.TabIndex = 15
        Me.chkOmitNulls.Text = "Omit Nulls"
        '
        'frmScriptToHTML
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 83)
        Me.ControlBox = False
        Me.Controls.Add(Me.chkOmitNulls)
        Me.Controls.Add(Me.chkMatchColors)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtBorder)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSelectFile)
        Me.Controls.Add(Me.txtFileName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmScriptToHTML"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Script Grid as an HTML Table"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub

    Private Sub btnSelectFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectFile.Click
        OFD.CheckFileExists = False
        If OFD.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtFileName.Text = OFD.FileName
        End If
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If txtFileName.Text = "" Then

            MsgBox("You must select a failename to save the resulting HTML Table to...", _
                    MsgBoxStyle.Information, "Save to HTML Table error")


        Else

            Dim bv As Integer = 0

            If IsNumeric(txtBorder.Text) Then
                bv = Integer.Parse(txtBorder.Text)
            End If

            Dim txtw As System.IO.TextWriter = New System.IO.StreamWriter(txtFileName.Text)
            txtw.Write(_Taig.CreateHTMLTableScript(bv, chkMatchColors.Checked, chkOmitNulls.Checked))
            txtw.Close()

            Me.Hide()

        End If
    End Sub
End Class
