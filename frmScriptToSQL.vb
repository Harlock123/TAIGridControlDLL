Public Class frmScriptToSQL
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Private _taig As TAIGridControl

    Public Sub New(ByVal TAIG As TAIGridControl)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        _taig = TAIG

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
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents btnSelectFile As System.Windows.Forms.Button
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents txtTableName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtFileName = New System.Windows.Forms.TextBox
        Me.btnSelectFile = New System.Windows.Forms.Button
        Me.OFD = New System.Windows.Forms.OpenFileDialog
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.txtTableName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(4, 44)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(344, 20)
        Me.txtFileName.TabIndex = 0
        Me.txtFileName.Text = ""
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(352, 44)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(24, 20)
        Me.btnSelectFile.TabIndex = 1
        Me.btnSelectFile.Text = "..."
        '
        'OFD
        '
        Me.OFD.CheckFileExists = False
        Me.OFD.DefaultExt = "sql"
        Me.OFD.Title = "Select file name to save the script to"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(260, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 20)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(320, 4)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(56, 20)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "Ok"
        '
        'txtTableName
        '
        Me.txtTableName.Location = New System.Drawing.Point(4, 20)
        Me.txtTableName.Name = "txtTableName"
        Me.txtTableName.Size = New System.Drawing.Size(152, 20)
        Me.txtTableName.TabIndex = 4
        Me.txtTableName.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(148, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Name the resulting table"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(188, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Save the resulting SQL script to"
        '
        'frmScriptToSQL
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(394, 88)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtTableName)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSelectFile)
        Me.Controls.Add(Me.txtFileName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmScriptToSQL"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Script Grid contents to SQL"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If txtFileName.Text = "" Or txtTableName.Text = "" Then

            MsgBox("You must select a file for the resulting Script to be written to" + vbCrLf + _
                   "as well as a name for the resulting Table that is crafted by that script", _
                    MsgBoxStyle.Information, "Script Grid to SQL error")
        Else

            Dim txtw As System.IO.TextWriter = New System.IO.StreamWriter(txtFileName.Text)
            txtw.Write(_taig.CreatePersistanceScript(txtTableName.Text))
            txtw.Close()

            Me.Hide()

        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub

    Private Sub btnSelectFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectFile.Click
        OFD.CheckFileExists = False
        If OFD.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtFileName.Text = OFD.FileName
        End If
    End Sub
End Class
