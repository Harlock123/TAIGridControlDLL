Public Class frmMultipleColumnTearAway
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal Headerlist() As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        chkList.Items.Clear()

        Dim t As Integer

        For t = Headerlist.GetLowerBound(0) To Headerlist.GetUpperBound(0)
            If Not Headerlist(t) Is Nothing Then
                chkList.Items.Add(Headerlist(t), False)
            End If
        Next

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
    Friend WithEvents chkList As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnAll As System.Windows.Forms.Button
    Friend WithEvents btnNone As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.chkList = New System.Windows.Forms.CheckedListBox
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnAll = New System.Windows.Forms.Button
        Me.btnNone = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'chkList
        '
        Me.chkList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkList.CheckOnClick = True
        Me.chkList.Location = New System.Drawing.Point(4, 24)
        Me.chkList.Name = "chkList"
        Me.chkList.Size = New System.Drawing.Size(232, 304)
        Me.chkList.TabIndex = 0
        '
        'btnOk
        '
        Me.btnOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOk.Location = New System.Drawing.Point(20, 335)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(76, 24)
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(136, 335)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(76, 24)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'btnAll
        '
        Me.btnAll.Location = New System.Drawing.Point(4, 0)
        Me.btnAll.Name = "btnAll"
        Me.btnAll.Size = New System.Drawing.Size(40, 20)
        Me.btnAll.TabIndex = 3
        Me.btnAll.Text = "All"
        '
        'btnNone
        '
        Me.btnNone.Location = New System.Drawing.Point(48, 0)
        Me.btnNone.Name = "btnNone"
        Me.btnNone.Size = New System.Drawing.Size(40, 20)
        Me.btnNone.TabIndex = 4
        Me.btnNone.Text = "None"
        '
        'frmMultipleColumnTearAway
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(240, 364)
        Me.Controls.Add(Me.btnNone)
        Me.Controls.Add(Me.btnAll)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.chkList)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(212, 184)
        Me.Name = "frmMultipleColumnTearAway"
        Me.Text = "Select Columns to Tear Away"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    '
    ' Internal Variable Declarations
    '

    Private m_Canceled As Boolean = True

    '
    ' Propertys are defined here
    '

    Property Canceled() As Boolean
        Get
            Return m_Canceled
        End Get
        Set(ByVal Value As Boolean)
            m_Canceled = Value
        End Set
    End Property

    ReadOnly Property SelectedIndices() As System.Windows.Forms.CheckedListBox.CheckedIndexCollection
        Get
            Return chkList.CheckedIndices
        End Get
    End Property

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        m_Canceled = False
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_Canceled = True
        Me.Hide()
    End Sub

    Private Sub btnAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAll.Click
        Dim t As Integer

        If chkList.Items.Count = 0 Then
            Exit Sub
        End If

        For t = 0 To chkList.Items.Count - 1
            chkList.SetItemChecked(t, True)
        Next
    End Sub

    Private Sub btnNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNone.Click
        Dim t As Integer

        If chkList.Items.Count = 0 Then
            Exit Sub
        End If

        For t = 0 To chkList.Items.Count - 1
            chkList.SetItemChecked(t, False)
        Next
    End Sub
End Class
