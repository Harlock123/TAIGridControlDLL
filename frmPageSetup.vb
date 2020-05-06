Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Management

Public Class frmPageSetup
    Inherits System.Windows.Forms.Form

    'Private _psets As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings
    'Private _PaperSize As System.Drawing.Printing.PaperSize = _psets.PaperSize

    Private _psets As System.Drawing.Printing.PageSettings
    Private _PaperSize As System.Drawing.Printing.PaperSize

    Private _MaxPage As Integer = 1
    Private _MinPage As Integer = 1

    Private _Canceled As Boolean = True
    Private _Print As Boolean = False
    Private _Preview As Boolean = False
    Private _PrintAllPages As Boolean = True
    Private _PrintOrientationLandscape As Boolean = False

    Public Event PageSizeChanged(ByVal psiz As System.Drawing.Printing.PaperSize)
    Public Event OrientationChanged(ByVal LandscapeOrientation As Boolean)
    Public Event PaperMetricsHaveChanged(ByVal psiz As System.Drawing.Printing.PaperSize, ByVal LandscapeOrientation As Boolean)


#Region " Windows Form Designer generated code "

    Private Sub LogThis(ByVal str As String)

        'Dim r As System.IO.StreamWriter = System.IO.File.AppendText("C:\TAIGRIDLOG.TXT")

        'r.WriteLine(str)

        'r.Flush()
        'r.Close()


    End Sub

    Public Sub New()

        MyBase.New()
        InitializeComponent()

    End Sub

    Public Sub New(ByVal pset As System.Drawing.Printing.PageSettings)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        ' MsgBox(pset.ToString())

        Try

            LogThis("In PageSetup Form")
            _psets = pset

            LogThis("1")

            _PaperSize = _psets.PaperSize

            LogThis("2")

            cmboPaperSize.Items.Clear()

            LogThis("3")

            For Each psiz As System.Drawing.Printing.PaperSize In _psets.PrinterSettings.PaperSizes

                LogThis("Looping in some paper size Crapola")

                cmboPaperSize.Items.Add(psiz)

                LogThis(psiz.PaperName)

                'Console.WriteLine(psiz.PaperName)

                'If psiz.PaperName.Split(" ".ToCharArray)(0) = _
                '    _psets.PaperSize.PaperName.Split(" ".ToCharArray)(0) Then

                '    LogThis("Inside The If")
                '    cmboPaperSize.SelectedIndex = cmboPaperSize.Items.Count - 1
                'End If

            Next

            LogThis("Clearing Printer List")

            cmboPrinter.Items.Clear()

            For Each printer As String In Printing.PrinterSettings.InstalledPrinters

                LogThis("Adding a Printer " + printer.ToString())

                cmboPrinter.Items.Add(printer)

            Next

            LogThis("Printers Added Selecting It now")

            cmboPrinter.SelectedItem = _psets.PrinterSettings.PrinterName

            If _psets.PrinterSettings.PrintRange = Printing.PrintRange.AllPages Then
                rbAllPages.Checked = True
            Else
                rbPageRange.Checked = True
            End If

            If _psets.Landscape Then
                rbLandscape.Checked = True
            Else
                rbProtriate.Checked = True
            End If

            txtStartPage.Text = _psets.PrinterSettings.FromPage.ToString()
            txtEndPage.Text = _psets.PrinterSettings.ToPage.ToString()

        Catch ex As Exception
            '' What to do here I don't know so lets do nothing.

        End Try

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
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents cmboPaperSize As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbAllPages As System.Windows.Forms.RadioButton
    Friend WithEvents rbPageRange As System.Windows.Forms.RadioButton
    Friend WithEvents txtStartPage As System.Windows.Forms.TextBox
    Friend WithEvents txtEndPage As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbLandscape As System.Windows.Forms.RadioButton
    Friend WithEvents rbProtriate As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmboPrinter As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.cmboPaperSize = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtEndPage = New System.Windows.Forms.TextBox
        Me.txtStartPage = New System.Windows.Forms.TextBox
        Me.rbPageRange = New System.Windows.Forms.RadioButton
        Me.rbAllPages = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rbLandscape = New System.Windows.Forms.RadioButton
        Me.rbProtriate = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmboPrinter = New System.Windows.Forms.ComboBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnPreview = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(360, 208)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(60, 20)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(296, 208)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(60, 20)
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "Accept"
        '
        'cmboPaperSize
        '
        Me.cmboPaperSize.Location = New System.Drawing.Point(4, 48)
        Me.cmboPaperSize.Name = "cmboPaperSize"
        Me.cmboPaperSize.Size = New System.Drawing.Size(420, 21)
        Me.cmboPaperSize.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Select Paper Size"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtEndPage)
        Me.GroupBox1.Controls.Add(Me.txtStartPage)
        Me.GroupBox1.Controls.Add(Me.rbPageRange)
        Me.GroupBox1.Controls.Add(Me.rbAllPages)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(160, 108)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print What?"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(72, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "to"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtEndPage
        '
        Me.txtEndPage.Location = New System.Drawing.Point(100, 76)
        Me.txtEndPage.Name = "txtEndPage"
        Me.txtEndPage.Size = New System.Drawing.Size(44, 20)
        Me.txtEndPage.TabIndex = 3
        Me.txtEndPage.Text = ""
        '
        'txtStartPage
        '
        Me.txtStartPage.Location = New System.Drawing.Point(20, 76)
        Me.txtStartPage.Name = "txtStartPage"
        Me.txtStartPage.Size = New System.Drawing.Size(44, 20)
        Me.txtStartPage.TabIndex = 2
        Me.txtStartPage.Text = ""
        '
        'rbPageRange
        '
        Me.rbPageRange.Location = New System.Drawing.Point(20, 52)
        Me.rbPageRange.Name = "rbPageRange"
        Me.rbPageRange.Size = New System.Drawing.Size(124, 20)
        Me.rbPageRange.TabIndex = 1
        Me.rbPageRange.Text = "A Range Of Pages"
        '
        'rbAllPages
        '
        Me.rbAllPages.Checked = True
        Me.rbAllPages.Location = New System.Drawing.Point(20, 28)
        Me.rbAllPages.Name = "rbAllPages"
        Me.rbAllPages.Size = New System.Drawing.Size(124, 20)
        Me.rbAllPages.TabIndex = 0
        Me.rbAllPages.TabStop = True
        Me.rbAllPages.Text = "All Pages"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbLandscape)
        Me.GroupBox2.Controls.Add(Me.rbProtriate)
        Me.GroupBox2.Location = New System.Drawing.Point(260, 96)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(160, 108)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Orientation"
        '
        'rbLandscape
        '
        Me.rbLandscape.Location = New System.Drawing.Point(20, 52)
        Me.rbLandscape.Name = "rbLandscape"
        Me.rbLandscape.Size = New System.Drawing.Size(124, 20)
        Me.rbLandscape.TabIndex = 1
        Me.rbLandscape.Text = "Landscape"
        '
        'rbProtriate
        '
        Me.rbProtriate.Checked = True
        Me.rbProtriate.Location = New System.Drawing.Point(20, 28)
        Me.rbProtriate.Name = "rbProtriate"
        Me.rbProtriate.Size = New System.Drawing.Size(124, 20)
        Me.rbProtriate.TabIndex = 0
        Me.rbProtriate.TabStop = True
        Me.rbProtriate.Text = "Portrait"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(4, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(152, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Select Printer to Print to"
        '
        'cmboPrinter
        '
        Me.cmboPrinter.Location = New System.Drawing.Point(4, 4)
        Me.cmboPrinter.Name = "cmboPrinter"
        Me.cmboPrinter.Size = New System.Drawing.Size(420, 21)
        Me.cmboPrinter.TabIndex = 6
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(232, 208)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(60, 20)
        Me.btnPrint.TabIndex = 8
        Me.btnPrint.Text = "Print"
        '
        'btnPreview
        '
        Me.btnPreview.Location = New System.Drawing.Point(168, 208)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(60, 20)
        Me.btnPreview.TabIndex = 9
        Me.btnPreview.Text = "Preview"
        '
        'frmPageSetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(430, 231)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmboPrinter)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmboPaperSize)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmPageSetup"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setup Printed Page Metrics"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub HandleFormShow(ByVal sender As Object, ByVal ea As EventArgs) Handles MyBase.Activated

        'RaiseEvent PaperMetricsHaveChanged(_PaperSize, _PrintOrientationLandscape)
        'Application.DoEvents()
        'RaiseEvent PaperMetricsHaveChanged(_PaperSize, _PrintOrientationLandscape)

    End Sub

    Public Property Psets() As System.Drawing.Printing.PageSettings
        Get
            Return _psets
        End Get
        Set(ByVal Value As System.Drawing.Printing.PageSettings)
            _psets = Value
        End Set
    End Property

    Public Property PaperSize() As System.Drawing.Printing.PaperSize
        Get
            Return _PaperSize
        End Get
        Set(ByVal Value As System.Drawing.Printing.PaperSize)
            _PaperSize = Value
            _psets.PaperSize = Value
            cmboPaperSize.SelectedItem = Value
        End Set
    End Property

    Public Property MaxPage() As Integer
        Get
            Return _MaxPage
        End Get
        Set(ByVal Value As Integer)
            _MaxPage = Value
            Me.txtEndPage.Text = _MaxPage.ToString()
            Me.txtStartPage.Text = _MinPage.ToString()
        End Set
    End Property

    Public Property MinPage() As Integer
        Get
            Return _MinPage
        End Get
        Set(ByVal Value As Integer)
            _MinPage = Value
            Me.txtEndPage.Text = _MaxPage.ToString()
            Me.txtStartPage.Text = _MinPage.ToString()
        End Set
    End Property

    Public Property Canceled() As Boolean
        Get
            Return _Canceled
        End Get
        Set(ByVal Value As Boolean)
            _Canceled = Value
        End Set
    End Property

    Public Property Print() As Boolean
        Get
            Return _Print
        End Get
        Set(ByVal Value As Boolean)
            _Print = Value
        End Set
    End Property

    Public Property Preview() As Boolean
        Get
            Return _Preview
        End Get
        Set(ByVal Value As Boolean)
            _Preview = Value
        End Set
    End Property

    Public Property PrintAllPages() As Boolean
        Get
            Return _PrintAllPages
        End Get
        Set(ByVal Value As Boolean)
            _PrintAllPages = Value
            rbAllPages.Checked = Value
        End Set
    End Property

    Public Property PrintOrientationLandscape() As Boolean
        Get
            Return _PrintOrientationLandscape
        End Get
        Set(ByVal Value As Boolean)
            _PrintOrientationLandscape = Value
            rbLandscape.Checked = Value
        End Set
    End Property

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        _Canceled = False

        If rbAllPages.Checked Then
            _psets.PrinterSettings.PrintRange = Printing.PrintRange.AllPages

            _psets.PrinterSettings.ToPage = 0
            _psets.PrinterSettings.FromPage = 0

        Else
            _psets.PrinterSettings.PrintRange = Printing.PrintRange.SomePages

            If IsNumeric(txtStartPage.Text) Then
                _MinPage = Val(txtStartPage.Text)
            Else
                _MinPage = 1
            End If

            If IsNumeric(txtEndPage.Text) Then
                _MaxPage = Val(txtEndPage.Text)
            Else
                _MaxPage = _MinPage
            End If

            _psets.PrinterSettings.ToPage = _MaxPage
            _psets.PrinterSettings.FromPage = _MinPage

        End If


        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        _Canceled = True

        Me.Hide()
    End Sub

    Private Sub cmboPaperSize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmboPaperSize.SelectedIndexChanged

        Me.Refresh()

        Application.DoEvents()

        Try
            _PaperSize = cmboPaperSize.SelectedItem
            _psets.PaperSize = cmboPaperSize.SelectedItem
        Catch ex As Exception
            '' What to do here? I don't know so lets do nothing

        End Try

        LogThis("Calling event handler with a papersize of " + _PaperSize.ToString())

        RaiseEvent PaperMetricsHaveChanged(_PaperSize, _PrintOrientationLandscape)

        Application.DoEvents()

    End Sub

    Private Sub DecodeOrientation()

        Me.Refresh()

        Application.DoEvents()

        Try
            If rbProtriate.Checked Then
                _PrintOrientationLandscape = False
                If cmboPaperSize.SelectedItem Is Nothing Then
                    RaiseEvent PaperMetricsHaveChanged(_psets.PaperSize, _PrintOrientationLandscape)
                Else
                    RaiseEvent PaperMetricsHaveChanged(_PaperSize, _PrintOrientationLandscape)
                End If

            Else
                _PrintOrientationLandscape = True
                If cmboPaperSize.SelectedItem Is Nothing Then
                    RaiseEvent PaperMetricsHaveChanged(_psets.PaperSize, _PrintOrientationLandscape)
                Else
                    RaiseEvent PaperMetricsHaveChanged(_PaperSize, _PrintOrientationLandscape)
                End If

            End If
        Catch ex As Exception
            '' What to do here? I don't know so lets do nothing

        End Try



        Application.DoEvents()

    End Sub

    Private Sub rbProtriate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbProtriate.CheckedChanged
        DecodeOrientation()
    End Sub

    Private Sub rbLandscape_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbLandscape.CheckedChanged
        DecodeOrientation()
    End Sub

    Private Sub DecodePageRanges()

        If IsNumeric(txtStartPage.Text) Then
            _MinPage = Val(txtStartPage.Text)
        Else
            _MinPage = 1
        End If

        If IsNumeric(txtEndPage.Text) Then
            _MaxPage = Val(txtEndPage.Text)
        Else
            _MaxPage = _MinPage
        End If

    End Sub

    Private Sub cmboPrinter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmboPrinter.SelectedIndexChanged
        _psets.PrinterSettings.PrinterName = cmboPrinter.Text

        cmboPaperSize.Items.Clear()

        For Each psiz As System.Drawing.Printing.PaperSize In _psets.PrinterSettings.PaperSizes

            Try

                cmboPaperSize.Items.Add(psiz)

                ' If psiz.PaperName.Split(" ".ToCharArray)(0) = _
                '_psets.PaperSize.PaperName.Split(" ".ToCharArray)(0) Then
                '     cmboPaperSize.SelectedIndex = cmboPaperSize.Items.Count - 1
                ' End If

            Catch ex As Exception
                ' we are a try'n to add a crap paper size to things so lets skip it

            End Try


        Next

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        _Canceled = False
        _Print = True
        _Preview = False

        If rbAllPages.Checked Then
            _psets.PrinterSettings.PrintRange = Printing.PrintRange.AllPages

            _psets.PrinterSettings.ToPage = 0
            _psets.PrinterSettings.FromPage = 0
            _psets.Landscape = rbLandscape.Checked

        Else
            _psets.PrinterSettings.PrintRange = Printing.PrintRange.SomePages

            If IsNumeric(txtStartPage.Text) Then
                _MinPage = Val(txtStartPage.Text)
            Else
                _MinPage = 1
            End If

            If IsNumeric(txtEndPage.Text) Then
                _MaxPage = Val(txtEndPage.Text)
            Else
                _MaxPage = _MinPage
            End If

            _psets.PrinterSettings.ToPage = _MaxPage
            _psets.PrinterSettings.FromPage = _MinPage
            _psets.Landscape = rbLandscape.Checked

        End If

        Me.Hide()
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        _Canceled = False
        _Print = False
        _Preview = True

        If rbAllPages.Checked Then
            _psets.PrinterSettings.PrintRange = Printing.PrintRange.AllPages

            _psets.PrinterSettings.ToPage = 0
            _psets.PrinterSettings.FromPage = 0
            _psets.Landscape = rbLandscape.Checked

        Else
            _psets.PrinterSettings.PrintRange = Printing.PrintRange.SomePages

            If IsNumeric(txtStartPage.Text) Then
                _MinPage = Val(txtStartPage.Text)
            Else
                _MinPage = 1
            End If

            If IsNumeric(txtEndPage.Text) Then
                _MaxPage = Val(txtEndPage.Text)
            Else
                _MaxPage = _MinPage
            End If

            _psets.PrinterSettings.ToPage = _MaxPage
            _psets.PrinterSettings.FromPage = _MinPage
            _psets.Landscape = rbLandscape.Checked

        End If

        Me.Hide()
    End Sub
End Class
