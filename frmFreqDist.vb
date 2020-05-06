Public Class frmFreqDist
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal TAIG As TAIGridControl, ByVal coltoCount As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        taigFreq.FrequencyDistribution(TAIG, coltoCount, True)

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
    Friend WithEvents taigFreq As TAIGridControl
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.taigFreq = New TAIGridControl
        Me.btnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'taigFreq
        '
        Me.taigFreq.AlternateColoration = False
        Me.taigFreq.AlternateColorationAltColor = System.Drawing.Color.MediumSpringGreen
        Me.taigFreq.AlternateColorationBaseColor = System.Drawing.Color.AntiqueWhite
        Me.taigFreq.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.taigFreq.BorderColor = System.Drawing.Color.Black
        Me.taigFreq.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.taigFreq.CellOutlines = True
        Me.taigFreq.ColBackColorEdit = System.Drawing.Color.Yellow
        Me.taigFreq.Cols = 0
        Me.taigFreq.DefaultBackgroundColor = System.Drawing.Color.AntiqueWhite
        Me.taigFreq.DefaultCellFont = New System.Drawing.Font("Arial", 9.0!)
        Me.taigFreq.DefaultForegroundColor = System.Drawing.Color.Black
        Me.taigFreq.Delimiter = ","
        Me.taigFreq.ExcelAlternateColoration = System.Drawing.Color.FromArgb(CType(204, Byte), CType(255, Byte), CType(204, Byte))
        Me.taigFreq.ExcelAutoFitColumn = True
        Me.taigFreq.ExcelAutoFitRow = True
        Me.taigFreq.ExcelFilename = ""
        Me.taigFreq.ExcelIncludeColumnHeaders = True
        Me.taigFreq.ExcelKeepAlive = True
        Me.taigFreq.ExcelMatchGridColorScheme = True
        Me.taigFreq.ExcelMaximized = True
        Me.taigFreq.ExcelOutlineCells = True
        Me.taigFreq.ExcelPageOrientation = 1
        Me.taigFreq.ExcelShowBorders = False
        Me.taigFreq.ExcelUseAlternateRowColor = True
        Me.taigFreq.ExcelWorksheetName = "Grid Output"
        Me.taigFreq.GridHeaderBackColor = System.Drawing.Color.LightBlue
        Me.taigFreq.GridHeaderFont = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.taigFreq.GridHeaderForeColor = System.Drawing.Color.Black
        Me.taigFreq.GridHeaderHeight = 16
        Me.taigFreq.GridheaderVisible = True
        Me.taigFreq.Location = New System.Drawing.Point(0, 4)
        Me.taigFreq.Name = "taigFreq"
        Me.taigFreq.PaginationSize = 0
        Me.taigFreq.Rows = 0
        Me.taigFreq.ScrollInterval = 5
        Me.taigFreq.SelectedColBackColor = System.Drawing.Color.MediumSlateBlue
        Me.taigFreq.SelectedColForeColor = System.Drawing.Color.LightGray
        Me.taigFreq.SelectedColumn = -1
        Me.taigFreq.SelectedRow = -1
        Me.taigFreq.SelectedRowBackColor = System.Drawing.Color.Blue
        Me.taigFreq.SelectedRowForeColor = System.Drawing.Color.White
        Me.taigFreq.Size = New System.Drawing.Size(392, 344)
        Me.taigFreq.TabIndex = 0
        Me.taigFreq.TitleBackColor = System.Drawing.Color.Blue
        Me.taigFreq.TitleFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.taigFreq.TitleForeColor = System.Drawing.Color.White
        Me.taigFreq.TitleText = "Frequency Distribution"
        Me.taigFreq.TitleVisible = True
        Me.taigFreq.XMLDataSetName = "Grid_Output"
        Me.taigFreq.XMLFileName = ""
        Me.taigFreq.XMLIncludeSchema = False
        Me.taigFreq.XMLNameSpace = "TAI_Grid_Ouptut"
        Me.taigFreq.XMLTableName = "Table"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(156, 352)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 20)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        '
        'frmFreqDist
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(394, 375)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.taigFreq)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "frmFreqDist"
        Me.ShowInTaskbar = False
        Me.Text = "Frequency Distribution in source grid"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Hide()
    End Sub
End Class
