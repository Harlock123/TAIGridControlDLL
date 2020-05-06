Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class frmColumnTearAway
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal title As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.Text = title
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)

        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
            If Not (m_gridParent Is Nothing) Then
                ' callback into the parent grid to kill it from the display list
                m_gridParent.KillTearAwayColumnWindow(m_colid)
            End If
        End If

            MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents vscroller As System.Windows.Forms.VScrollBar
    Friend WithEvents tt1 As System.Windows.Forms.ToolTip
    Friend WithEvents PulldownMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents miAutoArrange As System.Windows.Forms.MenuItem
    Friend WithEvents miPushtoBack As System.Windows.Forms.MenuItem
    Friend WithEvents miPullAllTearAwaysToTheFront As System.Windows.Forms.MenuItem
    Friend WithEvents miHideAllTearAways As System.Windows.Forms.MenuItem

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.vscroller = New System.Windows.Forms.VScrollBar
        Me.tt1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PulldownMenu = New System.Windows.Forms.ContextMenu
        Me.miAutoArrange = New System.Windows.Forms.MenuItem
        Me.miPushtoBack = New System.Windows.Forms.MenuItem
        Me.miPullAllTearAwaysToTheFront = New System.Windows.Forms.MenuItem
        Me.miHideAllTearAways = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'vscroller
        '
        Me.vscroller.Dock = System.Windows.Forms.DockStyle.Right
        Me.vscroller.Location = New System.Drawing.Point(256, 0)
        Me.vscroller.Name = "vscroller"
        Me.vscroller.Size = New System.Drawing.Size(12, 340)
        Me.vscroller.TabIndex = 0
        '
        'tt1
        '
        Me.tt1.ShowAlways = True
        '
        'PulldownMenu
        '
        Me.PulldownMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miAutoArrange, Me.miHideAllTearAways, Me.miPushtoBack, Me.miPullAllTearAwaysToTheFront})
        '
        'miAutoArrange
        '
        Me.miAutoArrange.Index = 0
        Me.miAutoArrange.Text = "Auto Arrange All Tear Away Columns"
        '
        'miPushtoBack
        '
        Me.miPushtoBack.Index = 2
        Me.miPushtoBack.Text = "Push All Tear Away's to the Back"
        '
        'miPullAllTearAwaysToTheFront
        '
        Me.miPullAllTearAwaysToTheFront.Index = 3
        Me.miPullAllTearAwaysToTheFront.Text = "Pull All Tear Away's to the Front"
        '
        'miHideAllTearAways
        '
        Me.miHideAllTearAways.Index = 1
        Me.miHideAllTearAways.Text = "Hide All Tear Away Columns"
        '
        'frmColumnTearAway
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(268, 340)
        Me.Controls.Add(Me.vscroller)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmColumnTearAway"
        Me.ShowInTaskbar = False
        Me.Text = "frmColumnTearAway"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Internal Declarations"

    Private _Painting As Boolean = False
    Private m_VertScrollIndex As Integer
    Private m_VertScrollMin As Integer
    Private m_VertScrollMax As Integer
    Private m_VertScrollVisible As Boolean
    Private m_colid As Integer
    Private m_RowCLicked As Integer = -1

    Private m_ListItems As ArrayList = New ArrayList

    Private _SelectedRow As Integer = -1

    Private m_gridParent As TAIGridControl

    Private m_DisplayFont As Font = New Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Point)

    Private _DefaultBackColor As Color = Color.AntiqueWhite
    Private _DefaultSelectionColor As Color = Color.Blue
    Private _DefaultSelectionForeColor As Color = Color.White
    Private _DefaultForeColor As Color = System.Drawing.Color.Black

#End Region

#Region "Properties"

    Property Colid() As Integer
        Get
            Return m_colid
        End Get
        Set(ByVal Value As Integer)
            m_colid = Value
        End Set
    End Property

    Property SelectedRow() As Integer
        Get
            Return _SelectedRow
        End Get
        Set(ByVal Value As Integer)
            If _SelectedRow <> Value Then

                _SelectedRow = Value
                Me.Invalidate()

            End If
        End Set
    End Property

    Property DefaultSelectionColor() As Color
        Get
            Return _DefaultSelectionColor
        End Get
        Set(ByVal Value As Color)
            _DefaultSelectionColor = Value
            Me.Invalidate()
        End Set
    End Property

    Property DefaultSelectionForColor() As Color
        Get
            Return _DefaultSelectionForeColor
        End Get
        Set(ByVal Value As Color)
            _DefaultSelectionForeColor = Value
            Me.Invalidate()
        End Set
    End Property

    Property GridDefaultBackColor() As Color
        Get
            Return _DefaultBackColor
        End Get
        Set(ByVal Value As Color)
            _DefaultBackColor = Value
            Me.Invalidate()

        End Set
    End Property

    Property GridDefaultForeColor() As Color
        Get
            Return _DefaultForeColor
        End Get
        Set(ByVal Value As Color)
            _DefaultForeColor = Value
            Me.Invalidate()
        End Set
    End Property

    Property DisplayFont() As Font
        Get
            Return (m_DisplayFont)
        End Get
        Set(ByVal Value As Font)
            m_DisplayFont = Value
            Me.Invalidate()
        End Set
    End Property

    Property ListItems() As ArrayList
        Get
            Return (m_ListItems)
        End Get
        Set(ByVal Value As ArrayList)
            m_ListItems = Value
            If Me.Visible Then
                Me.Invalidate()
            End If
        End Set
    End Property

    Property GridParent() As TAIGridControl
        Get
            Return m_gridParent
        End Get
        Set(ByVal Value As TAIGridControl)
            m_gridParent = Value
        End Set
    End Property

    Property VertScrollIndex() As Integer
        Get
            Return m_VertScrollIndex
        End Get
        Set(ByVal Value As Integer)
            m_VertScrollIndex = Value
            vscroller.Value = Value
        End Set
    End Property

    Property VertScrollMin() As Integer
        Get
            Return m_VertScrollMin
        End Get
        Set(ByVal Value As Integer)
            m_VertScrollMin = Value
            vscroller.Minimum = Value
        End Set
    End Property

    Property VertScrollMax() As Integer
        Get
            Return m_VertScrollMax
        End Get
        Set(ByVal Value As Integer)
            m_VertScrollMax = Value
            vscroller.Maximum = Value
        End Set
    End Property

    Property VertScrollVisible() As Boolean
        Get
            Return m_VertScrollVisible
        End Get
        Set(ByVal Value As Boolean)
            m_VertScrollVisible = Value
            vscroller.Visible = Value
        End Set
    End Property

    Public Sub KillMe(ByVal colid As Integer)

        If m_colid = colid Then
            Me.Close()
        End If

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub frmColumnTearAway_Paint(ByVal sender As Object, _
                                        ByVal e As System.Windows.Forms.PaintEventArgs) _
                                        Handles MyBase.Paint

        If _Painting Then
            Exit Sub
        End If

        _Painting = True

        If m_ListItems Is Nothing Then
            ' what display list
            Exit Sub
        End If

        If m_ListItems.Count = 0 Then
            ' the display list be empty 
            Exit Sub
        End If

        '' Ok we got something to display lets do this  

        Dim g As Graphics = e.Graphics

        Dim t As Integer = 0
        Dim y As Integer = 0
        Dim xmax As Integer = 0
        Dim li As Integer = 0
        Dim sz As SizeF = New SizeF(0, 0)
        Dim a As String

        For li = 0 To m_ListItems.Count - 1
            a = DirectCast(m_ListItems.Item(li), String)
            sz = g.MeasureString(a, m_DisplayFont)
            If xmax < sz.Width Then
                xmax = sz.Width
            End If
        Next

        sz = g.MeasureString(DirectCast(sender, frmColumnTearAway).Text, m_DisplayFont)
        If xmax < sz.Width Then
            xmax = sz.Width
        End If

        Me.Width = xmax + 12 + vscroller.Width

        If vscroller.Visible Then
            t = vscroller.Value
        Else
            t = 0
        End If

        If t > m_ListItems.Count - 1 Then

            t = m_ListItems.Count - 1
        End If

        For li = t To m_ListItems.Count - 1
            a = DirectCast(m_ListItems.Item(li), String)

            If a.Trim = "" Then
                sz = g.MeasureString("Wy", m_DisplayFont)
            Else
                sz = g.MeasureString(a, m_DisplayFont)
            End If

            If li = _SelectedRow Then

                g.FillRectangle(New SolidBrush(_DefaultSelectionColor), 2, y, xmax, sz.Height)

            Else
                g.FillRectangle(New SolidBrush(_DefaultBackColor), 2, y, xmax, sz.Height)

            End If

            g.DrawRectangle(New Pen(Color.Black), 2, y, xmax, sz.Height)

            If li = _SelectedRow Then
                g.DrawString(a, m_DisplayFont, New SolidBrush(_DefaultSelectionForeColor), 2, y)
            Else
                g.DrawString(a, m_DisplayFont, New SolidBrush(_DefaultForeColor), 2, y)
            End If

            y += sz.Height

            If y > Me.Height Then
                Exit For
            End If
        Next

        _Painting = False

    End Sub

    Private Sub vscroll_ValueChanged(ByVal sender As Object, _
                                     ByVal e As System.EventArgs) _
                                     Handles vscroller.ValueChanged

        If vscroller.Focus Then
            ' here we tell the rest of the world hat wwe are futzing about with the 
            ' scrollbar on a given tornaway window

            If Not m_gridParent Is Nothing Then

                m_gridParent.SetVertScrollbarPosition(vscroller.Value)

            End If

        End If

        Me.Invalidate()
    End Sub

    Private Sub frmColumnTearAway_MouseMove(ByVal sender As Object, _
                                            ByVal e As System.Windows.Forms.MouseEventArgs) _
                                            Handles MyBase.MouseMove

        If _Painting Then
            Exit Sub
        End If

        Me.BringToFront()

        Dim mp As Point = PointToClient(Windows.Forms.Control.MousePosition)

        Dim yoff, r, t, row As Integer
        Dim a As String = ""

        If vscroller.Visible Then
            r = vscroller.Value
        Else
            r = 0
        End If

        Dim g As Graphics = Me.CreateGraphics

        row = -1
        yoff = 0

        For t = r To m_ListItems.Count - 1
            a = DirectCast(m_ListItems.Item(t), String)
            Dim sz As SizeF = g.MeasureString(a, m_DisplayFont)
            If mp.Y >= yoff And mp.Y <= yoff + sz.Height Then
                row = r + t
                Exit For
            Else
                yoff += sz.Height
            End If
        Next

        If row <> -1 Then
            ' we have a row we be hovering over
            m_gridParent.RaiseGridHoverEvents(Me, row, m_colid, a)
        End If

    End Sub


    Private Sub frmColumnTearAway_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp

        If m_ListItems Is Nothing Then
            ' what display list
            m_RowCLicked = -1
            Exit Sub
        End If

        If m_ListItems.Count = 0 Then
            ' the display list be empty 
            m_RowCLicked = -1
            Exit Sub
        End If

        If e.Button = Windows.Forms.MouseButtons.Right Then
            ' bail on a right mousebutton
            m_RowCLicked = -1
            Exit Sub
        End If

        Dim x, y, yoff, r, t As Integer
        Dim a As String
        Dim sz As SizeF = New SizeF

        m_RowCLicked = -1

        x = e.X
        y = e.Y

        Console.WriteLine(x.ToString() + " - " + y.ToString())

        If vscroller.Visible Then
            yoff = vscroller.Value
        Else
            yoff = 0
        End If

        r = 0

        For t = yoff To m_ListItems.Count - 1
            a = DirectCast(m_ListItems.Item(t), String)

            If a.Trim = "" Then
                sz = Me.CreateGraphics.MeasureString("Wy", m_DisplayFont)
            Else
                sz = Me.CreateGraphics.MeasureString(a, m_DisplayFont)
            End If

            If y >= r And y <= r + sz.Height Then
                ' it be between so lets figure out the row

                m_gridParent.SelectedRows.Clear()
                m_gridParent.SelectedRow = t
                m_RowCLicked = t

                Exit For
            End If

            r += sz.Height

        Next

    End Sub

    Private Sub frmColumnTearAway_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize


    End Sub

    Private Sub frmColumnTearAway_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseWheel
        If m_ListItems.Count = 0 Then
            Exit Sub
        End If

        If _SelectedRow < 0 Or _SelectedRow > m_ListItems.Count - 1 Then
            Exit Sub
        End If

        If e.Delta > 0 Then
            _SelectedRow += -1
            If _SelectedRow < 0 Then
                _SelectedRow = 0
            End If
        Else
            _SelectedRow += 1
            If _SelectedRow = m_ListItems.Count Then
                _SelectedRow = m_ListItems.Count - 1
            End If
        End If

        m_gridParent.SelectedRow = _SelectedRow

        Me.Invalidate()

    End Sub

    Private Sub frmColumnTearAway_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp

        If m_ListItems.Count = 0 Then
            Exit Sub
        End If

        If _SelectedRow < 0 Or _SelectedRow > m_ListItems.Count - 1 Then
            Exit Sub
        End If


        If e.KeyCode = Windows.Forms.Keys.Up Then
            _SelectedRow += -1
            If _SelectedRow < 0 Then
                _SelectedRow = 0
            End If
        End If

        If e.KeyCode = Windows.Forms.Keys.Down Then
            _SelectedRow += 1
            If _SelectedRow = m_ListItems.Count Then
                _SelectedRow = m_ListItems.Count - 1
            End If
        End If

        m_gridParent.SelectedRow = _SelectedRow

        Me.Invalidate()

        e.Handled = True

    End Sub

    Private Sub frmColumnTearAway_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
        If m_gridParent Is Nothing Then
            Exit Sub
        End If

        m_gridParent.ResizeTearawayColumnsVertically(Me.Height)
    End Sub

    Private Sub frmColumnTearAway_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown

        Dim p As Point

        m_RowCLicked = -1

        If e.Button = Windows.Forms.MouseButtons.Right Then

            p = Me.PointToClient(Windows.Forms.Control.MousePosition)
            PulldownMenu.Show(Me, p)


        Else
            Dim x, y, yoff, r, t As Integer
            Dim a As String
            Dim sz As SizeF = New SizeF

            x = e.X
            y = e.Y

            If vscroller.Visible Then
                yoff = vscroller.Value
            Else
                yoff = 0
            End If

            r = 0

            For t = yoff To m_ListItems.Count - 1
                a = DirectCast(m_ListItems.Item(t), String)

                If a.Trim = "" Then
                    sz = Me.CreateGraphics.MeasureString("Wy", m_DisplayFont)
                Else
                    sz = Me.CreateGraphics.MeasureString(a, m_DisplayFont)
                End If

                If y >= r And y <= r + sz.Height Then
                    ' it be between so lets figure out the row

                    m_RowCLicked = t

                    Exit For
                End If

                r += sz.Height

            Next

        End If

    End Sub

    Private Sub miAutoArrange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miAutoArrange.Click
        m_gridParent.ArrangeTearAwayWindows()
    End Sub

    Private Sub miHideAllTearAways_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHideAllTearAways.Click
        m_gridParent.KillAllTearAwayColumnWindows()
    End Sub

    Private Sub miPushtoBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miPushtoBack.Click
        m_gridParent.PushAllTearAwaysToTheBack()
    End Sub

    Private Sub miPullAllTearAwaysToTheFront_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miPullAllTearAwaysToTheFront.Click
        m_gridParent.PullAllTearAwaysToTheFront()
    End Sub

    Private Sub frmColumnTearAway_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        If m_RowCLicked = -1 Then
            Exit Sub
        End If

        m_gridParent.RaiseCellClickedEvent(m_RowCLicked, m_colid)

    End Sub

    Private Sub frmColumnTearAway_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DoubleClick
        If m_RowCLicked = -1 Then
            Exit Sub
        End If

        m_gridParent.RaiseCellDoubleClickedEvent(m_RowCLicked, m_colid)

    End Sub

#End Region

#Region "Public Methods"

    Public Sub ShowToolTipOnForm(ByVal ttext As String)
        tt1.ShowAlways = True
        tt1.Active = True
        tt1.SetToolTip(Me, ttext)
    End Sub

    Public Sub HideToolTipOnForm()
        tt1.Active = False
    End Sub

    Public Function MaxRenderHeight() As Integer
        Dim res As Integer = 0
        Dim t As Integer
        Dim sz As SizeF
        Dim a As String

        If m_ListItems.Count = 0 Then
            ' we have no items in the list so lets return a 0
            Return res
            Exit Function ' not strictly necessary I believe
        End If

        Dim g As Graphics = Me.CreateGraphics()

        For t = 0 To m_ListItems.Count - 1
            a = DirectCast(m_ListItems.Item(t), String)

            If a.Trim = "" Then
                sz = g.MeasureString("Wy", m_DisplayFont)
            Else
                sz = g.MeasureString(a, m_DisplayFont)
            End If

            res += sz.Height
        Next

        Return res + System.Windows.Forms.SystemInformation.ToolWindowCaptionHeight
    End Function

#End Region

End Class
