using System.Diagnostics;
using System.Collections;
using System;
using System.Drawing;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

namespace TAIGridControl2
{
    public class frmColumnTearAway : System.Windows.Forms.Form
    {
        public frmColumnTearAway() : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();
        }

        public frmColumnTearAway(string title) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            Text = title;
        }

        // Form overrides dispose to clean up the component list.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!(components == null))
                    components.Dispose();
                if (!(m_gridParent == null))
                    // callback into the parent grid to kill it from the display list
                    m_gridParent.KillTearAwayColumnWindow(m_colid);
            }

            base.Dispose(disposing);
        }

        // Required by the Windows Form Designer
        private System.ComponentModel.IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        private System.Windows.Forms.VScrollBar _vscroller;

        internal System.Windows.Forms.VScrollBar vscroller
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _vscroller;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_vscroller != null)
                {
                    _vscroller.ValueChanged -= vscroll_ValueChanged;
                }

                _vscroller = value;
                if (_vscroller != null)
                {
                    _vscroller.ValueChanged += vscroll_ValueChanged;
                }
            }
        }

        private System.Windows.Forms.ToolTip _tt1;

        internal System.Windows.Forms.ToolTip tt1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _tt1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_tt1 != null)
                {
                }

                _tt1 = value;
                if (_tt1 != null)
                {
                }
            }
        }

        private System.Windows.Forms.ContextMenu _PulldownMenu;

        internal System.Windows.Forms.ContextMenu PulldownMenu
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _PulldownMenu;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_PulldownMenu != null)
                {
                }

                _PulldownMenu = value;
                if (_PulldownMenu != null)
                {
                }
            }
        }

        private System.Windows.Forms.MenuItem _miAutoArrange;

        internal System.Windows.Forms.MenuItem miAutoArrange
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miAutoArrange;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miAutoArrange != null)
                {
                    _miAutoArrange.Click -= miAutoArrange_Click;
                }

                _miAutoArrange = value;
                if (_miAutoArrange != null)
                {
                    _miAutoArrange.Click += miAutoArrange_Click;
                }
            }
        }

        private System.Windows.Forms.MenuItem _miPushtoBack;

        internal System.Windows.Forms.MenuItem miPushtoBack
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miPushtoBack;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miPushtoBack != null)
                {
                    _miPushtoBack.Click -= miPushtoBack_Click;
                }

                _miPushtoBack = value;
                if (_miPushtoBack != null)
                {
                    _miPushtoBack.Click += miPushtoBack_Click;
                }
            }
        }

        private System.Windows.Forms.MenuItem _miPullAllTearAwaysToTheFront;

        internal System.Windows.Forms.MenuItem miPullAllTearAwaysToTheFront
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miPullAllTearAwaysToTheFront;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miPullAllTearAwaysToTheFront != null)
                {
                    _miPullAllTearAwaysToTheFront.Click -= miPullAllTearAwaysToTheFront_Click;
                }

                _miPullAllTearAwaysToTheFront = value;
                if (_miPullAllTearAwaysToTheFront != null)
                {
                    _miPullAllTearAwaysToTheFront.Click += miPullAllTearAwaysToTheFront_Click;
                }
            }
        }

        private System.Windows.Forms.MenuItem _miHideAllTearAways;

        internal System.Windows.Forms.MenuItem miHideAllTearAways
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHideAllTearAways;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHideAllTearAways != null)
                {
                    _miHideAllTearAways.Click -= miHideAllTearAways_Click;
                }

                _miHideAllTearAways = value;
                if (_miHideAllTearAways != null)
                {
                    _miHideAllTearAways.Click += miHideAllTearAways_Click;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            _vscroller = new System.Windows.Forms.VScrollBar();
            _vscroller.ValueChanged += vscroll_ValueChanged;
            _tt1 = new System.Windows.Forms.ToolTip(components);
            _PulldownMenu = new System.Windows.Forms.ContextMenu();
            _miAutoArrange = new System.Windows.Forms.MenuItem();
            _miAutoArrange.Click += miAutoArrange_Click;
            _miPushtoBack = new System.Windows.Forms.MenuItem();
            _miPushtoBack.Click += miPushtoBack_Click;
            _miPullAllTearAwaysToTheFront = new System.Windows.Forms.MenuItem();
            _miPullAllTearAwaysToTheFront.Click += miPullAllTearAwaysToTheFront_Click;
            _miHideAllTearAways = new System.Windows.Forms.MenuItem();
            _miHideAllTearAways.Click += miHideAllTearAways_Click;
            SuspendLayout();
            // 
            // vscroller
            // 
            _vscroller.Dock = System.Windows.Forms.DockStyle.Right;
            _vscroller.Location = new Point(256, 0);
            _vscroller.Name = "vscroller";
            _vscroller.Size = new Size(12, 340);
            _vscroller.TabIndex = 0;
            // 
            // tt1
            // 
            _tt1.ShowAlways = true;
            // 
            // PulldownMenu
            // 
            _PulldownMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { _miAutoArrange, _miHideAllTearAways, _miPushtoBack, _miPullAllTearAwaysToTheFront });
            // 
            // miAutoArrange
            // 
            _miAutoArrange.Index = 0;
            _miAutoArrange.Text = "Auto Arrange All Tear Away Columns";
            // 
            // miPushtoBack
            // 
            _miPushtoBack.Index = 2;
            _miPushtoBack.Text = "Push All Tear Away's to the Back";
            // 
            // miPullAllTearAwaysToTheFront
            // 
            _miPullAllTearAwaysToTheFront.Index = 3;
            _miPullAllTearAwaysToTheFront.Text = "Pull All Tear Away's to the Front";
            // 
            // miHideAllTearAways
            // 
            _miHideAllTearAways.Index = 1;
            _miHideAllTearAways.Text = "Hide All Tear Away Columns";
            // 
            // frmColumnTearAway
            // 
            AutoScaleBaseSize = new Size(5, 13);
            ClientSize = new Size(268, 340);
            Controls.Add(_vscroller);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "frmColumnTearAway";
            ShowInTaskbar = false;
            Text = "frmColumnTearAway";
            TopMost = true;
            base.Paint += frmColumnTearAway_Paint;
            base.MouseMove += frmColumnTearAway_MouseMove;
            base.MouseUp += frmColumnTearAway_MouseUp;
            base.Resize += frmColumnTearAway_Resize;
            base.MouseWheel += frmColumnTearAway_MouseWheel;
            base.KeyUp += frmColumnTearAway_KeyUp;
            base.SizeChanged += frmColumnTearAway_SizeChanged;
            base.MouseDown += frmColumnTearAway_MouseDown;
            base.Click += frmColumnTearAway_Click;
            base.DoubleClick += frmColumnTearAway_DoubleClick;
            ResumeLayout(false);
        }



        private bool _Painting = false;
        private int m_VertScrollIndex;
        private int m_VertScrollMin;
        private int m_VertScrollMax;
        private bool m_VertScrollVisible;
        private int m_colid;
        private int m_RowCLicked = -1;

        private ArrayList m_ListItems = new ArrayList();

        private int _SelectedRow = -1;

        private TAIGridControl m_gridParent;

        private Font m_DisplayFont = new Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Point);

        private Color _DefaultBackColor = Color.AntiqueWhite;
        private Color _DefaultSelectionColor = Color.Blue;
        private Color _DefaultSelectionForeColor = Color.White;
        private Color _DefaultForeColor = Color.Black;



        public int Colid
        {
            get
            {
                return m_colid;
            }
            set
            {
                m_colid = value;
            }
        }

        public int SelectedRow
        {
            get
            {
                return _SelectedRow;
            }
            set
            {
                if (_SelectedRow != value)
                {
                    _SelectedRow = value;
                    Invalidate();
                }
            }
        }

        public Color DefaultSelectionColor
        {
            get
            {
                return _DefaultSelectionColor;
            }
            set
            {
                _DefaultSelectionColor = value;
                Invalidate();
            }
        }

        public Color DefaultSelectionForColor
        {
            get
            {
                return _DefaultSelectionForeColor;
            }
            set
            {
                _DefaultSelectionForeColor = value;
                Invalidate();
            }
        }

        public Color GridDefaultBackColor
        {
            get
            {
                return _DefaultBackColor;
            }
            set
            {
                _DefaultBackColor = value;
                Invalidate();
            }
        }

        public Color GridDefaultForeColor
        {
            get
            {
                return _DefaultForeColor;
            }
            set
            {
                _DefaultForeColor = value;
                Invalidate();
            }
        }

        public Font DisplayFont
        {
            get
            {
                return m_DisplayFont;
            }
            set
            {
                m_DisplayFont = value;
                Invalidate();
            }
        }

        public ArrayList ListItems
        {
            get
            {
                return m_ListItems;
            }
            set
            {
                m_ListItems = value;
                if (Visible)
                    Invalidate();
            }
        }

        public TAIGridControl GridParent
        {
            get
            {
                return m_gridParent;
            }
            set
            {
                m_gridParent = value;
            }
        }

        public int VertScrollIndex
        {
            get
            {
                return m_VertScrollIndex;
            }
            set
            {
                m_VertScrollIndex = value;
                vscroller.Value = value;
            }
        }

        public int VertScrollMin
        {
            get
            {
                return m_VertScrollMin;
            }
            set
            {
                m_VertScrollMin = value;
                vscroller.Minimum = value;
            }
        }

        public int VertScrollMax
        {
            get
            {
                return m_VertScrollMax;
            }
            set
            {
                m_VertScrollMax = value;
                vscroller.Maximum = value;
            }
        }

        public bool VertScrollVisible
        {
            get
            {
                return m_VertScrollVisible;
            }
            set
            {
                m_VertScrollVisible = value;
                vscroller.Visible = value;
            }
        }

        public void KillMe(int colid)
        {
            if (m_colid == colid)
                Close();
        }



        private void frmColumnTearAway_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            if (_Painting)
                return;

            _Painting = true;

            if (m_ListItems == null)
                // what display list
                return;

            if (m_ListItems.Count == 0)
                // the display list be empty 
                return;

            // ' Ok we got something to display lets do this  

            var g = e.Graphics;

            int t = 0;
            int y = 0;
            int xmax = 0;
            int li = 0;
            var sz = new SizeF(0, 0);
            string a;
            var loopTo = m_ListItems.Count - 1;
            for (li = 0; li <= loopTo; li++)
            {
                a = (string)m_ListItems[li];
                sz = g.MeasureString(a, m_DisplayFont);
                if (xmax < sz.Width)
                    xmax = Conversions.ToInteger(sz.Width);
            }

            sz = g.MeasureString(((frmColumnTearAway)sender).Text, m_DisplayFont);
            if (xmax < sz.Width)
                xmax = Conversions.ToInteger(sz.Width);

            Width = xmax + 12 + vscroller.Width;

            if (vscroller.Visible)
                t = vscroller.Value;
            else
                t = 0;

            if (t > m_ListItems.Count - 1)
                t = m_ListItems.Count - 1;
            var loopTo1 = m_ListItems.Count - 1;
            for (li = t; li <= loopTo1; li++)
            {
                a = (string)m_ListItems[li];

                if (string.IsNullOrEmpty(a.Trim()))
                    sz = g.MeasureString("Wy", m_DisplayFont);
                else
                    sz = g.MeasureString(a, m_DisplayFont);

                if (li == _SelectedRow)
                    g.FillRectangle(new SolidBrush(_DefaultSelectionColor), 2, y, xmax, sz.Height);
                else
                    g.FillRectangle(new SolidBrush(_DefaultBackColor), 2, y, xmax, sz.Height);

                g.DrawRectangle(new Pen(Color.Black), 2, y, xmax, sz.Height);

                if (li == _SelectedRow)
                    g.DrawString(a, m_DisplayFont, new SolidBrush(_DefaultSelectionForeColor), 2, y);
                else
                    g.DrawString(a, m_DisplayFont, new SolidBrush(_DefaultForeColor), 2, y);

                y += (int)sz.Height;

                if (y > Height)
                    break;
            }

            _Painting = false;
        }

        private void vscroll_ValueChanged(object sender, EventArgs e)
        {
            if (vscroller.Focus())
            {
                // here we tell the rest of the world hat wwe are futzing about with the 
                // scrollbar on a given tornaway window

                if (!(m_gridParent == null))
                    m_gridParent.SetVertScrollbarPosition(vscroller.Value);
            }

            Invalidate();
        }

        private void frmColumnTearAway_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (_Painting)
                return;

            BringToFront();

            var mp = PointToClient(MousePosition);

            int yoff, r, t, row;
            string a = "";

            if (vscroller.Visible)
                r = vscroller.Value;
            else
                r = 0;

            var g = CreateGraphics();

            row = -1;
            yoff = 0;
            var loopTo = m_ListItems.Count - 1;
            for (t = r; t <= loopTo; t++)
            {
                a = (string)m_ListItems[t];
                var sz = g.MeasureString(a, m_DisplayFont);
                if (mp.Y >= yoff & mp.Y <= yoff + sz.Height)
                {
                    row = r + t;
                    break;
                }
                else
                    yoff += (int)sz.Height;
            }

            if (row != -1)
                // we have a row we be hovering over
                m_gridParent.RaiseGridHoverEvents(this, row, m_colid, a);
        }


        private void frmColumnTearAway_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (m_ListItems == null)
            {
                // what display list
                m_RowCLicked = -1;
                return;
            }

            if (m_ListItems.Count == 0)
            {
                // the display list be empty 
                m_RowCLicked = -1;
                return;
            }

            if ((int)e.Button == (int)System.Windows.Forms.MouseButtons.Right)
            {
                // bail on a right mousebutton
                m_RowCLicked = -1;
                return;
            }

            int x, y, yoff, r, t;
            string a;
            var sz = new SizeF();

            m_RowCLicked = -1;

            x = e.X;
            y = e.Y;

            Console.WriteLine(x.ToString() + " - " + y.ToString());

            if (vscroller.Visible)
                yoff = vscroller.Value;
            else
                yoff = 0;

            r = 0;
            var loopTo = m_ListItems.Count - 1;
            for (t = yoff; t <= loopTo; t++)
            {
                a = (string)m_ListItems[t];

                if (string.IsNullOrEmpty(a.Trim()))
                    sz = CreateGraphics().MeasureString("Wy", m_DisplayFont);
                else
                    sz = CreateGraphics().MeasureString(a, m_DisplayFont);

                if (y >= r & y <= r + sz.Height)
                {
                    // it be between so lets figure out the row

                    m_gridParent.SelectedRows.Clear();
                    m_gridParent.SelectedRow = t;
                    m_RowCLicked = t;

                    break;
                }

                r += (int)sz.Height;
            }
        }

        private void frmColumnTearAway_Resize(object sender, EventArgs e)
        {
        }

        private void frmColumnTearAway_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (m_ListItems.Count == 0)
                return;

            if (_SelectedRow < 0 | _SelectedRow > m_ListItems.Count - 1)
                return;

            if (e.Delta > 0)
            {
                _SelectedRow += -1;
                if (_SelectedRow < 0)
                    _SelectedRow = 0;
            }
            else
            {
                _SelectedRow += 1;
                if (_SelectedRow == m_ListItems.Count)
                    _SelectedRow = m_ListItems.Count - 1;
            }

            m_gridParent.SelectedRow = _SelectedRow;

            Invalidate();
        }

        private void frmColumnTearAway_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (m_ListItems.Count == 0)
                return;

            if (_SelectedRow < 0 | _SelectedRow > m_ListItems.Count - 1)
                return;


            if ((int)e.KeyCode == (int)System.Windows.Forms.Keys.Up)
            {
                _SelectedRow += -1;
                if (_SelectedRow < 0)
                    _SelectedRow = 0;
            }

            if ((int)e.KeyCode == (int)System.Windows.Forms.Keys.Down)
            {
                _SelectedRow += 1;
                if (_SelectedRow == m_ListItems.Count)
                    _SelectedRow = m_ListItems.Count - 1;
            }

            m_gridParent.SelectedRow = _SelectedRow;

            Invalidate();

            e.Handled = true;
        }

        private void frmColumnTearAway_SizeChanged(object sender, EventArgs e)
        {
            if (m_gridParent == null)
                return;

            m_gridParent.ResizeTearawayColumnsVertically(Height);
        }

        private void frmColumnTearAway_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            Point p;

            m_RowCLicked = -1;

            if ((int)e.Button == (int)System.Windows.Forms.MouseButtons.Right)
            {
                p = PointToClient(MousePosition);
                PulldownMenu.Show(this, p);
            }
            else
            {
                int x, y, yoff, r, t;
                string a;
                var sz = new SizeF();

                x = e.X;
                y = e.Y;

                if (vscroller.Visible)
                    yoff = vscroller.Value;
                else
                    yoff = 0;

                r = 0;
                var loopTo = m_ListItems.Count - 1;
                for (t = yoff; t <= loopTo; t++)
                {
                    a = (string)m_ListItems[t];

                    if (string.IsNullOrEmpty(a.Trim()))
                        sz = CreateGraphics().MeasureString("Wy", m_DisplayFont);
                    else
                        sz = CreateGraphics().MeasureString(a, m_DisplayFont);

                    if (y >= r & y <= r + sz.Height)
                    {
                        // it be between so lets figure out the row

                        m_RowCLicked = t;

                        break;
                    }

                    r += (int)sz.Height;
                }
            }
        }

        private void miAutoArrange_Click(object sender, EventArgs e)
        {
            m_gridParent.ArrangeTearAwayWindows();
        }

        private void miHideAllTearAways_Click(object sender, EventArgs e)
        {
            m_gridParent.KillAllTearAwayColumnWindows();
        }

        private void miPushtoBack_Click(object sender, EventArgs e)
        {
            m_gridParent.PushAllTearAwaysToTheBack();
        }

        private void miPullAllTearAwaysToTheFront_Click(object sender, EventArgs e)
        {
            m_gridParent.PullAllTearAwaysToTheFront();
        }

        private void frmColumnTearAway_Click(object sender, EventArgs e)
        {
            if (m_RowCLicked == -1)
                return;

            m_gridParent.RaiseCellClickedEvent(m_RowCLicked, m_colid);
        }

        private void frmColumnTearAway_DoubleClick(object sender, EventArgs e)
        {
            if (m_RowCLicked == -1)
                return;

            m_gridParent.RaiseCellDoubleClickedEvent(m_RowCLicked, m_colid);
        }



        public void ShowToolTipOnForm(string ttext)
        {
            tt1.ShowAlways = true;
            tt1.Active = true;
            tt1.SetToolTip(this, ttext);
        }

        public void HideToolTipOnForm()
        {
            tt1.Active = false;
        }

        public int MaxRenderHeight()
        {
            int res = 0;
            int t;
            SizeF sz;
            string a;

            if (m_ListItems.Count == 0)
            {
                // we have no items in the list so lets return a 0
                return res;
                return default(int); // not strictly necessary I believe
            }

            var g = CreateGraphics();
            var loopTo = m_ListItems.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                a = (string)m_ListItems[t];

                if (string.IsNullOrEmpty(a.Trim()))
                    sz = g.MeasureString("Wy", m_DisplayFont);
                else
                    sz = g.MeasureString(a, m_DisplayFont);

                res += (int)sz.Height;
            }

            return res + System.Windows.Forms.SystemInformation.ToolWindowCaptionHeight;
        }
    }
}
