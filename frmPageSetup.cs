using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

namespace TAIGridControl2
{
    public class frmPageSetup : Form
    {

        // Private _psets As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings
        // Private _PaperSize As System.Drawing.Printing.PaperSize = _psets.PaperSize

        private System.Drawing.Printing.PageSettings _psets;
        private System.Drawing.Printing.PaperSize _PaperSize;

        private int _MaxPage = 1;
        private int _MinPage = 1;

        private bool _Canceled = true;
        private bool _Print = false;
        private bool _Preview = false;
        private bool _PrintAllPages = true;
        private bool _PrintOrientationLandscape = false;

        public event PageSizeChangedEventHandler PageSizeChanged;

        public delegate void PageSizeChangedEventHandler(System.Drawing.Printing.PaperSize psiz);

        public event OrientationChangedEventHandler OrientationChanged;

        public delegate void OrientationChangedEventHandler(bool LandscapeOrientation);

        public event PaperMetricsHaveChangedEventHandler PaperMetricsHaveChanged;

        public delegate void PaperMetricsHaveChangedEventHandler(System.Drawing.Printing.PaperSize psiz, bool LandscapeOrientation);

        private void LogThis(string str)
        {
        }

        public frmPageSetup() : base()
        {
            InitializeComponent();
        }

        public frmPageSetup(System.Drawing.Printing.PageSettings pset) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            // MsgBox(pset.ToString())

            try
            {
                LogThis("In PageSetup Form");
                _psets = pset;

                LogThis("1");

                _PaperSize = _psets.PaperSize;

                LogThis("2");

                cmboPaperSize.Items.Clear();

                LogThis("3");

                foreach (System.Drawing.Printing.PaperSize psiz in _psets.PrinterSettings.PaperSizes)
                {
                    LogThis("Looping in some paper size Crapola");

                    cmboPaperSize.Items.Add(psiz);

                    LogThis(psiz.PaperName);
                }

                LogThis("Clearing Printer List");

                cmboPrinter.Items.Clear();

                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    LogThis("Adding a Printer " + printer.ToString());

                    cmboPrinter.Items.Add(printer);
                }

                LogThis("Printers Added Selecting It now");

                cmboPrinter.SelectedItem = _psets.PrinterSettings.PrinterName;

                if (_psets.PrinterSettings.PrintRange == (int)System.Drawing.Printing.PrintRange.AllPages)
                    rbAllPages.Checked = true;
                else
                    rbPageRange.Checked = true;

                if (_psets.Landscape)
                    rbLandscape.Checked = true;
                else
                    rbProtriate.Checked = true;

                txtStartPage.Text = _psets.PrinterSettings.FromPage.ToString();
                txtEndPage.Text = _psets.PrinterSettings.ToPage.ToString();
            }
            catch (Exception ex)
            {
            }
        }

        // Form overrides dispose to clean up the component list.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!(components == null))
                    components.Dispose();
            }
            base.Dispose(disposing);
        }

        // Required by the Windows Form Designer
        private IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        private Button _btnCancel;

        internal Button btnCancel
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnCancel;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnCancel != null)
                {
                    _btnCancel.Click -= btnCancel_Click;
                }

                _btnCancel = value;
                if (_btnCancel != null)
                {
                    _btnCancel.Click += btnCancel_Click;
                }
            }
        }

        private Button _btnOk;

        internal Button btnOk
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnOk;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnOk != null)
                {
                    _btnOk.Click -= btnOk_Click;
                }

                _btnOk = value;
                if (_btnOk != null)
                {
                    _btnOk.Click += btnOk_Click;
                }
            }
        }

        private ComboBox _cmboPaperSize;

        internal ComboBox cmboPaperSize
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _cmboPaperSize;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_cmboPaperSize != null)
                {
                    _cmboPaperSize.SelectedIndexChanged -= cmboPaperSize_SelectedIndexChanged;
                }

                _cmboPaperSize = value;
                if (_cmboPaperSize != null)
                {
                    _cmboPaperSize.SelectedIndexChanged += cmboPaperSize_SelectedIndexChanged;
                }
            }
        }

        private Label _Label1;

        internal Label Label1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Label1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Label1 != null)
                {
                }

                _Label1 = value;
                if (_Label1 != null)
                {
                }
            }
        }

        private GroupBox _GroupBox1;

        internal GroupBox GroupBox1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _GroupBox1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_GroupBox1 != null)
                {
                }

                _GroupBox1 = value;
                if (_GroupBox1 != null)
                {
                }
            }
        }

        private RadioButton _rbAllPages;

        internal RadioButton rbAllPages
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _rbAllPages;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_rbAllPages != null)
                {
                }

                _rbAllPages = value;
                if (_rbAllPages != null)
                {
                }
            }
        }

        private RadioButton _rbPageRange;

        internal RadioButton rbPageRange
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _rbPageRange;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_rbPageRange != null)
                {
                }

                _rbPageRange = value;
                if (_rbPageRange != null)
                {
                }
            }
        }

        private TextBox _txtStartPage;

        internal TextBox txtStartPage
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtStartPage;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtStartPage != null)
                {
                }

                _txtStartPage = value;
                if (_txtStartPage != null)
                {
                }
            }
        }

        private TextBox _txtEndPage;

        internal TextBox txtEndPage
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtEndPage;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtEndPage != null)
                {
                }

                _txtEndPage = value;
                if (_txtEndPage != null)
                {
                }
            }
        }

        private Label _Label2;

        internal Label Label2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Label2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Label2 != null)
                {
                }

                _Label2 = value;
                if (_Label2 != null)
                {
                }
            }
        }

        private GroupBox _GroupBox2;

        internal GroupBox GroupBox2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _GroupBox2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_GroupBox2 != null)
                {
                }

                _GroupBox2 = value;
                if (_GroupBox2 != null)
                {
                }
            }
        }

        private RadioButton _rbLandscape;

        internal RadioButton rbLandscape
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _rbLandscape;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_rbLandscape != null)
                {
                    _rbLandscape.CheckedChanged -= rbLandscape_CheckedChanged;
                }

                _rbLandscape = value;
                if (_rbLandscape != null)
                {
                    _rbLandscape.CheckedChanged += rbLandscape_CheckedChanged;
                }
            }
        }

        private RadioButton _rbProtriate;

        internal RadioButton rbProtriate
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _rbProtriate;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_rbProtriate != null)
                {
                    _rbProtriate.CheckedChanged -= rbProtriate_CheckedChanged;
                }

                _rbProtriate = value;
                if (_rbProtriate != null)
                {
                    _rbProtriate.CheckedChanged += rbProtriate_CheckedChanged;
                }
            }
        }

        private Label _Label3;

        internal Label Label3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Label3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Label3 != null)
                {
                }

                _Label3 = value;
                if (_Label3 != null)
                {
                }
            }
        }

        private ComboBox _cmboPrinter;

        internal ComboBox cmboPrinter
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _cmboPrinter;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_cmboPrinter != null)
                {
                    _cmboPrinter.SelectedIndexChanged -= cmboPrinter_SelectedIndexChanged;
                }

                _cmboPrinter = value;
                if (_cmboPrinter != null)
                {
                    _cmboPrinter.SelectedIndexChanged += cmboPrinter_SelectedIndexChanged;
                }
            }
        }

        private Button _btnPrint;

        internal Button btnPrint
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnPrint;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnPrint != null)
                {
                    _btnPrint.Click -= btnPrint_Click;
                }

                _btnPrint = value;
                if (_btnPrint != null)
                {
                    _btnPrint.Click += btnPrint_Click;
                }
            }
        }

        private Button _btnPreview;

        internal Button btnPreview
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnPreview;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnPreview != null)
                {
                    _btnPreview.Click -= btnPreview_Click;
                }

                _btnPreview = value;
                if (_btnPreview != null)
                {
                    _btnPreview.Click += btnPreview_Click;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            _btnCancel = new Button();
            _btnCancel.Click += btnCancel_Click;
            _btnOk = new Button();
            _btnOk.Click += btnOk_Click;
            _cmboPaperSize = new ComboBox();
            _cmboPaperSize.SelectedIndexChanged += cmboPaperSize_SelectedIndexChanged;
            _Label1 = new Label();
            _GroupBox1 = new GroupBox();
            _Label2 = new Label();
            _txtEndPage = new TextBox();
            _txtStartPage = new TextBox();
            _rbPageRange = new RadioButton();
            _rbAllPages = new RadioButton();
            _GroupBox2 = new GroupBox();
            _rbLandscape = new RadioButton();
            _rbLandscape.CheckedChanged += rbLandscape_CheckedChanged;
            _rbProtriate = new RadioButton();
            _rbProtriate.CheckedChanged += rbProtriate_CheckedChanged;
            _Label3 = new Label();
            _cmboPrinter = new ComboBox();
            _cmboPrinter.SelectedIndexChanged += cmboPrinter_SelectedIndexChanged;
            _btnPrint = new Button();
            _btnPrint.Click += btnPrint_Click;
            _btnPreview = new Button();
            _btnPreview.Click += btnPreview_Click;
            _GroupBox1.SuspendLayout();
            _GroupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // btnCancel
            // 
            _btnCancel.Location = new Point(360, 208);
            _btnCancel.Name = "btnCancel";
            _btnCancel.Size = new Size(60, 20);
            _btnCancel.TabIndex = 0;
            _btnCancel.Text = "Cancel";
            // 
            // btnOk
            // 
            _btnOk.Location = new Point(296, 208);
            _btnOk.Name = "btnOk";
            _btnOk.Size = new Size(60, 20);
            _btnOk.TabIndex = 1;
            _btnOk.Text = "Accept";
            // 
            // cmboPaperSize
            // 
            _cmboPaperSize.Location = new Point(4, 48);
            _cmboPaperSize.Name = "cmboPaperSize";
            _cmboPaperSize.Size = new Size(420, 21);
            _cmboPaperSize.TabIndex = 2;
            // 
            // Label1
            // 
            _Label1.Location = new Point(4, 72);
            _Label1.Name = "Label1";
            _Label1.Size = new Size(120, 16);
            _Label1.TabIndex = 3;
            _Label1.Text = "Select Paper Size";
            // 
            // GroupBox1
            // 
            _GroupBox1.Controls.Add(_Label2);
            _GroupBox1.Controls.Add(_txtEndPage);
            _GroupBox1.Controls.Add(_txtStartPage);
            _GroupBox1.Controls.Add(_rbPageRange);
            _GroupBox1.Controls.Add(_rbAllPages);
            _GroupBox1.Location = new Point(8, 96);
            _GroupBox1.Name = "GroupBox1";
            _GroupBox1.Size = new Size(160, 108);
            _GroupBox1.TabIndex = 4;
            _GroupBox1.TabStop = false;
            _GroupBox1.Text = "Print What?";
            // 
            // Label2
            // 
            _Label2.Location = new Point(72, 80);
            _Label2.Name = "Label2";
            _Label2.Size = new Size(20, 16);
            _Label2.TabIndex = 4;
            _Label2.Text = "to";
            _Label2.TextAlign = ContentAlignment.TopCenter;
            // 
            // txtEndPage
            // 
            _txtEndPage.Location = new Point(100, 76);
            _txtEndPage.Name = "txtEndPage";
            _txtEndPage.Size = new Size(44, 20);
            _txtEndPage.TabIndex = 3;
            _txtEndPage.Text = "";
            // 
            // txtStartPage
            // 
            _txtStartPage.Location = new Point(20, 76);
            _txtStartPage.Name = "txtStartPage";
            _txtStartPage.Size = new Size(44, 20);
            _txtStartPage.TabIndex = 2;
            _txtStartPage.Text = "";
            // 
            // rbPageRange
            // 
            _rbPageRange.Location = new Point(20, 52);
            _rbPageRange.Name = "rbPageRange";
            _rbPageRange.Size = new Size(124, 20);
            _rbPageRange.TabIndex = 1;
            _rbPageRange.Text = "A Range Of Pages";
            // 
            // rbAllPages
            // 
            _rbAllPages.Checked = true;
            _rbAllPages.Location = new Point(20, 28);
            _rbAllPages.Name = "rbAllPages";
            _rbAllPages.Size = new Size(124, 20);
            _rbAllPages.TabIndex = 0;
            _rbAllPages.TabStop = true;
            _rbAllPages.Text = "All Pages";
            // 
            // GroupBox2
            // 
            _GroupBox2.Controls.Add(_rbLandscape);
            _GroupBox2.Controls.Add(_rbProtriate);
            _GroupBox2.Location = new Point(260, 96);
            _GroupBox2.Name = "GroupBox2";
            _GroupBox2.Size = new Size(160, 108);
            _GroupBox2.TabIndex = 5;
            _GroupBox2.TabStop = false;
            _GroupBox2.Text = "Orientation";
            // 
            // rbLandscape
            // 
            _rbLandscape.Location = new Point(20, 52);
            _rbLandscape.Name = "rbLandscape";
            _rbLandscape.Size = new Size(124, 20);
            _rbLandscape.TabIndex = 1;
            _rbLandscape.Text = "Landscape";
            // 
            // rbProtriate
            // 
            _rbProtriate.Checked = true;
            _rbProtriate.Location = new Point(20, 28);
            _rbProtriate.Name = "rbProtriate";
            _rbProtriate.Size = new Size(124, 20);
            _rbProtriate.TabIndex = 0;
            _rbProtriate.TabStop = true;
            _rbProtriate.Text = "Portrait";
            // 
            // Label3
            // 
            _Label3.Location = new Point(4, 28);
            _Label3.Name = "Label3";
            _Label3.Size = new Size(152, 16);
            _Label3.TabIndex = 7;
            _Label3.Text = "Select Printer to Print to";
            // 
            // cmboPrinter
            // 
            _cmboPrinter.Location = new Point(4, 4);
            _cmboPrinter.Name = "cmboPrinter";
            _cmboPrinter.Size = new Size(420, 21);
            _cmboPrinter.TabIndex = 6;
            // 
            // btnPrint
            // 
            _btnPrint.Location = new Point(232, 208);
            _btnPrint.Name = "btnPrint";
            _btnPrint.Size = new Size(60, 20);
            _btnPrint.TabIndex = 8;
            _btnPrint.Text = "Print";
            // 
            // btnPreview
            // 
            _btnPreview.Location = new Point(168, 208);
            _btnPreview.Name = "btnPreview";
            _btnPreview.Size = new Size(60, 20);
            _btnPreview.TabIndex = 9;
            _btnPreview.Text = "Preview";
            // 
            // frmPageSetup
            // 
            AutoScaleBaseSize = new Size(5, 13);
            ClientSize = new Size(430, 231);
            ControlBox = false;
            Controls.Add(_btnPreview);
            Controls.Add(_btnPrint);
            Controls.Add(_Label3);
            Controls.Add(_cmboPrinter);
            Controls.Add(_GroupBox2);
            Controls.Add(_GroupBox1);
            Controls.Add(_Label1);
            Controls.Add(_cmboPaperSize);
            Controls.Add(_btnOk);
            Controls.Add(_btnCancel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Name = "frmPageSetup";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Setup Printed Page Metrics";
            base.Activated += HandleFormShow;
            _GroupBox1.ResumeLayout(false);
            base.Activated += HandleFormShow;
            _GroupBox2.ResumeLayout(false);
            base.Activated += HandleFormShow;
            ResumeLayout(false);
        }


        private void HandleFormShow(object sender, EventArgs ea)
        {
        }

        public System.Drawing.Printing.PageSettings Psets
        {
            get
            {
                return _psets;
            }
            set
            {
                _psets = value;
            }
        }

        public System.Drawing.Printing.PaperSize PaperSize
        {
            get
            {
                return _PaperSize;
            }
            set
            {
                _PaperSize = value;
                _psets.PaperSize = value;
                cmboPaperSize.SelectedItem = value;
            }
        }

        public int MaxPage
        {
            get
            {
                return _MaxPage;
            }
            set
            {
                _MaxPage = value;
                txtEndPage.Text = _MaxPage.ToString();
                txtStartPage.Text = _MinPage.ToString();
            }
        }

        public int MinPage
        {
            get
            {
                return _MinPage;
            }
            set
            {
                _MinPage = value;
                txtEndPage.Text = _MaxPage.ToString();
                txtStartPage.Text = _MinPage.ToString();
            }
        }

        public bool Canceled
        {
            get
            {
                return _Canceled;
            }
            set
            {
                _Canceled = value;
            }
        }

        public bool Print
        {
            get
            {
                return _Print;
            }
            set
            {
                _Print = value;
            }
        }

        public bool Preview
        {
            get
            {
                return _Preview;
            }
            set
            {
                _Preview = value;
            }
        }

        public bool PrintAllPages
        {
            get
            {
                return _PrintAllPages;
            }
            set
            {
                _PrintAllPages = value;
                rbAllPages.Checked = value;
            }
        }

        public bool PrintOrientationLandscape
        {
            get
            {
                return _PrintOrientationLandscape;
            }
            set
            {
                _PrintOrientationLandscape = value;
                rbLandscape.Checked = value;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            _Canceled = false;

            if (rbAllPages.Checked)
            {
                _psets.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.AllPages;

                _psets.PrinterSettings.ToPage = 0;
                _psets.PrinterSettings.FromPage = 0;
            }
            else
            {
                _psets.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;

                if (Information.IsNumeric(txtStartPage.Text))
                    _MinPage = Conversions.ToInteger(Conversion.Val(txtStartPage.Text));
                else
                    _MinPage = 1;

                if (Information.IsNumeric(txtEndPage.Text))
                    _MaxPage = Conversions.ToInteger(Conversion.Val(txtEndPage.Text));
                else
                    _MaxPage = _MinPage;

                _psets.PrinterSettings.ToPage = _MaxPage;
                _psets.PrinterSettings.FromPage = _MinPage;
            }


            Hide();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _Canceled = true;

            Hide();
        }

        private void cmboPaperSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            Refresh();

            Application.DoEvents();

            try
            {
                _PaperSize = (System.Drawing.Printing.PaperSize)cmboPaperSize.SelectedItem;
                _psets.PaperSize = (System.Drawing.Printing.PaperSize)cmboPaperSize.SelectedItem;
            }
            catch (Exception ex)
            {
            }

            LogThis("Calling event handler with a papersize of " + _PaperSize.ToString());

            PaperMetricsHaveChanged?.Invoke(_PaperSize, _PrintOrientationLandscape);

            Application.DoEvents();
        }

        private void DecodeOrientation()
        {
            Refresh();

            Application.DoEvents();

            try
            {
                if (rbProtriate.Checked)
                {
                    _PrintOrientationLandscape = false;
                    if (cmboPaperSize.SelectedItem == null)
                        PaperMetricsHaveChanged?.Invoke(_psets.PaperSize, _PrintOrientationLandscape);
                    else
                        PaperMetricsHaveChanged?.Invoke(_PaperSize, _PrintOrientationLandscape);
                }
                else
                {
                    _PrintOrientationLandscape = true;
                    if (cmboPaperSize.SelectedItem == null)
                        PaperMetricsHaveChanged?.Invoke(_psets.PaperSize, _PrintOrientationLandscape);
                    else
                        PaperMetricsHaveChanged?.Invoke(_PaperSize, _PrintOrientationLandscape);
                }
            }
            catch (Exception ex)
            {
            }



            Application.DoEvents();
        }

        private void rbProtriate_CheckedChanged(object sender, EventArgs e)
        {
            DecodeOrientation();
        }

        private void rbLandscape_CheckedChanged(object sender, EventArgs e)
        {
            DecodeOrientation();
        }

        private void DecodePageRanges()
        {
            if (Information.IsNumeric(txtStartPage.Text))
                _MinPage = Conversions.ToInteger(Conversion.Val(txtStartPage.Text));
            else
                _MinPage = 1;

            if (Information.IsNumeric(txtEndPage.Text))
                _MaxPage = Conversions.ToInteger(Conversion.Val(txtEndPage.Text));
            else
                _MaxPage = _MinPage;
        }

        private void cmboPrinter_SelectedIndexChanged(object sender, EventArgs e)
        {
            _psets.PrinterSettings.PrinterName = cmboPrinter.Text;

            cmboPaperSize.Items.Clear();

            foreach (System.Drawing.Printing.PaperSize psiz in _psets.PrinterSettings.PaperSizes)
            {
                try
                {
                    cmboPaperSize.Items.Add(psiz);
                }

                // If psiz.PaperName.Split(" ".ToCharArray)(0) = _
                // _psets.PaperSize.PaperName.Split(" ".ToCharArray)(0) Then
                // cmboPaperSize.SelectedIndex = cmboPaperSize.Items.Count - 1
                // End If

                catch (Exception ex)
                {
                }
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            _Canceled = false;
            _Print = true;
            _Preview = false;

            if (rbAllPages.Checked)
            {
                _psets.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.AllPages;

                _psets.PrinterSettings.ToPage = 0;
                _psets.PrinterSettings.FromPage = 0;
                _psets.Landscape = rbLandscape.Checked;
            }
            else
            {
                _psets.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;

                if (Information.IsNumeric(txtStartPage.Text))
                    _MinPage = Conversions.ToInteger(Conversion.Val(txtStartPage.Text));
                else
                    _MinPage = 1;

                if (Information.IsNumeric(txtEndPage.Text))
                    _MaxPage = Conversions.ToInteger(Conversion.Val(txtEndPage.Text));
                else
                    _MaxPage = _MinPage;

                _psets.PrinterSettings.ToPage = _MaxPage;
                _psets.PrinterSettings.FromPage = _MinPage;
                _psets.Landscape = rbLandscape.Checked;
            }

            Hide();
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            _Canceled = false;
            _Print = false;
            _Preview = true;

            if (rbAllPages.Checked)
            {
                _psets.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.AllPages;

                _psets.PrinterSettings.ToPage = 0;
                _psets.PrinterSettings.FromPage = 0;
                _psets.Landscape = rbLandscape.Checked;
            }
            else
            {
                _psets.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;

                if (Information.IsNumeric(txtStartPage.Text))
                    _MinPage = Conversions.ToInteger(Conversion.Val(txtStartPage.Text));
                else
                    _MinPage = 1;

                if (Information.IsNumeric(txtEndPage.Text))
                    _MaxPage = Conversions.ToInteger(Conversion.Val(txtEndPage.Text));
                else
                    _MaxPage = _MinPage;

                _psets.PrinterSettings.ToPage = _MaxPage;
                _psets.PrinterSettings.FromPage = _MinPage;
                _psets.Landscape = rbLandscape.Checked;
            }

            Hide();
        }
    }
}
