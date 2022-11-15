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
            this._btnCancel = new System.Windows.Forms.Button();
            this._btnOk = new System.Windows.Forms.Button();
            this._cmboPaperSize = new System.Windows.Forms.ComboBox();
            this._Label1 = new System.Windows.Forms.Label();
            this._GroupBox1 = new System.Windows.Forms.GroupBox();
            this._Label2 = new System.Windows.Forms.Label();
            this._txtEndPage = new System.Windows.Forms.TextBox();
            this._txtStartPage = new System.Windows.Forms.TextBox();
            this._rbPageRange = new System.Windows.Forms.RadioButton();
            this._rbAllPages = new System.Windows.Forms.RadioButton();
            this._GroupBox2 = new System.Windows.Forms.GroupBox();
            this._rbLandscape = new System.Windows.Forms.RadioButton();
            this._rbProtriate = new System.Windows.Forms.RadioButton();
            this._Label3 = new System.Windows.Forms.Label();
            this._cmboPrinter = new System.Windows.Forms.ComboBox();
            this._btnPrint = new System.Windows.Forms.Button();
            this._btnPreview = new System.Windows.Forms.Button();
            this._GroupBox1.SuspendLayout();
            this._GroupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // _btnCancel
            // 
            this._btnCancel.Location = new System.Drawing.Point(360, 208);
            this._btnCancel.Name = "_btnCancel";
            this._btnCancel.Size = new System.Drawing.Size(60, 20);
            this._btnCancel.TabIndex = 0;
            this._btnCancel.Text = "Cancel";
            this._btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // _btnOk
            // 
            this._btnOk.Location = new System.Drawing.Point(296, 208);
            this._btnOk.Name = "_btnOk";
            this._btnOk.Size = new System.Drawing.Size(60, 20);
            this._btnOk.TabIndex = 1;
            this._btnOk.Text = "Accept";
            this._btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // _cmboPaperSize
            // 
            this._cmboPaperSize.Location = new System.Drawing.Point(4, 48);
            this._cmboPaperSize.Name = "_cmboPaperSize";
            this._cmboPaperSize.Size = new System.Drawing.Size(420, 21);
            this._cmboPaperSize.TabIndex = 2;
            // 
            // _Label1
            // 
            this._Label1.Location = new System.Drawing.Point(4, 72);
            this._Label1.Name = "_Label1";
            this._Label1.Size = new System.Drawing.Size(120, 16);
            this._Label1.TabIndex = 3;
            this._Label1.Text = "Select Paper Size";
            // 
            // _GroupBox1
            // 
            this._GroupBox1.Controls.Add(this._Label2);
            this._GroupBox1.Controls.Add(this._txtEndPage);
            this._GroupBox1.Controls.Add(this._txtStartPage);
            this._GroupBox1.Controls.Add(this._rbPageRange);
            this._GroupBox1.Controls.Add(this._rbAllPages);
            this._GroupBox1.Location = new System.Drawing.Point(8, 96);
            this._GroupBox1.Name = "_GroupBox1";
            this._GroupBox1.Size = new System.Drawing.Size(160, 108);
            this._GroupBox1.TabIndex = 4;
            this._GroupBox1.TabStop = false;
            this._GroupBox1.Text = "Print What?";
            // 
            // _Label2
            // 
            this._Label2.Location = new System.Drawing.Point(72, 80);
            this._Label2.Name = "_Label2";
            this._Label2.Size = new System.Drawing.Size(20, 16);
            this._Label2.TabIndex = 4;
            this._Label2.Text = "to";
            this._Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // _txtEndPage
            // 
            this._txtEndPage.Location = new System.Drawing.Point(100, 76);
            this._txtEndPage.Name = "_txtEndPage";
            this._txtEndPage.Size = new System.Drawing.Size(44, 20);
            this._txtEndPage.TabIndex = 3;
            // 
            // _txtStartPage
            // 
            this._txtStartPage.Location = new System.Drawing.Point(20, 76);
            this._txtStartPage.Name = "_txtStartPage";
            this._txtStartPage.Size = new System.Drawing.Size(44, 20);
            this._txtStartPage.TabIndex = 2;
            // 
            // _rbPageRange
            // 
            this._rbPageRange.Location = new System.Drawing.Point(20, 52);
            this._rbPageRange.Name = "_rbPageRange";
            this._rbPageRange.Size = new System.Drawing.Size(124, 20);
            this._rbPageRange.TabIndex = 1;
            this._rbPageRange.Text = "A Range Of Pages";
            // 
            // _rbAllPages
            // 
            this._rbAllPages.Checked = true;
            this._rbAllPages.Location = new System.Drawing.Point(20, 28);
            this._rbAllPages.Name = "_rbAllPages";
            this._rbAllPages.Size = new System.Drawing.Size(124, 20);
            this._rbAllPages.TabIndex = 0;
            this._rbAllPages.TabStop = true;
            this._rbAllPages.Text = "All Pages";
            // 
            // _GroupBox2
            // 
            this._GroupBox2.Controls.Add(this._rbLandscape);
            this._GroupBox2.Controls.Add(this._rbProtriate);
            this._GroupBox2.Location = new System.Drawing.Point(260, 96);
            this._GroupBox2.Name = "_GroupBox2";
            this._GroupBox2.Size = new System.Drawing.Size(160, 108);
            this._GroupBox2.TabIndex = 5;
            this._GroupBox2.TabStop = false;
            this._GroupBox2.Text = "Orientation";
            // 
            // _rbLandscape
            // 
            this._rbLandscape.Location = new System.Drawing.Point(20, 52);
            this._rbLandscape.Name = "_rbLandscape";
            this._rbLandscape.Size = new System.Drawing.Size(124, 20);
            this._rbLandscape.TabIndex = 1;
            this._rbLandscape.Text = "Landscape";
            // 
            // _rbProtriate
            // 
            this._rbProtriate.Checked = true;
            this._rbProtriate.Location = new System.Drawing.Point(20, 28);
            this._rbProtriate.Name = "_rbProtriate";
            this._rbProtriate.Size = new System.Drawing.Size(124, 20);
            this._rbProtriate.TabIndex = 0;
            this._rbProtriate.TabStop = true;
            this._rbProtriate.Text = "Portrait";
            // 
            // _Label3
            // 
            this._Label3.Location = new System.Drawing.Point(4, 28);
            this._Label3.Name = "_Label3";
            this._Label3.Size = new System.Drawing.Size(152, 16);
            this._Label3.TabIndex = 7;
            this._Label3.Text = "Select Printer to Print to";
            // 
            // _cmboPrinter
            // 
            this._cmboPrinter.Location = new System.Drawing.Point(4, 4);
            this._cmboPrinter.Name = "_cmboPrinter";
            this._cmboPrinter.Size = new System.Drawing.Size(420, 21);
            this._cmboPrinter.TabIndex = 6;
            // 
            // _btnPrint
            // 
            this._btnPrint.Location = new System.Drawing.Point(232, 208);
            this._btnPrint.Name = "_btnPrint";
            this._btnPrint.Size = new System.Drawing.Size(60, 20);
            this._btnPrint.TabIndex = 8;
            this._btnPrint.Text = "Print";
            this._btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // _btnPreview
            // 
            this._btnPreview.Location = new System.Drawing.Point(168, 208);
            this._btnPreview.Name = "_btnPreview";
            this._btnPreview.Size = new System.Drawing.Size(60, 20);
            this._btnPreview.TabIndex = 9;
            this._btnPreview.Text = "Preview";
            this._btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // frmPageSetup
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(430, 231);
            this.Controls.Add(this._btnPreview);
            this.Controls.Add(this._btnPrint);
            this.Controls.Add(this._Label3);
            this.Controls.Add(this._cmboPrinter);
            this.Controls.Add(this._GroupBox2);
            this.Controls.Add(this._GroupBox1);
            this.Controls.Add(this._Label1);
            this.Controls.Add(this._cmboPaperSize);
            this.Controls.Add(this._btnOk);
            this.Controls.Add(this._btnCancel);
            this.Name = "frmPageSetup";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Setup Printed Page Metrics";
            this._GroupBox1.ResumeLayout(false);
            this._GroupBox1.PerformLayout();
            this._GroupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

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
