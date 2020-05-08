using System.Diagnostics;
using System;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

namespace TAIGridControl2
{
    public class frmFreqDist : System.Windows.Forms.Form
    {
        public frmFreqDist(TAIGridControl TAIG, int coltoCount) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            taigFreq.FrequencyDistribution(TAIG, coltoCount, true);
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
        private System.ComponentModel.IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        private TAIGridControl _taigFreq;

        internal TAIGridControl taigFreq
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _taigFreq;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_taigFreq != null)
                {
                }

                _taigFreq = value;
                if (_taigFreq != null)
                {
                }
            }
        }

        private System.Windows.Forms.Button _btnClose;

        internal System.Windows.Forms.Button btnClose
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnClose;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnClose != null)
                {
                    _btnClose.Click -= btnClose_Click;
                }

                _btnClose = value;
                if (_btnClose != null)
                {
                    _btnClose.Click += btnClose_Click;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            System.Drawing.StringFormat stringFormat1 = new System.Drawing.StringFormat();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFreqDist));
            this._taigFreq = new TAIGridControl2.TAIGridControl();
            this._btnClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // _taigFreq
            // 
            this._taigFreq.AlternateColoration = false;
            this._taigFreq.AlternateColorationAltColor = System.Drawing.Color.MediumSpringGreen;
            this._taigFreq.AlternateColorationBaseColor = System.Drawing.Color.AntiqueWhite;
            this._taigFreq.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._taigFreq.BorderColor = System.Drawing.Color.Black;
            this._taigFreq.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this._taigFreq.CellOutlines = true;
            this._taigFreq.ColBackColorEdit = System.Drawing.Color.Yellow;
            this._taigFreq.Cols = 0;
            this._taigFreq.DefaultBackgroundColor = System.Drawing.Color.AntiqueWhite;
            this._taigFreq.DefaultCellFont = new System.Drawing.Font("Arial", 9F);
            this._taigFreq.DefaultForegroundColor = System.Drawing.Color.Black;
            this._taigFreq.Delimiter = ",";
            this._taigFreq.ExcelAlternateColoration = System.Drawing.Color.FromArgb(((int)(((byte)(204)))), ((int)(((byte)(255)))), ((int)(((byte)(204)))));
            this._taigFreq.ExcelAutoFitColumn = true;
            this._taigFreq.ExcelAutoFitRow = true;
            this._taigFreq.ExcelFilename = "";
            this._taigFreq.ExcelIncludeColumnHeaders = true;
            this._taigFreq.ExcelKeepAlive = true;
            this._taigFreq.ExcelMatchGridColorScheme = true;
            this._taigFreq.ExcelMaximized = true;
            this._taigFreq.ExcelMaxRowsPerSheet = 30000;
            this._taigFreq.ExcelOutlineCells = true;
            this._taigFreq.ExcelPageOrientation = 1;
            this._taigFreq.ExcelShowBorders = false;
            this._taigFreq.ExcelUseAlternateRowColor = true;
            this._taigFreq.ExcelWorksheetName = "Grid Output";
            this._taigFreq.GridEditMode = TAIGridControl2.TAIGridControl.GridEditModes.KeyReturn;
            this._taigFreq.GridHeaderBackColor = System.Drawing.Color.LightBlue;
            this._taigFreq.GridHeaderFont = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this._taigFreq.GridHeaderForeColor = System.Drawing.Color.Black;
            this._taigFreq.GridHeaderHeight = 16;
            stringFormat1.Alignment = System.Drawing.StringAlignment.Near;
            stringFormat1.HotkeyPrefix = System.Drawing.Text.HotkeyPrefix.None;
            stringFormat1.LineAlignment = System.Drawing.StringAlignment.Near;
            stringFormat1.Trimming = System.Drawing.StringTrimming.Character;
            this._taigFreq.GridHeaderStringFormat = stringFormat1;
            this._taigFreq.GridheaderVisible = true;
            this._taigFreq.Location = new System.Drawing.Point(0, 4);
            this._taigFreq.Name = "_taigFreq";
            this._taigFreq.PageSettings = null;
            this._taigFreq.PaginationSize = 0;
            this._taigFreq.Rows = 0;
            this._taigFreq.ScrollInterval = 5;
            this._taigFreq.SelectedColBackColor = System.Drawing.Color.MediumSlateBlue;
            this._taigFreq.SelectedColForeColor = System.Drawing.Color.LightGray;
            this._taigFreq.SelectedColumn = -1;
            this._taigFreq.SelectedRow = -1;
            this._taigFreq.SelectedRowBackColor = System.Drawing.Color.Blue;
            this._taigFreq.SelectedRowForeColor = System.Drawing.Color.White;
            this._taigFreq.SelectedRows = ((System.Collections.ArrayList)(resources.GetObject("_taigFreq.SelectedRows")));
            this._taigFreq.Size = new System.Drawing.Size(392, 344);
            this._taigFreq.TabIndex = 0;
            this._taigFreq.TitleBackColor = System.Drawing.Color.Blue;
            this._taigFreq.TitleFont = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._taigFreq.TitleForeColor = System.Drawing.Color.White;
            this._taigFreq.TitleText = "Frequency Distribution";
            this._taigFreq.TitleVisible = true;
            this._taigFreq.XMLDataSetName = "Grid_Output";
            this._taigFreq.XMLFileName = "";
            this._taigFreq.XMLIncludeSchema = false;
            this._taigFreq.XMLNameSpace = "TAI_Grid_Ouptut";
            this._taigFreq.XMLTableName = "Table";
            // 
            // _btnClose
            // 
            this._btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._btnClose.Location = new System.Drawing.Point(156, 352);
            this._btnClose.Name = "_btnClose";
            this._btnClose.Size = new System.Drawing.Size(88, 20);
            this._btnClose.TabIndex = 1;
            this._btnClose.Text = "Close";
            // 
            // frmFreqDist
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(394, 375);
            this.Controls.Add(this._btnClose);
            this.Controls.Add(this._taigFreq);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "frmFreqDist";
            this.ShowInTaskbar = false;
            this.Text = "Frequency Distribution in source grid";
            this.ResumeLayout(false);

        }


        private void btnClose_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
