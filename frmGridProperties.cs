using System.Diagnostics;
using System;
using System.Runtime.CompilerServices;

namespace TAIGridControl2
{
    public class frmGridProperties : System.Windows.Forms.Form
    {
        private TAIGridControl _pg;

        public frmGridProperties(TAIGridControl ParentGrid) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            _pg = ParentGrid;

            chkShowHeaders.Checked = _pg.GridheaderVisible;
            chkShowTitle.Checked = _pg.TitleVisible;
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
        private System.Windows.Forms.Label _Label1;

        internal System.Windows.Forms.Label Label1
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

        private System.Windows.Forms.Button _btnFontsSmaller;

        internal System.Windows.Forms.Button btnFontsSmaller
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnFontsSmaller;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnFontsSmaller != null)
                {
                    _btnFontsSmaller.Click -= btnFontsSmaller_Click;
                }

                _btnFontsSmaller = value;
                if (_btnFontsSmaller != null)
                {
                    _btnFontsSmaller.Click += btnFontsSmaller_Click;
                }
            }
        }

        private System.Windows.Forms.Button _btnFontsLarger;

        internal System.Windows.Forms.Button btnFontsLarger
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnFontsLarger;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnFontsLarger != null)
                {
                    _btnFontsLarger.Click -= btnFontsLarger_Click;
                }

                _btnFontsLarger = value;
                if (_btnFontsLarger != null)
                {
                    _btnFontsLarger.Click += btnFontsLarger_Click;
                }
            }
        }

        private System.Windows.Forms.Button _btnExit;

        internal System.Windows.Forms.Button btnExit
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnExit;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnExit != null)
                {
                    _btnExit.Click -= btnExit_Click;
                }

                _btnExit = value;
                if (_btnExit != null)
                {
                    _btnExit.Click += btnExit_Click;
                }
            }
        }

        private System.Windows.Forms.CheckBox _chkShowTitle;

        internal System.Windows.Forms.CheckBox chkShowTitle
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _chkShowTitle;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_chkShowTitle != null)
                {
                    _chkShowTitle.CheckedChanged -= chkShowTitle_CheckedChanged;
                }

                _chkShowTitle = value;
                if (_chkShowTitle != null)
                {
                    _chkShowTitle.CheckedChanged += chkShowTitle_CheckedChanged;
                }
            }
        }

        private System.Windows.Forms.CheckBox _chkShowHeaders;

        internal System.Windows.Forms.CheckBox chkShowHeaders
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _chkShowHeaders;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_chkShowHeaders != null)
                {
                    _chkShowHeaders.CheckedChanged -= chkShowHeaders_CheckedChanged;
                }

                _chkShowHeaders = value;
                if (_chkShowHeaders != null)
                {
                    _chkShowHeaders.CheckedChanged += chkShowHeaders_CheckedChanged;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            _Label1 = new System.Windows.Forms.Label();
            _btnFontsSmaller = new System.Windows.Forms.Button();
            _btnFontsSmaller.Click += btnFontsSmaller_Click;
            _btnFontsLarger = new System.Windows.Forms.Button();
            _btnFontsLarger.Click += btnFontsLarger_Click;
            _btnExit = new System.Windows.Forms.Button();
            _btnExit.Click += btnExit_Click;
            _chkShowTitle = new System.Windows.Forms.CheckBox();
            _chkShowTitle.CheckedChanged += chkShowTitle_CheckedChanged;
            _chkShowHeaders = new System.Windows.Forms.CheckBox();
            _chkShowHeaders.CheckedChanged += chkShowHeaders_CheckedChanged;
            SuspendLayout();
            // 
            // Label1
            // 
            _Label1.Location = new System.Drawing.Point(8, 8);
            _Label1.Name = "Label1";
            _Label1.Size = new System.Drawing.Size(76, 16);
            _Label1.TabIndex = 0;
            _Label1.Text = "Grids Fonts";
            // 
            // btnFontsSmaller
            // 
            _btnFontsSmaller.Location = new System.Drawing.Point(12, 28);
            _btnFontsSmaller.Name = "btnFontsSmaller";
            _btnFontsSmaller.Size = new System.Drawing.Size(16, 20);
            _btnFontsSmaller.TabIndex = 1;
            _btnFontsSmaller.Text = "<";
            // 
            // btnFontsLarger
            // 
            _btnFontsLarger.Location = new System.Drawing.Point(44, 28);
            _btnFontsLarger.Name = "btnFontsLarger";
            _btnFontsLarger.Size = new System.Drawing.Size(16, 20);
            _btnFontsLarger.TabIndex = 2;
            _btnFontsLarger.Text = ">";
            // 
            // btnExit
            // 
            _btnExit.Location = new System.Drawing.Point(364, 20);
            _btnExit.Name = "btnExit";
            _btnExit.Size = new System.Drawing.Size(72, 24);
            _btnExit.TabIndex = 3;
            _btnExit.Text = "Close";
            // 
            // chkShowTitle
            // 
            _chkShowTitle.Location = new System.Drawing.Point(144, 16);
            _chkShowTitle.Name = "chkShowTitle";
            _chkShowTitle.Size = new System.Drawing.Size(152, 16);
            _chkShowTitle.TabIndex = 4;
            _chkShowTitle.Text = "Show Grid Title Bar";
            // 
            // chkShowHeaders
            // 
            _chkShowHeaders.Location = new System.Drawing.Point(144, 36);
            _chkShowHeaders.Name = "chkShowHeaders";
            _chkShowHeaders.Size = new System.Drawing.Size(172, 16);
            _chkShowHeaders.TabIndex = 5;
            _chkShowHeaders.Text = "Show Grid Column Headers";
            // 
            // frmGridProperties
            // 
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            BackColor = System.Drawing.SystemColors.Info;
            ClientSize = new System.Drawing.Size(442, 68);
            Controls.Add(_chkShowHeaders);
            Controls.Add(_chkShowTitle);
            Controls.Add(_btnExit);
            Controls.Add(_btnFontsLarger);
            Controls.Add(_btnFontsSmaller);
            Controls.Add(_Label1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "frmGridProperties";
            Text = "Properties for the Grid";
            ResumeLayout(false);
        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void btnFontsSmaller_Click(object sender, EventArgs e)
        {
            _pg.AllFontsSmaller();
        }

        private void btnFontsLarger_Click(object sender, EventArgs e)
        {
            _pg.AllFontsLarger();
        }

        private void chkShowTitle_CheckedChanged(object sender, EventArgs e)
        {
            _pg.TitleVisible = chkShowTitle.Checked;
        }

        private void chkShowHeaders_CheckedChanged(object sender, EventArgs e)
        {
            _pg.GridheaderVisible = chkShowHeaders.Checked;
        }
    }
}
