using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using System.Runtime.CompilerServices;

namespace TAIGridControl2
{
    public class frmScriptToHTML : System.Windows.Forms.Form
    {
        private TAIGridControl _Taig;


        public frmScriptToHTML(TAIGridControl TAIG) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            _Taig = TAIG;
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
        private System.Windows.Forms.Label _Label2;

        internal System.Windows.Forms.Label Label2
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

        private System.Windows.Forms.Button _btnOk;

        internal System.Windows.Forms.Button btnOk
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

        private System.Windows.Forms.Button _btnCancel;

        internal System.Windows.Forms.Button btnCancel
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

        private System.Windows.Forms.Button _btnSelectFile;

        internal System.Windows.Forms.Button btnSelectFile
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnSelectFile;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnSelectFile != null)
                {
                    _btnSelectFile.Click -= btnSelectFile_Click;
                }

                _btnSelectFile = value;
                if (_btnSelectFile != null)
                {
                    _btnSelectFile.Click += btnSelectFile_Click;
                }
            }
        }

        private System.Windows.Forms.TextBox _txtFileName;

        internal System.Windows.Forms.TextBox txtFileName
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtFileName;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtFileName != null)
                {
                }

                _txtFileName = value;
                if (_txtFileName != null)
                {
                }
            }
        }

        private System.Windows.Forms.OpenFileDialog _OFD;

        internal System.Windows.Forms.OpenFileDialog OFD
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _OFD;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_OFD != null)
                {
                }

                _OFD = value;
                if (_OFD != null)
                {
                }
            }
        }

        private System.Windows.Forms.TextBox _txtBorder;

        internal System.Windows.Forms.TextBox txtBorder
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtBorder;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtBorder != null)
                {
                }

                _txtBorder = value;
                if (_txtBorder != null)
                {
                }
            }
        }

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

        private System.Windows.Forms.CheckBox _chkMatchColors;

        internal System.Windows.Forms.CheckBox chkMatchColors
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _chkMatchColors;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_chkMatchColors != null)
                {
                }

                _chkMatchColors = value;
                if (_chkMatchColors != null)
                {
                }
            }
        }

        private System.Windows.Forms.CheckBox _chkOmitNulls;

        internal System.Windows.Forms.CheckBox chkOmitNulls
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _chkOmitNulls;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_chkOmitNulls != null)
                {
                }

                _chkOmitNulls = value;
                if (_chkOmitNulls != null)
                {
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            this._Label2 = new System.Windows.Forms.Label();
            this._btnOk = new System.Windows.Forms.Button();
            this._btnCancel = new System.Windows.Forms.Button();
            this._btnSelectFile = new System.Windows.Forms.Button();
            this._txtFileName = new System.Windows.Forms.TextBox();
            this._OFD = new System.Windows.Forms.OpenFileDialog();
            this._txtBorder = new System.Windows.Forms.TextBox();
            this._Label1 = new System.Windows.Forms.Label();
            this._chkMatchColors = new System.Windows.Forms.CheckBox();
            this._chkOmitNulls = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // _Label2
            // 
            this._Label2.Location = new System.Drawing.Point(4, 68);
            this._Label2.Name = "_Label2";
            this._Label2.Size = new System.Drawing.Size(212, 16);
            this._Label2.TabIndex = 11;
            this._Label2.Text = "Save the resulting HTML Table script to";
            // 
            // _btnOk
            // 
            this._btnOk.Location = new System.Drawing.Point(316, 4);
            this._btnOk.Name = "_btnOk";
            this._btnOk.Size = new System.Drawing.Size(56, 20);
            this._btnOk.TabIndex = 10;
            this._btnOk.Text = "Ok";
            this._btnOk.Click += new System.EventHandler(this._btnOk_Click);
            // 
            // _btnCancel
            // 
            this._btnCancel.Location = new System.Drawing.Point(256, 4);
            this._btnCancel.Name = "_btnCancel";
            this._btnCancel.Size = new System.Drawing.Size(56, 20);
            this._btnCancel.TabIndex = 9;
            this._btnCancel.Text = "Cancel";
            this._btnCancel.Click += new System.EventHandler(this._btnCancel_Click);
            // 
            // _btnSelectFile
            // 
            this._btnSelectFile.Location = new System.Drawing.Point(348, 44);
            this._btnSelectFile.Name = "_btnSelectFile";
            this._btnSelectFile.Size = new System.Drawing.Size(24, 20);
            this._btnSelectFile.TabIndex = 8;
            this._btnSelectFile.Text = "...";
            // 
            // _txtFileName
            // 
            this._txtFileName.Location = new System.Drawing.Point(4, 44);
            this._txtFileName.Name = "_txtFileName";
            this._txtFileName.Size = new System.Drawing.Size(340, 20);
            this._txtFileName.TabIndex = 7;
            // 
            // _OFD
            // 
            this._OFD.CheckFileExists = false;
            this._OFD.DefaultExt = "html";
            this._OFD.Title = "Select file name to save the script to";
            // 
            // _txtBorder
            // 
            this._txtBorder.Location = new System.Drawing.Point(4, 20);
            this._txtBorder.Name = "_txtBorder";
            this._txtBorder.Size = new System.Drawing.Size(28, 20);
            this._txtBorder.TabIndex = 12;
            this._txtBorder.Text = "1";
            // 
            // _Label1
            // 
            this._Label1.Location = new System.Drawing.Point(0, 4);
            this._Label1.Name = "_Label1";
            this._Label1.Size = new System.Drawing.Size(60, 12);
            this._Label1.TabIndex = 13;
            this._Label1.Text = "Border";
            // 
            // _chkMatchColors
            // 
            this._chkMatchColors.Checked = true;
            this._chkMatchColors.CheckState = System.Windows.Forms.CheckState.Checked;
            this._chkMatchColors.Location = new System.Drawing.Point(40, 24);
            this._chkMatchColors.Name = "_chkMatchColors";
            this._chkMatchColors.Size = new System.Drawing.Size(96, 16);
            this._chkMatchColors.TabIndex = 14;
            this._chkMatchColors.Text = "Match Colors";
            // 
            // _chkOmitNulls
            // 
            this._chkOmitNulls.Checked = true;
            this._chkOmitNulls.CheckState = System.Windows.Forms.CheckState.Checked;
            this._chkOmitNulls.Location = new System.Drawing.Point(144, 24);
            this._chkOmitNulls.Name = "_chkOmitNulls";
            this._chkOmitNulls.Size = new System.Drawing.Size(96, 16);
            this._chkOmitNulls.TabIndex = 15;
            this._chkOmitNulls.Text = "Omit Nulls";
            // 
            // frmScriptToHTML
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(376, 110);
            this.ControlBox = false;
            this.Controls.Add(this._chkOmitNulls);
            this.Controls.Add(this._chkMatchColors);
            this.Controls.Add(this._Label1);
            this.Controls.Add(this._txtBorder);
            this.Controls.Add(this._Label2);
            this.Controls.Add(this._btnOk);
            this.Controls.Add(this._btnCancel);
            this.Controls.Add(this._btnSelectFile);
            this.Controls.Add(this._txtFileName);
            this.Name = "frmScriptToHTML";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Script Grid as an HTML Table";
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OFD.CheckFileExists = false;
            if ((int)OFD.ShowDialog() == (int)System.Windows.Forms.DialogResult.OK)
                txtFileName.Text = OFD.FileName;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text))
                Interaction.MsgBox("You must select a failename to save the resulting HTML Table to...", MsgBoxStyle.Information, "Save to HTML Table error");
            else
            {
                int bv = 0;

                if (Information.IsNumeric(txtBorder.Text))
                    bv = int.Parse(txtBorder.Text);

                System.IO.TextWriter txtw = new System.IO.StreamWriter(txtFileName.Text);
                txtw.Write(_Taig.CreateHTMLTableScript(bv, chkMatchColors.Checked, chkOmitNulls.Checked));
                txtw.Close();

                Hide();
            }
        }

        private void _btnCancel_Click(object sender, EventArgs e)
        {
            btnCancel_Click(sender, e);
        }

        private void _btnOk_Click(object sender, EventArgs e)
        {
            btnOk_Click(sender, e);
        }
    }
}
