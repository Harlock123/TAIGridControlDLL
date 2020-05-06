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
            _Label2 = new System.Windows.Forms.Label();
            _btnOk = new System.Windows.Forms.Button();
            _btnOk.Click += btnOk_Click;
            _btnCancel = new System.Windows.Forms.Button();
            _btnCancel.Click += btnCancel_Click;
            _btnSelectFile = new System.Windows.Forms.Button();
            _btnSelectFile.Click += btnSelectFile_Click;
            _txtFileName = new System.Windows.Forms.TextBox();
            _OFD = new System.Windows.Forms.OpenFileDialog();
            _txtBorder = new System.Windows.Forms.TextBox();
            _Label1 = new System.Windows.Forms.Label();
            _chkMatchColors = new System.Windows.Forms.CheckBox();
            _chkOmitNulls = new System.Windows.Forms.CheckBox();
            SuspendLayout();
            // 
            // Label2
            // 
            _Label2.Location = new System.Drawing.Point(4, 68);
            _Label2.Name = "Label2";
            _Label2.Size = new System.Drawing.Size(212, 16);
            _Label2.TabIndex = 11;
            _Label2.Text = "Save the resulting HTML Table script to";
            // 
            // btnOk
            // 
            _btnOk.Location = new System.Drawing.Point(316, 4);
            _btnOk.Name = "btnOk";
            _btnOk.Size = new System.Drawing.Size(56, 20);
            _btnOk.TabIndex = 10;
            _btnOk.Text = "Ok";
            // 
            // btnCancel
            // 
            _btnCancel.Location = new System.Drawing.Point(256, 4);
            _btnCancel.Name = "btnCancel";
            _btnCancel.Size = new System.Drawing.Size(56, 20);
            _btnCancel.TabIndex = 9;
            _btnCancel.Text = "Cancel";
            // 
            // btnSelectFile
            // 
            _btnSelectFile.Location = new System.Drawing.Point(348, 44);
            _btnSelectFile.Name = "btnSelectFile";
            _btnSelectFile.Size = new System.Drawing.Size(24, 20);
            _btnSelectFile.TabIndex = 8;
            _btnSelectFile.Text = "...";
            // 
            // txtFileName
            // 
            _txtFileName.Location = new System.Drawing.Point(4, 44);
            _txtFileName.Name = "txtFileName";
            _txtFileName.Size = new System.Drawing.Size(340, 20);
            _txtFileName.TabIndex = 7;
            _txtFileName.Text = "";
            // 
            // OFD
            // 
            _OFD.CheckFileExists = false;
            _OFD.DefaultExt = "html";
            _OFD.Title = "Select file name to save the script to";
            // 
            // txtBorder
            // 
            _txtBorder.Location = new System.Drawing.Point(4, 20);
            _txtBorder.Name = "txtBorder";
            _txtBorder.Size = new System.Drawing.Size(28, 20);
            _txtBorder.TabIndex = 12;
            _txtBorder.Text = "1";
            // 
            // Label1
            // 
            _Label1.Location = new System.Drawing.Point(0, 4);
            _Label1.Name = "Label1";
            _Label1.Size = new System.Drawing.Size(60, 12);
            _Label1.TabIndex = 13;
            _Label1.Text = "Border";
            // 
            // chkMatchColors
            // 
            _chkMatchColors.Checked = true;
            _chkMatchColors.CheckState = System.Windows.Forms.CheckState.Checked;
            _chkMatchColors.Location = new System.Drawing.Point(40, 24);
            _chkMatchColors.Name = "chkMatchColors";
            _chkMatchColors.Size = new System.Drawing.Size(96, 16);
            _chkMatchColors.TabIndex = 14;
            _chkMatchColors.Text = "Match Colors";
            // 
            // chkOmitNulls
            // 
            _chkOmitNulls.Checked = true;
            _chkOmitNulls.CheckState = System.Windows.Forms.CheckState.Checked;
            _chkOmitNulls.Location = new System.Drawing.Point(144, 24);
            _chkOmitNulls.Name = "chkOmitNulls";
            _chkOmitNulls.Size = new System.Drawing.Size(96, 16);
            _chkOmitNulls.TabIndex = 15;
            _chkOmitNulls.Text = "Omit Nulls";
            // 
            // frmScriptToHTML
            // 
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            ClientSize = new System.Drawing.Size(376, 83);
            ControlBox = false;
            Controls.Add(_chkOmitNulls);
            Controls.Add(_chkMatchColors);
            Controls.Add(_Label1);
            Controls.Add(_txtBorder);
            Controls.Add(_Label2);
            Controls.Add(_btnOk);
            Controls.Add(_btnCancel);
            Controls.Add(_btnSelectFile);
            Controls.Add(_txtFileName);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            Name = "frmScriptToHTML";
            ShowInTaskbar = false;
            StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            Text = "Script Grid as an HTML Table";
            ResumeLayout(false);
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
    }
}
