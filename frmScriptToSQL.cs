using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using System.Runtime.CompilerServices;

namespace TAIGridControl2
{
    public class frmScriptToSQL : System.Windows.Forms.Form
    {
        private TAIGridControl _taig;

        public frmScriptToSQL(TAIGridControl TAIG) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            _taig = TAIG;
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

        private System.Windows.Forms.TextBox _txtTableName;

        internal System.Windows.Forms.TextBox txtTableName
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtTableName;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtTableName != null)
                {
                }

                _txtTableName = value;
                if (_txtTableName != null)
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
                    _Label2.Click -= Label2_Click;
                }

                _Label2 = value;
                if (_Label2 != null)
                {
                    _Label2.Click += Label2_Click;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            this._txtFileName = new System.Windows.Forms.TextBox();
            this._btnSelectFile = new System.Windows.Forms.Button();
            this._OFD = new System.Windows.Forms.OpenFileDialog();
            this._btnCancel = new System.Windows.Forms.Button();
            this._btnOk = new System.Windows.Forms.Button();
            this._txtTableName = new System.Windows.Forms.TextBox();
            this._Label1 = new System.Windows.Forms.Label();
            this._Label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // _txtFileName
            // 
            this._txtFileName.Location = new System.Drawing.Point(4, 44);
            this._txtFileName.Name = "_txtFileName";
            this._txtFileName.Size = new System.Drawing.Size(344, 20);
            this._txtFileName.TabIndex = 0;
            // 
            // _btnSelectFile
            // 
            this._btnSelectFile.Location = new System.Drawing.Point(352, 44);
            this._btnSelectFile.Name = "_btnSelectFile";
            this._btnSelectFile.Size = new System.Drawing.Size(24, 20);
            this._btnSelectFile.TabIndex = 1;
            this._btnSelectFile.Text = "...";
            this._btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // _OFD
            // 
            this._OFD.CheckFileExists = false;
            this._OFD.DefaultExt = "sql";
            this._OFD.Title = "Select file name to save the script to";
            // 
            // _btnCancel
            // 
            this._btnCancel.Location = new System.Drawing.Point(260, 4);
            this._btnCancel.Name = "_btnCancel";
            this._btnCancel.Size = new System.Drawing.Size(56, 20);
            this._btnCancel.TabIndex = 2;
            this._btnCancel.Text = "Cancel";
            this._btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // _btnOk
            // 
            this._btnOk.Location = new System.Drawing.Point(320, 4);
            this._btnOk.Name = "_btnOk";
            this._btnOk.Size = new System.Drawing.Size(56, 20);
            this._btnOk.TabIndex = 3;
            this._btnOk.Text = "Ok";
            this._btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // _txtTableName
            // 
            this._txtTableName.Location = new System.Drawing.Point(4, 20);
            this._txtTableName.Name = "_txtTableName";
            this._txtTableName.Size = new System.Drawing.Size(152, 20);
            this._txtTableName.TabIndex = 4;
            // 
            // _Label1
            // 
            this._Label1.Location = new System.Drawing.Point(8, 4);
            this._Label1.Name = "_Label1";
            this._Label1.Size = new System.Drawing.Size(148, 12);
            this._Label1.TabIndex = 5;
            this._Label1.Text = "Name the resulting table";
            // 
            // _Label2
            // 
            this._Label2.Location = new System.Drawing.Point(8, 68);
            this._Label2.Name = "_Label2";
            this._Label2.Size = new System.Drawing.Size(188, 16);
            this._Label2.TabIndex = 6;
            this._Label2.Text = "Save the resulting SQL script to";
            // 
            // frmScriptToSQL
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(394, 110);
            this.ControlBox = false;
            this.Controls.Add(this._Label2);
            this.Controls.Add(this._Label1);
            this.Controls.Add(this._txtTableName);
            this.Controls.Add(this._btnOk);
            this.Controls.Add(this._btnCancel);
            this.Controls.Add(this._btnSelectFile);
            this.Controls.Add(this._txtFileName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "frmScriptToSQL";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Script Grid contents to SQL";
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        private void Label2_Click(object sender, EventArgs e)
        {
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text) | string.IsNullOrEmpty(txtTableName.Text))
                Interaction.MsgBox("You must select a file for the resulting Script to be written to" + Constants.vbCrLf + "as well as a name for the resulting Table that is crafted by that script", MsgBoxStyle.Information, "Script Grid to SQL error");
            else
            {
                System.IO.TextWriter txtw = new System.IO.StreamWriter(txtFileName.Text);
                txtw.Write(_taig.CreatePersistanceScript(txtTableName.Text));
                txtw.Close();

                Hide();
            }
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
    }
}
