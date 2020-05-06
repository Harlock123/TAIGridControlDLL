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
            _txtFileName = new System.Windows.Forms.TextBox();
            _btnSelectFile = new System.Windows.Forms.Button();
            _btnSelectFile.Click += btnSelectFile_Click;
            _OFD = new System.Windows.Forms.OpenFileDialog();
            _btnCancel = new System.Windows.Forms.Button();
            _btnCancel.Click += btnCancel_Click;
            _btnOk = new System.Windows.Forms.Button();
            _btnOk.Click += btnOk_Click;
            _txtTableName = new System.Windows.Forms.TextBox();
            _Label1 = new System.Windows.Forms.Label();
            _Label2 = new System.Windows.Forms.Label();
            _Label2.Click += Label2_Click;
            SuspendLayout();
            // 
            // txtFileName
            // 
            _txtFileName.Location = new System.Drawing.Point(4, 44);
            _txtFileName.Name = "txtFileName";
            _txtFileName.Size = new System.Drawing.Size(344, 20);
            _txtFileName.TabIndex = 0;
            _txtFileName.Text = "";
            // 
            // btnSelectFile
            // 
            _btnSelectFile.Location = new System.Drawing.Point(352, 44);
            _btnSelectFile.Name = "btnSelectFile";
            _btnSelectFile.Size = new System.Drawing.Size(24, 20);
            _btnSelectFile.TabIndex = 1;
            _btnSelectFile.Text = "...";
            // 
            // OFD
            // 
            _OFD.CheckFileExists = false;
            _OFD.DefaultExt = "sql";
            _OFD.Title = "Select file name to save the script to";
            // 
            // btnCancel
            // 
            _btnCancel.Location = new System.Drawing.Point(260, 4);
            _btnCancel.Name = "btnCancel";
            _btnCancel.Size = new System.Drawing.Size(56, 20);
            _btnCancel.TabIndex = 2;
            _btnCancel.Text = "Cancel";
            // 
            // btnOk
            // 
            _btnOk.Location = new System.Drawing.Point(320, 4);
            _btnOk.Name = "btnOk";
            _btnOk.Size = new System.Drawing.Size(56, 20);
            _btnOk.TabIndex = 3;
            _btnOk.Text = "Ok";
            // 
            // txtTableName
            // 
            _txtTableName.Location = new System.Drawing.Point(4, 20);
            _txtTableName.Name = "txtTableName";
            _txtTableName.Size = new System.Drawing.Size(152, 20);
            _txtTableName.TabIndex = 4;
            _txtTableName.Text = "";
            // 
            // Label1
            // 
            _Label1.Location = new System.Drawing.Point(8, 4);
            _Label1.Name = "Label1";
            _Label1.Size = new System.Drawing.Size(148, 12);
            _Label1.TabIndex = 5;
            _Label1.Text = "Name the resulting table";
            // 
            // Label2
            // 
            _Label2.Location = new System.Drawing.Point(8, 68);
            _Label2.Name = "Label2";
            _Label2.Size = new System.Drawing.Size(188, 16);
            _Label2.TabIndex = 6;
            _Label2.Text = "Save the resulting SQL script to";
            // 
            // frmScriptToSQL
            // 
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            ClientSize = new System.Drawing.Size(394, 88);
            ControlBox = false;
            Controls.Add(_Label2);
            Controls.Add(_Label1);
            Controls.Add(_txtTableName);
            Controls.Add(_btnOk);
            Controls.Add(_btnCancel);
            Controls.Add(_btnSelectFile);
            Controls.Add(_txtFileName);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            Name = "frmScriptToSQL";
            ShowInTaskbar = false;
            StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            Text = "Script Grid contents to SQL";
            ResumeLayout(false);
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
