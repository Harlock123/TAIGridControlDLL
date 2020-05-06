using System.Diagnostics;
using System;
using System.Runtime.CompilerServices;

namespace TAIGridControl2
{
    public class frmMultipleColumnTearAway : System.Windows.Forms.Form
    {
        public frmMultipleColumnTearAway(string[] Headerlist) : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

            chkList.Items.Clear();

            int t;
            var loopTo = Headerlist.GetUpperBound(0);
            for (t = Headerlist.GetLowerBound(0); t <= loopTo; t++)
            {
                if (!(Headerlist[t] == null))
                    chkList.Items.Add(Headerlist[t], false);
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
        private System.ComponentModel.IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        private System.Windows.Forms.CheckedListBox _chkList;

        internal System.Windows.Forms.CheckedListBox chkList
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _chkList;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_chkList != null)
                {
                }

                _chkList = value;
                if (_chkList != null)
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

        private System.Windows.Forms.Button _btnAll;

        internal System.Windows.Forms.Button btnAll
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnAll;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnAll != null)
                {
                    _btnAll.Click -= btnAll_Click;
                }

                _btnAll = value;
                if (_btnAll != null)
                {
                    _btnAll.Click += btnAll_Click;
                }
            }
        }

        private System.Windows.Forms.Button _btnNone;

        internal System.Windows.Forms.Button btnNone
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _btnNone;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_btnNone != null)
                {
                    _btnNone.Click -= btnNone_Click;
                }

                _btnNone = value;
                if (_btnNone != null)
                {
                    _btnNone.Click += btnNone_Click;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            _chkList = new System.Windows.Forms.CheckedListBox();
            _btnOk = new System.Windows.Forms.Button();
            _btnOk.Click += btnOk_Click;
            _btnCancel = new System.Windows.Forms.Button();
            _btnCancel.Click += btnCancel_Click;
            _btnAll = new System.Windows.Forms.Button();
            _btnAll.Click += btnAll_Click;
            _btnNone = new System.Windows.Forms.Button();
            _btnNone.Click += btnNone_Click;
            SuspendLayout();
            // 
            // chkList
            // 
            _chkList.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom
                        | System.Windows.Forms.AnchorStyles.Left
                        | System.Windows.Forms.AnchorStyles.Right;
            _chkList.CheckOnClick = true;
            _chkList.Location = new System.Drawing.Point(4, 24);
            _chkList.Name = "chkList";
            _chkList.Size = new System.Drawing.Size(232, 304);
            _chkList.TabIndex = 0;
            // 
            // btnOk
            // 
            _btnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
            _btnOk.Location = new System.Drawing.Point(20, 335);
            _btnOk.Name = "btnOk";
            _btnOk.Size = new System.Drawing.Size(76, 24);
            _btnOk.TabIndex = 1;
            _btnOk.Text = "Ok";
            // 
            // btnCancel
            // 
            _btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            _btnCancel.Location = new System.Drawing.Point(136, 335);
            _btnCancel.Name = "btnCancel";
            _btnCancel.Size = new System.Drawing.Size(76, 24);
            _btnCancel.TabIndex = 2;
            _btnCancel.Text = "Cancel";
            // 
            // btnAll
            // 
            _btnAll.Location = new System.Drawing.Point(4, 0);
            _btnAll.Name = "btnAll";
            _btnAll.Size = new System.Drawing.Size(40, 20);
            _btnAll.TabIndex = 3;
            _btnAll.Text = "All";
            // 
            // btnNone
            // 
            _btnNone.Location = new System.Drawing.Point(48, 0);
            _btnNone.Name = "btnNone";
            _btnNone.Size = new System.Drawing.Size(40, 20);
            _btnNone.TabIndex = 4;
            _btnNone.Text = "None";
            // 
            // frmMultipleColumnTearAway
            // 
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            ClientSize = new System.Drawing.Size(240, 364);
            Controls.Add(_btnNone);
            Controls.Add(_btnAll);
            Controls.Add(_btnCancel);
            Controls.Add(_btnOk);
            Controls.Add(_chkList);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            MaximizeBox = false;
            MinimizeBox = false;
            MinimumSize = new System.Drawing.Size(212, 184);
            Name = "frmMultipleColumnTearAway";
            Text = "Select Columns to Tear Away";
            TopMost = true;
            ResumeLayout(false);
        }


        // 
        // Internal Variable Declarations
        // 

        private bool m_Canceled = true;

        // 
        // Propertys are defined here
        // 

        public bool Canceled
        {
            get
            {
                return m_Canceled;
            }
            set
            {
                m_Canceled = value;
            }
        }

        public System.Windows.Forms.CheckedListBox.CheckedIndexCollection SelectedIndices
        {
            get
            {
                return chkList.CheckedIndices;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            m_Canceled = false;
            Hide();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            m_Canceled = true;
            Hide();
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            int t;

            if (chkList.Items.Count == 0)
                return;
            var loopTo = chkList.Items.Count - 1;
            for (t = 0; t <= loopTo; t++)
                chkList.SetItemChecked(t, true);
        }

        private void btnNone_Click(object sender, EventArgs e)
        {
            int t;

            if (chkList.Items.Count == 0)
                return;
            var loopTo = chkList.Items.Count - 1;
            for (t = 0; t <= loopTo; t++)
                chkList.SetItemChecked(t, false);
        }
    }
}
