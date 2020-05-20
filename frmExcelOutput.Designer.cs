namespace TAIGridControl2
{
    partial class frmExcelOutput
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this._Label2 = new System.Windows.Forms.Label();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.OFD = new System.Windows.Forms.OpenFileDialog();
            this.Label1 = new System.Windows.Forms.Label();
            this.chkOmitNulls = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // _Label2
            // 
            this._Label2.Location = new System.Drawing.Point(1, 69);
            this._Label2.Name = "_Label2";
            this._Label2.Size = new System.Drawing.Size(188, 16);
            this._Label2.TabIndex = 13;
            this._Label2.Text = "Save the resulting EXCEL (xlsx) file here";
            // 
            // txtTableName
            // 
            this.txtTableName.Location = new System.Drawing.Point(2, 21);
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(152, 20);
            this.txtTableName.TabIndex = 11;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(318, 5);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(56, 20);
            this.btnOk.TabIndex = 10;
            this.btnOk.Text = "Ok";
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(258, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(56, 20);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(350, 45);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(24, 20);
            this.btnSelectFile.TabIndex = 8;
            this.btnSelectFile.Text = "...";
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(2, 45);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(344, 20);
            this.txtFileName.TabIndex = 7;
            // 
            // OFD
            // 
            this.OFD.CheckFileExists = false;
            this.OFD.DefaultExt = "xlsx";
            this.OFD.Title = "Select file name to save the script to";
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(1, 5);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(173, 13);
            this.Label1.TabIndex = 12;
            this.Label1.Text = "Name of the resulting WorkSheet";
            // 
            // chkOmitNulls
            // 
            this.chkOmitNulls.AutoSize = true;
            this.chkOmitNulls.Checked = true;
            this.chkOmitNulls.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOmitNulls.Location = new System.Drawing.Point(161, 22);
            this.chkOmitNulls.Name = "chkOmitNulls";
            this.chkOmitNulls.Size = new System.Drawing.Size(81, 17);
            this.chkOmitNulls.TabIndex = 14;
            this.chkOmitNulls.Text = "Omit {Nulls}";
            this.chkOmitNulls.UseVisualStyleBackColor = true;
            // 
            // frmExcelOutput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(389, 88);
            this.ControlBox = false;
            this.Controls.Add(this.chkOmitNulls);
            this.Controls.Add(this._Label2);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.txtTableName);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtFileName);
            this.Name = "frmExcelOutput";
            this.Text = "Export Grid contents to Excel (xlsx)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label _Label2;
        private System.Windows.Forms.TextBox txtTableName;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.OpenFileDialog OFD;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.CheckBox chkOmitNulls;
    }
}