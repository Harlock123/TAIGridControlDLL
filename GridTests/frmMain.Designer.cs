namespace GridTests
{
    partial class frmMain
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
            System.Drawing.StringFormat stringFormat1 = new System.Drawing.StringFormat();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.taiGridControl1 = new TAIGridControl2.TAIGridControl();
            this.SuspendLayout();
            // 
            // taiGridControl1
            // 
            this.taiGridControl1.AlternateColoration = false;
            this.taiGridControl1.AlternateColorationAltColor = System.Drawing.Color.MediumSpringGreen;
            this.taiGridControl1.AlternateColorationBaseColor = System.Drawing.Color.AntiqueWhite;
            this.taiGridControl1.BorderColor = System.Drawing.Color.Black;
            this.taiGridControl1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.taiGridControl1.CellOutlines = true;
            this.taiGridControl1.ColBackColorEdit = System.Drawing.Color.Yellow;
            this.taiGridControl1.Cols = 0;
            this.taiGridControl1.DefaultBackgroundColor = System.Drawing.Color.AntiqueWhite;
            this.taiGridControl1.DefaultCellFont = new System.Drawing.Font("Arial", 9F);
            this.taiGridControl1.DefaultForegroundColor = System.Drawing.Color.Black;
            this.taiGridControl1.Delimiter = ",";
            this.taiGridControl1.ExcelAlternateColoration = System.Drawing.Color.FromArgb(((int)(((byte)(204)))), ((int)(((byte)(255)))), ((int)(((byte)(204)))));
            this.taiGridControl1.ExcelAutoFitColumn = true;
            this.taiGridControl1.ExcelAutoFitRow = true;
            this.taiGridControl1.ExcelFilename = "";
            this.taiGridControl1.ExcelIncludeColumnHeaders = true;
            this.taiGridControl1.ExcelKeepAlive = true;
            this.taiGridControl1.ExcelMatchGridColorScheme = true;
            this.taiGridControl1.ExcelMaximized = true;
            this.taiGridControl1.ExcelMaxRowsPerSheet = 30000;
            this.taiGridControl1.ExcelOutlineCells = true;
            this.taiGridControl1.ExcelPageOrientation = 1;
            this.taiGridControl1.ExcelShowBorders = false;
            this.taiGridControl1.ExcelUseAlternateRowColor = true;
            this.taiGridControl1.ExcelWorksheetName = "Grid Output";
            this.taiGridControl1.GridEditMode = TAIGridControl2.TAIGridControl.GridEditModes.KeyReturn;
            this.taiGridControl1.GridHeaderBackColor = System.Drawing.Color.LightBlue;
            this.taiGridControl1.GridHeaderFont = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.taiGridControl1.GridHeaderForeColor = System.Drawing.Color.Black;
            this.taiGridControl1.GridHeaderHeight = 16;
            stringFormat1.Alignment = System.Drawing.StringAlignment.Near;
            stringFormat1.HotkeyPrefix = System.Drawing.Text.HotkeyPrefix.None;
            stringFormat1.LineAlignment = System.Drawing.StringAlignment.Near;
            stringFormat1.Trimming = System.Drawing.StringTrimming.Character;
            this.taiGridControl1.GridHeaderStringFormat = stringFormat1;
            this.taiGridControl1.GridheaderVisible = true;
            this.taiGridControl1.Location = new System.Drawing.Point(12, 12);
            this.taiGridControl1.Name = "taiGridControl1";
            this.taiGridControl1.PageSettings = null;
            this.taiGridControl1.PaginationSize = 0;
            this.taiGridControl1.Rows = 0;
            this.taiGridControl1.ScrollInterval = 5;
            this.taiGridControl1.SelectedColBackColor = System.Drawing.Color.MediumSlateBlue;
            this.taiGridControl1.SelectedColForeColor = System.Drawing.Color.LightGray;
            this.taiGridControl1.SelectedColumn = -1;
            this.taiGridControl1.SelectedRow = -1;
            this.taiGridControl1.SelectedRowBackColor = System.Drawing.Color.Blue;
            this.taiGridControl1.SelectedRowForeColor = System.Drawing.Color.White;
            this.taiGridControl1.SelectedRows = ((System.Collections.ArrayList)(resources.GetObject("taiGridControl1.SelectedRows")));
            this.taiGridControl1.Size = new System.Drawing.Size(986, 438);
            this.taiGridControl1.TabIndex = 0;
            this.taiGridControl1.TitleBackColor = System.Drawing.Color.Blue;
            this.taiGridControl1.TitleFont = new System.Drawing.Font("Arial", 16F);
            this.taiGridControl1.TitleForeColor = System.Drawing.Color.White;
            this.taiGridControl1.TitleText = "Grid Title";
            this.taiGridControl1.TitleVisible = true;
            this.taiGridControl1.XMLDataSetName = "Grid_Output";
            this.taiGridControl1.XMLFileName = "";
            this.taiGridControl1.XMLIncludeSchema = false;
            this.taiGridControl1.XMLNameSpace = "TAI_Grid_Ouptut";
            this.taiGridControl1.XMLTableName = "Table";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1010, 512);
            this.Controls.Add(this.taiGridControl1);
            this.Name = "frmMain";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private TAIGridControl2.TAIGridControl taiGridControl1;
    }
}

