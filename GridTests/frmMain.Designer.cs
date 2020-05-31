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
            this.btnWQL = new System.Windows.Forms.Button();
            this.taig = new TAIGridControl2.TAIGridControl();
            this.btnSQLPopulate = new System.Windows.Forms.Button();
            this.btnBigSqlPopulate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnWQL
            // 
            this.btnWQL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnWQL.Location = new System.Drawing.Point(10, 371);
            this.btnWQL.Margin = new System.Windows.Forms.Padding(2);
            this.btnWQL.Name = "btnWQL";
            this.btnWQL.Size = new System.Drawing.Size(91, 34);
            this.btnWQL.TabIndex = 1;
            this.btnWQL.Text = "WQL Populate";
            this.btnWQL.UseVisualStyleBackColor = true;
            this.btnWQL.Click += new System.EventHandler(this.btnWQL_Click);
            // 
            // taig
            // 
            this.taig.AlternateColoration = false;
            this.taig.AlternateColorationAltColor = System.Drawing.Color.MediumSpringGreen;
            this.taig.AlternateColorationBaseColor = System.Drawing.Color.AntiqueWhite;
            this.taig.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.taig.BorderColor = System.Drawing.Color.Black;
            this.taig.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.taig.CellOutlines = true;
            this.taig.ColBackColorEdit = System.Drawing.Color.Yellow;
            this.taig.Cols = 0;
            this.taig.DefaultBackgroundColor = System.Drawing.Color.AntiqueWhite;
            this.taig.DefaultCellFont = new System.Drawing.Font("Arial", 9F);
            this.taig.DefaultForegroundColor = System.Drawing.Color.Black;
            this.taig.Delimiter = ",";
            this.taig.ExcelAlternateColoration = System.Drawing.Color.FromArgb(((int)(((byte)(204)))), ((int)(((byte)(255)))), ((int)(((byte)(204)))));
            this.taig.ExcelAutoFitColumn = true;
            this.taig.ExcelAutoFitRow = true;
            this.taig.ExcelFilename = "";
            this.taig.ExcelIncludeColumnHeaders = true;
            this.taig.ExcelKeepAlive = true;
            this.taig.ExcelMatchGridColorScheme = true;
            this.taig.ExcelMaximized = true;
            this.taig.ExcelMaxRowsPerSheet = 30000;
            this.taig.ExcelOutlineCells = true;
            this.taig.ExcelPageOrientation = 1;
            this.taig.ExcelShowBorders = false;
            this.taig.ExcelUseAlternateRowColor = true;
            this.taig.ExcelWorksheetName = "Grid Output";
            this.taig.GridEditMode = TAIGridControl2.TAIGridControl.GridEditModes.KeyReturn;
            this.taig.GridHeaderBackColor = System.Drawing.Color.LightBlue;
            this.taig.GridHeaderFont = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.taig.GridHeaderForeColor = System.Drawing.Color.Black;
            this.taig.GridHeaderHeight = 16;
            stringFormat1.Alignment = System.Drawing.StringAlignment.Near;
            stringFormat1.HotkeyPrefix = System.Drawing.Text.HotkeyPrefix.None;
            stringFormat1.LineAlignment = System.Drawing.StringAlignment.Near;
            stringFormat1.Trimming = System.Drawing.StringTrimming.Character;
            this.taig.GridHeaderStringFormat = stringFormat1;
            this.taig.GridheaderVisible = true;
            this.taig.Location = new System.Drawing.Point(9, 10);
            this.taig.Margin = new System.Windows.Forms.Padding(2);
            this.taig.Name = "taig";
            this.taig.OmitNulls = true;
            this.taig.PageSettings = null;
            this.taig.PaginationSize = 0;
            this.taig.Rows = 0;
            this.taig.ScrollInterval = 5;
            this.taig.SelectedColBackColor = System.Drawing.Color.MediumSlateBlue;
            this.taig.SelectedColForeColor = System.Drawing.Color.LightGray;
            this.taig.SelectedColumn = -1;
            this.taig.SelectedRow = -1;
            this.taig.SelectedRowBackColor = System.Drawing.Color.Blue;
            this.taig.SelectedRowForeColor = System.Drawing.Color.White;
            this.taig.SelectedRows = ((System.Collections.ArrayList)(resources.GetObject("taig.SelectedRows")));
            this.taig.Size = new System.Drawing.Size(740, 356);
            this.taig.TabIndex = 0;
            this.taig.TitleBackColor = System.Drawing.Color.Blue;
            this.taig.TitleFont = new System.Drawing.Font("Arial", 16F);
            this.taig.TitleForeColor = System.Drawing.Color.White;
            this.taig.TitleText = "Grid Title";
            this.taig.TitleVisible = true;
            this.taig.XMLDataSetName = "Grid_Output";
            this.taig.XMLFileName = "";
            this.taig.XMLIncludeSchema = false;
            this.taig.XMLNameSpace = "TAI_Grid_Ouptut";
            this.taig.XMLTableName = "Table";
            // 
            // btnSQLPopulate
            // 
            this.btnSQLPopulate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSQLPopulate.Location = new System.Drawing.Point(105, 371);
            this.btnSQLPopulate.Margin = new System.Windows.Forms.Padding(2);
            this.btnSQLPopulate.Name = "btnSQLPopulate";
            this.btnSQLPopulate.Size = new System.Drawing.Size(91, 34);
            this.btnSQLPopulate.TabIndex = 2;
            this.btnSQLPopulate.Text = "SQL Populate";
            this.btnSQLPopulate.UseVisualStyleBackColor = true;
            this.btnSQLPopulate.Click += new System.EventHandler(this.btnSQLPopulate_Click);
            // 
            // btnBigSqlPopulate
            // 
            this.btnBigSqlPopulate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnBigSqlPopulate.Location = new System.Drawing.Point(200, 371);
            this.btnBigSqlPopulate.Margin = new System.Windows.Forms.Padding(2);
            this.btnBigSqlPopulate.Name = "btnBigSqlPopulate";
            this.btnBigSqlPopulate.Size = new System.Drawing.Size(109, 34);
            this.btnBigSqlPopulate.TabIndex = 3;
            this.btnBigSqlPopulate.Text = "BIG SQL Populate";
            this.btnBigSqlPopulate.UseVisualStyleBackColor = true;
            this.btnBigSqlPopulate.Click += new System.EventHandler(this.btnBigSqlPopulate_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(758, 416);
            this.Controls.Add(this.btnBigSqlPopulate);
            this.Controls.Add(this.btnSQLPopulate);
            this.Controls.Add(this.btnWQL);
            this.Controls.Add(this.taig);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmMain";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private TAIGridControl2.TAIGridControl taig;
        private System.Windows.Forms.Button btnWQL;
        private System.Windows.Forms.Button btnSQLPopulate;
        private System.Windows.Forms.Button btnBigSqlPopulate;
    }
}

