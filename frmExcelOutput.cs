using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TAIGridControl2
{
    public partial class frmExcelOutput : Form
    {

        TAIGridControl _taig;

        public string SELECTEDWORKBOOKNAME = "";
        public string SELECTEDPATH = "";
        public bool FRMOK = false;
        public bool OMITNULLS = true;
        public bool OPENWHENSAVED = true;

        public frmExcelOutput()
        {
            InitializeComponent();

            txtTableName.Text = "GRID OUTPUT";
            txtFileName.Text = System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EXCELOUTPUT.xlsx");

        }

        public frmExcelOutput(TAIGridControl TAIG)
        {
            InitializeComponent();
            _taig = TAIG;

            var fname = "EXCELOUTPUT_" + DateTime.Now.ToString("MMddyyyy_HHmmss") + ".xlsx";

            if (TAIG.ExcelWorksheetName.ToUpper() == "GRID OUTPUT")
                txtTableName.Text = TAIG.TitleText;
            else
                txtTableName.Text = TAIG.ExcelWorksheetName;

            txtFileName.Text = System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fname);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text) || string.IsNullOrEmpty(txtTableName.Text))
                MessageBox.Show("You must select a file for the resulting Excel Document to be written to\n" +
                                "as well as a name for the resulting Worksheet name that is crafted in the Excel document", "Export to Excel error", MessageBoxButtons.OK, MessageBoxIcon.Information);

            else
            {
                SELECTEDPATH = txtFileName.Text;
                SELECTEDWORKBOOKNAME = txtTableName.Text.Replace(":","_")
                    .Replace(@"\","_")
                    .Replace(@"/","_")
                    .Replace("?","_")
                    .Replace("*","_")
                    .Replace("[","_")
                    .Replace("]","_");
                FRMOK = true;
                OMITNULLS = chkOmitNulls.Checked;
                OPENWHENSAVED = chkOpenFileWhenSaved.Checked;
                Hide();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            SELECTEDPATH = "";
            SELECTEDWORKBOOKNAME = "";
            FRMOK = false;
            OMITNULLS = false;
            OPENWHENSAVED = false;
            Hide();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OFD.CheckFileExists = false;
            if ((int)OFD.ShowDialog() == (int)System.Windows.Forms.DialogResult.OK)
            {
                txtFileName.Text = OFD.FileName;
                if (!txtFileName.Text.ToUpper().Trim().EndsWith(".XLSX"))
                {
                    txtFileName.Text += ".xlsx";
                }
            }
        }
    }
}
