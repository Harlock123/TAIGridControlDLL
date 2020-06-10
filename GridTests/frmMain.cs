using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GridTests
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnWQL_Click(object sender, EventArgs e)
        {
            string wql = "SELECT * FROM Win32_Printer ";

            taig.PopulateFromWQL(wql);
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

        }

        private void btnSQLPopulate_Click(object sender, EventArgs e)
        {
            //string cn = "Server=(local);Database=HIDATA;Trusted_Connection=True;"; // My test env has this running locally

            string cn = "Server=(local);User ID=sa;password=P@ssw0rd;Database=HIDATA;"; // My test env in Windows VM in Linux has SQL on a docker container


            string sql = "SELECT TOP 1000 * from tblMEMBERMAIN";

            taig.PopulateGridWithData(cn, sql);
        }

        private void btnBigSqlPopulate_Click(object sender, EventArgs e)
        {
            //string cn = "Server=(local);Database=HIDATA;Trusted_Connection=True;"; // My test env has this running locally

            string cn = "Server=(local);User ID=sa;password=P@ssw0rd;Database=HIDATA;"; // My test env in Windows VM in Linux has SQL on a docker container


            string sql = "SELECT * from tblMEMBERMAIN"; // In my test env it shouild grab over 14000 rows

            taig.PopulateGridWithData(cn, sql);
        }
    }
}
