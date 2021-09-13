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

        private void btnSmallSQL_Click(object sender, EventArgs e)
        {
            //string cn = "Server=(local);Database=HIDATA;Trusted_Connection=True;"; // My test env has this running locally

            string cn = "Server=(local);User ID=sa;password=P@ssw0rd;Database=HIDATA;"; // My test env in Windows VM in Linux has SQL on a docker container


            string sql = "SELECT TOP 1 MMID,FIRSTNAME,MIDDLENAME,LASTNAME,DOB," +
                         "(SELECT TOP 1 DESCRIPTION FROM TBLLOOKUPGENDER B where B.CODE=A.GENDER) as 'GENDER'," +
                         "SSN,CREATEDDATE,CREATEDBY,UPDATEDBY,UPDATEDDATE,Phone1,Phone1Type,Phone1Ext," +
                         "Phone2,Phone2Type,Phone2Ext,Email,ParentGuardian,ParentGuardPhone,CLIENTID,STATEMODDATE," +
                         "STATEMODUSER,STATEMODTIME,NOMSID,LOCATIONDATE from tblMEMBERMAIN A";

            taig.PopulateGridWithData(cn, sql);
        }

        private void btnDirPop_Click(object sender, EventArgs e)
        {
            taig.PopulateFromADirectory(@"C:\Windows");
            taig.WordWrapColumn(0, 30);
        }
    }
}
