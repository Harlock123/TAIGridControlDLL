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
    }
}
