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

        private void HandleGridColorChange(object sender, int row, int col, Color TheColorChangedTo)
        {
            MessageBox.Show("Grids colors were context menu changed\n " +
                            "The Row selected was " + row.ToString() + "\n" +
                            "The color changed to was R-" + TheColorChangedTo.R.ToString() + " G-" +
                            TheColorChangedTo.G.ToString() + " B-" + TheColorChangedTo.B.ToString());
        }

        private void btnAzureSQLTest_Click(object sender, EventArgs e)
        {
			string cn = "Put AZURE SQL connection string here";

			var thestring = "";
			thestring += "select RecentDas, ";
			thestring += "	[Note Date], ";
			thestring += "	[FinalDate], ";
			thestring += "	ID, ";
			thestring += "	[Is Note Void ?], ";
			thestring += "	[Originator], ";
			thestring += "	Type, ";
			thestring += "	NoteDivision, ";
			thestring += "	[Type of Contact], ";
			thestring += "	Note, ";
			thestring += "	NoteType, ";
			thestring += "	Locked, ";
			thestring += "	[OriginatorID], ";
			thestring += "	[CaseManagerID], ";
			thestring += "	[SupervisorID], ";
			thestring += "	[OriginatorManagerID], ";
			thestring += "	[Currentuserid], ";
			thestring += "	[CreatedDate], ";
			thestring += "	LastNoteViewed from(SELECT d.dasTimeIn as [TIME], ROW_NUMBER() OVER(PARTITION by cn.ID  ";
			thestring += "		order by d.dasTimeIn DESC ";
			thestring += "		) RecentDas, cn.ID as ID, Convert(varchar(40), d.dasTimeIn, 22) as [Note Date], cn.NoteStatus as [Is Note Void ?], cn.NoteDate, Convert(varchar(40), cn.NoteDate, 22) as [ThisDate],  ";
			thestring += "		case  ";
			thestring += "				when Convert(varchar(40), d.dasTimeIn, 22) <> '' and Convert(varchar(40), d.dasTimeIn, 22) is not NULL then Convert(varchar(40), d.dasTimeIn, 22)  ";
			thestring += "				else Convert(varchar(40), cn.NoteDate, 22)  ";
			thestring += "			end as [FinalDate], u.UserName as Originator, lcnt.Description as Type, cn.NoteType as NoteType, NoteDivision, lcnTitle.TitleDescription as [Type of Contact], cn.NoteText as Note, cn.NoteLock as Locked, cn.memberid as mmID, cn.NoteType as nTYPE, Coalesce(u.tblusersID, 0) as[OriginatorID], Coalesce( ";
			thestring += "			( ";
			thestring += "			select tblUsersID  ";
			thestring += "			from tblUsers  ";
			thestring += "			where cn.CaseManagerID = tblUsersID ";
			thestring += "			), 0) AS[CaseManagerID], Coalesce( ";
			thestring += "			( ";
			thestring += "			select DirectSupervisor  ";
			thestring += "			from tblUsers  ";
			thestring += "			where tblUsersID = cn.CaseManagerID ";
			thestring += "			), 0) AS[SupervisorID], Coalesce( ";
			thestring += "			( ";
			thestring += "			select DirectSupervisor  ";
			thestring += "			from tblUsers  ";
			thestring += "			where cn.NoteOriginator = tblUsersID ";
			thestring += "			), 0) as [OriginatorManagerID],  ";
			thestring += "		( ";
			thestring += "		select tblUsersID  ";
			thestring += "		from tblUsers  ";
			thestring += "		where UserName = 'harlo' ";
			thestring += "		) as [Currentuserid], cn.CreatedBy, Convert(varchar(40), cn.CreatedDate, 22) as [CreatedDate], cn.LastNoteViewed  ";
			thestring += "	from ClinicalNotes cn  ";
			thestring += "		left outer join tblUsers u on cn.NoteOriginator = u.tblUsersID  ";
			thestring += "		left outer join tblLookupClinicalNoteType lcnt on cn.NoteType = lcnt.ID  ";
			thestring += "		left outer join tblLookupClinicalNotesTitle lcnTitle on cn.NoteTitle = lcnTitle.ID  ";
			thestring += "		left outer join DAS d on d.noteID = cn.id  ";
			thestring += "	where cn.Memberid = 6313232 ";
			thestring += "	) t  ";
			thestring += "where 1 = 1  ";
			thestring += "	and [Is Note Void ?] <> 1  ";
			thestring += "	and RecentDas = 1  ";
			thestring += "order by Cast([FinalDate] as datetime) DESC, NoteType asc ";

			taig.PopulateGridWithData(cn, thestring);

			taig.WordWrapColumn(taig.GetColumnIDByName("Note"), 50);
			taig.set_ColWidth(taig.GetColumnIDByName("Note"), 300);
		}

	}
}
