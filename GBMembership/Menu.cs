using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;           // access to drawing tools
using System.Windows.Forms;     // access to forms controls
using System.Data;              // to access dataset
using System.Data.OleDb;        // to access database


namespace GBMembership
{
    public partial class Form1
    {
        // button placement variables
        public int MemberButtonSub = 0;
        public int xx = 20, x_int = 80, x_max = 700, x_Left = 20;
        public int yy = 40, y_int = 80, y_top = 402;

        //public void Form1_Load(object sender, EventArgs e)
        //{
        //    Form1_Load();
        //}
        public void Form1_Load()
            { 
            DataSet dtSet = new DataSet();          // to hold list of members

            // populate list of members
            Utils MyConn = new Utils();
            string mySelectQuery = "SELECT Id, FirstN, Inits, LastN, Leader FROM Member";

            OleDbConnection myConnection = new OleDbConnection(MyConn.myConnectionString);
            OleDbDataAdapter myCmd = new OleDbDataAdapter(mySelectQuery, myConnection);
            System.Data.DataTable dTable = null;
            try
            {
                myConnection.Open();
                myCmd.Fill(dtSet, "Member");
                dTable = dtSet.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Get Person details failed");
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (myConnection != null)
                    myConnection.Close();
            }

            // place buttons for Members on Form1:  Max of 50 Members + controls
            Button[] MemberButtons = new Button[50];
            // initialise variables for this build
            MemberButtonSub = 0;
            xx = 20;
            yy = 40;

            // build a button for each member
            foreach (DataRow dtRow in dTable.Rows)
            {
                MemberButtons[MemberButtonSub] = new Button();
                MemberButtons[MemberButtonSub].Name = dtRow["Id"].ToString();
                MemberButtons[MemberButtonSub].Text = dtRow["FirstN"] +
                    Environment.NewLine + dtRow["LastN"];
                if (dtRow["Leader"].ToString() == "0")
                {
                    MemberButtons[MemberButtonSub].BackColor = Color.CadetBlue;
                }
                else
                {
                    MemberButtons[MemberButtonSub].BackColor = Color.Salmon;
                }
                MemberButtons[MemberButtonSub].Size = new Size(70, 70);
                MemberButtons[MemberButtonSub].Location = new Point(xx, yy);
                MemberButtons[MemberButtonSub].Click += new EventHandler(MemberButton_Click);
                this.Controls.Add(MemberButtons[MemberButtonSub]);
                nextLocation();
            }

            xx = 701;           // Force Standard buttons onto their own line
            nextLocation();
            #region add new member button
            MemberButtons[MemberButtonSub] = new Button();      // add button to add new member
            MemberButtons[MemberButtonSub].Name = "0";
            MemberButtons[MemberButtonSub].Text = "New" + Environment.NewLine + "Member";
            MemberButtons[MemberButtonSub].Size = new Size(70, 70);
            MemberButtons[MemberButtonSub].BackColor = Color.LightGreen;
            MemberButtons[MemberButtonSub].Location = new Point(xx, yy);
            MemberButtons[MemberButtonSub].Click += new EventHandler(MemberButton_Click);
            this.Controls.Add(MemberButtons[MemberButtonSub]);
            #endregion add new member
            #region Refresh screen button
            nextLocation();
            MemberButtons[MemberButtonSub] = new Button();      // add button to refresh members form
            MemberButtons[MemberButtonSub].Name = "R";
            MemberButtons[MemberButtonSub].Text = "Refresh" + Environment.NewLine + "Members";
            MemberButtons[MemberButtonSub].Size = new Size(70, 70);
            MemberButtons[MemberButtonSub].BackColor = Color.LightGreen;
            MemberButtons[MemberButtonSub].Location = new Point(xx, yy);
            MemberButtons[MemberButtonSub].Click += new EventHandler(RefreshButton_Click);
            this.Controls.Add(MemberButtons[MemberButtonSub]);
            #endregion refresh
            #region Print All Button
            nextLocation();
            MemberButtons[MemberButtonSub] = new Button();      // add button to Print all Members
            MemberButtons[MemberButtonSub].Name = "P";
            MemberButtons[MemberButtonSub].Text = "Print All" + Environment.NewLine + "Details";
            MemberButtons[MemberButtonSub].Size = new Size(70, 70);
            MemberButtons[MemberButtonSub].BackColor = Color.LightGreen;
            MemberButtons[MemberButtonSub].Location = new Point(xx, yy);
            MemberButtons[MemberButtonSub].Click += new EventHandler(PrintAllButton_Click);
            this.Controls.Add(MemberButtons[MemberButtonSub]);
            #endregion Print All
            #region Quit Button
            nextLocation();
            MemberButtons[MemberButtonSub] = new Button();      // add button to quit application
            MemberButtons[MemberButtonSub].Name = "Q";
            MemberButtons[MemberButtonSub].Text = "Quit" + Environment.NewLine + "Application";
            MemberButtons[MemberButtonSub].Size = new Size(70, 70);
            MemberButtons[MemberButtonSub].BackColor = Color.LightGreen;
            MemberButtons[MemberButtonSub].Location = new Point(xx, yy);
            MemberButtons[MemberButtonSub].Click += new EventHandler(QuitButton_Click);
            this.Controls.Add(MemberButtons[MemberButtonSub]);
            #endregion Quit
        }

        // update button location values
        public void nextLocation()
        {
            xx += x_int;
            if (xx > x_max)
            {
                xx = x_Left;
                yy += y_int;
            }
            MemberButtonSub++;
        }

        #region Button Click events
        // Call Member details screen based on button number, 0 = add new member
        public void MemberButton_Click(object sender, System.EventArgs e)       // member button clicked
        {                                                                       // identify and display
            Button mb = (Button)sender;
            MemberForm MemberDetails = new MemberForm(mb.Name);
            MemberDetails.Text = "Member form";
            MemberDetails.ShowDialog();
            this.Controls.Clear();
            this.InitializeComponent();
            this.Form1_Load(sender, e); this.Controls.Clear();
            this.InitializeComponent();
            this.Form1_Load(sender, e);
        }
        public void RefreshButton_Click(object sender, System.EventArgs e)      // refresh members form
        {
            this.Controls.Clear();
            this.InitializeComponent();
            this.Form1_Load(sender, e);
        }
        public void PrintAllButton_Click(object sender, EventArgs e)            // print all the records (Mailmerge)
        {
            System.Diagnostics.Process.Start(@"C:\ProgramData\GBRecords\GBMembersMailMergeBody.docx".ToString());
        }
        public void QuitButton_Click(object sender, System.EventArgs e)         // quit application
        {
            System.Windows.Forms.Application.Exit();
        }
        #endregion Button Click events
    }

    public class Utils
    {
        public string myConnectionString
        {
            get
            {
                return "Provider=Microsoft.Jet.OLEDB.4.0;" + "User ID=;Password =;" +
                    @"Data Source=C:\ProgramData\GBRecords\GBMembers.mdb";
            }
        }
    }
}
