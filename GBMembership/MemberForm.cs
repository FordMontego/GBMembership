using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;        // to access database
// add reference microsoft office 16.0 object library to get microsoft.office.core
// add reference Microsoft word 16.0 Object Library for Microsoft.office.interop.word
using Microsoft.Office.Interop.Word;
// used in print member
using Mword = Microsoft.Office.Interop.Word;

namespace GBMembership
{
    public partial class MemberForm : Form
    {
        public Member MData = new Member();
        bool screenLoaded = false;

        public MemberForm()
        {
            InitializeComponent();
        }
        public MemberForm(string mb)
        {
            InitializeComponent();
            if (mb == "0")
                InitializeNewComponent();
            else
                InitializeComponent(mb);
        }
        public void InitializeNewComponent()
        {
            screenLoaded = false;
            MData.PD = new Person[4];

            // get next id number
            #region read database
            Utils MyConn = new Utils();
            string myQuery = "SELECT MAX(Id)+1 FROM member;";
            OleDbConnection myConnection = new OleDbConnection(MyConn.myConnectionString);
            OleDbCommand myCommand = new OleDbCommand(myQuery, myConnection);
            try
            {
                myConnection.Open();
                OleDbDataReader myReader = null;
                myReader = myCommand.ExecuteReader();
                myReader.Read();
                MData.IdNbr = myReader.GetInt32(0).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Get next member number failed" + ex.Message);
                MessageBox.Show(ex.Source);
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (myConnection != null)
                    myConnection.Close();
            }
            #endregion read database

            MemNo.Text = "New Id: " + MData.IdNbr;
            dtppSigned.Text = "01/01/2001";
            dtp1Signed.Text = "01/01/2001";
            dtp2Signed.Text = "01/01/2001";
            screenLoaded = true;
        }

        //Populate fields with details
        public void InitializeComponent(string mb)
        {
            screenLoaded = false;
            MData.PD = new Person[4];

            #region Members details
            GetMemberDetails(mb);
            MemNo.Text = "Id: " + mb;
            if (MData.Leader == "1")
                cbLeader.Checked = true;
            else
                cbLeader.Checked = false;
            tbTitle.Text = MData.Title;
            tbName.Text = MData.First;
            tbLast.Text = MData.Last;
            tbAddress.Text = MData.Address;
            tbPostcode.Text = MData.PostCode;
            dtpDOB.Text = MData.DOB;
            tbsigned.Text = MData.SignedBy;
            dtpSigned.Text = MData.SignedOn;
            if (MData.photo == "1")
                cbPhoto.Checked = true;
            else
                cbPhoto.Checked = false;
            tbaSigned.Text = MData.PhotoSigned;
            dtpaSigned.Text = MData.PhotoOn;
            tbMedical.Text = MData.Conditions;
            tbNotes.Text = MData.Notes;
            if (MData.MedicalPermission == "1")
                cbMed.Checked = true;
            else
                cbMed.Checked = false;
            tbMedicalSign.Text = MData.MedicalSigned;
            dtpMedicalSign.Text = MData.MedicalOn;
            tbaSigned.Text = MData.AuthorisedBy;
            dtpaSigned.Text = MData.AuthorisedOn;
            tbGP.Text = MData.Surgery;
            if(MData.StorePhone=="1")
                cbpPermission.Checked = true;
            else
                cbpPermission.Checked = false;


            #endregion Members details
            #region Parent/guardian details
            Get_Person_Details(mb, 1);
            tbpTitle.Text = MData.PD[0].Title;
            tbpName.Text = MData.PD[0].First;
            tbpSurname.Text = MData.PD[0].Last;
            tbpAddress.Text = MData.PD[0].Address;
            tbpPostcode.Text = MData.PD[0].PostCode;
            tbpPhone.Text = MData.PD[0].Telephone;
            tbpMobile.Text = MData.PD[0].Mobile;
            tbpEmail.Text = MData.PD[0].email;
            if (MData.PD[0].StorePhone == "1")
                cbpPermission.Checked = true;
            else
                cbpPermission.Checked = false;
            tbpSigned.Text = MData.PD[0].SignedBy;
            dtppSigned.Text = MData.PD[0].DateSigned;
            #endregion Parent/guardian details
            #region emergency 1
            Get_Person_Details(mb, 2);
            tb1Title.Text = MData.PD[1].Title;
            tb1Name.Text = MData.PD[1].First;
            tb1Surname.Text = MData.PD[1].Last;
            tb1Address.Text = MData.PD[1].Address;
            tb1Postcode.Text = MData.PD[1].PostCode;
            tb1Phone.Text = MData.PD[1].Telephone;
            tb1Mobile.Text = MData.PD[1].Mobile;
            tb1Email.Text = MData.PD[1].email;
            if (MData.PD[1].StorePhone == "1")
                cbpPermission.Checked = true;
            else
                cbpPermission.Checked = false;
            tb1Signed.Text = MData.PD[1].SignedBy;
            dtp1Signed.Text = MData.PD[1].DateSigned;
            tb1Relation.Text = MData.PD[1].Relation;
            #endregion emergency 1
            #region emergency 2
            Get_Person_Details(mb, 3);
            tb2Title.Text = MData.PD[2].Title;
            tb2Name.Text = MData.PD[2].First;
            tb2Surname.Text = MData.PD[2].Last;
            tb2Address.Text = MData.PD[2].Address;
            tb2Postcode.Text = MData.PD[2].PostCode;
            tb2Phone.Text = MData.PD[2].Telephone;
            tb2Mobile.Text = MData.PD[2].Mobile;
            tb2Email.Text = MData.PD[2].email;
            if (MData.PD[2].StorePhone == "1")
                cbpPermission.Checked = true;
            else
                cbpPermission.Checked = false;
            tb2Signed.Text = MData.PD[2].SignedBy;
            dtp2Signed.Text = MData.PD[2].DateSigned;
            tb2Relation.Text = MData.PD[2].Relation;
            #endregion emergency 2
            screenLoaded = true;
        }

        public struct Member
        {
            public string IdNbr,
                Title,
                First,
                Inits,
                Last,
                DOB,
                Leader,
                SignedOn,
                SignedBy,
                Address,
                PostCode,
                photo,
                PhotoSigned,
                PhotoOn,
                Conditions,
                Notes,
                MedicalPermission,
                MedicalSigned,
                MedicalOn,
                AuthorisedBy,
                AuthorisedOn,
                StorePhone,
                Surgery;

            public Person[] PD;
        }
        public struct Person
        {
            // Contact Details
            //  1 = Parent/guardian
            //  2 = Emergency 1
            //  3 = Emergency 2
            public string Title;
            public string First;
            public string Inits;
            public string Last;
            public string Address;
            public string PostCode;
            public string Telephone;
            public string Mobile;
            public string email;
            public string StorePhone;
            public string SignedBy;
            public string DateSigned;
            public string Relation;
        }

        public void GetMemberDetails(string MId)
        {
            #region read database : member
            DataSet dtSet = new DataSet();


            // string myConnectionString = ////
            Utils MyConn = new Utils();

            string mySelectQuery = "SELECT Id, Title, FirstN, Inits, LastN, DOB, Leader, SignedOn, " +
                "SignedBy, Address, PostCode, PhotoConsent, PhotoSigned, PhotoOn, MedicalConditions, Notes,  " +
                "MedicalPermission, MedicalSigned, MedicalOn, AuthorisedBy, AuthorisedOn, Surgery, StorePhone " +
                "FROM member WHERE Id=" + MId;

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
                MessageBox.Show("Get Member details for " + MId.ToString() + " failed");
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (myConnection != null)
                    myConnection.Close();
            }
            #endregion read database : member
            foreach (DataRow dtRow in dTable.Rows)      // unpack member record
            {
                MData.IdNbr = dtRow["Id"].ToString();
                MData.Title = dtRow["Title"].ToString();
                MData.First = dtRow["FirstN"].ToString();
                MData.Inits = dtRow["Inits"].ToString();
                MData.Last = dtRow["LastN"].ToString();
                MData.DOB = dtRow["DOB"].ToString();
                MData.Leader = dtRow["Leader"].ToString();
                MData.SignedOn = dtRow["SignedOn"].ToString();
                MData.SignedBy = dtRow["SignedBy"].ToString();
                MData.Address = dtRow["Address"].ToString();
                MData.PostCode = dtRow["PostCode"].ToString();
                MData.photo = dtRow["PhotoConsent"].ToString();
                MData.PhotoSigned = dtRow["PhotoSigned"].ToString();
                MData.PhotoOn = dtRow["PhotoOn"].ToString();
                MData.Conditions = dtRow["MedicalConditions"].ToString();
                MData.Notes = dtRow["Notes"].ToString();
                MData.MedicalPermission = dtRow["MedicalPermission"].ToString();
                MData.MedicalSigned = dtRow["MedicalSigned"].ToString();
                MData.MedicalOn = dtRow["MedicalOn"].ToString();
                MData.AuthorisedOn = dtRow["AuthorisedOn"].ToString();
                MData.AuthorisedBy = dtRow["AuthorisedBy"].ToString();
                MData.Surgery = dtRow["Surgery"].ToString();
                MData.StorePhone = dtRow["StorePhone"].ToString();
            }       // unpack

        }
        public void Get_Person_Details(string MId, int PType)
        {
            #region Read database : person
            DataSet dtSet = new DataSet();

            Utils MyConn = new Utils();

            string mySelectQuery = "SELECT Id, Type,Title, FirstN, Inits, LastN, Address, PostCode, " +
                "Telephone, Mobile, email, StorePhone, SignedBy, DateSigned, Relation " +
                "FROM Person WHERE Id=" + MId.ToString() + " AND Type=" + PType.ToString();

            OleDbConnection myConnection = new OleDbConnection(MyConn.myConnectionString);
            OleDbDataAdapter myCmd = new OleDbDataAdapter(mySelectQuery, myConnection);
            System.Data.DataTable dTable = null;
            try
            {
                myConnection.Open();
                myCmd.Fill(dtSet, "Person");
                dTable = dtSet.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Get Person details for " + MId.ToString() + "/" + PType.ToString() + " failed");
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (myConnection != null)
                    myConnection.Close();
            }
            myConnection.Close();
            #endregion read db person

            int sub = PType - 1;                        // types 1-4 map to subscripts 0-3
                                                        //  1 = Parent/guardian
                                                        //  2 = Emergency 1                                                        
                                                        //  3 = Emergency 2
            foreach (DataRow dtRow in dTable.Rows)      // unpack each person record
            {
                MData.PD[sub].Title = dtRow["Title"].ToString();
                MData.PD[sub].First = dtRow["FirstN"].ToString();
                MData.PD[sub].Inits = dtRow["Inits"].ToString();
                MData.PD[sub].Last = dtRow["LastN"].ToString();
                MData.PD[sub].Address = dtRow["Address"].ToString();
                MData.PD[sub].PostCode = dtRow["PostCode"].ToString();
                MData.PD[sub].Telephone = dtRow["Telephone"].ToString();
                MData.PD[sub].Mobile = dtRow["Mobile"].ToString();
                MData.PD[sub].email = dtRow["email"].ToString();
                MData.PD[sub].StorePhone = dtRow["StorePhone"].ToString();
                MData.PD[sub].SignedBy = dtRow["SignedBy"].ToString();
                MData.PD[sub].DateSigned = dtRow["DateSigned"].ToString();
                MData.PD[sub].Relation = dtRow["Relation"].ToString();
            }
        }
        public void DataChanged(object sender, EventArgs e) // change detected, hide close and unhide cancel and save
        {
            if (screenLoaded == true)
            {
                this.btnClose.Visible = false;
                this.btnClose.Enabled = false;
                this.btnCancel.Visible = true;
                this.btnCancel.Enabled = true;
                this.btnSave.Visible = true;
                this.btnSave.Enabled = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Click 'OK' to lose the changes, otherwise click 'Cancel' to return to form", "Confirm", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                Close();
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            object oTemplate = @"C:\ProgramData\GBRecords\GBMemberTemplate.dotx";

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Mword._Application oWord;
            Mword._Document oDoc;
            oWord = new Mword.Application();
            oWord.Visible = true;
            //oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing);
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
                ref oMissing, ref oMissing);
            object oBookMark;
            string YesNo;
            oBookMark = "MTitle"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Title;
            oBookMark = "MFirst"; oDoc.Bookmarks[oBookMark].Range.Text = MData.First;
            oBookMark = "MLast"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Last;
            oBookMark = "MAddress"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Address;
            oBookMark = "MPostCode"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PostCode;
            oBookMark = "MDOB"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.DOB).ToString("d");
            oBookMark = "MSignedBy"; oDoc.Bookmarks[oBookMark].Range.Text = MData.SignedBy;
            oBookMark = "MSignedOn"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.SignedOn).ToString("d");
            if (MData.photo == "1") { YesNo = "Yes"; } else { YesNo = "No"; }
            oBookMark = "MPhoto"; oDoc.Bookmarks[oBookMark].Range.Text = YesNo;
            //oBookMark = "MPhotoSigned"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PhotoSigned;    // not on form
            //oBookMark = "MPhotoOn"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.PhotoOn).ToString("d");    // not on form
            oBookMark = "MConditions"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Conditions;
            oBookMark = "MNotes"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Notes;
            oBookMark = "MMedicalSigned"; oDoc.Bookmarks[oBookMark].Range.Text = MData.MedicalSigned;
            oBookMark = "MMedicalOn"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.MedicalOn).ToString("d");
            oBookMark = "MAuthorisedBy"; oDoc.Bookmarks[oBookMark].Range.Text = MData.AuthorisedBy;
            oBookMark = "MAuthorisedOn"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.AuthorisedOn).ToString("d");
            oBookMark = "MSurgery"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Surgery;
            //oBookMark = "MLeader"; oDoc.Bookmarks[oBookMark].Range.Text = MData.Leader; // Not on form
            if(MData.MedicalPermission=="1") { YesNo = "Yes I"; } else { YesNo = "No I Do NOT"; }
            oBookMark = "MMedicalPermission"; oDoc.Bookmarks[oBookMark].Range.Text = YesNo;
            if (MData.StorePhone == "1") { YesNo = "Yes"; } else { YesNo = "No"; }
            oBookMark = "MStorePhone";oDoc.Bookmarks[oBookMark].Range.Text = YesNo;

            oBookMark = "P0Title"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].Title;
            oBookMark = "P0First"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].First;
            oBookMark = "P0Last"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].Last;
            oBookMark = "P0Address"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].Address;
            oBookMark = "P0PostCode"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].PostCode;
            oBookMark = "P0Telephone"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].Telephone;
            oBookMark = "P0Mobile"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].Mobile;
            oBookMark = "P0email"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].email;
            oBookMark = "P0SignedBy"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].SignedBy;
            oBookMark = "P0DateSigned"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.PD[0].DateSigned).ToString("d");
            oBookMark = "P1Title"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].Title;
            oBookMark = "P1First"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].First;
            oBookMark = "P1Last"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].Last;
            oBookMark = "P1Address"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].Address;
            oBookMark = "P1PostCode"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].PostCode;
            oBookMark = "P1Telephone"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].Telephone;
            oBookMark = "P1Mobile"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].Mobile;
            oBookMark = "P1email"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].email;
            oBookMark = "P1SignedBy"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].SignedBy;
            oBookMark = "P1SignedOn"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.PD[1].DateSigned).ToString("d");
            oBookMark = "P1Relation"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].Relation;
            oBookMark = "P2Title"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].Title;
            oBookMark = "P2First"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].First;
            oBookMark = "P2Last"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].Last;
            oBookMark = "P2Address"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].Address;
            oBookMark = "P2PostCode"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].PostCode;
            oBookMark = "P2Telephone"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].Telephone;
            oBookMark = "P2Mobile"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].Mobile;
            oBookMark = "P2email"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].email;
            oBookMark = "P2SignedBy"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].SignedBy;
            oBookMark = "P2SignedOn"; oDoc.Bookmarks[oBookMark].Range.Text = Convert.ToDateTime(MData.PD[2].DateSigned).ToString("d");
            oBookMark = "P2Relation"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].Relation;
            //oBookMark = "P0StorePhone"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[0].StorePhone;    // not on form
            //oBookMark = "P1StorePhone"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[1].StorePhone;    // not on form
            //oBookMark = "P2StorePhone"; oDoc.Bookmarks[oBookMark].Range.Text = MData.PD[2].StorePhone;    // not on form

            //Close this form.
            this.Close();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            #region  Capture Screen details
            MData.Title = tbTitle.Text;
            MData.First = tbName.Text;
            MData.Last = tbLast.Text;
            MData.Address = tbAddress.Text;
            MData.PostCode = tbPostcode.Text;
            MData.DOB = dtpDOB.Text;
            MData.SignedBy = tbsigned.Text;
            MData.SignedOn = dtpSigned.Text;
            MData.PhotoSigned = tbaSigned.Text;
            MData.PhotoOn = dtpaSigned.Text;
            MData.Conditions = tbMedical.Text;
            MData.Notes = tbNotes.Text;
            MData.MedicalSigned = tbMedicalSign.Text;
            MData.MedicalOn = dtpMedicalSign.Text;
            MData.AuthorisedBy = tbaSigned.Text;
            MData.AuthorisedOn = dtpaSigned.Text;
            MData.Surgery = tbGP.Text;
            MData.PD[0].Title = tbpTitle.Text;
            MData.PD[0].First = tbpName.Text;
            MData.PD[0].Last = tbpSurname.Text;
            MData.PD[0].Address = tbpAddress.Text;
            MData.PD[0].PostCode = tbpPostcode.Text;
            MData.PD[0].Telephone = tbpPhone.Text;
            MData.PD[0].Mobile = tbpMobile.Text;
            MData.PD[0].email = tbpEmail.Text;
            MData.PD[0].SignedBy = tbpSigned.Text;
            if (dtppSigned.Text != "")
                MData.PD[0].DateSigned = dtppSigned.Text;
            else
                MData.PD[0].DateSigned = "01/01/2001";
            MData.PD[1].Title = tb1Title.Text;
            MData.PD[1].First = tb1Name.Text;
            MData.PD[1].Last = tb1Surname.Text;
            MData.PD[1].Address = tb1Address.Text;
            MData.PD[1].PostCode = tb1Postcode.Text;
            MData.PD[1].Telephone = tb1Phone.Text;
            MData.PD[1].Mobile = tb1Mobile.Text;
            MData.PD[1].email = tb1Email.Text;
            MData.PD[1].SignedBy = tb1Signed.Text;
            if (dtp1Signed.Text != "")
                MData.PD[1].DateSigned = dtp1Signed.Text;
            else
                MData.PD[1].DateSigned = "01/01/2001";
            MData.PD[1].Relation = tb1Relation.Text;
            MData.PD[2].Title = tb2Title.Text;
            MData.PD[2].First = tb2Name.Text;
            MData.PD[2].Last = tb2Surname.Text;
            MData.PD[2].Address = tb2Address.Text;
            MData.PD[2].PostCode = tb2Postcode.Text;
            MData.PD[2].Telephone = tb2Phone.Text;
            MData.PD[2].Mobile = tb2Mobile.Text;
            MData.PD[2].email = tb2Email.Text;
            MData.PD[2].SignedBy = tb2Signed.Text;
            if (dtp2Signed.Text != "")
                MData.PD[2].DateSigned = dtp2Signed.Text;
            else
                MData.PD[2].DateSigned = "01/01/2001";
            MData.PD[2].Relation = tb2Relation.Text;
            if (cbLeader.Checked == true)
                MData.Leader = "1";
            else
                MData.Leader = "0";
            if (cbpPermission.Checked == true)
                MData.StorePhone = "1";
            else
                MData.StorePhone = "0";
            if (cbPhoto.Checked==true)
                MData.photo = "1";
            else
                MData.photo = "0";
            if (cbMed.Checked == true)
                MData.MedicalPermission = "1";
            else
                MData.MedicalPermission = "0";
            if (cbpPermission.Checked == true)
                MData.PD[0].StorePhone = "1";
            else
                MData.PD[0].StorePhone = "0";
            if (cbpPermission.Checked == true)
                MData.PD[1].StorePhone = "1";
            else
                MData.PD[1].StorePhone = "0";
            if (cbpPermission.Checked == true)
                MData.PD[2].StorePhone = "1";
            else
                MData.PD[2].StorePhone = "0";
            #endregion capture screen

            if (MemNo.Text == "New Id: " + MData.IdNbr)
                InsertRecord();
            else
                UpdateRecord();
            // close form and reload
            this.Close();
        }
        private void InsertRecord()
        {
            int rv = 0;
            #region write records
            #region write Member
            string myInsertQuery = "INSERT INTO member (" +
                "Id, Title, FirstN, Inits, LastN, DOB, Leader, SignedOn, SignedBy, Address, " +
                "PostCode, PhotoConsent, PhotoSigned, PhotoOn, MedicalConditions, Notes,  " +
                "MedicalPermission, MedicalSigned, MedicalOn, AuthorisedBy, AuthorisedOn, Surgery, StorePhone " +
                ") VALUES ( " +
                MData.IdNbr + ",\"" +
                MData.Title + "\",\"" +
                MData.First + "\",\"" +
                MData.Inits + "\",\"" +
                MData.Last + "\",\'" +
                MData.DOB + "\',\"" +
                MData.Leader + "\",\'" +
                MData.SignedOn + "\',\"" +
                MData.SignedBy + "\",\"" +
                MData.Address + "\",\"" +
                MData.PostCode + "\",\"" +
                MData.photo + "\",\"" +
                MData.PhotoSigned + "\",\'" +
                MData.PhotoOn + "\',\"" +
                MData.Conditions + "\",\"" +
                MData.Notes + "\",\"" +
                MData.MedicalPermission + "\",\"" +
                MData.MedicalSigned + "\",\'" +
                MData.MedicalOn + "\',\"" +
                MData.AuthorisedBy + "\",\"" +
                MData.AuthorisedOn + "\",\"" +
                MData.Surgery + "\", \"" +
                MData.StorePhone +
                "\" );";

            Utils MyConn = new Utils();
            OleDbConnection myConnection = new OleDbConnection(MyConn.myConnectionString);
            OleDbCommand myCommand = new OleDbCommand(myInsertQuery, myConnection);
            try
            {
                myConnection.Open();
                rv = myCommand.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Insert Member Failed due to" + ex.Message);
                MessageBox.Show(ex.Source);
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (myConnection != null)
                    myConnection.Close();
            }
            #endregion write member
            #region write addresses
            rv = 0;
            for (int ityp = 0; ityp < 3; ityp++)
            {
                myInsertQuery = "INSERT INTO Person (" +
                "Id, Type,Title, FirstN, Inits, LastN, Address, PostCode, " +
                "Telephone, Mobile, email, StorePhone, SignedBy, DateSigned, Relation " +
                ") VALUES ( " +
                    MData.IdNbr + ", " + (ityp + 1).ToString() + ",\"" +
                    MData.PD[ityp].Title + "\",\"" +
                    MData.PD[ityp].First + "\",\"" +
                    MData.PD[ityp].Inits + "\",\"" +
                    MData.PD[ityp].Last + "\",\"" +
                    MData.PD[ityp].Address + "\",\"" +
                    MData.PD[ityp].PostCode + "\",\"" +
                    MData.PD[ityp].Telephone + "\",\"" +
                    MData.PD[ityp].Mobile + "\",\"" +
                    MData.PD[ityp].email + "\",\"" +
                    MData.PD[ityp].StorePhone + "\",\"" +
                    MData.PD[ityp].SignedBy + "\",\'" +
                    MData.PD[ityp].DateSigned + "\',\"" +
                    MData.PD[ityp].Relation +
                    "\" );";

                myCommand = new OleDbCommand(myInsertQuery, myConnection);
                try
                {
                    myConnection.Open();
                    rv = myCommand.ExecuteNonQuery();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Insert person type " + (ityp + 1).ToString() + " Failed due to" + ex.Message);
                    MessageBox.Show(ex.Source);
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (myConnection != null)
                        myConnection.Close();
                }
            }
            #endregion write addresses
            MessageBox.Show(" New Member Added");
            #endregion write recs
        }
        private void UpdateRecord()
        {
            int rv = 0;
            #region write records
            #region write Member
            string myInsertQuery = "UPDATE member SET " +
                "Title = \"" + MData.Title + "\", " +
                "FirstN = \"" + MData.First + "\", " +
                "Inits = \"" + MData.Inits + "\", " +
                "LastN = \"" + MData.Last + "\", " +
                "DOB = \'" + MData.DOB + "\', " +
                "Leader = \"" + MData.Leader + "\", " +
                "SignedOn = \'" + MData.SignedOn + "\', " +
                "SignedBy = \"" + MData.SignedBy + "\", " +
                "Address = \"" + MData.Address + "\", " +
                "PostCode = \"" + MData.PostCode + "\", " +
                "PhotoConsent = \"" + MData.photo + "\", " +
                "PhotoSigned = \"" + MData.PhotoSigned + "\", " +
                "PhotoOn = \'" + MData.PhotoOn + "\', " +
                "MedicalConditions = \"" + MData.Conditions + "\", " +
                "Notes = \"" + MData.Notes + "\", " +
                "MedicalPermission = \"" + MData.MedicalPermission + "\", " +
                "MedicalSigned = \"" + MData.MedicalSigned + "\", " +
                "MedicalOn = \'" + MData.MedicalOn + "\', " +
                "AuthorisedBy = \"" + MData.AuthorisedBy + "\", " +
                "AuthorisedOn = \'" + MData.AuthorisedOn + "\', " +
                "Surgery = \"" + MData.Surgery + "\", " +
                "StorePhone = \"" + MData.StorePhone+"\" " +
                " WHERE Id = " + MData.IdNbr + " ; ";

            Utils MyConn = new Utils();
            OleDbConnection myConnection = new OleDbConnection(MyConn.myConnectionString);
            OleDbCommand myCommand = new OleDbCommand(myInsertQuery, myConnection);
            try
            {
                myConnection.Open();
                rv = myCommand.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Update Member Failed due to" + ex.Message);
                MessageBox.Show(ex.Source);
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (myConnection != null)
                    myConnection.Close();
            }
            #endregion write member
            #region write addresses
            rv = 0;
            for (int ityp = 0; ityp < 3; ityp++)
            {
                myInsertQuery = "UPDATE Person SET " +
                    "Title = \"" + MData.PD[ityp].Title + "\", " +
                    "FirstN = \"" + MData.PD[ityp].First + "\", " +
                    "Inits = \"" + MData.PD[ityp].Inits + "\", " +
                    "LastN = \"" + MData.PD[ityp].Last + "\", " +
                    "Address = \"" + MData.PD[ityp].Address + "\", " +
                    "PostCode = \"" + MData.PD[ityp].PostCode + "\", " +
                    "Telephone = \"" + MData.PD[ityp].Telephone + "\", " +
                    "Mobile = \"" + MData.PD[ityp].Mobile + "\", " +
                    "email = \"" + MData.PD[ityp].email + "\", " +
                    "StorePhone = \"" + MData.PD[ityp].StorePhone + "\", " +
                    "SignedBy = \"" + MData.PD[ityp].SignedBy + "\", " +
                    "DateSigned = \'" + MData.PD[ityp].DateSigned + "\', " +
                    "Relation = \"" + MData.PD[ityp].Relation + "\" " +
                    "WHERE Id = " + MData.IdNbr + " AND Type = " + (ityp + 1).ToString() + "; ";

                myCommand = new OleDbCommand(myInsertQuery, myConnection);
                try
                {
                    myConnection.Open();
                    rv = myCommand.ExecuteNonQuery();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Update person type " + (ityp + 1).ToString() + " Failed due to" + ex.Message);
                    MessageBox.Show(ex.Source);
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (myConnection != null)
                        myConnection.Close();
                }
            }
            #endregion write addresses
            MessageBox.Show("Member Updated");
            #endregion write recs

        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Click 'OK' to confirm Delete of " + MData.First + " " + MData.Last + ", otherwise click 'Cancel' to return to form", "Confirm Delete", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                int rvSum = 0;
                int rv = 0;

                Utils MyConn = new Utils();
                OleDbConnection myConnection = new OleDbConnection(MyConn.myConnectionString);
                string myDeleteQuery = "DELETE from Person " +
                        "WHERE Id = " + MData.IdNbr + "; ";
                OleDbCommand myCommand = new OleDbCommand(myDeleteQuery, myConnection);

                try
                {
                    myConnection.Open();
                    rv = myCommand.ExecuteNonQuery();
                    rvSum += rv;
                    myConnection.Close();
                    myDeleteQuery = "DELETE FROM member " +
                       "WHERE Id = " + MData.IdNbr + ";";

                    myCommand = new OleDbCommand(myDeleteQuery, myConnection);
                    try
                    {
                        myConnection.Open();
                        rv = myCommand.ExecuteNonQuery();
                        myConnection.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Delete Member Failed due to" + ex.Message);
                        MessageBox.Show(ex.Source);
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        if (myConnection != null)
                            myConnection.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Delete Person record Failed due to" + ex.Message);
                    MessageBox.Show(ex.Source);
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (myConnection != null)
                        myConnection.Close();
                }
                Close();            // close member form now it's deleted
                MessageBox.Show("Member Deleted");
            }
        }
    }
}