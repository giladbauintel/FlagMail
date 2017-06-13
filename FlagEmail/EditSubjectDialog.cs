using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FlagEmail
{
    public partial class EditSubjectDialog : Form
    {
        #region Properties

        public string Subject
        {
            get { return tbSubject.Text; }
        }

        public bool IncludeBody
        {
            get { return cbIncludeBody.Checked; }
        }

        public bool IncludeAttachments
        {
            get { return cbIncludeAttachments.Checked; }
        }
        public bool IncludeMembers
        {
            get { return cbMembers.Checked; }
        }
        #endregion

        public EditSubjectDialog(string currentSubject)
        {
            InitializeComponent();

            if (currentSubject != null)
                tbSubject.Text = currentSubject;

            cbIncludeBody.Checked        = Properties.Settings.Default.IncludeBody;
            cbIncludeAttachments.Checked = Properties.Settings.Default.IncludeAttachments;

        }

        public void LoadMembers()
        {
            membersDB = MembersDB.Instance;

            if (!membersDB.Initialized && Properties.Settings.Default.JSONPath != "")
                membersDB.loadJSON(Properties.Settings.Default.JSONPath);
           
            lbMembers.DataSource = membersDB.getTable();
            lbMembers.DisplayMember = "FullName";                    
        }

        public string getMembersString()
        {
            string membersString = "";

            foreach (var item in lbMembers.SelectedItems)
            {
                DataRowView drv = (DataRowView)item;

                membersString += " @" + drv.Row["Username"];
            }

            return membersString;
        }

        private void cbMembers_CheckStateChanged(object sender, EventArgs e)
        {
            if (cbMembers.Checked)
            {
                lbMembers.Enabled = true;
            }
            else
            {
                lbMembers.Enabled = false;
            }
        }

    }
}
