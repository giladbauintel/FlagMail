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

        #endregion

        public EditSubjectDialog(string currentSubject)
        {
            InitializeComponent();

            if (currentSubject != null)
                tbSubject.Text = currentSubject;

            cbIncludeBody.Checked        = Properties.Settings.Default.IncludeBody;
            cbIncludeAttachments.Checked = Properties.Settings.Default.IncludeAttachments;

        }
    }
}
