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

        #endregion

        public EditSubjectDialog(string currentSubject)
        {
            InitializeComponent();

            if (currentSubject != null)
                tbSubject.Text = currentSubject;
        }
    }
}
