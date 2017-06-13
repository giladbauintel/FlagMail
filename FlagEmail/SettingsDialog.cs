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
    public partial class SettingsDialog : Form
    {
        #region Properties

        public string Email
        {
            get { return tbEmail.Text; }
        }

        public string JSONPath
        {
            get { return tbJSONPath.Text; }
        }

        #endregion

        public SettingsDialog()
        {
            InitializeComponent();

            tbEmail.Text = Properties.Settings.Default.Email;
            tbJSONPath.Text = Properties.Settings.Default.JSONPath;
        }

        private void btnOpenFile_KeyPressed(object sender, EventArgs e)
        {
            if (openJSON.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tbJSONPath.Text = openJSON.FileName;
            }  
        }
    }
}
