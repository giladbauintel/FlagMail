using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new FlagEmailRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace FlagEmail
{
    [ComVisible(true)]
    public class FlagEmailRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public FlagEmailRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FlagEmail.FlagEmailRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public bool SettingsCheckboxGetPressed(IRibbonControl control)
        {
            CheckBox checkboxControl = control as CheckBox;

            switch (control.Id)
            {
                case "cbIncludeBody":
                    return Properties.Settings.Default.IncludeBody;
                case "cbIncludeAttachments":
                    return Properties.Settings.Default.IncludeAttachments;
                case "cbEditSubject":
                    return Properties.Settings.Default.EditSubject;
            }

            return false;
        }

        public void OnSettingsCheckboxAction(IRibbonControl control, bool isChecked)
        {
            switch(control.Id)
            {
                case "cbIncludeBody":
                    Properties.Settings.Default.IncludeBody = isChecked;
                    break;
                case "cbIncludeAttachments":
                    Properties.Settings.Default.IncludeAttachments = isChecked;
                    break;
                case "cbEditSubject":
                    Properties.Settings.Default.EditSubject = isChecked;
                    break;
                default:
                    return;
            }
            
            Properties.Settings.Default.Save();

        }

        public string OnSettingsAction(IRibbonControl control)
        {
            var dlg = new SettingsDialog();
            var dlgRes = dlg.ShowDialog();

            if (dlgRes == DialogResult.OK)
            {
                // Check for email
                if (dlg.Email != null && dlg.Email != "" && dlg.Email != " " &&
                    dlg.Email.Contains("@") && dlg.Email.Contains("."))
                {
                    Properties.Settings.Default.Email = dlg.Email;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    MessageBox.Show("Email must be a valid address!", "Invalid Email!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return dlgRes.ToString();
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
