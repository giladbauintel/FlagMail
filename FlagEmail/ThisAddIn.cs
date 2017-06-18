using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Diagnostics;

namespace FlagEmail
{
    public partial class ThisAddIn
    {
        private Outlook.Items mToDoItems = null;

        public MembersDB membersDB;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.Folder todoFolder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo) as Outlook.Folder;
            mToDoItems = todoFolder.Items;
            mToDoItems.ItemAdd += Items_ItemAdd;
        }

        void Items_ItemAdd(object Item)
        {
            Debug.WriteLine("Adding task");

            var mailItem = Item as Outlook.MailItem;
            if (mailItem != null &&
                mailItem.FlagStatus != Outlook.OlFlagStatus.olFlagComplete &&
                Properties.Settings.Default.Email != "")
            {
                var newItem = Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                newItem.To = Properties.Settings.Default.Email;
                newItem.Subject = mailItem.Subject;

                bool includeBody = Properties.Settings.Default.IncludeBody;
                bool includeAttachments = Properties.Settings.Default.IncludeAttachments;
                bool editSubject = Properties.Settings.Default.EditSubject;

                if (editSubject)
                {
                    var subjectDialog = new EditSubjectDialog(newItem.Subject);


                    Debug.WriteLine("Loading members");

                    subjectDialog.LoadMembers();

                    Debug.WriteLine("Finished loading members");
                    
                    var dialogRes = subjectDialog.ShowDialog();


                    if (dialogRes == System.Windows.Forms.DialogResult.OK)
                    {
                        newItem.Subject = subjectDialog.Subject;

                        if (subjectDialog.IncludeMembers)
                            newItem.Subject += subjectDialog.getMembersString();

                        includeBody = subjectDialog.IncludeBody;
                        includeAttachments = subjectDialog.IncludeAttachments;
                    }
                    else
                        return;
                }

                if (includeBody)
                {
                    newItem.Body = mailItem.Body;
                }

                if (includeAttachments)
                {
                    foreach (Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        if (attachment.Type == Outlook.OlAttachmentType.olByValue)
                        {
                            var tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), attachment.FileName);
                            attachment.SaveAsFile(tempPath);
                            newItem.Attachments.Add(tempPath, attachment.Type, attachment.Position, attachment.DisplayName);
                        }
                    }
                }

                ((Microsoft.Office.Interop.Outlook._MailItem)newItem).Send();
            }
            else
            {
                if (mailItem == null)
                {
                    MessageBox.Show("Failed to add task! mailItem is null", "Failed task!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Failed to add task! email: " + Properties.Settings.Default.Email, "Failed task!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }                
            }
        }
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            mToDoItems.ItemAdd -= Items_ItemAdd;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new FlagEmailRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
