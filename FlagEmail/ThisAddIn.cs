using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace FlagEmail
{
    public partial class ThisAddIn
    {
        private Outlook.Items mToDoItems = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.Folder todoFolder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo) as Outlook.Folder;
            mToDoItems = todoFolder.Items;
            mToDoItems.ItemAdd += Items_ItemAdd;
        }

        void Items_ItemAdd(object Item)
        {
            var mailItem = Item as Outlook.MailItem;
            if (mailItem != null &&
                mailItem.FlagStatus != Outlook.OlFlagStatus.olFlagComplete &&
                Properties.Settings.Default.Email != "")
            {
                var forwardedItem = mailItem.Forward();
                forwardedItem.To = Properties.Settings.Default.Email;
                forwardedItem.Send();
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
