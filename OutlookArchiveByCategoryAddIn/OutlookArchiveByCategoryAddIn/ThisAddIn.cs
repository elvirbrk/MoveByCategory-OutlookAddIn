using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookArchiveByCategoryAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        public void ArchiveItem()
        {
            Outlook.Selection conversations = this.Application.ActiveExplorer().Selection;

            foreach (Outlook.ConversationHeader convHeader in conversations.GetSelection(Outlook.OlSelectionContents.olConversationHeaders))
            {
                foreach (Outlook.MailItem item in convHeader.GetItems())
                {
                    string id="";

                    try
                    {
                        id = Config.GetFolderIDByCategoryConfig(item.Categories);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show("Error reading archive folder path for category:" + item.Categories + " from configuration./n/r" + e.ToString());
                        continue;
                    }

                    if (string.IsNullOrEmpty(id))
                    {
                        System.Windows.Forms.MessageBox.Show("Archive folder for category:"+item.Categories+" not defined.");
                        continue;
                    }
                    else
                    {
                        Outlook.MAPIFolder folder;

                        try
                        {
                            folder = Application.Session.GetFolderFromID(id);
                        }
                        catch (Exception e)
                        {
                            System.Windows.Forms.MessageBox.Show("Archive folder with ID:" + id + " does not exist./n/r" + e.ToString());
                            continue;
                        }

                        try
                        {
                            item.Move(folder);
                        }
                        catch (Exception e)
                        {
                            System.Windows.Forms.MessageBox.Show("Unable to move item:" + item.Subject + " /n/r" + e.ToString());
                            continue;
                        }

                        System.Windows.Forms.MessageBox.Show("S:"+item.Sender.Name+" R:"+item.Recipients[1].Name+" SU:"+item.Subject + " RT:" + item.ReceivedTime + " ST:" + item.SentOn + "  Move to: " + folder.FolderPath);
                    }
                    
                } 
            }
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
