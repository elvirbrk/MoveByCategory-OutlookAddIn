using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookArchiveByCategoryAddIn
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMoveByCategory
    {
        void ArchiveItem();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    class MoveByCategory : IMoveByCategory
    {
        private static MoveByCategory instance;

        public static MoveByCategory GetInstance()
        {
            if (instance == null)
                instance = new MoveByCategory();

            return instance;
        }
        public void ArchiveItem()
        {
            Outlook.Selection conversations = Globals.ThisAddIn.Application.ActiveExplorer().Selection;

            foreach (Outlook.ConversationHeader convHeader in conversations.GetSelection(Outlook.OlSelectionContents.olConversationHeaders))
            {
                foreach (Outlook.MailItem item in convHeader.GetItems())
                {
                    string id = "";

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
                        System.Windows.Forms.MessageBox.Show("Archive folder for category:" + item.Categories + " not defined.");
                        continue;
                    }
                    else
                    {
                        Outlook.MAPIFolder folder;

                        try
                        {
                            folder = Globals.ThisAddIn.Application.Session.GetFolderFromID(id);
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

                        //Debug
                        //System.Windows.Forms.MessageBox.Show("S:"+item.Sender.Name+" R:"+item.Recipients[1].Name+" SU:"+item.Subject + " RT:" + item.ReceivedTime + " ST:" + item.SentOn + "  Move to: " + folder.FolderPath);
                    }

                }
            }
        }
    }
}
