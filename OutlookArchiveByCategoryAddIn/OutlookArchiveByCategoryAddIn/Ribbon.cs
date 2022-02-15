using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ContextMenu();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookArchiveByCategoryAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public void OnArchiveClick(Office.IRibbonControl control)
        {
            MoveByCategory.GetInstance().ArchiveItem();
        }

        public System.Drawing.Bitmap GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "MyContextMenuMultipleItems":
                    {
                        return new System.Drawing.Bitmap(Properties.Resources.icon2);
                    }
                case "MyContextMenuMailItem":
                    {
                        return new System.Drawing.Bitmap(Properties.Resources.icon2);
                    }
                case "archiveButton":
                    {
                        return new System.Drawing.Bitmap(Properties.Resources.icon2);
                    }
            }
            return null;

        }

        public string GetLabel(Office.IRibbonControl control)
        {
            int ver = int.Parse(Globals.ThisAddIn.Application.Version.Substring(0,2));

            if (ver < 15 || ver >= 16)
            {
                return "Archive";
            }
            else
            {
                return "ARCHIVE";
            }
        }

        public void OnConfigClick(Office.IRibbonControl control)
        {
            Config form = new Config();
            form.ShowDialog();
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookArchiveByCategoryAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
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
