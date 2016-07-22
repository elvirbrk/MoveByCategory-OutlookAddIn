using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookArchiveByCategoryAddIn
{
    public partial class RibbonArchive : Microsoft.Office.Core.IRibbonExtensibility
    {
        private void Config_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Config form = new Config();
            form.ShowDialog();
        }

        private void btnArchive_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ArchiveItem();
        }

        string IRibbonExtensibility.GetCustomUI(string RibbonID)
        {
            throw new NotImplementedException();
        }
    }
}
