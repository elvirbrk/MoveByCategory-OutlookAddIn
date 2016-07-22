using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookArchiveByCategoryAddIn
{
    public partial class Config : Form
    {
        public Config()
        {
            InitializeComponent();

            LoadConfig();
        }

        private void LoadConfig()
        {
            LoadCategories();

            LoadConfiguredFolders();
        }

        private void LoadCategories()
        {
            Outlook.Categories categories = Globals.ThisAddIn.Application.Session.Categories;
            foreach (Outlook.Category cat in categories)
            {
                DataGridViewRow r = (DataGridViewRow)dgConfig.RowTemplate.Clone();
                r.CreateCells(dgConfig);

                r.Cells[0].Value = cat.Name;

                dgConfig.Rows.Add(r);
            }
        }

        private void LoadConfiguredFolders()
        {
            foreach (DataGridViewRow row in dgConfig.Rows)
            {
                try
                {
                    row.Cells[1].Value = GetFolderByCategoryConfig(row.Cells[0].Value.ToString());
                    row.Cells[2].Value = GetFolderIDByCategoryConfig(row.Cells[0].Value.ToString());
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error reading archive folder path for category:" + row.Cells[0].Value.ToString() + " from configuration. /n/r"+e.ToString());
                }
            }
        }

        public static string GetFolderByCategoryConfig(string cat)
        {
            return GetConfigByID(cat, 0);


        }

        public static string GetFolderIDByCategoryConfig(string cat)
        {
            return GetConfigByID(cat, 1);

        }

        private static string GetConfigByID(string key, int index)
        {
            try
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains(key))
                {
                    string val = ConfigurationManager.AppSettings[key];
                    return (val.Split('|'))[index].Trim();
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        private void SetFolderByCategoryConfig(string cat, string fol, string id)
        {
            try
            {
                // Open App.Config of executable
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                if (config.AppSettings.Settings.AllKeys.Contains(cat))
                {
                    config.AppSettings.Settings.Remove(cat);
                }

                // Add an Application Setting.
                config.AppSettings.Settings.Add(cat, fol + "|" + id);


                // Save the configuration file.
                config.Save(ConfigurationSaveMode.Modified);

                // Force a reload of a changed section.
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception e)
            {
                MessageBox.Show("Unable to save configuration. /n/r"+e.ToString());
            }

        }

        private void dgConfig_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.PickFolder();
            DataGridViewRow r = dgConfig.Rows[e.RowIndex];
            r.Cells["Folder"].Value = folder.FolderPath;
            r.Cells["ID"].Value = folder.EntryID;

            SetFolderByCategoryConfig(r.Cells[0].Value.ToString(), folder.FolderPath, folder.EntryID);
        }

    }
}
