using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Collections.Specialized;

namespace FirstDocumentCustomization
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void Settings_Load(object sender, EventArgs e)
        {

        }

        private void buttonSaveSettings_Click(object sender, EventArgs e)
        {
            var items = comboBoxSelectionFont.Items;
            foreach(var item in items)
            {
                System.Configuration.Configuration currentConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                currentConfig.AppSettings.Settings["nameFontOfOST"].Value = comboBoxSelectionFont.Items.ToString();
            }

        }
    }
}
