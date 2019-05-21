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
using System.Xml;
using System.IO;
using System.Reflection;

namespace FirstDocumentCustomization
{
    public partial class Settings : Form
    {
        static String ReadSetting(string key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string result = appSettings[key] ?? "Not Found";
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error reading app settings");
                return "wtf";
            }
        }



        private static string GetPropertyOfConfig(string property, string tagName, string configName)
        {
            string m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            string reference = "not found value in XML";
            System.Xml.XmlDocument _XmlDocument = new System.Xml.XmlDocument();
            _XmlDocument.Load(m_exePath + "\\" + configName);

            foreach (XmlElement _XmlElement in _XmlDocument.GetElementsByTagName(tagName))
            {
                foreach (XmlElement XmlElementChild in _XmlElement)
                {
                    if (XmlElementChild.Name == property) { reference = XmlElementChild.InnerText; }
                }
            
            }

            return reference;
        }


        static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                configFile.Sections.Add("fsfsdf", new MySettings
                {
                    MyProperty = "LOL"
                });
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }


        public Settings()
        {
            InitializeComponent();
        }

        private void Settings_Load(object sender, EventArgs e)
        {

        }

        private void buttonSaveSettings_Click(object sender, EventArgs e)
        {


        }
    }
}
