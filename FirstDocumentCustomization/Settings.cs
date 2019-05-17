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


            //AddUpdateAppSettings("nameFontOfOST", "Andreyonelove");
            //var keyOfupdate = ReadSetting("nameFontOfOST");

            // Create the XmlDocument.
            XmlDocument doc = new XmlDocument();
            string m_progectPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            //doc.Load("FirstDocumentCustomization.dll.config");
            doc.Load(m_progectPath + "\\"+ "FirstDocumentCustomization.dll.config");

            // Add a price element.
            var newElem = doc.GetElementsByTagName("appSettings"); // /doc.CreateElement("price");
            
            
            
            newElem.InnerText = "10.95";
            doc.DocumentElement.AppendChild(newElem);

            // Save the document to a file. White space is
            // preserved (no white space).
            doc.PreserveWhitespace = true;
            doc.Save("data.xml");




        }
    }
}
