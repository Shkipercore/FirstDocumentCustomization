using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Configuration;
using System.Collections.Specialized;
using System.Xml;

namespace FirstDocumentCustomization
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            fontDialog1.ShowColor = true;
        }

        private void buttonApply_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;
            var nameFont = ConfigurationManager.AppSettings.Get("nameFontOfOST");
            var pointOfCentimetrLine = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            var widthSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            var hightSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            ChekerGOST checker = new ChekerGOST(currentDocument, nameFont, WdColor.wdColorBlack, pointOfCentimetrLine, 14, WdParagraphAlignment.wdAlignParagraphJustify, widthSpacing, hightSpacing, "TimesNewRome", WdParagraphAlignment.wdAlignParagraphCenter);
            checker.Check();
        }

        private void buttonSettings_Click(object sender, RibbonControlEventArgs e)
        {
            Settings newForm = new Settings();
            newForm.Show();
        }

        private void buttonFont_Click(object sender, RibbonControlEventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            System.Configuration.Configuration currentConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            currentConfig.AppSettings.Settings["nameFontOfOST"].Value = fontDialog1.Font.Name.ToString();
            currentConfig.AppSettings.Settings["sizeFontOfOST"].Value = fontDialog1.Font.Size.ToString();
            currentConfig.AppSettings.Settings["colorFontOfOST"].Value = colorDialog1.Color.Name.ToString();
        }

        private void comboBoxAlignmentText_TextChanged(object sender, RibbonControlEventArgs e)
        {
            System.Configuration.Configuration currentConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var items = comboBoxAlignmentText.Items;

            var aligmentTextValue = items[0];
            if (aligmentTextValue.Label.Contains("По левому краю"))
                currentConfig.AppSettings.Settings["alignmentText"].Value = "0";
            
        }

        private void buttonSaveSettings_Click(object sender, RibbonControlEventArgs e)
        {

            XmlDocument xDoc = new XmlDocument();
            xDoc.Load("Config.xml");

            XmlElement xRoot = xDoc.DocumentElement;
            // создаем объект settings
            XmlElement settingsElem = xDoc.CreateElement("settings");
            // создаем атрибут name
            XmlAttribute nameAttr = xDoc.CreateAttribute("name");
            // создаем элементы
            XmlElement nameFontOfOSTElem = xDoc.CreateElement("nameFontOfOST");
            XmlElement colorFontOfOSTElem = xDoc.CreateElement("colorFontOfOST");
            XmlElement sizeFontOfOSTElem = xDoc.CreateElement("sizeFontOfOST");
            XmlElement alignmentTextElem = xDoc.CreateElement("alignmentText");
            //XmlElement widthOfOSTElem = xDoc.CreateElement("widthOfOST");
            //XmlElement hightOfOSTElem = xDoc.CreateElement("hightOfOST");
            //XmlElement nameFontForFooterOfOSTElem = xDoc.CreateElement("nameFontForFooterOfOST");
            //XmlElement alignmentFooterElem = xDoc.CreateElement("alignmentFooter");
            //XmlElement alignmentHeaderElem = xDoc.CreateElement("alignmentHeader");

            XmlText nameText = xDoc.CreateTextNode("Диплом");
            nameAttr.AppendChild(nameText);
            settingsElem.Attributes.Append(nameAttr);
            xRoot.AppendChild(settingsElem);
            xDoc.Save("Config.xml");

            foreach (XmlElement xnode in xRoot)
            {
                Settings settings = new Settings();
                XmlNode attr = xnode.Attributes.GetNamedItem("name");
                if (attr == null)
                {
                    XmlText nameText = xDoc.CreateTextNode("Диплом");
                    nameAttr.AppendChild(nameText);
                    settingsElem.Attributes.Append(nameAttr);
                    xRoot.AppendChild(settingsElem);
                    xDoc.Save("Config.xml");

                }

                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    if (childnode.Name == "nameFontOfOST")

                    {
                        XmlText nameFontOfOSTText = xDoc.CreateTextNode("Calibri");
                        nameFontOfOSTElem.AppendChild(nameFontOfOSTText);
                        xDoc.Save("Config.xml");
                    }
                    if (childnode.Name == "colorFontOfOST")
                    {
                        XmlText colorFontOfOSTText = xDoc.CreateTextNode("Red");
                        nameFontOfOSTElem.AppendChild(colorFontOfOSTText);
                        xDoc.Save("Config.xml");
                    }
                }

            }

        }
    }
}
