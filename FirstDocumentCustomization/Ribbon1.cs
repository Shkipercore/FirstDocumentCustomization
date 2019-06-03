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
            var colorFont = ConfigurationManager.AppSettings.Get("colorFontOfOST");
            var lineSpacing = ConfigurationManager.AppSettings.Get("lineSpacing");
            var sizeFont = ConfigurationManager.AppSettings.Get("sizeFont");
            var width = ConfigurationManager.AppSettings.Get("wight");
            var highest = ConfigurationManager.AppSettings.Get("highest");
            var leftIndent = ConfigurationManager.AppSettings.Get("leftIndent");
            var firstLineIndent = ConfigurationManager.AppSettings.Get("firstLineIndent");
            var fontFooter = ConfigurationManager.AppSettings.Get("fontFooter");
            var alignment = ConfigurationManager.AppSettings.Get("alignment");
            var alignmentHeader = ConfigurationManager.AppSettings.Get("alignmentHeader");
            var alignmentFooter = ConfigurationManager.AppSettings.Get("alignmentFooter");
            var pointOfCentimetrLine = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            var widthSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            var hightSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);

            GostOptions gostOptions = new GostOptions(currentDocument, 
                                                      nameFont,  
                                                      colorFont, 
                                                      Convert.ToInt32(lineSpacing),
                                                      Convert.ToInt32(sizeFont), 
                                                      Convert.ToInt32(width),
                                                      Convert.ToInt32(highest),
                                                      Convert.ToInt32(leftIndent),
                                                      Convert.ToInt32(firstLineIndent),
                                                      fontFooter,
                                                      alignment,
                                                      alignmentHeader,
                                                      alignmentFooter);


            Checker checker = new Checker(gostOptions);
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

            EditorXML editorXML = new EditorXML();
            editorXML.AddElement(comboBoxSelectionWork.Text,
                                 fontDialog1.Font.Name.ToString(),
                                 fontDialog1.Color.Name.ToString(),
                                 fontDialog1.Font.Size.ToString(),
                                 editBoxLineSpacing.Text,
                                 editBoxLeftIndent.Text,
                                 editBoxFirstLineIndent.Text
                                 );

        }

        private void buttonAddWork_Click(object sender, RibbonControlEventArgs e)
        {
            //comboBoxSelectionWork.Items.Add(editBoxAddWork.Text);
            var tagName = comboBoxSelectionWork.Text;
            editBoxLeftIndent.Text = getValueOFXMLForBoxies(tagName, "leftIndentOfOST");

        }

        private string getValueOFXMLForBoxies(string tagName, string elementName)
        {
            ReaderXML readerXML = new ReaderXML();
            var dictionary = readerXML.LoadPropertyOfXML(tagName);

            string valueOfDictionary = "";

            if (dictionary.TryGetValue(elementName, out valueOfDictionary))
            {
                ///write to log
            }

            return valueOfDictionary;
        }


        private void comboBoxSelectionWork_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void comboBoxSelectionWork_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
