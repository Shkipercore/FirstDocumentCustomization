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
            //Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;
            //var nameFont = ConfigurationManager.AppSettings.Get("nameFontOfOST");
            //var colorFont = ConfigurationManager.AppSettings.Get("colorFontOfOST");
            //var lineSpacing = ConfigurationManager.AppSettings.Get("lineSpacing");
            //var sizeFont = ConfigurationManager.AppSettings.Get("sizeFont");
            //var width = ConfigurationManager.AppSettings.Get("wight");
            //var highest = ConfigurationManager.AppSettings.Get("highest");
            //var leftIndent = ConfigurationManager.AppSettings.Get("leftIndent");
            //var firstLineIndent = ConfigurationManager.AppSettings.Get("firstLineIndent");
            //var fontFooter = ConfigurationManager.AppSettings.Get("fontFooter");
            //var alignment = ConfigurationManager.AppSettings.Get("alignment");
            //var alignmentHeader = ConfigurationManager.AppSettings.Get("alignmentHeader");
            //var alignmentFooter = ConfigurationManager.AppSettings.Get("alignmentFooter");
            //var pointOfCentimetrLine = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            //var widthSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);
            //var hightSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(1.5f);

            //GostOptions gostOptions = new GostOptions(currentDocument,
            //                                          nameFont,
            //                                          colorFont,
            //                                          Convert.ToInt32(lineSpacing),
            //                                          Convert.ToInt32(sizeFont),
            //                                          Convert.ToInt32(width),
            //                                          Convert.ToInt32(highest),
            //                                          Convert.ToInt32(leftIndent),
            //                                          Convert.ToInt32(firstLineIndent),
            //                                          fontFooter,
            //                                          alignment,
            //                                          alignmentHeader,
            //                                          alignmentFooter);


            //Checker checker = new Checker(gostOptions);


        }

        private void buttonFont_Click(object sender, RibbonControlEventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.Cancel)
                return;
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
                                 editBoxLineSpacing.Text,
                                 fontDialog1.Font.Size.ToString(),
                                 editBoxLeftIndent.Text,
                                 editBoxFirstLineIndent.Text
                                 );

        }

        private void buttonAddWork_Click(object sender, RibbonControlEventArgs e)
        {
            if (editBoxAddWork.Text.Length > 0)
            {
                RibbonDropDownItem item1 = Factory.CreateRibbonDropDownItem();
                comboBoxSelectionWork.Items.Add(item1);
                item1.Label = editBoxAddWork.Text;

                EditorXML editorXML = new EditorXML();
                editorXML.CreateNode(editBoxAddWork.Text);

            }

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

        private void buttonLoadProperty_Click(object sender, RibbonControlEventArgs e)
        {
            var tagName = comboBoxSelectionWork.Text;
            editBoxLineSpacing.Text = getValueOFXMLForBoxies(tagName, "lineSpacingOfOST");
            editBoxLeftIndent.Text = getValueOFXMLForBoxies(tagName, "leftIndentOfOST");
            editBoxFirstLineIndent.Text = getValueOFXMLForBoxies(tagName, "firstLineIndentOfOST");

        }

        private void buttonTest_Click(object sender, RibbonControlEventArgs e)
        {

        }

        public GostOptions IniinitializeGostOptions()
        {
            GostOptions gostOptions = new GostOptions(Globals.ThisAddIn.Application.ActiveDocument,
                                                      fontDialog1.Font.Name.ToString(),
                                                      fontDialog1.Color.Name.ToString(),
                                                      Convert.ToInt32(editBoxLineSpacing.Text),
                                                      Convert.ToInt32(fontDialog1.Font.Size),
                                                      43,
                                                      87,
                                                      Convert.ToInt32(editBoxLeftIndent.Text),
                                                      Convert.ToInt32(editBoxFirstLineIndent.Text),
                                                      fontDialog1.Color.Name.ToString(),
                                                      "0",
                                                      "0",
                                                      "0");

            return gostOptions;
        }
        
    }
}
