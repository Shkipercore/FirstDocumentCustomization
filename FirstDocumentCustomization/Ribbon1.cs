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
        private Dictionary<string, Dictionary<string, string>> cashOFXML;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            fontDialog1.ShowColor = true;
        }

        private void buttonApply_Click(object sender, RibbonControlEventArgs e)
        {
            var options = IniinitializeGostOptions();

            Checker checker = new Checker(options);
            checker.Check();

        }

        private void buttonFont_Click(object sender, RibbonControlEventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.Cancel)
                return;
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
                                 editBoxRightIndent.Text,
                                 editBoxFirstLineIndent.Text,
                                 editorXML.ConvertedComboBoxAlignmentTextForIndex(comboBoxAlignmentText.Text)
                                 );
        }

        private void buttonAddWork_Click(object sender, RibbonControlEventArgs e)
        {
            if (editBoxAddWork.Text.Length > 0)
            {
                RibbonDropDownItem insertItem = Factory.CreateRibbonDropDownItem();
                insertItem.Label = editBoxAddWork.Text;
                bool isItemNotPresent = true;
                foreach (var item in comboBoxSelectionWork.Items)
                {
                    if (item.Label.Equals(insertItem.Label))
                        isItemNotPresent = false;
                }
                if (isItemNotPresent)
                {
                    comboBoxSelectionWork.Items.Add(insertItem);

                    EditorXML editorXML = new EditorXML();
                    editorXML.CreateNode(editBoxAddWork.Text);
                }
            }
        }

        private string getValueOFXMLForBoxies(Dictionary<string, string> dictionary, string elementName)
        {
            string valueOfDictionary = "";

            if (dictionary.TryGetValue(elementName, out valueOfDictionary))
            {
                ///write to log
            }

            return valueOfDictionary;
        }


        private void comboBoxSelectionWork_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var tagName = comboBoxSelectionWork.Text;
            LoadOfXMLForCash();
   
            if (cashOFXML.Keys.Contains(tagName))
            {
                var propertyOfXML = cashOFXML[tagName];
                editBoxLineSpacing.Text = propertyOfXML["lineSpacingOfOST"];
                editBoxLeftIndent.Text = propertyOfXML["leftIndentOfOST"];
                editBoxRightIndent.Text = propertyOfXML["rightIndentOfOST"];
                editBoxFirstLineIndent.Text = propertyOfXML["firstLineIndentOfOST"];
                string myCurrentlySelectedFont = propertyOfXML["nameFontOfOST"];
                string myCurrentlySelectedSize = propertyOfXML["sizeFontOfOST"];
                fontDialog1.Font = new System.Drawing.Font(myCurrentlySelectedFont, (float)Convert.ToInt32(myCurrentlySelectedSize));
                string selectedColorFromXML = propertyOfXML["colorFontOfOST"];
                System.Drawing.Color myCurrentlySelectedColor = System.Drawing.Color.FromName(selectedColorFromXML);
                fontDialog1.Color = myCurrentlySelectedColor;
                string selectedAligmentText = propertyOfXML["alignmentTextOfOST"];
                comboBoxAlignmentText.Text = ConvertedIndexForComboBoxAlignmentText(selectedAligmentText);
            }
        }

        private void comboBoxSelectionWork_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void buttonTest_Click(object sender, RibbonControlEventArgs e)
        {

        }

        public GostOptions IniinitializeGostOptions()
        {
            GostOptions gostOptions = new GostOptions(Globals.ThisAddIn.Application.ActiveDocument,
                                                        fontDialog1.Font.Name.ToString(),
                                                        fontDialog1.Color.Name.ToString(),
                                                        (float)Convert.ToDouble(editBoxLineSpacing.Text),
                                                        (float)Convert.ToInt32(fontDialog1.Font.Size),
                                                        (float)595.3,
                                                        (float)841.9,
                                                        (float)Convert.ToDouble(editBoxLeftIndent.Text),
                                                        (float)Convert.ToDouble(editBoxRightIndent.Text),
                                                        (float)Convert.ToDouble(editBoxFirstLineIndent.Text),
                                                        fontDialog1.Font.Name.ToString(),
                                                        "0",
                                                        "0",
                                                        "0");

            //GostOptions gostOptions = new GostOptions();
            //gostOptions.SetCurrentDocument(Globals.ThisAddIn.Application.ActiveDocument);
            //gostOptions.SetNameFontOfOST(fontDialog1.Font.Name.ToString());
            //gostOptions.SetColorFontOfOST(fontDialog1.Color.Name.ToString());
            //gostOptions.SetLineSpacingOFOST((float)Convert.ToDouble(editBoxLineSpacing.Text));
            //gostOptions.SetSizeFontOfOST((float)Convert.ToInt32(fontDialog1.Font.Size));
            //gostOptions.SetWidthOfOST((float)43);
            //gostOptions.SetHightOfOST((float)87);
            //gostOptions.SetLeftIndent((float)Convert.ToDouble(editBoxLeftIndent.Text));
            //gostOptions.SetFirstLineIndent((float)Convert.ToDouble(editBoxFirstLineIndent.Text));
            //gostOptions.SetNameFontForFooterOfOST(fontDialog1.Font.Name.ToString());
            //gostOptions.alignmentText = "0";
            //gostOptions.alignmentFooter = "0";
            //gostOptions.alignmentHeader = "0";

            return gostOptions;
        }

        private void editBoxLeftIndent_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxLeftIndent.Text, "^([0-9]+([,]{1}[0-9]+)?)$"))
            {
                editBoxLeftIndent.Text = string.Empty;
            }
        }

        public string ConvertedIndexForComboBoxAlignmentText(string items)
        {

            switch (items)
            {
                case "0":
                    return "По левому краю";

                case "1":
                    return "По центру";

                case "2":
                    return "По правому краю";

                case "3":
                    return "По ширине";
            }

            return items;

        }

        public void LoadOfXMLForCash()
        {
            var readerXML = new ReaderXML();

            List<string> listTagNames = new List<string>();
            foreach (var item in comboBoxSelectionWork.Items)
            {
                listTagNames.Add(item.Label);
            }
            cashOFXML = readerXML.GetDictionaryPropertyOfXML(listTagNames);
        }
    }
}
