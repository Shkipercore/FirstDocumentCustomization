using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Media;
using System.IO;

namespace FirstDocumentCustomization
{
    public partial class Ribbon1
    {
        private Dictionary<string, Dictionary<string, string>> cashOFXML;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            fontDialog1.ShowColor = true;
            LoadTypeWorkForRibbon();
        }

        private void buttonApply_Click(object sender, RibbonControlEventArgs e)
        {
            if (editBoxLeftIndent.Text != "" && editBoxRightIndent.Text != "" && editBoxFirstLineIndent.Text != "" && editBoxLineSpacing.Text != "" && editBoxIntervalBefore.Text != "" && editBoxIntervalAfter.Text != "" && editBoxNumberOfTitlePages.Text != "")
            {
                var options = IniinitializeGostOptions();

                Checker checker = new Checker(options);
                checker.Check();
 
            }
            else
            {
                if (editBoxLeftIndent.Text == "")
                {
                    editBoxLeftIndent.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxLeftIndent.OfficeImageId = "IndentClassic";
                }

                if (editBoxRightIndent.Text == "")
                {
                    editBoxRightIndent.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxRightIndent.OfficeImageId = "IndentRTL";
                }

                if (editBoxFirstLineIndent.Text == "")
                {
                    editBoxFirstLineIndent.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxFirstLineIndent.OfficeImageId = "AlignJustifyMedium";
                }

                if (editBoxLineSpacing.Text == "")
                {
                    editBoxLineSpacing.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxLineSpacing.OfficeImageId = "LineSpacing";
                }

                if (editBoxIntervalBefore.Text == "")
                {
                    editBoxIntervalBefore.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxIntervalBefore.OfficeImageId = "ParagraphSpacingBefore";
                }

                if (editBoxIntervalAfter.Text == "")
                {
                    editBoxIntervalAfter.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxIntervalAfter.OfficeImageId = "ParagraphSpacingAfter";
                }

                if (editBoxNumberOfTitlePages.Text == "")
                {
                    editBoxNumberOfTitlePages.OfficeImageId = "DeclineTask";
                    playSimpleSound();
                }
                else
                {
                    editBoxNumberOfTitlePages.OfficeImageId = "Thesaurus";
                }

            }
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
                                 ConvertedComboBoxAlignmentTextForIndex(comboBoxAlignmentText.Text),
                                 editBoxIntervalBefore.Text,
                                 editBoxIntervalAfter.Text
                                 );

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
                editBoxIntervalBefore.Text = propertyOfXML["intervalBeforeOfOST"];
                editBoxIntervalAfter.Text = propertyOfXML["intervalAfterOfOST"];
            }
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
                                                        ConvertedComboBoxAlignmentTextForIndex(comboBoxAlignmentText.Text),
                                                        "0",
                                                        "0",
                                                        (float)Convert.ToDouble(editBoxIntervalBefore.Text),
                                                        (float)Convert.ToDouble(editBoxIntervalAfter.Text));

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
                case "wdAlignParagraphLeft":
                    return "По левому краю";

                case "wdAlignParagraphCenter":
                    return "По центру";

                case "wdAlignParagraphRight":
                    return "По правому краю";

                case "wdAlignParagraphJustify":
                    return "По ширине";
            }
            return items;
        }

        public string ConvertedComboBoxAlignmentTextForIndex(string items)
        {
            switch (items)
            {
                case "По левому краю":
                    return "wdAlignParagraphLeft";

                case "По центру":
                    return "wdAlignParagraphCenter";

                case "По правому краю":
                    return "wdAlignParagraphRight";

                case "По ширине":
                    return "wdAlignParagraphJustify";
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

        private void buttonEditWork_Click(object sender, RibbonControlEventArgs e)
        {
            FormEditWork formEditWork = new FormEditWork();
            formEditWork.Show();
        }

        public void LoadTypeWorkForRibbon()
        {

            string path = Environment.CurrentDirectory + "\\Config.xml";
            MessageBox.Show(path);
            if (!File.Exists(path))
            {
                File.Create(path);
            }

            XDocument xdoc = XDocument.Load(path);

            foreach (XElement settingsElement in xdoc.Element("ConfigSettings").Elements("Settings"))
            {
                XAttribute nameAttribute = settingsElement.Attribute("name");
                if (nameAttribute != null)
                {
                    RibbonDropDownItem insertItem = Factory.CreateRibbonDropDownItem();
                    insertItem.Label = nameAttribute.Value;

                    comboBoxSelectionWork.Items.Add(insertItem);
                }
            }
        }

        private void editBoxRightIndent_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxRightIndent.Text, "^([0-9]+([,]{1}[0-9]+)?)$"))
            {
                editBoxRightIndent.Text = string.Empty;
            }
        }

        private void editBoxFirstLineIndent_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxFirstLineIndent.Text, "^([0-9]+([,]{1}[0-9]+)?)$"))
            {
                editBoxFirstLineIndent.Text = string.Empty;
            }
        }

        private void editBoxIntervalBefore_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxIntervalBefore.Text, "^([0-9]+([,]{1}[0-9]+)?)$"))
            {
                editBoxIntervalBefore.Text = string.Empty;
            }
        }

        private void editBoxIntervalAfter_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxIntervalAfter.Text, "^([0-9]+([,]{1}[0-9]+)?)$"))
            {
                editBoxIntervalAfter.Text = string.Empty;
            }
        }

        private void editBoxLineSpacing_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxLineSpacing.Text, "^([0-9]+([,]{1}[0-9]+)?)$"))
            {
                editBoxLineSpacing.Text = string.Empty;
            }
        }

        private void editBoxNumberOfTitlePages_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (!Regex.IsMatch(editBoxNumberOfTitlePages.Text, "^[0-9]+$"))
            {
                editBoxNumberOfTitlePages.Text = string.Empty;
            }
        }

        private void playSimpleSound()
        {
            SoundPlayer simpleSound = new SoundPlayer(@"c:\Windows\Media\Windows Background.wav");
            simpleSound.Play();
        }
    }
}
