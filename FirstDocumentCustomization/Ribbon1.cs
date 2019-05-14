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
            if (comboBoxAlignmentText. == 0)
            {
                currentConfig.AppSettings.Settings["nameFontOfOST"].Value = fontDialog1.Font.Name.ToString();
            }
        }
    }
}
