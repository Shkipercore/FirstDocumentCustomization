using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace FirstDocumentCustomization
{

  
    public partial class ThisAddIn
    {


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //dynamic dialog = Application.Dialogs[Word.WdWordDialog.wdDialogFileOpen];
            //dialog.Show();
            //Word.Document currentDocument = this.Application.ActiveDocument;
            ////// currentDocument.Paragraphs[1].Range.InsertParagraphBefore();
            //////    currentDocument.Paragraphs[1].Range.Select();
            //var curre = currentDocument.Content;

            //var width = curre.PageSetup.PageWidth;
            //var height = curre.PageSetup.PageHeight;
            //var rightmargin = curre.PageSetup.RightMargin;
            //var leftMargin = curre.PageSetup.LeftMargin;
            //var topMargine = curre.PageSetup.TopMargin;
            //var bottommargin = curre.PageSetup.BottomMargin;
            //var mirrorMargin = curre.PageSetup.MirrorMargins;
            //var orientation = curre.Orientation;
            //var verticalAlign = curre.PageSetup.VerticalAlignment;

            //var vare = Application.PointsToCentimeters(leftMargin);

            //var pointOfCentimetrLine = Application.CentimetersToPoints(1.5f);
            //var widthSpacing = Application.CentimetersToPoints(1.5f);
            //var hightSpacing = Application.CentimetersToPoints(1.5f);
            //ChekerGOST checker = new ChekerGOST(currentDocument, "TimesNewRome", WdColor.wdColorBlack, pointOfCentimetrLine, 14, WdParagraphAlignment.wdAlignParagraphJustify, widthSpacing, hightSpacing, "TimesNewRome", WdParagraphAlignment.wdAlignParagraphCenter);

            //foreach (Word.Paragraph p in currentDocument.Paragraphs)
            //{
            //    // MessageBox.Show(p.Range.Text);
            //    Word.Range range = p.Range;
            //    if (range.OMaths.Count <= 0)
            //    {
            //        var boldText = p.Range.Font.Bold;
            //        var aligment = p.Alignment;
            //        Word.Style style = p.get_Style() as Word.Style;
            //        var srt = p.Range.Font.Name;
            //        var str2 = p.Range.Font.Color;
            //        var styles = style.Description;
            //        var Linespacing = p.LineSpacing;
            //        var text = p.Range.Text;
            //        // var str2 = p.Range.;

            //    }

            //}

        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private Microsoft.Office.Tools.Word.ContentControl richTextControl1;


        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    internal class Heading
    {
        public Heading()
        {
        }

        public string Text { get; set; }
        public object Start { get; set; }
    }
}
