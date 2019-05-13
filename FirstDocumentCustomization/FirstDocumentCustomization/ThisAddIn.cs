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
using System.Configuration;
using System.Collections.Specialized;


namespace FirstDocumentCustomization
{

    class ChekerGOST
    {
        private readonly String nameFontOfOST;
        private readonly String nameFontForFooterOfOST;
        private readonly String nameFontForHeaderOfOST;
        private readonly WdColor colorFontOfOST;
        private readonly float lineSpacingOFOST;
        private readonly Word.Document currentDocument;
        private readonly float sizeFontOfOST;
        private readonly float widthOfOST;
        private readonly float hightOfOST;
        private readonly WdParagraphAlignment alignmentText;
        private readonly WdParagraphAlignment alignmentFooter;
        private readonly WdParagraphAlignment alignmentHeader;


        public ChekerGOST(Word.Document document)
        {
            this.currentDocument = document;
        }

        public ChekerGOST(Word.Document document, String nameFont, WdColor colorFont, float lineSpacing, float sizeFont, WdParagraphAlignment alignment, float width, float highest, String fontFooter, WdParagraphAlignment alignmentFooter)
        {
            this.currentDocument = document;
            this.nameFontOfOST = nameFont;
            this.colorFontOfOST = colorFont;
            this.sizeFontOfOST = sizeFont;
            this.alignmentText = alignment;
            this.widthOfOST = width;
            this.hightOfOST = highest;
            this.nameFontForFooterOfOST = fontFooter;
            this.alignmentFooter = alignmentFooter;
        }

        public void Check()
        {
            bool titlePageflag = true;
            // MessageBox.Show("Проверено инфа сотка");
            Word.Paragraph paragraph = currentDocument.Paragraphs[1];
            CheckPageProperty(paragraph.Range);
            CheckHeaderPage();
            CheckAllImage();
            CheckTable();

            foreach (Word.Paragraph p in currentDocument.Paragraphs)
            {


                if (titlePageflag && IsTitlePage(p))
                {
                    CheckTitlePageDocument(p);
                }
                
                else
                    {

                    titlePageflag = false;

                    if (IsContentPage(p))
                    {
                        CheckDocumentContents(p);
                    }
                    else
                    {
                        if (IsHeader(p))
                        {
                            CheckHeaderText(p);
                        }
                        else
                        {
                            if (IsMathText(p))
                            {
                                CheckMathText(p);
                            }
                            else
                                CheckText(p);
                        }
                    }

                }
            }

        }
        private bool IsTitlePage(Word.Paragraph p)
        {
            var numberPage = GetPageNumber(p.Range);
            if (numberPage == 1)
            {
                return true;
            }
            return false;
        }

        private void CheckTitlePageDocument(Word.Paragraph p)
        {
            if (p.Range.Text != "\r")
            {
                if (p.Range.Text == "Федеральное государственное бюджетное образовательное учреждение высшего образования")
                {
                    var fontName = p.Range.Font.Name;
                    var fontSize = p.Range.Font.Size;
                    if (fontName != nameFontOfOST || fontSize != sizeFontOfOST)
                    {
                        if (fontSize != sizeFontOfOST)
                        {
                            AddCommet("Некорректен размер шрифта, должен стоять" + sizeFontOfOST, p.Range);
                        }
                        if (fontName != nameFontOfOST)
                        {
                            AddCommet("Не корректен тип шрифта, должен стоять " + nameFontOfOST, p.Range);
                        }
                    }

                }
            }

        }

        private void CheckText(Word.Paragraph p)
        {
            Word.Range range = p.Range;
            if (range.OMaths.Count <= 0)
            {

                var nameFont = range.Font.Name;
                var colorFont = range.Font.Color;
                var lineSpacing = p.LineSpacing;
                var fontSize = range.Font.Size;
                var aligment = p.Alignment;
                if (nameFont != nameFontOfOST && colorFont != colorFontOfOST && lineSpacing != lineSpacingOFOST && sizeFontOfOST != fontSize)
                {
                    StringBuilder texForComment = new StringBuilder("Текст не корректно оформлен согласно OST TUSUR: \n");
                    if (nameFont != nameFontOfOST)
                    {
                        texForComment.Append("\n Текущие имя шрифта " + nameFont + ", а должен быть установлен" + nameFontOfOST);
                    }
                    if (colorFont != colorFontOfOST)
                    {
                        texForComment.Append("\n Не коректен цвет шрифта, должен быть " + colorFontOfOST.ToString());
                    }
                    if (lineSpacing != lineSpacingOFOST)
                    {
                        texForComment.Append("\n Не корректно установлен межстрочный интервал");
                    }
                    if (sizeFontOfOST != fontSize)
                    {
                        texForComment.Append("\n Не коррректен размер шрифта, установлен" + fontSize + ", а должен быть установлен " + sizeFontOfOST);
                    }
                    AddCommet(texForComment.ToString(), range);
                }

            }

        }
        private void AddCommet(String textComment, Word.Range range)
        {
            var textRange = range.Text;
            if (range.Text != "\r\a")
            {
                currentDocument.Comments.Add(range, textComment);
            }
        }
        private void CheckPageProperty(Range rangeOfFirstParagrap)
        {
            var widthPage = currentDocument.PageSetup.PageWidth;
            var highestPage = currentDocument.PageSetup.PageHeight;

            if (widthPage != widthOfOST && hightOfOST != highestPage)
            {
                AddCommet("Не корректно указаны размеры страницы, указан Ширина" + widthOfOST + " Высота " + highestPage + "а нужна Ширина" + widthOfOST + " Высота " + hightOfOST, rangeOfFirstParagrap);
            }
        }
        private bool IsHeader(Word.Paragraph p)
        {
            var boldText = p.Range.Font.Bold;
            if (boldText == -1 && p.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
            {
                return true;
            }
            return false;
        }

        private void CheckHeaderText(Word.Paragraph p)
        {
            if (p.Range.Font.Name != nameFontOfOST && p.Range.Font.Size != sizeFontOfOST)
            {
                StringBuilder textForComment = new StringBuilder("Не корректно задан заголовок ");
                if (p.Range.Font.Name != nameFontOfOST)
                {
                    textForComment.Append(" \n  Не корректен выбранный шрифт " + p.Range.Font.Name + ", а необходим " + nameFontOfOST);
                }
                if (p.Range.Font.Size != sizeFontOfOST)
                {
                    textForComment.Append("\n Не корректно указан размер шрифта " + p.Range.Font.Size + ", необходимо установить " + sizeFontOfOST);
                }
                AddCommet(textForComment.ToString(), p.Range);
            }

        }


        private void CheckTable()
        {

            foreach (Word.Table table in currentDocument.Tables)
            {
                var fontNameTable = table.Range.Font.Name;
                var sizeFontTable = table.Range.Font.Size;

                if (fontNameTable != nameFontOfOST)
                {
                    AddCommet("Не коректно выбран шрифт для таблицы", table.Range);
                }
                if (sizeFontTable != sizeFontOfOST)
                {
                    AddCommet("Не коректно выбран шрифт для таблицы", table.Range);
                }


            }

        }
        private void CheckAllImage()
        {

            for (int i = 1; i <= currentDocument.InlineShapes.Count; i++)
            {
                var shap = currentDocument.InlineShapes[i];
                if (shap.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    var alignmentImage = shap.Range.Paragraphs.Alignment;
                    if (alignmentImage != WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        AddCommet("Необходимо установить орентацию по центру для данного рисунка", shap.Range);
                    }
                }
            }
        }
        private void CheckHeaderPage()
        {
            foreach (Word.Section section in currentDocument.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                var fontName = headerRange.Font.Name;
                WdParagraphAlignment aligment = headerRange.Paragraphs.Alignment;
                if (nameFontForHeaderOfOST != fontName && alignmentFooter != aligment)
                {
                    headerRange.Text = headerRange.Text + "  " + " Не корректно офрмлен верхний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + nameFontForFooterOfOST + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + alignmentFooter;
                    //AddCommet(" Не корректно офрмлен верхний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + nameFontForFooterOfOST + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + alignmentFooter, headerRange);
                }
            }

            //Нижний колонтитул
            foreach (Word.Section wordSection in currentDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                var fontName = footerRange.Font.Name;
                var aligment = footerRange.Paragraphs.Alignment;
                if (nameFontForFooterOfOST != fontName && alignmentFooter != aligment)
                {
                    footerRange.Text = footerRange.Text + "  " + " Не корректно офрмлен нижний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + nameFontForFooterOfOST + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + alignmentFooter;
                    // AddCommet(" Не корректно офрмлен нижний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + nameFontForFooterOfOST + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + alignmentFooter, footerRange);
                }
            }

        }

        private bool IsContentPage(Word.Paragraph p)
        {
            return false;
        }
        private void CheckDocumentContents(Word.Paragraph p)
        {

        }

        private bool IsMathText(Word.Paragraph p)
        {
            if (p.Range.OMaths.Count > 0)
            {
                return true;
            }
            else
                return false;

        }
        private void CheckMathText(Word.Paragraph p)
        {

        }
        private static string GetHeading(Word.Paragraph paragraph)
        {
            string heading = "";

            // Try to get the list number, otherwise just take the entire heading text
            heading = paragraph.Range.ListFormat.ListString;
            if (string.IsNullOrEmpty(heading))
            {
                heading = paragraph.Range.Text;
                heading = Regex.Replace(heading, "\\s+$", "");
            }
            return heading;
        }


        static int GetPageNumber(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
        }




    }
    public partial class ThisAddIn
    {


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //dynamic dialog = Application.Dialogs[Word.WdWordDialog.wdDialogFileOpen];
            //dialog.Show();
            //Word.Document currentDocument = this.Application.ActiveDocument;
            //// currentDocument.Paragraphs[1].Range.InsertParagraphBefore();
            ////    currentDocument.Paragraphs[1].Range.Select();
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
