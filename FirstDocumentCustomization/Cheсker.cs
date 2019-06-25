using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Drawing;

namespace FirstDocumentCustomization
{
    public class Checker
    {
        private GostOptions gostOptions;
        private Word.Document currentDocument;

        public Checker(GostOptions options)
        {
            gostOptions = options;
            currentDocument = options.GetCurrentDocument();
        }

        public void Check()
        {
            CheckPageProperty();
            CheckHeaderPage();
            CheckAllImage();
            CheckTable();
            CheckText();
        }
        private void CheckPageProperty()
        {
            int indexFirstParagraph = 1;
            var paragraphForCheckSetupPage = currentDocument.Paragraphs[indexFirstParagraph];
            var rangeForAddCommit = paragraphForCheckSetupPage.Range;


            var widthPage = currentDocument.PageSetup.PageWidth;
            var highestPage = currentDocument.PageSetup.PageHeight;

            if (widthPage != gostOptions.GetWidthOfOST())
            {
                AddComment("Не корректно указаны размеры страницы, указан Ширина " + widthPage + ",а нужна Ширина " + gostOptions.GetWidthOfOST(), rangeForAddCommit);
            }

            if (gostOptions.GetHightOfOST() != highestPage)
            {
                AddComment("Не корректно указаны размеры страницы, указан  Высота " + highestPage + ", а нужна Высота " + gostOptions.GetHightOfOST(), rangeForAddCommit);
            }
        }

        private void CheckHeaderPage()
        {
            foreach (Word.Section section in currentDocument.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                var fontName = headerRange.Font.Name;
                string aligment = headerRange.Paragraphs.Alignment.ToString();

                if (gostOptions.alignmentHeader != null && !gostOptions.GetNameFontForHeaderOfOST().Equals(fontName) && !gostOptions.alignmentHeader.Equals(aligment))
                {

                    headerRange.Text = headerRange.Text + "  " + " Не корректно офрмлен верхний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + gostOptions.GetNameFontForHeaderOfOST() + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + gostOptions.alignmentHeader;

                }
            }

            foreach (Word.Section wordSection in currentDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                var fontName = footerRange.Font.Name;
                var aligment = footerRange.Paragraphs.Alignment;
                if (!gostOptions.GetNameFontForFooterOfOST().Equals(fontName) && !gostOptions.alignmentFooter.Equals(aligment))
                {
                    footerRange.Text = footerRange.Text + "  " + " Не корректно оформлен нижний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + gostOptions.GetNameFontForFooterOfOST() + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + gostOptions.alignmentFooter;

                }
            }

        }

        private void CheckAllImage()
        {
            var currentDocument = gostOptions.GetCurrentDocument();

            for (int i = 1; i <= currentDocument.InlineShapes.Count; i++)
            {
                var shap = currentDocument.InlineShapes[i];
                if (shap.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    var alignmentImage = shap.Range.Paragraphs.Alignment;
                    if (alignmentImage != WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        AddComment("Необходимо установить выравнивание по центру для данного рисунка", shap.Range);
                    }
                }
            }
        }

        private void CheckTable()
        {

            foreach (Word.Table table in currentDocument.Tables)
            {
                var fontNameTable = table.Range.Font.Name;
                var sizeFontTable = table.Range.Font.Size;

                if (fontNameTable != gostOptions.GetNameFontOfOST())
                {
                    AddComment("Не корректно выбран шрифт для таблицы", table.Range);
                }
                if (sizeFontTable != gostOptions.GetSizeFontOfOST())
                {
                    AddComment("Не корректно размер шрифта для таблицы", table.Range);
                }

            }

        }

        private void CheckText()
        {
            var paragraphs = gostOptions.GetCurrentDocument().Paragraphs;

            bool titlePageflag = true;

            foreach (Word.Paragraph p in paragraphs)
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
                       if (p.Range.Tables.Count == 0)
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
                                CheckTextPargraph(p);
                        }
                    }
                }
            }
        }

        private void CheckTitlePageDocument(Word.Paragraph p)
        {

            if (p.Range.Text != "\r")
            {
                if (p.Range.Text == "Федеральное государственное бюджетное образовательное учреждение высшего образования")
                {
                    var fontName = p.Range.Font.Name;
                    var fontSize = p.Range.Font.Size;
                    if (fontName != gostOptions.GetNameFontOfOST() || fontSize != gostOptions.GetSizeFontOfOST())
                    {
                        if (fontSize != gostOptions.GetSizeFontOfOST())
                        {
                            AddComment("Некорректен размер шрифта, должен стоять" + gostOptions.GetSizeFontOfOST(), p.Range);
                        }
                        if (fontName != gostOptions.GetNameFontOfOST())
                        {
                            AddComment("Не корректен тип шрифта, должен стоять " + gostOptions.GetNameFontOfOST(), p.Range);
                        }
                    }

                }
            }

        }

        private void CheckTextPargraph(Word.Paragraph p)
        {
            Ribbon1 ribbon = new Ribbon1();

            Word.Range range = p.Range;
            if (range.OMaths.Count <= 0)
            {
                var leftIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.LeftIndent), 1);
                var firstLineIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.FirstLineIndent), 1);
                var rightIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.RightIndent), 1);
                var intervalBefore = p.SpaceBefore;
                var intervalAfter = p.SpaceAfter;
                // var leftIndent = Range.Paragraphs.LeftIndent;
                //  var firstLineIndent = curre.Paragraphs.FirstLineIndent;
                var nameFont = range.Font.Name;
                var colorFont = range.Font.ColorIndex.ToString();
                var lineSpacing = p.LineSpacing / 12;
                var fontSize = range.Font.Size;
                var aligmentText = p.Alignment;
                var text = range.Text;

                if ((nameFont != gostOptions.GetNameFontOfOST() ||
                    (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight") ||
                    lineSpacing != gostOptions.GetLineSpacingOFOST() ||
                    fontSize != gostOptions.GetSizeFontOfOST() ||
                    leftIndent != gostOptions.GetLeftIndent() ||
                    firstLineIndent != gostOptions.GetFirstLineIndent() ||
                    rightIndent != gostOptions.GetRightIndent() ||
                    intervalBefore != gostOptions.GetIntervalBefore() ||
                    intervalAfter != gostOptions.GetIntervalAfter() ||
                    aligmentText.ToString() != gostOptions.alignmentText) && !(text == "/\r" || text == "\r"))
                {
                    StringBuilder textForComment = new StringBuilder("Текст не корректно оформлен согласно ОС ТУСУР 01-2013: \n");
                    if (leftIndent != gostOptions.GetLeftIndent())
                    {
                        textForComment.Append("\n Отступ слева установлен не корректно " + leftIndent + " необходимо установить отступ слева равный = " + gostOptions.GetLeftIndent());
                    }
                    if (firstLineIndent != gostOptions.GetFirstLineIndent())
                    {
                        textForComment.Append("\n Отступ первой строки установлен не корректно " + firstLineIndent + " необходимо установить отступ первой строки равный = " + gostOptions.GetFirstLineIndent());
                    }
                    if (rightIndent != gostOptions.GetRightIndent())
                    {
                        textForComment.Append("\n Отступ справа установлен не корректно " + rightIndent + " необходимо установить отступ справа равный = " + gostOptions.GetRightIndent());
                    }
                    if (nameFont != gostOptions.GetNameFontOfOST())
                    {
                        textForComment.Append("\n Текущие имя шрифта " + nameFont + ", а должен быть установлен " + gostOptions.GetNameFontOfOST());
                    }
                    if (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight")
                    {
                        textForComment.Append("\n Не корректен цвет шрифта, должен быть " + gostOptions.GetColorFontOfOST().ToString());
                    }
                    if (lineSpacing != gostOptions.GetLineSpacingOFOST())
                    {
                        textForComment.Append("\n Не корректно установлен межстрочный интервал");
                    }
                    if (gostOptions.GetSizeFontOfOST() != fontSize)
                    {
                        textForComment.Append("\n Не корректен размер шрифта, установлен " + fontSize + ", а должен быть установлен " + gostOptions.GetSizeFontOfOST());
                    }
                    if (intervalBefore != gostOptions.GetIntervalBefore())
                    {
                        textForComment.Append("\n Интервал перед установлен не корректно " + intervalBefore + " необходимо установить интервал перед равный = " + gostOptions.GetIntervalBefore());
                    }
                    if (intervalAfter != gostOptions.GetIntervalAfter())
                    {
                        textForComment.Append("\n Интервал после установлен не корректно " + intervalAfter + " необходимо установить интервал после равный = " + gostOptions.GetIntervalAfter());
                    }
                    if (aligmentText.ToString() != gostOptions.alignmentText)
                    {
                        textForComment.Append("\n Выравнивание текста установлено " + ribbon.ConvertedIndexForComboBoxAlignmentText(aligmentText.ToString()) + " необходимо установить выравнивание текста " + ribbon.ConvertedIndexForComboBoxAlignmentText(gostOptions.alignmentText));
                    }

                    AddComment(textForComment.ToString(), range);
                }

            }

        }

        private void AddComment(String textComment, Word.Range range)
        {
            if (range.Text != "\r\a")
            {
                currentDocument.Comments.Add(range, textComment);
            }
        }

        private void CheckHeaderText(Word.Paragraph p)
        {
            if (p.Range.Font.Name != gostOptions.GetNameFontOfOST() && p.Range.Font.Size != gostOptions.GetSizeFontOfOST())
            {
                StringBuilder textForComment = new StringBuilder("Не корректно задан заголовок ");
                if (p.Range.Font.Name != gostOptions.GetNameFontOfOST())
                {
                    textForComment.Append(" \n  Не корректен выбранный шрифт " + p.Range.Font.Name + ", а необходим " + gostOptions.GetNameFontOfOST());
                }
                if (p.Range.Font.Size != gostOptions.GetSizeFontOfOST())
                {
                    textForComment.Append("\n Не корректно указан размер шрифта " + p.Range.Font.Size + ", необходимо установить " + gostOptions.GetSizeFontOfOST());
                }
                AddComment(textForComment.ToString(), p.Range);
            }

        }

        private void CheckDocumentContents(Word.Paragraph p)
        {
            //проверить текст
        }

        private void CheckMathText(Word.Paragraph p)
        {
            //придумать проверки
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

        private bool IsTitlePage(Word.Paragraph p)
        {
            var numberPage = GetPageNumber(p.Range);
            if (numberPage == 1)
            {
                return true;
            }
            return false;
        }

        private bool IsContentPage(Word.Paragraph p)
        {
            return false; // проверка оглавления
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

        private string GetHeading(Word.Paragraph paragraph)
        {
            string heading = "";

            heading = paragraph.Range.ListFormat.ListString;
            if (string.IsNullOrEmpty(heading))
            {
                heading = paragraph.Range.Text;
                heading = Regex.Replace(heading, "\\s+$", "");
            }
            return heading;
        }

        private int GetPageNumber(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
        }

    }
}
