using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

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
                AddComment("Некорректно указаны размеры страницы, указана ширина " + widthPage + ", необходимо установить ширину " + gostOptions.GetWidthOfOST(), rangeForAddCommit);
            }

            if (gostOptions.GetHightOfOST() != highestPage)
            {
                AddComment("Некорректно указаны размеры страницы, указана высота " + highestPage + ", необходимо установить высоту " + gostOptions.GetHightOfOST(), rangeForAddCommit);
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

                    headerRange.Text = headerRange.Text + "  " + " Некорректно офрмлен верхний колонтитул: установлен шрифт " + fontName + ", необходимо установить " + gostOptions.GetNameFontForHeaderOfOST() + ", " + "установлена ориентация " + aligment.ToString() + ", необходимо установить " + gostOptions.alignmentHeader;

                }
            }

            foreach (Word.Section wordSection in currentDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                var fontName = footerRange.Font.Name;
                var aligment = footerRange.Paragraphs.Alignment;
                if (!gostOptions.GetNameFontForFooterOfOST().Equals(fontName) && !gostOptions.alignmentFooter.Equals(aligment))
                {
                    footerRange.Text = footerRange.Text + "  " + " Некорректно оформлен нижний колонтитул: установлен шрифт " + fontName + ", необходимо установить " + gostOptions.GetNameFontForFooterOfOST() + ", " + "установлена ориентация " + aligment.ToString() + ", необходимо установить " + gostOptions.alignmentFooter;

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

                if (fontNameTable != gostOptions.GetNameFontOfOST() && fontNameTable != "")
                {
                    AddComment("Некорректно установлен шрифт таблицы", table.Range);
                }
                if (sizeFontTable != gostOptions.GetSizeFontOfOST() && sizeFontTable != 9999999)
                {
                    AddComment("Некорректно установлен размер шрифта таблицы", table.Range);
                }
            }
        }

        private void CheckText()
        {
            var paragraphs = gostOptions.GetCurrentDocument().Paragraphs;

            FormProgressBar formProgressBar = new FormProgressBar();
            formProgressBar.Show();

            formProgressBar.progressBar1.Maximum = paragraphs.Count;
            formProgressBar.progressBar1.Value = 0;

            foreach (Word.Paragraph p in paragraphs)
            {

                if (IsTitlePage(p))
                {
                    CheckTitlePageDocument(p);
                }
                else
                {
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
                            {
                                if (IsSignatureTable(p))
                                {
                                    CheckSignatureTable(p);
                                }
                                else
                                {
                                    if (IsSignatureImage(p))
                                    {
                                        CheckSignatureImage(p);
                                    }
                                    else
                                    {
                                        CheckTextPargraph(p);
                                    }
                                }
                            }
                       }
                    }
                }

                formProgressBar.progressBar1.Value++;
            }

            if (formProgressBar.progressBar1.Value == formProgressBar.progressBar1.Maximum)
            {
                formProgressBar.Close();
            }
        }

        private void CheckTitlePageDocument(Word.Paragraph p)
        {
            var text = p.Range.Text;

            if (!(text == "/\r" || text == "\r"))
            {
                    var fontName = p.Range.Font.Name;
                    var fontSize = p.Range.Font.Size;
                    if (fontName != gostOptions.GetNameFontOfOST() || fontSize != gostOptions.GetSizeFontOfOST())
                    {
                        if (fontSize != gostOptions.GetSizeFontOfOST() && fontSize != 9999999)
                        {
                            AddComment("Некорректен размер шрифта, необходимо установить " + gostOptions.GetSizeFontOfOST(), p.Range);
                        }
                        if (fontName != gostOptions.GetNameFontOfOST() && fontName != "")
                        {
                            AddComment("Некорректен тип шрифта, необходимо установить " + gostOptions.GetNameFontOfOST(), p.Range);
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

                //TestRegMatch();

                if ((nameFont != gostOptions.GetNameFontOfOST() ||
                    (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight" && colorFont != "9999999") ||
                    lineSpacing != gostOptions.GetLineSpacingOFOST() ||
                    fontSize != gostOptions.GetSizeFontOfOST() ||
                    leftIndent != gostOptions.GetLeftIndent() ||
                    firstLineIndent != gostOptions.GetFirstLineIndent() ||
                    rightIndent != gostOptions.GetRightIndent() ||
                    intervalBefore != gostOptions.GetIntervalBefore() ||
                    intervalAfter != gostOptions.GetIntervalAfter() ||
                    aligmentText.ToString() != gostOptions.alignmentText) &&
                    !(text == "/\r" || text == "\r" || text == "\f\r" || text =="\u0001\r"))
                {
                    StringBuilder textForComment = new StringBuilder("Оформление текста не соответствует ОС ТУСУР 01-2013: \n");
                    if (leftIndent != gostOptions.GetLeftIndent())
                    {
                        textForComment.Append("\n Отступ слева установлен некорректно " + leftIndent + " необходимо установить отступ слева равный " + gostOptions.GetLeftIndent());
                    }
                    if (firstLineIndent != gostOptions.GetFirstLineIndent())
                    {
                        textForComment.Append("\n Отступ первой строки установлен некорректно " + firstLineIndent + " необходимо установить отступ первой строки равный " + gostOptions.GetFirstLineIndent());
                    }
                    if (rightIndent != gostOptions.GetRightIndent())
                    {
                        textForComment.Append("\n Отступ справа установлен некорректно " + rightIndent + " необходимо установить отступ справа равный " + gostOptions.GetRightIndent());
                    }
                    if (nameFont != gostOptions.GetNameFontOfOST())
                    {
                        textForComment.Append("\n Текущее имя шрифта " + nameFont + ", необходимо установить " + gostOptions.GetNameFontOfOST());
                    }
                    if (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight" && colorFont != "9999999")
                    {
                        textForComment.Append("\n Некорректен цвет шрифта, необходимо установить " + gostOptions.GetColorFontOfOST().ToString());
                    }
                    if (lineSpacing != gostOptions.GetLineSpacingOFOST())
                    {
                        textForComment.Append("\n Некорректно установлен межстрочный интервал");
                    }
                    if (gostOptions.GetSizeFontOfOST() != fontSize)
                    {
                        textForComment.Append("\n Некорректен размер шрифта, установлен " + fontSize + ", необходимо установить " + gostOptions.GetSizeFontOfOST());
                    }
                    if (intervalBefore != gostOptions.GetIntervalBefore())
                    {
                        textForComment.Append("\n Интервал перед установлен некорректно " + intervalBefore + " необходимо установить интервал перед равный " + gostOptions.GetIntervalBefore());
                    }
                    if (intervalAfter != gostOptions.GetIntervalAfter())
                    {
                        textForComment.Append("\n Интервал после установлен некорректно " + intervalAfter + " необходимо установить интервал после равный " + gostOptions.GetIntervalAfter());
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
                StringBuilder textForComment = new StringBuilder("Некорректно задан заголовок ");
                if (p.Range.Font.Name != gostOptions.GetNameFontOfOST())
                {
                    textForComment.Append(" \n  Некорректен выбранный шрифт " + p.Range.Font.Name + ", необходимо установить " + gostOptions.GetNameFontOfOST());
                }
                if (p.Range.Font.Size != gostOptions.GetSizeFontOfOST())
                {
                    textForComment.Append("\n Некорректно указан размер шрифта " + p.Range.Font.Size + ", необходимо установить " + gostOptions.GetSizeFontOfOST());
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
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;
            int numberOfTitlePages = Convert.ToInt32(ribbon.editBoxNumberOfTitlePages.Text);
            var numberPage = GetPageNumber(p.Range);

            if (numberPage <= numberOfTitlePages)
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

        private bool IsSignatureImage(Word.Paragraph p)
        {

            if (Regex.IsMatch(p.Range.Text, "^Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*") || Regex.IsMatch(p.Range.Text, "^/ Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*"))
            {
                return true;
            }

            return false;
        }

        private void CheckSignatureImage(Word.Paragraph p)
        {
            Word.Range range = p.Range;
            Ribbon1 ribbon = new Ribbon1();

            var leftIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.LeftIndent), 1);
            var firstLineIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.FirstLineIndent), 1);
            var rightIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.RightIndent), 1);
            var intervalBefore = p.SpaceBefore;
            var intervalAfter = p.SpaceAfter;
            var nameFont = range.Font.Name;
            var colorFont = range.Font.ColorIndex.ToString();
            var lineSpacing = p.LineSpacing / 12;
            var fontSize = range.Font.Size;
            var aligmentText = p.Alignment;
            var text = range.Text;

            if (nameFont != gostOptions.GetNameFontOfOST() ||
               (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight" && colorFont != "9999999") ||
               lineSpacing != gostOptions.GetLineSpacingOFOST() ||
               fontSize != gostOptions.GetSizeFontOfOST() ||
               leftIndent != gostOptions.GetLeftIndent() ||
               firstLineIndent != 0 ||
               rightIndent != gostOptions.GetRightIndent() ||
               intervalBefore != gostOptions.GetIntervalBefore() ||
               intervalAfter != gostOptions.GetIntervalAfter() ||
               aligmentText != WdParagraphAlignment.wdAlignParagraphCenter ||
               (Regex.IsMatch(p.Range.Text, "^Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*([.]{1}\r)$") || Regex.IsMatch(p.Range.Text, "^/ Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*([.]{1}\r)$")))
            {
                StringBuilder textForComment = new StringBuilder("Оформление подписи к рисунку не соответствует ОС ТУСУР 01-2013: \n");
                if (leftIndent != gostOptions.GetLeftIndent())
                {
                    textForComment.Append("\n Отступ слева установлен некорректно " + leftIndent + " необходимо установить отступ слева равный " + gostOptions.GetLeftIndent());
                }
                if (firstLineIndent != 0)
                {
                    textForComment.Append("\n Отступ первой строки установлен некорректно " + firstLineIndent + " необходимо установить отступ первой строки равный 0");
                }
                if (rightIndent != gostOptions.GetRightIndent())
                {
                    textForComment.Append("\n Отступ справа установлен некорректно " + rightIndent + " необходимо установить отступ справа равный " + gostOptions.GetRightIndent());
                }
                if (nameFont != gostOptions.GetNameFontOfOST())
                {
                    textForComment.Append("\n Текущие имя шрифта " + nameFont + ", необходимо установить " + gostOptions.GetNameFontOfOST());
                }
                if (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight" && colorFont != "9999999")
                {
                    textForComment.Append("\n Некорректен цвет шрифта, необходимо установить " + gostOptions.GetColorFontOfOST().ToString());
                }
                if (lineSpacing != gostOptions.GetLineSpacingOFOST())
                {
                    textForComment.Append("\n Некорректно установлен межстрочный интервал");
                }
                if (gostOptions.GetSizeFontOfOST() != fontSize)
                {
                    textForComment.Append("\n Некорректен размер шрифта, установлен " + fontSize + ", необходимо установить " + gostOptions.GetSizeFontOfOST());
                }
                if (intervalBefore != gostOptions.GetIntervalBefore())
                {
                    textForComment.Append("\n Интервал перед установлен некорректно " + intervalBefore + " необходимо установить интервал перед равный " + gostOptions.GetIntervalBefore());
                }
                if (intervalAfter != gostOptions.GetIntervalAfter())
                {
                    textForComment.Append("\n Интервал после установлен некорректно " + intervalAfter + " необходимо установить интервал после равный " + gostOptions.GetIntervalAfter());
                }
                if (aligmentText != WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    textForComment.Append("\n Выравнивание подписи к рисунку установлено " + ribbon.ConvertedIndexForComboBoxAlignmentText(aligmentText.ToString()) + ", необходимо установить выравнивание текста по центру");
                }
                if (Regex.IsMatch(p.Range.Text, "^Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*([.]{1}\r)$") || Regex.IsMatch(p.Range.Text, "^/ Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*([.]{1}\r)$"))
                {
                    textForComment.Append("\n В конце подписи к рисунку установлена точка");
                }


                AddComment(textForComment.ToString(), range);
            }
        }

        private bool IsSignatureTable(Word.Paragraph p)
        {

            if (Regex.IsMatch(p.Range.Text, "^Таблица ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*") || Regex.IsMatch (p.Range.Text, "^Продолжение таблицы ([0-9]+)([.]{1}[0-9]+).*"))
            {
                return true;
            }

            return false;
        }

        private void CheckSignatureTable(Word.Paragraph p)
        {
            Ribbon1 ribbon = new Ribbon1();
            Word.Range range = p.Range;

            var leftIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.LeftIndent), 1);
            var firstLineIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.FirstLineIndent), 1);
            var rightIndent = Math.Round(Globals.ThisAddIn.Application.PointsToCentimeters(p.RightIndent), 1);
            var intervalBefore = p.SpaceBefore;
            var intervalAfter = p.SpaceAfter;
            var nameFont = range.Font.Name;
            var colorFont = range.Font.ColorIndex.ToString();
            var lineSpacing = p.LineSpacing / 12;
            var fontSize = range.Font.Size;
            var aligmentText = p.Alignment;
            var text = range.Text;

            if ((nameFont != gostOptions.GetNameFontOfOST() ||
                (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight" && colorFont != "9999999") ||
                lineSpacing != gostOptions.GetLineSpacingOFOST() ||
                fontSize != gostOptions.GetSizeFontOfOST() ||
                leftIndent != gostOptions.GetLeftIndent() ||
                firstLineIndent != 0 ||
                rightIndent != gostOptions.GetRightIndent() ||
                intervalBefore != gostOptions.GetIntervalBefore() ||
                intervalAfter != gostOptions.GetIntervalAfter() ||
                (aligmentText != WdParagraphAlignment.wdAlignParagraphLeft && aligmentText != WdParagraphAlignment.wdAlignParagraphJustify)))
            {
                StringBuilder textForComment = new StringBuilder("Оформление подписи к таблице не соответствует ОС ТУСУР 01-2013: \n");
                if (leftIndent != gostOptions.GetLeftIndent())
                {
                    textForComment.Append("\n Отступ слева установлен некорректно " + leftIndent + " необходимо установить отступ слева равный " + gostOptions.GetLeftIndent());
                }
                if (firstLineIndent != 0)
                {
                    textForComment.Append("\n Отступ первой строки установлен некорректно " + firstLineIndent + " необходимо установить отступ первой строки равный 0");
                }
                if (rightIndent != gostOptions.GetRightIndent())
                {
                    textForComment.Append("\n Отступ справа установлен некорректно " + rightIndent + " необходимо установить отступ справа равный " + gostOptions.GetRightIndent());
                }
                if (nameFont != gostOptions.GetNameFontOfOST())
                {
                    textForComment.Append("\n Текущие имя шрифта " + nameFont + ", необходимо установить " + gostOptions.GetNameFontOfOST());
                }
                if (colorFont != ("wd" + gostOptions.GetColorFontOfOST()) && colorFont != "wdNoHighlight" && colorFont != "9999999")
                {
                    textForComment.Append("\n Некорректен цвет шрифта, необходимо установить " + gostOptions.GetColorFontOfOST().ToString());
                }
                if (lineSpacing != gostOptions.GetLineSpacingOFOST())
                {
                    textForComment.Append("\n Некорректно установлен межстрочный интервал");
                }
                if (gostOptions.GetSizeFontOfOST() != fontSize)
                {
                    textForComment.Append("\n Некорректен размер шрифта, установлен " + fontSize + ", необходимо установить " + gostOptions.GetSizeFontOfOST());
                }
                if (intervalBefore != gostOptions.GetIntervalBefore())
                {
                    textForComment.Append("\n Интервал перед установлен некорректно " + intervalBefore + " необходимо установить интервал перед равный " + gostOptions.GetIntervalBefore());
                }
                if (intervalAfter != gostOptions.GetIntervalAfter())
                {
                    textForComment.Append("\n Интервал после установлен некорректно " + intervalAfter + " необходимо установить интервал после равный " + gostOptions.GetIntervalAfter());
                }
                if (aligmentText != WdParagraphAlignment.wdAlignParagraphLeft && aligmentText != WdParagraphAlignment.wdAlignParagraphJustify)
                {
                    textForComment.Append("\n Выравнивание подписи к таблице установлено некорректно");
                }

                AddComment(textForComment.ToString(), range);
            }
        }

        public void TestRegMatch()
        {
            string st = "Рисунок 3.12 - Вкладка «Главный.\r";
            Regex.IsMatch(st, @"^Рисунок ([0-9]+)([.]{1}[0-9]+) ([–]|[-]){1} .*([.]{1}\r)$");
        }
    }
}
