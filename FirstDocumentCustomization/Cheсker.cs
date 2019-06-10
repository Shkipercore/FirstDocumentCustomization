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
                    footerRange.Text = footerRange.Text + "  " + " Не корректно офрмлен нижний колонтитул : стоит шрифт " + fontName + "  ,а должен стоять " + gostOptions.GetNameFontForFooterOfOST() + ", " + " стоит ориентация " + aligment.ToString() + ", должна стоять " + gostOptions.alignmentFooter;

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
                        AddComment("Необходимо установить орентацию по " + WdParagraphAlignment.wdAlignParagraphCenter.ToString() + " для данного рисунка", shap.Range);
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
                    AddComment("Не коректно выбран шрифт для таблицы", table.Range);
                }
                if (sizeFontTable != gostOptions.GetSizeFontOfOST())
                {
                    AddComment("Не коректно выбран шрифт для таблицы", table.Range);
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
            Word.Range range = p.Range;
            if (range.OMaths.Count <= 0)
            {
                var leftIndent = p.LeftIndent;  // дописать
                var firstLineIndent = p.FirstLineIndent;
                // var leftIndent = Range.Paragraphs.LeftIndent;
                //  var firstLineIndent = curre.Paragraphs.FirstLineIndent;
                var nameFont = range.Font.Name;
                var colorFont = range.Font.Color.ToString();
                var lineSpacing = p.LineSpacing;
                var fontSize = range.Font.Size;
                var aligment = p.Alignment;
                if (nameFont != gostOptions.GetNameFontOfOST() &&
                    colorFont != gostOptions.GetColorFontOfOST() &&
                    lineSpacing != gostOptions.GetLineSpacingOFOST() &&
                    fontSize != gostOptions.GetSizeFontOfOST())
                {
                    StringBuilder texForComment = new StringBuilder("Текст не корректно оформлен согласно OST TUSUR: \n");
                    if (leftIndent != gostOptions.GetLeftIndent())
                    {
                        texForComment.Append("\n Выступ выстовлен не корректно " + leftIndent + " необходимо установить выступ равный = " + gostOptions.GetLeftIndent());
                    }
                    if (firstLineIndent != gostOptions.GetFirstLineIndent())
                    {
                        texForComment.Append("\n Отступ первой строки выстовлен не корректно " + leftIndent + " необходимо установить отступ первой строки равный = " + gostOptions.GetLeftIndent());
                    }
                    if (nameFont != gostOptions.GetNameFontOfOST())
                    {
                        texForComment.Append("\n Текущие имя шрифта " + nameFont + ", а должен быть установлен" + gostOptions.GetNameFontOfOST());
                    }
                    if (colorFont != gostOptions.GetColorFontOfOST())
                    {
                        texForComment.Append("\n Не коректен цвет шрифта, должен быть " + gostOptions.GetColorFontOfOST().ToString());
                    }
                    if (lineSpacing != gostOptions.GetLineSpacingOFOST())
                    {
                        texForComment.Append("\n Не корректно установлен межстрочный интервал");
                    }
                    if (gostOptions.GetSizeFontOfOST() != fontSize)
                    {
                        texForComment.Append("\n Не коррректен размер шрифта, установлен" + fontSize + ", а должен быть установлен " + gostOptions.GetSizeFontOfOST());
                    }

                    AddComment(texForComment.ToString(), range);
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
                    textForComment.Append("\n Не корректно указан размер шрифта " + p.Range.Font.Size + ", необходимо установить " + gostOptions.GetNameFontOfOST());
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
