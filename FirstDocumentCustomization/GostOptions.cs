using Word = Microsoft.Office.Interop.Word;
using System;
using Microsoft.Office.Interop.Word;

namespace FirstDocumentCustomization
{
    public class GostOptions
    {
        private Document currentDocument;

        public Document GetCurrentDocument()
        {
            return currentDocument;
        }

        public void SetCurrentDocument(Document value)
        {
            currentDocument = value;
        }

        private string nameFontOfOST;

        public string GetNameFontOfOST()
        {
            return nameFontOfOST;
        }

        public void SetNameFontOfOST(string value)
        {
            nameFontOfOST = value;
        }

        private string nameFontForFooterOfOST;

        public string GetNameFontForFooterOfOST()
        {
            return nameFontForFooterOfOST;
        }

        public void SetNameFontForFooterOfOST(string value)
        {
            nameFontForFooterOfOST = value;
        }

        private string nameFontForHeaderOfOST;

        public string GetNameFontForHeaderOfOST()
        {
            return nameFontForHeaderOfOST;
        }

        public void SetNameFontForHeaderOfOST(string value)
        {
            nameFontForHeaderOfOST = value;
        }

        private string colorFontOfOST;

        public string GetColorFontOfOST()
        {
            return colorFontOfOST;
        }

        public void SetColorFontOfOST(string value)
        {
            colorFontOfOST = value;
        }

        private float lineSpacingOFOST;

        public float GetLineSpacingOFOST()
        {
            return lineSpacingOFOST;
        }

        public void SetLineSpacingOFOST(float value)
        {
            lineSpacingOFOST = value;
        }

        private float sizeFontOfOST;

        public float GetSizeFontOfOST()
        {
            return sizeFontOfOST;
        }

        public void SetSizeFontOfOST(float value)
        {
            sizeFontOfOST = value;
        }

        private float widthOfOST;

        public float GetWidthOfOST()
        {
            return widthOfOST;
        }

        public void SetWidthOfOST(float value)
        {
            widthOfOST = value;
        }

        private float hightOfOST;

        public float GetHightOfOST()
        {
            return hightOfOST;
        }

        public void SetHightOfOST(float value)
        {
            hightOfOST = value;
        }

        private float leftIndentOfOST;

        public float GetLeftIndent()
        {
            return leftIndentOfOST;
        }

        public void SetLefttIndent(float value)
        {
            leftIndentOfOST = value;
        }

        private float rightIndentOfOST;

        public float GetRightIndent()
        {
            return rightIndentOfOST;
        }

        public void SetRightIndent(float value)
        {
            rightIndentOfOST = value;
        }

        private float firstLineIndentOfOST;

        public float GetFirstLineIndent()
        {
            return firstLineIndentOfOST;
        }

        public void SetFirstLineIndent(float value)
        {
            firstLineIndentOfOST = value;
        }

        public string alignmentText;

        public string alignmentFooter;

        public string alignmentHeader;

        private float intervalBeforeOfOST;

        public float GetIntervalBefore()
        {
            return intervalBeforeOfOST;
        }

        public void SetIntervalBefore(float value)
        {
            intervalBeforeOfOST = value;
        }

        private float intervalAfterOfOST;

        public float GetIntervalAfter()
        {
            return intervalBeforeOfOST;
        }

        public void SetIntervalAfter(float value)
        {
            intervalBeforeOfOST = value;
        }

        public GostOptions() { }
        public GostOptions(Word.Document document,
                            String nameFont,
                            String colorFont,
                            float lineSpacing,
                            float sizeFont,
                            float width,
                            float highest,
                            float leftIndent,
                            float rightIndent,
                            float firstLineIndent,
                            String fontFooter,
                            String alignment,
                            String alignmentHeader,
                            String alignmentFooter,
                            float intervalBefore,
                            float intervalAfter)
        {
            this.currentDocument = document;
            this.nameFontOfOST = nameFont;
            this.colorFontOfOST = colorFont;
            this.sizeFontOfOST = sizeFont;
            this.widthOfOST = width;
            this.hightOfOST = highest;
            this.leftIndentOfOST = leftIndent;
            this.rightIndentOfOST = rightIndent;
            this.firstLineIndentOfOST = firstLineIndent;
            this.nameFontForFooterOfOST = fontFooter;
            this.alignmentText = alignment;
            this.alignmentFooter = alignmentFooter;
            this.intervalBeforeOfOST = intervalBefore;
            this.intervalAfterOfOST = intervalAfter;
        }
    }
}
