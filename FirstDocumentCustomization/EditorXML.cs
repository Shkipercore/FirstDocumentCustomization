using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace FirstDocumentCustomization
{
    public class EditorXML
    {
        //private XmlDocument xDoc;
        //string m_exePath = Environment.CurrentDirectory;
        string userName = Environment.UserName;

        public void CreateNode(string nodeAttributeName)
        {
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load("C:\\Users\\" + userName + "\\AppData\\Local\\FirstDocumentCustomization\\Config.xml");

            XmlElement xRoot = xDoc.DocumentElement;
            // создаем объект settings
            XmlElement settingsElem = xDoc.CreateElement("Settings");
            // создаем атрибут name
            XmlAttribute nameAttr = xDoc.CreateAttribute("name");

            // создаем элементы
            XmlElement nameFontOfOSTElem = xDoc.CreateElement("nameFontOfOST");
            XmlElement colorFontOfOSTElem = xDoc.CreateElement("colorFontOfOST");
            XmlElement lineSpacingElem = xDoc.CreateElement("lineSpacingOfOST");
            XmlElement sizeFontOfOSTElem = xDoc.CreateElement("sizeFontOfOST");
            XmlElement widthOfOSTElem = xDoc.CreateElement("widthOfOST");
            XmlElement hightOfOSTElem = xDoc.CreateElement("hightOfOST");
            XmlElement leftIndentElem = xDoc.CreateElement("leftIndentOfOST");
            XmlElement rightIndentElem = xDoc.CreateElement("rightIndentOfOST");
            XmlElement firstLineIndent = xDoc.CreateElement("firstLineIndentOfOST");
            XmlElement nameFontForFooterOfOSTElem = xDoc.CreateElement("nameFontForFooterOfOST");
            XmlElement alignmentTextElem = xDoc.CreateElement("alignmentTextOfOST");
            XmlElement alignmentFooterElem = xDoc.CreateElement("alignmentFooter");
            XmlElement alignmentHeaderElem = xDoc.CreateElement("alignmentHeader");
            XmlElement intervalBeforeElem = xDoc.CreateElement("intervalBeforeOfOST");
            XmlElement intervalAfterElem = xDoc.CreateElement("intervalAfterOfOST");

            //создаем текстовые значения для элементов и атрибута
            XmlText nameText = xDoc.CreateTextNode(nodeAttributeName);
            XmlText nameFontText = xDoc.CreateTextNode("Microsoft Sans Serif");
            XmlText colorFontText = xDoc.CreateTextNode("Black");
            XmlText sizeFontText = xDoc.CreateTextNode("8");

            //добавляем узлы
            nameAttr.AppendChild(nameText);
            nameFontOfOSTElem.AppendChild(nameFontText);
            colorFontOfOSTElem.AppendChild(colorFontText);
            sizeFontOfOSTElem.AppendChild(sizeFontText);

            settingsElem.Attributes.Append(nameAttr);

            settingsElem.AppendChild(nameFontOfOSTElem);
            settingsElem.AppendChild(colorFontOfOSTElem);
            settingsElem.AppendChild(lineSpacingElem);
            settingsElem.AppendChild(sizeFontOfOSTElem);
            settingsElem.AppendChild(widthOfOSTElem);
            settingsElem.AppendChild(hightOfOSTElem);
            settingsElem.AppendChild(leftIndentElem);
            settingsElem.AppendChild(rightIndentElem);
            settingsElem.AppendChild(firstLineIndent);
            settingsElem.AppendChild(nameFontForFooterOfOSTElem);
            settingsElem.AppendChild(alignmentTextElem);
            settingsElem.AppendChild(alignmentFooterElem);
            settingsElem.AppendChild(alignmentHeaderElem);
            settingsElem.AppendChild(intervalBeforeElem);
            settingsElem.AppendChild(intervalAfterElem);

            xRoot.AppendChild(settingsElem);
            xDoc.Save("C:\\Users\\" + userName + "\\AppData\\Local\\FirstDocumentCustomization\\Config.xml");
        }

        public string AddElement(string tagName,
                                 string nameFont,
                                 string colorFont,
                                 string lineSpacing,
                                 string sizeFont,
                                 //string width,
                                 //string hight,
                                 string leftIndent,
                                 string rightIndent,
                                 string firstLineIndent,
                                 //string nameFontForFooter,
                                 string alignmentText,
                                 //string alignmentFooter,
                                 //string alignmentHeader
                                 string intervalBefore,
                                 string intervalAfter
                                                            )
        {

            XDocument xdoc = XDocument.Load("C:\\Users\\" + userName + "\\AppData\\Local\\FirstDocumentCustomization\\Config.xml");
            XElement root = xdoc.Element("ConfigSettings");

            foreach (XElement xe in root.Elements("Settings").ToList())
            {

                if (xe.Attribute("name").Value == tagName)

                {
                    xe.Element("nameFontOfOST").Value = nameFont;
                    xe.Element("colorFontOfOST").Value = colorFont;
                    xe.Element("lineSpacingOfOST").Value = lineSpacing;
                    xe.Element("sizeFontOfOST").Value = sizeFont;
                    //xe.Element("widthOfOST").Value = width;
                    //xe.Element("hightOfOST").Value = hight;
                    xe.Element("leftIndentOfOST").Value = leftIndent;
                    xe.Element("rightIndentOfOST").Value = rightIndent;
                    xe.Element("firstLineIndentOfOST").Value = firstLineIndent;
                    //xe.Element("nameFontForFooterOfOST").Value = nameFontForFooter;
                    xe.Element("alignmentTextOfOST").Value = alignmentText;
                    //xe.Element("alignmentFooterOfOST").Value = alignmentFooter;
                    //xe.Element("alignmentHeaderOfOST").Value = alignmentHeader;
                    xe.Element("alignmentTextOfOST").Value = alignmentText;
                    xe.Element("intervalBeforeOfOST").Value = intervalBefore;
                    xe.Element("intervalAfterOfOST").Value = intervalAfter;
                }
            }

            xdoc.Save("C:\\Users\\" + userName + "\\AppData\\Local\\FirstDocumentCustomization\\Config.xml");
            return nameFont;
        }

        public void RemoveElement(string tagName)
        {
            XDocument xdoc = XDocument.Load("C:\\Users\\" + userName + "\\AppData\\Local\\FirstDocumentCustomization\\Config.xml");
            XElement root = xdoc.Element("ConfigSettings");

            foreach (XElement xe in root.Elements("Settings").ToList())
            { 
                if (xe.Attribute("name").Value == tagName)
                {
                    xe.Remove();
                }
            }
            xdoc.Save("C:\\Users\\" + userName + "\\AppData\\Local\\FirstDocumentCustomization\\Config.xml");
        }
    }
}
