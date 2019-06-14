using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace FirstDocumentCustomization
{
    public class EditorXML
    {
        private XmlDocument xDoc;
        //string m_exePath = Environment.CurrentDirectory;

        public bool CreateNode(string nodeAttributeName)
        {
            if (1 > 0)
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load("Config.xml");

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

                //создаем текстовые значения для элементов
                XmlText nameText = xDoc.CreateTextNode(nodeAttributeName);

                //добавляем узлы
                nameAttr.AppendChild(nameText);
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

                xRoot.AppendChild(settingsElem);
                xDoc.Save("Config.xml");

            }
            return false;
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
                                 string alignmentText
                                 //string alignmentFooter,
                                 //string alignmentHeader
                                                            )
        {


            XDocument xdoc = XDocument.Load("Config.xml");
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

                }
            }

            xdoc.Save("Config.xml");
            return nameFont;

        }

        public bool RemoveElement(string tagName, string tagProperty, string nodeAttributeName)
        {
            XDocument xdoc = XDocument.Load("Config.xml");
            XElement root = xdoc.Element("ConfigSettings");

            foreach (XElement xe in root.Elements("Settings").ToList())

                if (xe.Element("name").Value == tagName)
                {
                    xe.Remove();
                }

            xdoc.Save("Config.xml");

            // выбираем узел у которого атрибут name имеет значение nodeAttributeName
            //XmlNode DeletedNode = xRoot.SelectSingleNode("Settings[@name="nodeAttributeName);
            //xRoot.RemoveChild(DeletedNode);
            //xDoc.Save("users.xml");

            return false;

        }

        public string ConvertedComboBoxAlignmentTextForIndex(string items)
        {

            switch (items)
            {
                case "По левому краю":
                    return "0";

                case "По центру":
                    return "1";

                case "По правому краю":
                    return "2";

                case "По ширине":
                    return "3";
            }

            return items;

        }


    }

}
