﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace FirstDocumentCustomization
{
    public class EditorXML
    {
        private XmlDocument xDoc;

        public bool CreateNode(string nodeAttributeName)
        {
            if(1>0)
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load("Config.xml");

                XmlElement xRoot = xDoc.DocumentElement;
                // создаем объект settings
                XmlElement settingsElem = xDoc.CreateElement("settings");
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
                XmlElement firstLineIndent = xDoc.CreateElement("firstLineIndentOfOST");
                XmlElement nameFontForFooterOfOSTElem = xDoc.CreateElement("nameFontForFooterOfOST");
                XmlElement alignmentTextElem = xDoc.CreateElement("alignmentText");
                XmlElement alignmentFooterElem = xDoc.CreateElement("alignmentFooter");
                XmlElement alignmentHeaderElem = xDoc.CreateElement("alignmentHeader");

                XmlText nameText = xDoc.CreateTextNode(nodeAttributeName);
                nameAttr.AppendChild(nameText);
                settingsElem.Attributes.Append(nameAttr);
                xRoot.AppendChild(settingsElem);
                xDoc.Save("Config.xml");

            }

            return false;
        }


        public string AddElement(//string tagName,
                                 string nameFont,
                                 string colorFont,
                                 //string lineSpacing, 
                                 string sizeFont
                                 //string width,
                                 //string hight,
                                 //string leftIndent,
                                 //string firstLineIndent,
                                 //string nameFontForFooter,
                                 //string alignmentText,
                                 //string alignmentFooter,
                                 //string alignmentHeader
                                                            )
        {
            XDocument xdoc = XDocument.Load("Config.xml");
            XElement root = xdoc.Element("ConfigSettings");

            foreach (XElement xe in root.Elements("Settings").ToList())

                if (xe.Attribute("name").Value == "Курсовая")

                {
                    xe.Element("nameFontOfOST").Value = nameFont;
                    xe.Element("colorFontOfOST").Value = colorFont;
                    //xe.Element("lineSpacingOfOST").Value = lineSpacing;
                    xe.Element("sizeFontOfOST").Value = sizeFont;
                    //xe.Element("widthOfOST").Value = width;
                    //xe.Element("hightOfOST").Value = hight;
                    //xe.Element("leftIndentOfOST").Value = leftIndent;
                    //xe.Element("firstLineIndentOfOST").Value = firstLineIndent;
                    //xe.Element("nameFontForFooterOfOST").Value = nameFontForFooter;
                    //xe.Element("alignmentTextOfOST").Value = alignmentText;
                    //xe.Element("alignmentFooterOfOST").Value = alignmentFooter;
                    //xe.Element("alignmentHeaderOfOST").Value = alignmentHeader;

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
    }
}
