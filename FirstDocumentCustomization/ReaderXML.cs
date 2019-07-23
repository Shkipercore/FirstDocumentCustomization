using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace FirstDocumentCustomization
{
    public class ReaderXML
    {
        private XmlDocument xDoc;

        public Dictionary<string,Dictionary<string, string>> GetDictionaryPropertyOfXML( List<String>  listTagName)
        {
            var dictionaryDictionaries = new Dictionary<string, Dictionary<string, string>>();

            string path = Environment.CurrentDirectory + "\\Config.xml";
            MessageBox.Show(path);
            if (!File.Exists(path))
            {
                File.Create(path);
            }

            XDocument xdoc = XDocument.Load(path);

            //XDocument xdoc = XDocument.Load("Config.xml");
            XElement root = xdoc.Element("ConfigSettings");
        
            foreach (XElement xe in root.Elements("Settings").ToList())
            {
                foreach(var tagName in listTagName)
                {
                    if (xe.Attribute("name").Value.Equals(tagName))
                    {
                        var propertyDictionary = new Dictionary<string, string>();

                        propertyDictionary.Add("nameFontOfOST", xe.Element("nameFontOfOST").Value);
                        propertyDictionary.Add("sizeFontOfOST", xe.Element("sizeFontOfOST").Value);
                        propertyDictionary.Add("lineSpacingOfOST", xe.Element("lineSpacingOfOST").Value);
                        propertyDictionary.Add("leftIndentOfOST", xe.Element("leftIndentOfOST").Value);
                        propertyDictionary.Add("rightIndentOfOST", xe.Element("rightIndentOfOST").Value);
                        propertyDictionary.Add("firstLineIndentOfOST", xe.Element("firstLineIndentOfOST").Value);
                        propertyDictionary.Add("colorFontOfOST", xe.Element("colorFontOfOST").Value);
                        propertyDictionary.Add("alignmentTextOfOST", xe.Element("alignmentTextOfOST").Value);
                        propertyDictionary.Add("intervalBeforeOfOST", xe.Element("intervalBeforeOfOST").Value);
                        propertyDictionary.Add("intervalAfterOfOST", xe.Element("intervalAfterOfOST").Value);

                        //propertyDictionary.Add("name", xe.Element("name").Value);

                        dictionaryDictionaries.Add(tagName, propertyDictionary);
                    }
                }
                
            }

            return dictionaryDictionaries;
        }
    }
}
