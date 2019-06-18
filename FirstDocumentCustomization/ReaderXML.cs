using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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


            XDocument xdoc = XDocument.Load("Config.xml");
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

                        dictionaryDictionaries.Add(tagName, propertyDictionary);
                    }
                }
                
            }

            return dictionaryDictionaries;

        }

    }
}
