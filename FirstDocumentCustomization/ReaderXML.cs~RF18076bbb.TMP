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

        public Dictionary<string, string> LoadPropertyOfXML(string tagName)
        {
            Dictionary<string,string> property = new Dictionary<string, string>();
            XDocument xdoc = XDocument.Load("Config.xml");
            XElement root = xdoc.Element("ConfigSettings");
        
            foreach (XElement xe in root.Elements("Settings").ToList())
            {
                if (xe.Attribute("name").Value == tagName)
                {
                    property.Add("nameFontOfOST", xe.Element("nameFontOfOST").Value);
                    property.Add("sizeFontOfOST", xe.Element("sizeFontOfOST").Value);
                    property.Add("lineSpacingOfOST", xe.Element("lineSpacingOfOST").Value);
                    property.Add("leftIndentOfOST", xe.Element("leftIndentOfOST").Value);
                    property.Add("firstLineIndentOfOST", xe.Element("firstLineIndentOfOST").Value);


                }
            }

            return property;

        }
    }
}
