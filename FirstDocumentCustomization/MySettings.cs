using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FirstDocumentCustomization
{
    public class MySettings: ConfigurationSection
    {
        [ConfigurationProperty("Propa")]
        public string MyProperty { get; set; }
    }
}
