using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Ntreev.Crema.VisualStudio.ResXGeneration
{
    [XmlRoot("Settings")]
    public struct ResXInfo
    {
        public string Address
        {
            get; set;
        }

        [XmlElement("DataBaseName")]
        public string DataBase
        {
            get; set;
        }

        public string ProjectInfo
        {
            get; set;
        }
    }
}
