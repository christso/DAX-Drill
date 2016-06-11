using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace DG2NTT.DaxDrill
{
    [Serializable]
    public class SelectedColumn
    {
        [XmlElement("name")]
        public string Name;
        [XmlElement("expression")]
        public string Expression;

        public SelectedColumn()
        {
        }
    }
}
