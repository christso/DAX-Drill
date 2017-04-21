using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaxDrill
{
    public class Constants
    {
        public const string DaxDrillXmlSchemaSpace = "http://schemas.microsoft.com/daxdrill";
        public const string AppName = "DAX Drill";
        public const string TableXpath = "/x:daxdrill/x:table";
        public const string MeasureXpath = "/x:daxdrill/x:measure";
        public const string RootXmlNode = "daxdrill";
        public const int DefaultMaxDrillThroughRecords = 99999;
    }
}
