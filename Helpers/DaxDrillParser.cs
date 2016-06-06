using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.Helpers
{
    public class DaxDrillParser
    {
        public Dictionary<string, string> ConvertExcelDrillToDaxFilterDiciontary(
            Dictionary<string, string> inputDic)
        {

            var outputDic = new Dictionary<string, string>();

            return outputDic;
        }

        // [Usage].[Inbound or Outbound].[Inbound or Outbound]
        public string ParseExcelPivotField(string input)
        {
            string[] split = input.Split('.');
            
            return split[2];
        }

        // "[Usage].[Inbound or Outbound].&[Inbound]
        public string ParseExcelPivotItem(string input)
        {
            return "";
        }
    }
}
