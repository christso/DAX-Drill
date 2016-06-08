using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.Helpers
{
    public class DaxDrillParser
    {
        public string BuildFilterCommandText(Dictionary<string, string> excelDic)
        {
            var daxFilter = ConvertExcelDrillToDaxFilter(excelDic);

            string commandText = "";
            foreach (var item in daxFilter)
            {
                if (commandText != "")
                    commandText += ",\n";
                commandText += string.Format("{0}[{1}] = \"{2}\"", item.TableName, item.ColumnName, item.Value);
            }
            return commandText;
        }

        public List<DaxFilter> ConvertExcelDrillToDaxFilter(
            Dictionary<string, string> inputDic)
        {

            var output = new List<DaxFilter>();

            foreach (var pair in inputDic)
            {
                string column = ParseExcelPivotFieldColumn(pair.Key);
                string table = ParseExcelPivotFieldTable(pair.Key);
                string value = ParseExcelPivotItem(pair.Value);
                output.Add(new DaxFilter() { TableName = table, ColumnName = column, Value = value });
            }
            return output;
        }

        public string ParseExcelPivotFieldTable(string input)
        {
            string[] split = input.Split('.');
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // [Usage].[Inbound or Outbound].[Inbound or Outbound]
        public string ParseExcelPivotFieldColumn(string input)
        {
            string[] split = input.Split('.');
            string output = split[1];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // "[Usage].[Inbound or Outbound].&[Inbound]
        public string ParseExcelPivotItem(string input)
        {
            var itemIndex = input.IndexOf('&');
            string output = input.Substring(itemIndex, input.Length - itemIndex);
            output = output.Substring(2, output.Length - 3);
            return output;
        }
    }
}
