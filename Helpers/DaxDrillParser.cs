using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tabular = Microsoft.AnalysisServices.Tabular;

namespace DG2NTT.DaxDrill.Helpers
{
    public class DaxDrillParser
    {
        public string BuildFilterCommandText(Dictionary<string, string> excelDic, string serverName, string databaseName)
        {
            using (var tabular = new TabularHelper(serverName, databaseName))
            {
                tabular.Connect();

                var daxFilter = ConvertExcelDrillToDaxFilter(excelDic);

                string commandText = "";
                foreach (var item in daxFilter)
                {
                    if (commandText != "")
                        commandText += ",\n";
                    var table = tabular.FindTable(item.TableName);
                    var column = table.Columns.Find(item.ColumnName);
                    commandText += BuildColumnCommandText(column, item);
                }

                tabular.Disconnect();
                return commandText;
            }
        }

        public string BuildColumnCommandText(Tabular.Column column, DaxFilter item)
        {
            string commandText;
            switch (column.DataType)
            {
                case Tabular.DataType.String:
                    commandText = string.Format("{0}[{1}] = \"{2}\"", item.TableName, item.ColumnName, item.Value);
                    break;
                case Tabular.DataType.Int64:
                case Tabular.DataType.Decimal:
                case Tabular.DataType.Double:
                    commandText = string.Format("{0}[{1}] = {2}", item.TableName, item.ColumnName, item.Value);
                    break;
                default:
                    commandText = string.Format("{0}[{1}] = \"{2}\"", item.TableName, item.ColumnName, item.Value);
                    break;
            }
            return commandText;
        }

        public List<DaxFilter> ConvertExcelDrillToDaxFilter(
            Dictionary<string, string> inputDic)
        {

            var output = new List<DaxFilter>();

            foreach (var pair in inputDic)
            {
                string column = GetColumnFromPivotFIeld(pair.Key);
                string table = GetTableFromPivotField(pair.Key);
                string value = GetValueFromPivotItem(pair.Value);
                output.Add(new DaxFilter() { TableName = table, ColumnName = column, Value = value });
            }
            return output;
        }

        public string GetTableFromPivotField(string input)
        {
            string[] split = input.Split('.');
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // [Usage].[Inbound or Outbound].[Inbound or Outbound]
        public string GetColumnFromPivotFIeld(string input)
        {
            string[] split = input.Split('.');
            string output = split[1];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // "[Usage].[Inbound or Outbound].&[Inbound]
        public string GetValueFromPivotItem(string input)
        {
            var itemIndex = input.IndexOf('&');
            string output = input.Substring(itemIndex, input.Length - itemIndex);
            output = output.Substring(2, output.Length - 3);
            return output;
        }

        //Gross Billed Sum
        public string GetTableFromPivotCell(string input)
        {
            string output = "";
            return output;
        }
    }
}
