using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tabular = Microsoft.AnalysisServices.Tabular;

namespace DG2NTT.DaxDrill.DaxHelpers
{
    public class DaxDrillParser
    {
        public string BuildQueryText(TabularHelper tabular, Dictionary<string, string> excelDic, string measureName,
            IEnumerable<SelectedColumn> selectedColumns)
        {
            string filterText = BuildFilterCommandText(excelDic, tabular);
            var measure = tabular.GetMeasure(measureName);

            string commandText = string.Format("TOPN ( 99999, {0} )", measure.Table.Name);

            if (selectedColumns != null)
            {
                commandText = string.Format("SELECTCOLUMNS ( {0}, {{0}} )", commandText);

                string selectColumnsText = BuildSelectText(selectedColumns);

                commandText = string.Format(commandText, selectColumnsText);
            }

            if (!string.IsNullOrWhiteSpace(filterText))
                commandText += string.Format(",\n{0}", filterText);

            commandText = string.Format("EVALUATE CALCULATETABLE ( {0} )", commandText);

            return commandText;
        }

        public string BuildSelectText(IEnumerable<SelectedColumn> selectedColumns)
        {
            string result = string.Empty;
            foreach (var column in selectedColumns)
            {
                if (result != string.Empty)
                    result += ",";

                result += string.Format("\n\"{0}\", {1}", column.Name, column.Expression);
            }
            return result;
        }

        public string BuildQueryText(TabularHelper tabular, Dictionary<string, string> excelDic, string measureName)
        {
            return BuildQueryText(tabular, excelDic, measureName, null);
        }
        public string BuildFilterCommandText(Dictionary<string, string> excelDic, TabularHelper tabular)
        {
            var daxFilter = ConvertExcelDrillToDaxFilter(excelDic);

            string commandText = "";
            foreach (var item in daxFilter)
            {
                if (commandText != "")
                    commandText += ",\n";
                var table = tabular.GetTable(item.TableName);
                var column = table.Columns.Find(item.ColumnName);
                commandText += BuildColumnCommandText(column, item);
            }

            return commandText;
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
                string column = GetColumnFromPivotField(pair.Key);
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
        public string GetColumnFromPivotField(string input)
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

        // example input: [Measures].[Gross Billed Sum]
        public string GetMeasureFromPivotItem(string input)
        {
            return GetColumnFromPivotField(input);
        }
    }
}
