using DG2NTT.DaxDrill.ExcelHelpers;
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
        public static string BuildQueryText(TabularHelper tabular, PivotCellDictionary pivotCellDic, string measureName, int maxRecords)
        {
            return BuildQueryText(tabular, pivotCellDic, measureName, maxRecords, null);
        }

        /// <summary>
        /// Builds DAX query based on location on pivot table (specified in parameters)
        /// </summary>
        /// <param name="tabular">Tabular connection helper</param>
        /// <param name="pivotCellDic">Dictionary representing Pivot Table context filters</param>
        /// <param name="measureName">Name of DAX measure to be used in drillthrough</param>
        /// <param name="maxRecords">Maximum records to be retrieved</param>
        /// <param name="detailColumns">List of columns to be included in drill-through</param>
        /// <returns></returns>
        public static string BuildQueryText(TabularHelper tabular, PivotCellDictionary pivotCellDic, string measureName,
            int maxRecords, IEnumerable<DetailColumn> detailColumns)
        {
            var measure = tabular.GetMeasure(measureName);
            string commandText = BuildCustomQueryText(tabular, pivotCellDic, measure.Table.Name, maxRecords, detailColumns);
            return commandText;
        }

        public static string BuildCustomQueryText(TabularHelper tabular, PivotCellDictionary pivotCellDic, string tableQuery,
            int maxRecords, IEnumerable<DetailColumn> detailColumns)
        {
            string filterText = BuildFilterCommandText(pivotCellDic, tabular);

            // create inner clause
            string commandText = string.Format("TOPN ( {1}, {0} )", tableQuery, maxRecords);

            // nest into SELECTCOLUMNS function
            if (detailColumns != null && detailColumns.Count() > 0)
            {
                commandText = string.Format("SELECTCOLUMNS ( {0}, {{0}} )", commandText);
                string selectColumnsText = BuildSelectText(detailColumns);
                commandText = string.Format(commandText, selectColumnsText);
            }

            // add filter arguments
            if (!string.IsNullOrWhiteSpace(filterText))
                commandText += string.Format(",\n{0}", filterText);

            // nest into CALCULATETABLE function
            commandText = string.Format("EVALUATE CALCULATETABLE ( {0} )", commandText);

            return commandText;
        }

        /// <summary>
        /// Creates a comma-delimited string of column filter arguments
        /// </summary>
        public static string BuildSelectText(IEnumerable<DetailColumn> detailColumns)
        {
            string result = string.Empty;
            foreach (var column in detailColumns)
            {
                if (result != string.Empty)
                    result += ",";

                result += string.Format("\n\"{0}\", {1}", column.Name, column.Expression);
            }
            return result;
        }

        public static string BuildFilterCommandText(PivotCellDictionary excelDic, TabularHelper tabular)
        {
            string singCmdText = BuildSingleSelectFilterCommandText(excelDic.SingleSelectDictionary, tabular);
            string multiCmdText = BuildMultiSelectFilterCommandText(excelDic.MultiSelectDictionary, tabular);

            string result = singCmdText;

            if (!string.IsNullOrEmpty(multiCmdText))
            {
                if (!string.IsNullOrEmpty(singCmdText))
                    result += ",\n";
                result += multiCmdText;
            }
            return result;
        }

        private static string BuildSingleSelectFilterCommandText(Dictionary<string, string> excelDic, TabularHelper tabular)
        {
            var daxFilter = ConvertSingleExcelDrillToDaxFilter(excelDic);

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

        private static string BuildMultiSelectFilterCommandText(Dictionary<string, List<string>> excelDic, TabularHelper tabular)
        {
            string commandText = "";
            foreach (KeyValuePair<string, List<string>> pair in excelDic)
            {
                if (commandText != "")
                    commandText += ",\n";

                var daxFilter = ConvertMultiExcelDrillToDaxFilter(pair.Key, pair.Value);

                string childCommandText = "";
                foreach (var item in daxFilter)
                {
                    if (childCommandText != "")
                        childCommandText += " || ";
                    var table = tabular.GetTable(item.TableName);
                    var column = table.Columns.Find(item.ColumnName);
                    childCommandText += BuildColumnCommandText(column, item);
                }

                commandText += childCommandText;
            }
            return commandText;
        }

        public static string BuildColumnCommandText(Tabular.Column column, DaxFilter item)
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
                case Tabular.DataType.Boolean:
                    if (item.Value.ToLower() == "true")
                        commandText = string.Format("{0}[{1}] = {2}", item.TableName, item.ColumnName, "TRUE");
                    else
                        commandText = string.Format("{0}[{1}] = {2}", item.TableName, item.ColumnName, "FALSE");
                    break;
                default:
                    commandText = string.Format("{0}[{1}] = \"{2}\"", item.TableName, item.ColumnName, item.Value);
                    break;
            }
            return commandText;
        }

        public static List<DaxFilter> ConvertSingleExcelDrillToDaxFilter(
            Dictionary<string, string> inputDic)
        {

            var output = new List<DaxFilter>();

            foreach (var pair in inputDic)
            {
                output.Add(CreateDaxFilter(pair.Key, pair.Value));
            }
            return output;
        }

        public static List<DaxFilter> ConvertMultiExcelDrillToDaxFilter(string key, List<string> listValues)
        {
            var output = new List<DaxFilter>();

            string column = GetColumnFromPivotField(key);
            string table = GetTableFromPivotField(key);

            foreach (string listValue in listValues)
            {
                string value = GetValueFromPivotItem(listValue);
                output.Add(new DaxFilter() { TableName = table, ColumnName = column, Value = value });
            }
            return output;
        }

        public static Dictionary<string, List<string>> ConvertDaxFilterListToDictionary(List<DaxFilter> daxFilters)
        {
            var dic = new Dictionary<string, List<string>>();
            return ConvertDaxFilterListToDictionary(daxFilters, dic);
        }

        public static Dictionary<string, List<string>> ConvertDaxFilterListToDictionary(List<DaxFilter> daxFilters, Dictionary<string, List<string>> dic)
        {
            foreach (var df in daxFilters)
            {
                string key = "[" + df.TableName + "].[" + df.ColumnName + "]";
                List<string> value = null;

                if (!dic.TryGetValue(key, out value))
                {
                    value = new List<string>();
                    dic.Add(key, value);
                }
                value.Add(key + ".&[" + df.Value + "]");
            }
            return dic;
        }

        public static List<DaxFilter> ConvertExcelMdxToDaxFilter(string mdxString)
        {
            const string pattern = "FROM (SELECT (";

            var result = new List<DaxFilter>();

            // start reading from the end of the pattern
            int startIndex = mdxString.IndexOf(pattern);
            if (startIndex < 0) return result;
            startIndex += pattern.Length;

            mdxString = mdxString.Substring(startIndex, mdxString.Length - startIndex);

            // stop reading after the first occurrence of ")"
            int endIndex = mdxString.IndexOf(')');
            mdxString = mdxString.Substring(0, endIndex);

            // remove the outer character "{" and "}"
            mdxString = mdxString.Replace("{", "").Replace("}", "");

            string[] itemStringArray = mdxString.Split(',');


            foreach (string itemString in itemStringArray)
            {
                var daxFilter = DaxDrillParser.CreateDaxFilter(itemString.Trim());
                result.Add(daxFilter);
            }
            return result;
        }

        public static DaxFilter CreateDaxFilter(string piKey, string piValue)
        {
            string column = GetColumnFromPivotField(piKey);
            string table = GetTableFromPivotField(piKey);
            string value = GetValueFromPivotItem(piValue);
            return new DaxFilter() { TableName = table, ColumnName = column, Value = value };
        }

        public static DaxFilter CreateDaxFilter(string piValue)
        {
            string column = GetColumnFromPivotField(piValue);
            string table = GetTableFromPivotField(piValue);
            string value = GetValueFromPivotItem(piValue);
            return new DaxFilter() { TableName = table, ColumnName = column, Value = value };
        }

        public static string GetTableFromPivotField(string input)
        {
            string[] split = input.Split('.');
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // input: [Usage].[Inbound or Outbound].[Inbound or Outbound]
        public static string GetColumnFromPivotField(string input)
        {
            string[] split = input.Split('.');
            string output = split[1];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // "[Usage].[Inbound or Outbound].&[Inbound]
        public static string GetValueFromPivotItem(string input)
        {
            try
            {
                string[] split = input.Split('.');
                string output = split[2];
                if (output.Length == 1) return string.Empty;
                output = output.Replace("&", "");
                output = output.Substring(1, output.Length - 2);
                return output;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + string.Format("\r\nCould not parse '{0}'", input), ex);
            }
        }

        public static string CreatePivotFieldPageName(string pivotFieldName, string currentPageName)
        {
            string column = DaxDrillParser.GetColumnFromPivotField(pivotFieldName);
            string table = DaxDrillParser.GetTableFromPivotField(pivotFieldName);
            return string.Format("[{0}].[{1}].&[{2}]", table, column, currentPageName);
        }

        public static bool IsAllItems(string input)
        {
            string[] split = input.Split('.');
            string output = split[2];
            output = output.Substring(1, output.Length - 2);
            return output == "All";
        }

        // example input: [Measures].[Gross Billed Sum]
        public static string GetMeasureFromPivotItem(string input)
        {
            return GetColumnFromPivotField(input);
        }
    }
}
