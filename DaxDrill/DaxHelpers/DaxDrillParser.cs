using DG2NTT.DaxDrill.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SSAS = Microsoft.AnalysisServices.Tabular;

namespace DG2NTT.DaxDrill.DaxHelpers
{
    public class DaxDrillParser
    {
        public static string BuildQueryText(DG2NTT.DaxDrill.Tabular.TabularHelper tabular, PivotCellDictionary pivotCellDic, string measureName, int maxRecords)
        {
            return BuildQueryText(tabular, pivotCellDic, measureName, maxRecords);
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
        public static string BuildQueryText(DG2NTT.DaxDrill.Tabular.TabularHelper tabular, PivotCellDictionary pivotCellDic, string measureName,
            int maxRecords, IEnumerable<DetailColumn> detailColumns,
            IEnumerable<string> pivotFieldNames)
        {
            var measure = tabular.GetMeasure(measureName);
            string commandText = BuildCustomQueryText(tabular, pivotCellDic, measure.TableName, maxRecords, detailColumns, pivotFieldNames);
            return commandText;
        }

        public static string BuildCustomQueryText(DG2NTT.DaxDrill.Tabular.TabularHelper tabular, PivotCellDictionary pivotCellDic, string tableQuery,
            int maxRecords, IEnumerable<DetailColumn> detailColumns,
            IEnumerable<string> pivotFieldNames)
        {
            string filterText = BuildFilterCommandText(pivotCellDic, tabular, pivotFieldNames);

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

        public static string BuildFilterCommandText(PivotCellDictionary excelDic, DG2NTT.DaxDrill.Tabular.TabularHelper tabular, 
            IEnumerable<string> pivotFieldNames)
        {
            string singCmdText = BuildSingleSelectFilterCommandText(excelDic.SingleSelectDictionary, tabular,
                pivotFieldNames);
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

        private static string BuildSingleSelectFilterCommandText(Dictionary<string, string> excelDic, DG2NTT.DaxDrill.Tabular.TabularHelper tabular, IEnumerable<string> pivotFieldNames)
        {
            List<DaxFilter> daxFilterList = ConvertSingleExcelDrillToDaxFilterList(excelDic, pivotFieldNames);

            string commandText = "";
            foreach (var item in daxFilterList)
            {
                if (commandText != "")
                {
                    commandText += ",\n";
                }
                if (item.IsHierarchy)
                    commandText += "FILTER ( " + item.TableName + ",";
                var table = tabular.GetTable(item.TableName);
                commandText += BuildColumnCommandText(table, item);
                if (item.IsHierarchy)
                    commandText += ")";
            }
            return commandText;
        }

        private static string BuildMultiSelectFilterCommandText(DaxFilterCollection daxFilterDic, DG2NTT.DaxDrill.Tabular.TabularHelper tabular)
        {
            string commandText = "";

            foreach (var pair in daxFilterDic)
            {
                List<DaxFilter> daxFilterList = pair.Value;
                string tableName = daxFilterList[0].TableName;

                if (commandText != "")
                    commandText += ",\n";

                string childCommandText = "";
                foreach (var item in daxFilterList)
                {
                    if (childCommandText != "")
                        childCommandText += " || ";
                    else
                        childCommandText = "FILTER ( " + tableName + ", ";

                    var table = tabular.GetTable(item.TableName);

                    childCommandText += BuildColumnCommandText(table, item);
                }

                commandText += childCommandText + " )";
            }
            return commandText;
        }

        public static string BuildColumnCommandText(TabularItems.Table table, DaxFilter daxFilter)
        {
            //TODO: problem is that column name = "Tran_YearMonthDay" 
            // this needs to be Tran_Year and TranMonth
            TabularItems.Column column = null;

            string commandText = "";

            if (daxFilter.IsHierarchy)
            {
                for (int i = 0; i < daxFilter.ValueHierarchy.Length; i++) 
                {
                    DaxColumn dc = daxFilter.ColumnNameHierarchy[i];
                    column = table.Columns.Find(dc.ColumnName);
                    if (!string.IsNullOrEmpty(commandText))
                        commandText += " && ";
                    commandText += BuildColumnCommandText(column, dc.TableName, dc.ColumnName,
                         daxFilter.ValueHierarchy[i]);
                }
            }
            else
            {
                column = table.Columns.Find(daxFilter.ColumnName);
                commandText = BuildColumnCommandText(column, daxFilter.TableName, daxFilter.ColumnName, daxFilter.Value);
            }
            return commandText;
        }

        public static string BuildColumnCommandText(TabularItems.Column column, string tableName, string columnName, string value)
        {
            if (column == null)
                throw new InvalidOperationException(
                    string.Format("Column '{0} was not found in Tabular database", columnName));

            string commandText;
            switch (column.DataType)
            {
                case SSAS.DataType.String:
                    commandText = string.Format("{0}[{1}] = \"{2}\"", tableName, columnName, value);
                    break;
                case SSAS.DataType.Int64:
                case SSAS.DataType.Decimal:
                case SSAS.DataType.Double:
                    commandText = string.Format("{0}[{1}] = {2}", tableName, columnName, value);
                    break;
                case SSAS.DataType.Boolean:
                    if (value.ToLower() == "true")
                        commandText = string.Format("{0}[{1}] = {2}", tableName, columnName, "TRUE");
                    else
                        commandText = string.Format("{0}[{1}] = {2}", tableName, columnName, "FALSE");
                    break;
                default:
                    commandText = string.Format("{0}[{1}] = \"{2}\"", tableName, columnName, value);
                    break;
            }
            return commandText;
        }

        public static List<DaxFilter> ConvertSingleExcelDrillToDaxFilterList(
            Dictionary<string, string> inputDic, IEnumerable<string> pivotFieldNames)
        {

            var output = new List<DaxFilter>();

            foreach (var pair in inputDic)
            {
                output.Add(CreateDaxFilterFromHierarchy(pair.Value, pivotFieldNames));
            }
            return output;
        }

        public static List<DaxFilter> ConvertMultiExcelDrillToDaxFilterList(string key, List<DaxFilter> listValues)
        {
            var output = new List<DaxFilter>();

            string column = GetColumnFromPivotField(key);
            string table = GetTableFromPivotField(key);

            foreach (DaxFilter listValue in listValues)
            {
                string value = GetValueFromPivotItem(listValue);
                output.Add(new DaxFilter() { TableName = table, ColumnName = column, Value = value });
            }
            return output;
        }

        public static Dictionary<string, List<DaxFilter>> ConvertDaxFilterListToDictionary(List<DaxFilter> daxFilters)
        {
            var dic = new DaxFilterCollection();
            return ConvertDaxFilterListToDictionary(daxFilters, dic);
        }

        // converts DaxFilter list to dictionary. Duplicate entries are removed.
        // This is done by moving each Dax Filter under a dictionary key
        // The flat data structure is converted to a tree structure
        public static DaxFilterCollection ConvertDaxFilterListToDictionary(
            List<DaxFilter> daxFilters, DaxFilterCollection dic)
        {
            foreach (var df in daxFilters)
            {
                List<DaxFilter> dicValue = null;

                // create dictionary element if it doesn't exist
                if (!dic.TryGetValue(df.Key, out dicValue))
                {
                    dicValue = new List<DaxFilter>();
                    dic.Add(df.Key, dicValue);
                }

                // add DaxFilter to dictionary element
                dicValue.Add(df);
            }
            return dic;
        }

        public static List<DaxFilter> ConvertPivotTableMdxToDaxFilterList(string mdxString)
        {
            return ConvertPivotTableMdxToDaxFilterList(mdxString, null);
        }

        public static List<DaxFilter> ConvertPivotTableMdxToDaxFilterList(string mdxString, IEnumerable<string> pivotFieldNames)
        {
            var mdxParser = new TableMdxParser(mdxString);
            
            string[] columnStringArray = mdxParser.ConvertColumnMdxToArray();

            string[] rowStringArray = mdxParser.ConvertRowMdxToArray();

            var result = new List<DaxFilter>();

            foreach (string itemString in columnStringArray)
            {
                var daxFilter = DaxDrillParser.CreateDaxFilterFromHierarchy(itemString.Trim(), pivotFieldNames);
                result.Add(daxFilter);
            }
            foreach (string itemString in rowStringArray)
            {
                var daxFilter = DaxDrillParser.CreateDaxFilterFromHierarchy(itemString.Trim(), pivotFieldNames);
                result.Add(daxFilter);
            }

            return result;
        }


        public static List<DaxFilter> ConvertPivotCellMdxToDaxFilterList(string mdxString)
        {
            mdxString = mdxString.Trim();
            mdxString = mdxString.Substring(1, mdxString.Length - 2);
            string[] itemStringArray = mdxString.Split(',');
            for (int i = 0; i < itemStringArray.Length; i++)
                itemStringArray[i] = itemStringArray[i].Trim();

            itemStringArray = itemStringArray.Where(x =>
            {
                // exclude the measure as it's not a DAX filter
                if (x.Substring(0, 10) == "[Measures]")
                    return false;
                return true;
            }).ToArray();

            var result = new List<DaxFilter>();
            foreach (string itemString in itemStringArray)
            {
                var daxFilter = DaxDrillParser.CreateDaxFilterFromHierarchy(itemString, null);
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


        public static DaxFilter CreateDaxFilterFromColumn(string piValue)
        {
            string column = GetColumnFromPivotField(piValue);
            string table = GetTableFromPivotField(piValue);
            string value = GetValueFromPivotItem(piValue);
            return new DaxFilter() { TableName = table, ColumnName = column, Value = value };
        }

        public static DaxFilter CreateDaxFilterFromHierarchy(string piValue, IEnumerable<string> pivotFieldNames)
        {
            var processor = new DaxFilterCreator(piValue, pivotFieldNames);
            var daxFilter = processor.CreateDaxFilter();
            return daxFilter;
        }

        public static string GetTableFromPivotField(string input)
        {
            string[] split = input.Split('.');
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }
        public static string GetTableFromPivotFieldElement(string[] split)
        {
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        // input: [Usage].[Inbound or Outbound].[Inbound or Outbound]
        public static string GetColumnFromPivotField(string input)
        {
            string[] split = input.Split('.');
            return GetColumnFromPivotFieldElement(split);
        }

        public static string GetColumnFromPivotFieldElement(string[] split)
        {
            string output = split[1];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        public static string GetHierarchyFromPivotField(string input)
        {
            string[] split = input.Split('.');
            return string.Empty;
        }

        // "[Usage].[Inbound or Outbound].&[Inbound]
        public static string GetValueFromPivotItem(string input)
        {
            try
            {
                input = input.Replace("&[", "[");
                string[] split = input.Split(new string[] { ".[" }, StringSplitOptions.None);
                if (split.Length <= 2)
                    return string.Empty;
                string output = split[2];
                output = output.Substring(0, output.Length - 1);
                return output;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + string.Format("\r\nCould not parse '{0}'", input), ex);
            }
        }

        public static string GetValueFromPivotItem(DaxFilter df)
        {
            try
            {
                return df.Value;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + string.Format("\r\nCould not parse '{0}'", df.MDX), ex);
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
        public static string RemoveBrackets(string columnText)
        {
            int first = columnText.IndexOf('[');
            if (first >= 0 && columnText.Substring(columnText.Length - 1) == "]")
            {
                columnText = columnText.Substring(first + 1, columnText.Length - first - 2);
            }
            return columnText;
        }
    }
}
