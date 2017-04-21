using DaxDrill.DaxHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaxDrill.DaxHelpers
{
    public class DaxFilterCreator
    {
        public DaxFilterCreator(string pivotItemValue,
            IEnumerable<string> pivotFieldNames)
        {
            this.pivotFieldNames = pivotFieldNames;
            this.pivotItemValue = pivotItemValue;
        }

        private string pivotItemValue;
        private IEnumerable<string> pivotFieldNames;
        private DaxFilter daxFilter;

        public DaxFilter CreateDaxFilter()
        {
            daxFilter = new DaxFilter();
            daxFilter.IsHierarchy = PivotItemIsHierarchy(pivotItemValue);
            daxFilter.ColumnName = GetColumnName();
            daxFilter.TableName = GetTableName();
            daxFilter.ValueHierarchy = GetValue(daxFilter.IsHierarchy);
            daxFilter.Value = daxFilter.ValueHierarchy == null ?
                null : daxFilter.ValueHierarchy[0];

            daxFilter.ColumnNameHierarchy = GetColumnNameHierarchy();

            return daxFilter;
        }


        private IList<DaxColumn> GetColumnNameHierarchy()
        {
            var columnName = daxFilter.ColumnName;
            var daxColumns = new List<DaxColumn>();
            if (pivotFieldNames == null)
                return null;

            int i = 0;
            foreach (var pfValue in pivotFieldNames)
            {
                if (pfValue == "Values")
                    continue;
                daxColumns.Add(new DaxColumn(pfValue, daxFilter.IsHierarchy, i));
                i++;
            }
            var matched = daxColumns.Where(x => x.HierarchyName == columnName);
            return matched.ToList<DaxColumn>();
        }

        private string GetColumnName()
        {
            string[] split = pivotItemValue.Split('.');
            string output = split[1];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        private string GetTableName()
        {
            string[] split = pivotItemValue.Split('.');
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        private string[] GetValue(bool isHierarchy)
        {

            if (isHierarchy)
            {
                return GetHierarchyValue();
            }
            else
            {
                return GetScalarValue();
            }

        }

        private string[] GetScalarValue()
        {
            pivotItemValue = pivotItemValue.Replace("&[", "[");
            string[] split = pivotItemValue.Split(new string[] { ".[" }, StringSplitOptions.None);

            try
            {
                if (split.Length <= 2)
                    return null;
                string output = split[2];
                output = output.Substring(0, output.Length - 1);
                return new string[] { output };
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + string.Format("\r\nCould not parse scalar '{0}'", pivotItemValue), ex);
            }
        }

        private string[] GetHierarchyValue()
        {
            pivotItemValue = pivotItemValue.Replace("&[", "[");
            string[] split = pivotItemValue.Split(new string[] { ".[" }, StringSplitOptions.None);

            try
            {
                int fieldNameEndIdx = 3;

                if (split.Length <= fieldNameEndIdx)
                    return null;

                var output = new string[split.Length - fieldNameEndIdx];

                int j = 0;
                for (int i = fieldNameEndIdx; i < split.Length; i++)
                {
                    output[j] = split[i].Substring(0, split[i].Length - 1);
                    j++;
                }
                return output;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + string.Format("\r\nCould not parse hierarchy '{0}'", pivotItemValue), ex);
            }
        }

        public static bool PivotFieldIsHierarchy(string pivotFieldValue)
        {
            var split = pivotFieldValue.Split('.');
            bool isHierarchy = split[1] != split[2];
            return isHierarchy;
        }

        public static bool PivotItemIsHierarchy(string pivotItemValue)
        {
            var split = pivotItemValue.Split('.');
            bool isHierarchy = split[2].Substring(0, 1) != "&";
            return isHierarchy;
        }

    }
}
