using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.DaxHelpers
{
    public class DaxColumn
    {
        public DaxColumn(string pivotFieldName, bool isHierarchyMember, int index)
        {
            this.pivotFieldName = pivotFieldName;
            this.tableName = GetTableName();
            this.isHierarchyMember = isHierarchyMember;
            this.columnName = GetColumnName(isHierarchyMember);
            this.hierarchyName = GetColumnName(false);
            this.index = index;
        }

        private string pivotFieldName;
        private string tableName;
        private string columnName;
        private string hierarchyName;
        private int index;
        private bool isHierarchyMember;
        
        public int Index
        {
            get
            {
                return index;
            }
        }
        public string TableName
        {
            get
            {
                return tableName;
            }
        }

        public string ColumnName
        {
            get
            {
                return columnName;
            }
        }

        public string HierarchyName
        {
            get
            {
                return hierarchyName;
            }
        }

        public bool IsHierarchyMember
        {
            get
            {
                return isHierarchyMember;
            }
        }

        public string DAX
        {
            get
            {
                return tableName + "[" + columnName + "]";
            }
        }

        private string GetTableName()
        {
            string[] split = pivotFieldName.Split('.');
            string output = split[0];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

        private string GetColumnName(bool isHierarchyMember)
        {
            string[] split = pivotFieldName.Split('.');

            int offsetIndex = 0;
            if (isHierarchyMember && split.Length > 2)
                offsetIndex = 1;
            string output = split[1 + offsetIndex];
            output = output.Substring(1, output.Length - 2);
            return output;
        }

    }
}
