using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaxDrill.TabularItems
{
    public class ColumnCollection : List<Column>
    {
        public ColumnCollection()
        {
            this.stringSet = new HashSet<string>(); // enforce unique column names
        }

        public new void Add(Column column)
        {
            if (!stringSet.Add(column.Name))
                throw new InvalidOperationException(
                    string.Format("Column {0} already exists in collection.", column.Name));
            base.Add(column);
        }

        private HashSet<string> stringSet;
        public Column Find(string columnName)
        {
            return this.Where(c => c.Name == columnName).FirstOrDefault();
        }
    }
}
