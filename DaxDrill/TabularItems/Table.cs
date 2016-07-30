using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SSAS14 = Microsoft.AnalysisServices.Tabular;

namespace DG2NTT.DaxDrill.TabularItems
{
    public class Table
    {
        public Table(SSAS14.Table table)
        {
            this.table14 = table;
        }

        private SSAS14.Table table14;
        private ColumnCollection columns;

        public ColumnCollection Columns
        {
            get
            {

                // lazily instantiate column collection
                if (columns == null && this.table14 != null)
                {
                    this.columns = new ColumnCollection();

                    foreach (var baseColumn in table14.Columns)
                    {
                        var column = new Column();
                        column.DataType = baseColumn.DataType;
                        column.Name = baseColumn.Name;
                        this.columns.Add(column);
                    }
                }

                return this.columns;
            }
        }
    }
}
