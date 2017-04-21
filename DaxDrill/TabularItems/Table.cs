extern alias AnalysisServices2014;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SSAS14 = Microsoft.AnalysisServices.Tabular;
using SSAS12 = AnalysisServices2014::Microsoft.AnalysisServices;

namespace DaxDrill.TabularItems
{
    public class Table
    {
        public Table(SSAS14.Table table)
        {
            this.table14 = table;
        }

        public Table(SSAS12.CubeDimension table)
        {
            this.table12 = table;
        }

        private SSAS12.CubeDimension table12;
        private SSAS14.Table table14;
        private ColumnCollection columns;

        public SSAS12.CubeDimension BaseTable12
        {
            get
            {
                return this.table12;
            }
        }

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
                else if (columns == null && this.table12 != null)
                {
                    this.columns = new ColumnCollection();

                    foreach (SSAS12.CubeAttribute cattr in table12.Attributes)
                    {
                        SSAS12.DimensionAttribute dattr = cattr.Attribute;
                        var dataType = dattr.KeyColumns[0].DataType;
                        var column = new Column();
                        column.DataType = MapDataType(dataType);
                        column.Name = dattr.Name;
                        this.columns.Add(column);
                    }
                }
                else if (this.table14 == null && this.table12 == null)
                {
                    throw new InvalidOperationException("Table cannot be null");
                }

                return this.columns;
            }
        }

        private SSAS14.DataType MapDataType(System.Data.OleDb.OleDbType dataType)
        {
            SSAS14.DataType destDataType;

            switch (dataType)
            {
                case System.Data.OleDb.OleDbType.Double:
                case System.Data.OleDb.OleDbType.Decimal:
                case System.Data.OleDb.OleDbType.Integer:
                case System.Data.OleDb.OleDbType.BigInt:
                    destDataType = SSAS14.DataType.Double;
                    break;
                case System.Data.OleDb.OleDbType.Boolean:
                    destDataType = SSAS14.DataType.Boolean;
                    break;
                case System.Data.OleDb.OleDbType.Date:
                    destDataType = SSAS14.DataType.DateTime;
                    break;
                case System.Data.OleDb.OleDbType.DBDate:
                    destDataType = SSAS14.DataType.DateTime;
                    break;
                default:
                    destDataType = SSAS14.DataType.String;
                    break;
            }

            return destDataType;
        }
    }
}
