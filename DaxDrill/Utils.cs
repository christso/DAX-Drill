using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace DaxDrill
{
    public class Utils
    {
        public static object[,] CreateArray(System.Data.DataTable dataTable)
        {
            var rowCount = dataTable.Rows.Count;
            var columnCount = dataTable.Columns.Count;
            const int headerFlag = 1;
            object[,] result = new object[rowCount + headerFlag, columnCount];
            
            // header
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                result[0, c] = dataTable.Columns[c].Caption;
            }

            // records
            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    result[headerFlag + r, c] = dataTable.Rows[r][c];
                }
            }

            return result;
        }

        public void ShowException(Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
}
