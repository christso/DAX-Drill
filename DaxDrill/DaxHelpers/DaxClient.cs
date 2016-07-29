using Microsoft.AnalysisServices;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;
using Microsoft.AnalysisServices.Tabular;
using SSAS = Microsoft.AnalysisServices;

namespace DG2NTT.DaxDrill.DaxHelpers
{
    public class DaxClient
    {
        public DataTable ExecuteTable(string qry, ADOMD.AdomdConnection cnn)
        {
            ADOMD.AdomdDataAdapter dataAdapter = new ADOMD.AdomdDataAdapter(qry, cnn);
            DataTable tabularResults = new DataTable();
            dataAdapter.Fill(tabularResults);
            foreach (System.Data.DataColumn column in tabularResults.Columns)
            {
                column.ColumnName = DaxDrillParser.RemoveBrackets(column.ColumnName);
            }

            return tabularResults;
        }

        public ADOMD.AdomdDataReader ExecuteReader(string qry, ADOMD.AdomdConnection cnn)
        {
            var cmd = new ADOMD.AdomdCommand(qry, cnn);
            ADOMD.AdomdDataReader reader = cmd.ExecuteReader();
            return reader;
        }

    }
}
