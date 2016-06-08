using Microsoft.AnalysisServices;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;

namespace DG2NTT.DaxDrill
{
    public class DaxClient
    {
        public DataTable ExecuteTable(string qry, ADOMD.AdomdConnection cnn)
        {
            ADOMD.AdomdDataAdapter dataAdapter = new ADOMD.AdomdDataAdapter(qry, cnn);
            DataTable tabularResults = new DataTable();
            dataAdapter.Fill(tabularResults);
            
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
