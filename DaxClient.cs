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
        public DataTable ExecuteQuery(string qry, ADOMD.AdomdConnection cnx)
        {
            ADOMD.AdomdDataAdapter dataAdapter = new ADOMD.AdomdDataAdapter(qry, cnx);
            DataTable tabularResults = new DataTable();
            dataAdapter.Fill(tabularResults);
            
            return tabularResults;
        }
    }
}
