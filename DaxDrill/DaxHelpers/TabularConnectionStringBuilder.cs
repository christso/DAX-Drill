using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.DaxHelpers
{
    public class TabularConnectionStringBuilder : DbConnectionStringBuilder
    {
        public TabularConnectionStringBuilder(string connectionString)
        {
            base.ConnectionString = connectionString;
        }
        public string DataSource
        {
            get
            {
                return this["data source"].ToString();
            }
        }

        public string InitialCatalog
        {
            get
            {
                return this["initial catalog"].ToString();
            }
        }

        public string StrippedConnectionString
        {
            get
            {
                return string.Format(
                    "Integrated Security=SSPI;Persist Security Info=True;Initial Catalog={1};Data Source={0};",
                    DataSource, InitialCatalog);
            }
        }
    }
}
