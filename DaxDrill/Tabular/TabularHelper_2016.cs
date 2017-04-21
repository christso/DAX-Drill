using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.Tabular;
using SSAS = Microsoft.AnalysisServices;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;

namespace DaxDrill.Tabular
{
    public class TabularHelper_2016 : IDisposable, ITabularHelper
    {
        private const int MinCompatibilityLevel = 1200; // MSAS 2016 and above
        private readonly string serverName;
        private readonly string databaseName;
        private readonly Server server;
        private readonly string connectionString;

        public TabularHelper_2016(string serverName, string databaseName)
        {
            this.serverName = serverName;
            this.databaseName = databaseName;
            this.connectionString = string.Format(
"Integrated Security=SSPI;Persist Security Info=True;Initial Catalog={1};Data Source={0};", serverName, databaseName);
            this.server = new Server();
        }

        public string ServerName
        {
            get
            {
                return this.serverName;
            }
        }

        public string DatabaseName
        {
            get
            {
                return this.databaseName;
            }
        }

        public string ConnectionString
        {
            get
            {
                return this.connectionString;
            }
        }

        public void Connect()
        {
            server.Connect(this.connectionString);
        }

        public void Disconnect()
        {
            server.Disconnect();
        }

        public TabularItems.Measure GetMeasure(string measureName)
        {
            if (!server.Connected)
            {
                throw new InvalidOperationException("You must be connected to the server");
            }

            Database database = GetDatabase(databaseName);
            if (database.Model == null)
                return GetMeasureFromDMV(measureName); // use DMV method if default method fails

            Measure measure = null;
            foreach (var table in database.Model.Tables)
            {
                measure = table.Measures.Find(measureName);
                if (measure != null)
                    break;
            }
            if (measure == null)
                throw new InvalidOperationException("Measure " + measureName + " was not found in database " + this.databaseName);

            return new TabularItems.Measure(measure);
        }

        public TabularItems.Measure GetMeasureFromDMV(string measureName)
        {
            if (!server.Connected)
            {
                throw new InvalidOperationException("You must be connected to the server");
            }

            var dmv = string.Format(
@"SELECT [CATALOG_NAME] as [DATABASE],
    CUBE_NAME AS [CUBE],[MEASUREGROUP_NAME] AS [TABLE],[MEASURE_CAPTION] AS [MEASURE],
    [MEASURE_IS_VISIBLE]
FROM $SYSTEM.MDSCHEMA_MEASURES
WHERE CUBE_NAME  ='Model'
	AND MEASURE_CAPTION = '{0}'",
                    measureName);

            var daxClient = new DaxHelpers.DaxClient();
            System.Data.DataTable dtResult = null;
            using (var cnn = new ADOMD.AdomdConnection(connectionString))
            {
                cnn.Open();
                dtResult = daxClient.ExecuteTable(dmv, cnn);
                cnn.Close();
            }

            TabularItems.Measure measure = null;
            foreach (System.Data.DataRow drow in dtResult.Rows)
            {
                if (Convert.ToString(drow["MEASURE"]) == measureName)
                {
                    measure = new TabularItems.Measure(Convert.ToString(drow["TABLE"]), measureName);
                    break;
                }
            }

            if (measure == null)
                throw new InvalidOperationException("Measure " + measureName + " was not found in database " + this.databaseName);
            return measure;
        }

        public bool IsDatabaseCompatible
        {
            get
            {
                if (!server.Connected)
                {
                    throw new InvalidOperationException("You must be connected to the server");
                }

                Database database = GetDatabase(databaseName);
                return database.CompatibilityLevel >= MinCompatibilityLevel
                    && database.ModelType == SSAS.ModelType.Tabular;
            }
        }
        
        public Database GetDatabase(string databaseName)
        {
            if (!server.Connected)
            {
                throw new InvalidOperationException("You must be connected to the server");
            }

            Database database = server.Databases.FindByName(databaseName);

            if (database == null)
                throw new InvalidOperationException(string.Format(
                    "Error retrieving database '{0}' because it does not exist on server '{1}'",
                    databaseName, server.Name));

            return database;
        }

        public TabularItems.Table GetTable(string tableName)
        {
            Database database = GetDatabase(this.databaseName);

            var table = database.Model.Tables.Find(tableName);
            if (table == null)
                throw new InvalidOperationException(string.Format(
                    "Error retrieving table '{0}' because it does not exist in database '{1}'",
                    tableName, databaseName));

            return new TabularItems.Table(table);
        }
        
        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // dispose managed state (managed objects).
                    if (server != null)
                    {
                        if (server.Connected)
                            server.Disconnect();
                        server.Dispose();
                    }
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~TabularHelper() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
