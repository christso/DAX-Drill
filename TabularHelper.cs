using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.Tabular;
using SSAS = Microsoft.AnalysisServices;

namespace DG2NTT.DaxDrill
{
    public class TabularHelper : IDisposable
    {
        private readonly string serverName;
        private readonly string databaseName;
        private readonly Server server;
        private readonly string connectionString;

        public TabularHelper(string serverName, string databaseName)
        {
            this.serverName = serverName;
            this.databaseName = databaseName;
            this.connectionString = string.Format(
"Integrated Security=SSPI; Data source = {0};", serverName);
            this.server = new Server();
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

        public Measure GetMeasure(string measureName)
        {
            if (!server.Connected)
            {
                throw new InvalidOperationException("You must be connect to the server");
            }

            Database database = GetDatabase(databaseName);

            
            foreach (var table in database.Model.Tables)
            {
                Measure measure = table.Measures.Find(measureName);
                if (measure != null)
                    return measure;
            }
            return null;
        }

        public Database GetDatabase(string databaseName)
        {
            if (!server.Connected)
            {
                throw new InvalidOperationException("You must be connect to the server");
            }

            Database database = server.Databases.FindByName(databaseName);

            if (database == null)
                throw new InvalidOperationException(string.Format(
                    "Error retrieving database '{0}' because it does not exist on server '{1}'",
                    databaseName, server.Name));

            return database;
        }

        public Table GetTable(string tableName)
        {
            Database database = GetDatabase(this.databaseName);

            var table = database.Model.Tables.Find(tableName);
            if (table == null)
                throw new InvalidOperationException(string.Format(
                    "Error retrieving table '{0}' because it does not exist in database '{1}'",
                    tableName, databaseName));

            return table;
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
