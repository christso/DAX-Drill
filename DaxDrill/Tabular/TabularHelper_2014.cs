extern alias AnalysisServices2014;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AnalysisServices2014::Microsoft.AnalysisServices;

namespace DaxDrill.Tabular
{
    public class TabularHelper_2014 : IDisposable
    {
        private const string cubeName = "Model";
        private const int MaxCompatibilityLevel = 1199; // MSAS 2016 and above
        private readonly string serverName;
        private readonly string databaseName;
        private readonly Server server;
        private readonly string connectionString;

        public TabularHelper_2014(string serverName, string databaseName)
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
            if (server.Connected)
                server.Disconnect();
        }

        public TabularItems.Measure GetMeasure(string measureName)
        {
            if (!server.Connected)
            {
                throw new InvalidOperationException("You must be connect to the server");
            }

            Database database = server.Databases.FindByName(databaseName);

            if (database == null)
                throw new InvalidOperationException(string.Format(
                    "Database '{0}' does not exist on server '{1}'",
                    databaseName, server.Name));

            Cube cube = database.Cubes.FindByName(cubeName);
            if (cube == null)
                throw new InvalidOperationException(string.Format(
                    "Cube '{0}' does not exist in database '{1}'",
                    cubeName, database));

            TabularItems.Measure measure = new TabularItems.Measure(cube, measureName);

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

                System.Reflection.MemberInfo databaseMemInf = typeof(Database).GetMember("ModelType").FirstOrDefault();
                var modelType = (Microsoft.AnalysisServices.ModelType)((System.Reflection.PropertyInfo)databaseMemInf).GetValue(database);

                bool isServerCompatible = database.CompatibilityLevel <= MaxCompatibilityLevel;
                bool isDatabaseCompatible = modelType == Microsoft.AnalysisServices.ModelType.Tabular;
                return isServerCompatible && isDatabaseCompatible;
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

            Cube cube = database.Cubes.FindByName(cubeName);
            if (cube == null)
                throw new InvalidOperationException(string.Format(
                    "Cube  '{0}' does not exist in database '{1}'",
                    cubeName, database));

            CubeDimension table = cube.Dimensions.FindByName(tableName);
            if (table == null)
                throw new InvalidOperationException(string.Format(
                    "Table '{0}' because it does not exist in cube '{1}'",
                    tableName, cubeName));

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
        // ~TabularHelper_2014() {
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
