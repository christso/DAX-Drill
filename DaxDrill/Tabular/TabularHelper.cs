using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.Tabular;
using SSAS = Microsoft.AnalysisServices;

namespace DaxDrill.Tabular
{
    public class TabularHelper : IDisposable, ITabularHelper
    {

        public TabularHelper(string serverName, string databaseName)
        {
            this.tabular14 = new TabularHelper_2014(serverName, databaseName);
            this.tabular16 = new TabularHelper_2016(serverName, databaseName);
        }

        private TabularHelper_2014 tabular14;
        private TabularHelper_2016 tabular16;

        public string ServerName
        {
            get
            {
                return tabular16.ServerName;
            }
        }

        public string DatabaseName
        {
            get
            {
                return tabular16.DatabaseName;
            }
        }
        
        public string ConnectionString
        {
            get
            {
                return tabular16.ConnectionString;
            }
        }

        public void Connect()
        {
            tabular16.Connect();
            if (!tabular16.IsDatabaseCompatible)
                tabular14.Connect();
        }

        public void Disconnect()
        {
            tabular16.Disconnect();
            tabular14.Disconnect();
        }

        public TabularItems.Measure GetMeasure(string measureName)
        {
            if (tabular16.IsDatabaseCompatible)
                return tabular16.GetMeasure(measureName);
            else
                return tabular16.GetMeasureFromDMV(measureName);
        }

        public bool IsDatabaseCompatible
        {
            get
            {
                Database database = tabular16.GetDatabase(this.DatabaseName);
                return database.ModelType == SSAS.ModelType.Tabular;
            }
        }

        public TabularItems.Table GetTable(string tableName)
        {
            if (tabular16.IsDatabaseCompatible)
                return tabular16.GetTable(tableName);
            else
                return tabular14.GetTable(tableName);
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (tabular14 != null)
                        tabular14.Dispose();
                    if (tabular16 != null)
                        tabular16.Dispose();
                }
        
                disposedValue = true;
            }
        }
        public void Dispose()
        {
            Dispose(true);
        }
        #endregion

    }
}
