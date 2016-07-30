using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.Tabular
{
    public interface ITabularHelper
    {
        string ServerName { get; }
        string DatabaseName { get; }
        string ConnectionString { get; }
        void Connect();
        void Disconnect();
        TabularItems.Measure GetMeasure(string measureName);
        bool IsDatabaseCompatible { get; }
        TabularItems.Table GetTable(string tableName);
    }
}
