using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.Tabular;

namespace DG2NTT.DaxDrill.TabularItems
{
    public class Column
    {
        public DataType DataType
        {
            get; set;
        }

        public string Name
        {
            get; set;
        }
    }
}
