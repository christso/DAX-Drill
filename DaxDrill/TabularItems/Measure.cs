extern alias AnalysisServices2014;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SSAS14 = Microsoft.AnalysisServices.Tabular;
using SSAS12 = AnalysisServices2014::Microsoft.AnalysisServices;

namespace DG2NTT.DaxDrill.TabularItems
{
    public class Measure
    {
        public Measure()
        {
            
        }

        public Measure(SSAS14.Measure measure)
        {
            this.tableName = measure.Table.Name;
            this.name = measure.Name;
        }

        private string tableName;
        private string name;
        public string TableName
        {
            get
            {
                return this.tableName;
            }
            set
            {
                this.tableName = value;
            }
        }

        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }
    }
}
