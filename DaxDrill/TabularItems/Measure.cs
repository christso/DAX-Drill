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
        public Measure(SSAS12.Cube cube, string measureName)
        {
            foreach (SSAS12.Command command in cube.DefaultMdxScript.Commands)
            {
                System.Reflection.MemberInfo member = typeof(SSAS12.Command).GetMember("Annotations").FirstOrDefault();
                SSAS12.AnnotationCollection annotations = (SSAS12.AnnotationCollection)((System.Reflection.PropertyInfo)member).GetValue(command);
                if (annotations.Count == 2
                    && annotations[0].Value.Value == measureName)
                {
                    this.name = measureName;
                    this.tableName = annotations[1].Value.Value;
                    break;
                }
            }
        }

        public Measure(string tableName, string measureName)
        {
            this.name = measureName;
            this.tableName = tableName;
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
