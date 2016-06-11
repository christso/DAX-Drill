using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace DG2NTT.DaxDrill.Helpers
{
    public class DaxDrillConfig
    {
        public static List<SelectedColumn> GetColumns(string xmlString, string nsString)
        {
            var columns = new List<SelectedColumn>();

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            XmlNode columnsNode = root.SelectSingleNode("/x:columns", nsmgr);
            foreach (XmlNode columnNode in columnsNode)
            {
                XmlNode nameNode = columnNode.SelectSingleNode("./x:name", nsmgr);
                XmlNode exprNode = columnNode.SelectSingleNode("./x:expression", nsmgr);
                columns.Add(new SelectedColumn()
                {
                    Name = nameNode.InnerText,
                    Expression = exprNode.InnerText
                });
            }
            return columns;
        }
    }
}
