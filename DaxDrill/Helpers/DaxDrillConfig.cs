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
        public static List<SelectedColumn> GetColumnsFromColumnsXml(string xmlString, string nsString)
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

        public static List<SelectedColumn> GetColumnsFromColumnsXmlNode(XmlNode columnsNode, XmlNamespaceManager nsmgr)
        {
            var columns = new List<SelectedColumn>();

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

        public static List<SelectedColumn> GetColumnsFromTableXml(string tableName, string xmlString, string nsString)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            string xpath = string.Empty;
            if (root.Name == "columns")
                xpath = "/x:columns";
            else if (root.Name == "table")
                xpath = string.Format("/x:table[@id=\"{0}\"]", tableName);

            XmlNode columnsNode = root.SelectSingleNode(xpath, nsmgr);
            var columns = GetColumnsFromColumnsXmlNode(columnsNode, nsmgr);
            return columns;
        }

       
    }
}
