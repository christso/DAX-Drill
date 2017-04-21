using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace DaxDrill.DaxHelpers
{
    public class DaxDrillConfig
    {
        public static List<DetailColumn> GetColumnsFromColumnsXml(string xmlString, string nsString)
        {
            var columns = new List<DetailColumn>();

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
                columns.Add(new DetailColumn()
                {
                    Name = nameNode.InnerText,
                    Expression = exprNode.InnerText
                });
            }
            return columns;
        }

        public static List<DetailColumn> GetColumnsFromColumnsXmlNode(XmlNode columnsNode, XmlNamespaceManager nsmgr)
        {
            var columns = new List<DetailColumn>();

            foreach (XmlNode columnNode in columnsNode)
            {
                XmlNode nameNode = columnNode.SelectSingleNode("./x:name", nsmgr);
                XmlNode exprNode = columnNode.SelectSingleNode("./x:expression", nsmgr);
                columns.Add(new DetailColumn()
                {
                    Name = nameNode.InnerText,
                    Expression = exprNode.InnerText
                });
            }
            return columns;
        }

        public static List<DetailColumn> GetColumnsFromTableXml(string nsString,
            string xmlString, string connectionName, string tableName)
        {
            XmlDocument doc = new XmlDocument();
            if (string.IsNullOrWhiteSpace(xmlString))
                return null;
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            string xpath = GetColumnsXpathFromTableXml(root, connectionName, tableName);

            XmlNode columnsNode = null;
            if (!string.IsNullOrEmpty(xpath))
                columnsNode = root.SelectSingleNode(xpath, nsmgr);
            if (columnsNode == null)
                return null;

            var columns = GetColumnsFromColumnsXmlNode(columnsNode, nsmgr);
            return columns;
        }

        public static string GetColumnsXpathFromTableXml(XmlNode root, string connectionName,
            string tableName)
        {

            string xpath = string.Empty;
            if (root.Name == "columns")
                xpath = "/x:columns";
            else if (root.Name == "table" && !string.IsNullOrWhiteSpace(connectionName))
                xpath = string.Format("/x:table[@id=\"{0}\" and @connection_id=\"{1}\"]/x:columns",
                    tableName, connectionName);
            else if (root.Name == "table")
                xpath = string.Format("/x:table[@id=\"{0}\"]/x:columns", tableName);
            else if (root.Name == Constants.RootXmlNode)
                xpath = string.Format("/x:{1}/x:table[@id=\"{0}\"]/x:columns", 
                    tableName, Constants.RootXmlNode);
            return xpath;
        }
    }
}
