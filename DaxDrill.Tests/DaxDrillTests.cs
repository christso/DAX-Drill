using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;
using DG2NTT.DaxDrill.DaxHelpers;
using Microsoft.AnalysisServices.Tabular;
using DG2NTT.DaxDrill.ExcelHelpers;

namespace DG2NTT.DaxDrill.Tests
{
    [TestFixture]
    public class DaxDrillTests
    {
        /*
EVALUATE
CALCULATETABLE (
TOPN( 100, Usage ),
Usage[Inbound or Outbound] = "Inbound",
Usage[Call Type] = "MOC",
UsageDate[Usage_Year] = "2014",
UsageDate[Usage_MonthAbbrev] = "May" 
)
         */
        [Test]
        public void ParseCellDictionary()
        {
            #region Arrange

            var excelDic = new Dictionary<string, string>();
            excelDic.Add("[Usage].[Inbound or Outbound].[Inbound or Outbound]",
                "[Usage].[Inbound or Outbound].&[Inbound]");
            excelDic.Add("[Usage].[Call Type].[Call Type]",
                "[Usage].[Call Type].&[MOC]");
            excelDic.Add("[UsageDate].[Usage_Year].[Usage_Year]",
                "[UsageDate].[Usage_Year].&[2014]");
            excelDic.Add("[UsageDate].[Usage_MonthAbbrev].[Usage_MonthAbbrev]",
                "[UsageDate].[Usage_MonthAbbrev].&[May]");
            #endregion

            #region Parse
            var pivotCellDic = new PivotCellDictionary();
            pivotCellDic.SingleSelectDictionary = excelDic;
            string commandText;
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                commandText = DaxDrillParser.BuildFilterCommandText(pivotCellDic, tabular);
                tabular.Disconnect();
            }
            #endregion

            #region Assert
            Console.WriteLine(commandText);

            #endregion
        }

        [Test]
        public void ParseAll()
        {
            #region Arrange

            var singDic = new Dictionary<string, string>();
            singDic.Add("[Usage].[Inbound or Outbound].[Inbound or Outbound]",
                "[Usage].[Inbound or Outbound].&[Inbound]");
            singDic.Add("[Usage].[Call Type].[Call Type]",
                "[Usage].[Call Type].&[MOC]");
            singDic.Add("[UsageDate].[Usage_Year].[Usage_Year]",
                "[UsageDate].[Usage_Year].&[2014]");
            singDic.Add("[UsageDate].[Usage_MonthAbbrev].[Usage_MonthAbbrev]",
                "[UsageDate].[Usage_MonthAbbrev].&[May]");
            #endregion

            #region Parse
            string commandText;

            var pivotCellDic = new PivotCellDictionary();
            pivotCellDic.SingleSelectDictionary = singDic;

            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                commandText = DaxDrillParser.BuildQueryText(tabular, pivotCellDic, "Gross Billed Sum", 99999);
                tabular.Disconnect();
            }
            #endregion

            #region Assert
            Console.WriteLine(commandText);

            #endregion
        }

        [Test]
        public void ParseSelectedColumns()
        {
            #region Arrange

            var singDic = new Dictionary<string, string>();
            singDic.Add("[Usage].[Inbound or Outbound].[Inbound or Outbound]",
                "[Usage].[Inbound or Outbound].&[Inbound]");
            singDic.Add("[Usage].[Call Type].[Call Type]",
                "[Usage].[Call Type].&[MOC]");
            singDic.Add("[UsageDate].[Usage_Year].[Usage_Year]",
                "[UsageDate].[Usage_Year].&[2014]");
            singDic.Add("[UsageDate].[Usage_MonthAbbrev].[Usage_MonthAbbrev]",
                "[UsageDate].[Usage_MonthAbbrev].&[May]");
            #endregion

            #region Parse
            var pivotCellDic = new PivotCellDictionary();
            pivotCellDic.SingleSelectDictionary = singDic;

            var parser = new DaxDrillParser();
            string commandText;
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                var selectedColumns = new List<DetailColumn>();
                selectedColumns.Add(new DetailColumn() { Name = "Call Type", Expression = "Usage[Call Type]" });
                selectedColumns.Add(new DetailColumn() { Name = "Call Type Description", Expression = "Usage[Call Type Description]" });
                commandText = DaxDrillParser.BuildQueryText(tabular, pivotCellDic, "Gross Billed Sum", 99999, selectedColumns);
                tabular.Disconnect();
            }
            #endregion

            #region Assert
            Console.WriteLine(commandText);

            #endregion
        }

        [Test]
        public void ParsePivotText()
        {
            #region Parse
            var columnCommandText = DaxDrillParser.GetColumnFromPivotField("[Usage].[Inbound or Outbound].[Inbound or Outbound]");
            Assert.AreEqual("Inbound or Outbound", columnCommandText);

            var tableCommandText = DaxDrillParser.GetTableFromPivotField("[Usage].[Inbound or Outbound].[Inbound or Outbound]");
            Assert.AreEqual("Usage", tableCommandText);

            var pivotItemCommandText = DaxDrillParser.GetValueFromPivotItem("[Usage].[Inbound or Outbound].&[Inbound]");
            Assert.AreEqual("Inbound", pivotItemCommandText);

            #endregion
        }

        [Test]
        public void GetMeasure()
        {
            string measureName = "Gross Billed Sum";
            Measure measure = null;
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                measure = tabular.GetMeasure(measureName);
                tabular.Disconnect();
            }
            Console.WriteLine("Measure = {0}, Table = {1}", measure.Name, measure.Table.Name);
        }

        [Test]
        public void ParseXml()
        {
            #region Arrange

            const string nsString = "dg2ntt.daxdrill";
            var xmlString =
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<table id=""Usage"" connection_id=""localhost Roaming Model"" xmlns=""{0}"">
	<columns>
	   <column>
		  <name>Call Type</name>
		  <expression>Usage[Call Type]</expression>
	   </column>
	   <column>
		  <name>Call Type Description</name>
		  <expression>Usage[Call Type Description]</expression>
	   </column>
	   <column>
		  <name>Gross Billed</name>
		  <expression>Usage[Gross Billed]</expression>
	   </column>
	</columns>
</table>"
            .Replace("{0}", nsString);

            #endregion

            #region Act

            List<DetailColumn> columns = DaxDrillConfig.GetColumnsFromTableXml(
                "localhost Roaming Model", "Usage", xmlString, nsString);

            #endregion

            #region Assert
            foreach (var column in columns)
            {
                Console.WriteLine(column.Name + " = " + column.Expression);
            }

            Assert.AreEqual(3, columns.Count);
            Assert.AreEqual("Usage[Call Type]", columns.Where(x => x.Name == "Call Type").Select(x => x.Expression).FirstOrDefault());
            Assert.AreEqual("Usage[Call Type Description]", columns.Where(x => x.Name == "Call Type Description").Select(x => x.Expression).FirstOrDefault());
            Assert.AreEqual("Usage[Gross Billed]", columns.Where(x => x.Name == "Gross Billed").Select(x => x.Expression).FirstOrDefault());
            #endregion
        }

        public void XmlTest()
        {
            #region Arrange

            const string nsString = "dg2ntt.daxdrill";
            var xmlString =
@"<table id=""Usage"" connection_id=""localhost Roaming Model"" xmlns=""{0}"">
	<columns>
	   <column>
		  <name>Call Type</name>
		  <expression>Usage[Call Type]</expression>
	   </column>
	   <column>
		  <name>Call Type Description</name>
		  <expression>Usage[Call Type Description]</expression>
	   </column>
	   <column>
		  <name>Gross Billed</name>
		  <expression>Usage[Gross Billed]</expression>
	   </column>
	</columns>
</table>"
            .Replace("{0}", nsString);

            #endregion

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            XmlNode node = root.SelectSingleNode("..", nsmgr);

            Console.WriteLine(node.InnerXml);
        }
    }
}
