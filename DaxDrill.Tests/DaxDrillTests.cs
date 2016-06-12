using DG2NTT.DaxDrill.Helpers;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;

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
            var parser = new DaxDrillParser();
            string commandText;
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                commandText = parser.BuildFilterCommandText(excelDic, tabular);
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
            var parser = new DaxDrillParser();
            string commandText;
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                commandText = parser.BuildQueryText(tabular, excelDic, "Gross Billed Sum");
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
            var parser = new DaxDrillParser();
            string commandText;
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                var selectedColumns = new List<SelectedColumn>();
                selectedColumns.Add(new SelectedColumn() { Name = "Call Type", Expression = "Usage[Call Type]" });
                selectedColumns.Add(new SelectedColumn() { Name = "Call Type Description", Expression = "Usage[Call Type Description]" });
                commandText = parser.BuildQueryText(tabular, excelDic, "Gross Billed Sum", selectedColumns);
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
            var parser = new DaxDrillParser();

            var columnCommandText = parser.GetColumnFromPivotField("[Usage].[Inbound or Outbound].[Inbound or Outbound]");
            Assert.AreEqual("Inbound or Outbound", columnCommandText);

            var tableCommandText = parser.GetTableFromPivotField("[Usage].[Inbound or Outbound].[Inbound or Outbound]");
            Assert.AreEqual("Usage", tableCommandText);

            var pivotItemCommandText = parser.GetValueFromPivotItem("[Usage].[Inbound or Outbound].&[Inbound]");
            Assert.AreEqual("Inbound", pivotItemCommandText);

            #endregion
        }

        [Test]
        public void GetMeasure()
        {
            string measureName = "Gross Billed Sum";
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {
                tabular.Connect();
                var measure = tabular.GetMeasure(measureName);
                Console.WriteLine("Measure = {0}, Table = {1}", measure.Name, measure.Table.Name);
                tabular.Disconnect();
            }
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

            List<SelectedColumn> columns = DaxDrillConfig.GetColumnsFromTableXml(
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
    }
}
