extern alias AnalysisServices2014;

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
using DG2NTT.DaxDrill.Tabular;
using SSAS12 = AnalysisServices2014::Microsoft.AnalysisServices;

namespace DG2NTT.DaxDrill.Tests
{
    [TestFixture]
    public class ParseXmlTests
    {
        [Test]
        public void ParseXmlFromTable()
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

            List<DetailColumn> columns = DaxDrillConfig.GetColumnsFromTableXml(nsString, xmlString, "localhost Roaming Model", "Usage");

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

        [Test]
        public void ParseXmlFromRoot()
        {
            #region Arrange

            const string nsString = "http://schemas.microsoft.com/daxdrill";
            var xmlString =
@"<daxdrill xmlns=""http://schemas.microsoft.com/daxdrill"">
	<table id=""Usage"" connection_id=""localhost Roaming Model"" xmlns=""http://schemas.microsoft.com/daxdrill"">
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
	</table>
	<table id=""RoamingMeasure"" connection_id=""localhost Roaming Model"" xmlns=""http://schemas.microsoft.com/daxdrill"">
		<query>Usage</query>
	</table>
</daxdrill>";

            #endregion

            #region Act

            List<DetailColumn> columns = DaxDrillConfig.GetColumnsFromTableXml(nsString, xmlString, "localhost Roaming Model", "Usage");

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

        [Test]
        public void XmlTest()
        {
            #region Arrange

            const string nsString = "http://schemas.microsoft.com/daxdrill";
            var xmlString =
@"<daxdrill xmlns=""http://schemas.microsoft.com/daxdrill"">
	<table id=""Usage"" connection_id=""localhost Roaming Model"" xmlns=""http://schemas.microsoft.com/daxdrill"">
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
	</table>
	<table id=""RoamingMeasure"" connection_id=""localhost Roaming Model"" xmlns=""http://schemas.microsoft.com/daxdrill"">
		<query>FILTER
(
	UNION (
		SELECTCOLUMNS (
			DiscRelease,
			""Roaming Measure"", ""Discount Release"",
			""Accrual Period"", DiscRelease[Accrual Period],
			""TAP Code"", DiscRelease[TAP Code],
			""Amount Aud"", DiscRelease[Net Release Aud]
		),
		SELECTCOLUMNS (
			Usage,
			""Roaming Measure"", ""Discount Accrual"",
			""Accrual Period"", Usage[Usage Date],
			""TAP Code"", Usage[Their PMN TADIG Code],
			""Amount Aud"", Usage[Gross Billed]		
		)
	),
	[Roaming Measure] = VALUES ( RoamingMeasure[Roaming Measure] )
)</query>
	</table>
</daxdrill>";

            #endregion

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            //string xpath = "..";
            string xpath = string.Format("/x:daxdrill/x:table[@id='{0}']/x:query", "RoamingMeasure");

            XmlNode node = root.SelectSingleNode(xpath, nsmgr);

            Console.WriteLine(node.InnerXml);
        }

    }
}
