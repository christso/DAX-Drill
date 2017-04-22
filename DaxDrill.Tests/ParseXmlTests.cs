extern alias AnalysisServices2014;

using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;
using DaxDrill.DaxHelpers;
using Microsoft.AnalysisServices.Tabular;
using DaxDrill.ExcelHelpers;
using DaxDrill.Tabular;
using SSAS12 = AnalysisServices2014::Microsoft.AnalysisServices;

namespace DaxDrill.Tests
{
    [TestFixture]
    public class ParseXmlTests
    {
        [Test]
        public void ParseXmlFromTable()
        {
            #region Arrange

            const string nsString = "daxdrill";
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


        [Test]
        public void XmlTest2()
        {
            #region Arrange

            const string nsString = "http://schemas.microsoft.com/daxdrill";
            var xmlString =
@"<daxdrill xmlns=""http://schemas.microsoft.com/daxdrill"">
	<measure id=""Counter Amt Current UL Sum"" connection_id=""CashFlow"" xmlns=""http://schemas.microsoft.com/daxdrill"">
		<query>
SELECTCOLUMNS (
	FILTER (
		CALCULATETABLE (
			CashFlow,
			ALL ( Snapshot )
		),
		SWITCH ( TRUE(),
			HASONEVALUE ( SnapshotReported[Compare Name] ),
			
			CONTAINS ( ALL(SnapshotReported),
			SnapshotReported[Compare Name], ""Current"",
			SnapshotReported[Snapshot OID], CashFlow[Snapshot OID],
			SnapshotReported[Snapshot Type], VALUES ( SnapshotReported[Snapshot Type] ) ),
			
			CONTAINS ( ALL(SnapshotReported),
			SnapshotReported[Compare Name], ""Current"",
			SnapshotReported[Snapshot OID], CashFlow[Snapshot OID])
		)
	),
	""Snapshot Type"", ""Current"",
	""Cash Flow OID"", CashFlow[Cash Flow OID],
	""Status"", CashFlow[Status],			
	""Snapshot"", RELATED(Snapshot[Snapshot]),
	""Tran Date"", CashFlow[Tran Date],
	""Account"", RELATED(Account[Account]),
	""Activity"", RELATED(Activity[Activity]),
	""Counterparty"", RELATED(Counterparty[Counterparty]),
	""Account Ccy Amt"", CashFlow[Account Ccy Amt],
	""Description"", CashFlow[Description],
	""Source"", CashFlow[Source],
	""Functional Ccy Amt"", CashFlow[Functional Ccy Amt],
	""Counter Ccy Amt"", CashFlow[Counter Ccy Amt],
	""Counter Ccy"", CashFlow[Counter Ccy],		
	""Fix"", CashFlow[Fix],		
	""Is Reclass"", CashFlow[Is Reclass],		
	""Fix Activity"", RELATED(FixActivity[Fix Activity]),
	""Activity L1"", RELATED(Activity[Activity L1]),
	""Activity L2"", RELATED(Activity[Activity L2]),
	""Activity L3"", RELATED(Activity[Activity L3]),
	""Activity L4"", RELATED(Activity[Activity L4]),
	""Snapshot OID"", CashFlow[Snapshot OID]
)		
		</query>
	</measure>
	<table id=""CashFlow"" connection_id=""CashFlow"" xmlns=""http://schemas.microsoft.com/daxdrill"">
		<query>
FILTER
(
	UNION
	(
		/* Current Snapshot */
		SELECTCOLUMNS (
			FILTER (
				CALCULATETABLE (
					CashFlow,
					ALL ( Snapshot )
				),
				SWITCH ( TRUE(),
					HASONEVALUE ( SnapshotReported[Compare Name] ),
					
					CONTAINS ( ALL(SnapshotReported),
					SnapshotReported[Compare Name], ""Current"",
					SnapshotReported[Snapshot OID], CashFlow[Snapshot OID],
					SnapshotReported[Snapshot Type], VALUES ( SnapshotReported[Snapshot Type] ) ),
					
					CONTAINS ( ALL(SnapshotReported),
					SnapshotReported[Compare Name], ""Current"",
					SnapshotReported[Snapshot OID], CashFlow[Snapshot OID])
				)
			),
			""Snapshot Type"", ""Current"",
			""Cash Flow OID"", CashFlow[Cash Flow OID],
			""Status"", CashFlow[Status],			
			""Snapshot"", RELATED(Snapshot[Snapshot]),
			""Tran Date"", CashFlow[Tran Date],
			""Account"", RELATED(Account[Account]),
			""Activity"", RELATED(Activity[Activity]),
			""Counterparty"", RELATED(Counterparty[Counterparty]),
			""Account Ccy Amt"", CashFlow[Account Ccy Amt],
			""Description"", CashFlow[Description],
			""Source"", CashFlow[Source],
			""Functional Ccy Amt"", CashFlow[Functional Ccy Amt],
			""Counter Ccy Amt"", CashFlow[Counter Ccy Amt],
			""Counter Ccy"", CashFlow[Counter Ccy],		
			""Fix"", CashFlow[Fix],		
			""Is Reclass"", CashFlow[Is Reclass],		
			""Fix Activity"", RELATED(FixActivity[Fix Activity]),
			""Activity L1"", RELATED(Activity[Activity L1]),
			""Activity L2"", RELATED(Activity[Activity L2]),
			""Activity L3"", RELATED(Activity[Activity L3]),
			""Activity L4"", RELATED(Activity[Activity L4]),
			""Snapshot OID"", CashFlow[Snapshot OID]
		),
		/* Previous Snapshot */
		SELECTCOLUMNS (
			FILTER (
				CALCULATETABLE (
					CashFlow,
					ALL ( Snapshot )
				),
				IF ( HASONEVALUE ( SnapshotReported[Compare Name] ),
					/* Current amounts to be prepended to Previous snapshot */
					CashFlow[Tran Date] &lt; CALCULATE (	
						MIN ( SnapshotReported[First Date] ),
						FILTER (
							ALLEXCEPT ( SnapshotReported, SnapshotReported[Compare Name] ),
							SnapshotReported[Snapshot Type] = ""Previous""
						)
					)
					&amp;&amp; CONTAINS ( ALL(SnapshotReported),
					SnapshotReported[Compare Name], ""Current"",
					SnapshotReported[Snapshot OID], CashFlow[Snapshot OID]) 
					/* Amounts in active snapshot */
					|| CashFlow[Snapshot OID] = 
					CALCULATE (
						VALUES ( SnapshotReported[Snapshot OID] ),
						SnapshotReported[Snapshot Type] = ""Previous"",
						SnapshotReported[Compare Name] = VALUES ( SnapshotReported[Compare Name] )
					)
					&amp;&amp; CONTAINS ( ALL(SnapshotReported),
					SnapshotReported[Compare Name], VALUES ( SnapshotReported[Compare Name] ),
					SnapshotReported[Snapshot OID], CashFlow[Snapshot OID]),
					TRUE()
				)
			),
			""Snapshot Type"", ""Previous"",
			""Cash Flow OID"", CashFlow[Cash Flow OID],
			""Status"", CashFlow[Status],			
			""Snapshot"", RELATED(Snapshot[Snapshot]),
			""Tran Date"", CashFlow[Tran Date],
			""Account"", RELATED(Account[Account]),
			""Activity"", RELATED(Activity[Activity]),
			""Counterparty"", RELATED(Counterparty[Counterparty]),
			""Account Ccy Amt"", CashFlow[Account Ccy Amt],
			""Description"", CashFlow[Description],
			""Source"", CashFlow[Source],
			""Functional Ccy Amt"", CashFlow[Functional Ccy Amt],
			""Counter Ccy Amt"", CashFlow[Counter Ccy Amt],
			""Counter Ccy"", CashFlow[Counter Ccy],		
			""Fix"", CashFlow[Fix],		
			""Is Reclass"", CashFlow[Is Reclass],		
			""Fix Activity"", RELATED(FixActivity[Fix Activity]),
			""Activity L1"", RELATED(Activity[Activity L1]),
			""Activity L2"", RELATED(Activity[Activity L2]),
			""Activity L3"", RELATED(Activity[Activity L3]),
			""Activity L4"", RELATED(Activity[Activity L4]),
			""Snapshot OID"", CashFlow[Snapshot OID]
		)
	),
	IF ( VALUES ( SnapshotReported[Snapshot Type] ) = ""Var"",
		TRUE(),
		[Snapshot Type] = VALUES ( SnapshotReported[Snapshot Type] )
	) 
	/* This works if snapshot is selected, but doesn't if SnapshotReported is selected */
	/*
	&amp;&amp; IF(ISFILTERED(SnapshotReported[Compare Name]) &amp;&amp; HASONEVALUE(SnapshotReported[Snapshot OID]), 
		[Snapshot OID] = VALUES ( SnapshotReported[Snapshot OID] ),
		TRUE()	
	)
	*/
)
		</query>
	</table>
	<table id=""BankStmt"" connection_id=""BankStmt"" xmlns=""http://schemas.microsoft.com/daxdrill"">
		<query>
	SELECTCOLUMNS (
		BankStmt,	
		""BankStmt OID"", BankStmt[BankStmt OID],
		""Tran Date"", BankStmt[Tran Date],
		""Account"", RELATED(Account[Account]),
		""Activity"", RELATED(Activity[Activity]),
		""Counterparty"", RELATED(Counterparty[Counterparty]),
		""Tran Amount"", BankStmt[Tran Amount],
		""Tran Description"", BankStmt[Tran Description],
		""Functional Ccy Amt"", BankStmt[Functional Ccy Amt],
		""Counter Ccy Amt"", BankStmt[Counter Ccy Amt],
		""Counter Ccy"", BankStmt[Counter Ccy]
	)
		</query>
	</table>	
</daxdrill>";

            #endregion

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            //string xpath = "..";
            string xpath = string.Format("/x:daxdrill/x:measure[@id='{0}']/x:query", "Counter Amt Current UL Sum");

            XmlNode node = root.SelectSingleNode(xpath, nsmgr);

            Console.WriteLine(node.InnerXml);
        }

        // Test logic for retrieving query in referenced measure
        [Test]
        public void MeasureRefXmlTest()
        {
            #region Arrange

            const string nsString = "http://schemas.microsoft.com/daxdrill";
            var xmlString =
@"<daxdrill xmlns=""http://schemas.microsoft.com/daxdrill"">
	<measure id=""Counter Amt Current UL Sum"" xmlns=""http://schemas.microsoft.com/daxdrill"">
        <query>CALCULATE ( CashFlow )</query>
	</measure>
	<measure id=""Func Amt Current UL Sum"" ref=""Counter Amt Current UL Sum"" xmlns=""http://schemas.microsoft.com/daxdrill"">
	</measure>
</daxdrill>";

            #endregion

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);
            XmlNode root = doc.DocumentElement;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", nsString);

            // find measure
            string xpath = string.Format("/x:daxdrill/x:measure[@id='{0}']", "Func Amt Current UL Sum");

            // find referenced measure and get Xml Path
            XmlNode node = root.SelectSingleNode(xpath, nsmgr);
            XmlAttribute refMeasure = null;
            if (node != null)
                refMeasure = node.Attributes["ref"];
            if (refMeasure != null)
                xpath = string.Format("/x:daxdrill/x:measure[@id='{0}']/x:query", refMeasure.Value);

            // return node for referenced measure
            node = root.SelectSingleNode(xpath, nsmgr);

            Console.WriteLine("--- MEASURE " + refMeasure.Value + " ---");
            Console.WriteLine(node.InnerXml);
        }
    }
}
