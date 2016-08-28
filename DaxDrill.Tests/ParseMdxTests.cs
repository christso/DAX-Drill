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
using System.Diagnostics;

namespace DG2NTT.DaxDrill.Tests
{
    [TestFixture]
    public class ParseMdxTests
    {
        [Test]
        public void ParsePivotText()
        {
            #region Parse
            var columnCommandText = DaxDrillParser.GetColumnFromPivotField("[Usage].[Inbound or Outbound].[Inbound or Outbound]");
            Assert.AreEqual("Inbound or Outbound", columnCommandText);

            var tableCommandText = DaxDrillParser.GetTableFromPivotField("[Usage].[Inbound or Outbound].[Inbound or Outbound]");
            Assert.AreEqual("Usage", tableCommandText);

            var pivotItemCommandText1 = DaxDrillParser.GetValueFromPivotItem("[Usage].[Inbound or Outbound].&[Inbound]");
            Assert.AreEqual("Inbound", pivotItemCommandText1);

            var pivotItemCommandText2 = DaxDrillParser.GetValueFromPivotItem("[Usage].[Inbound or Outbound].[Inbound]");
            Assert.AreEqual("Inbound", pivotItemCommandText2);

            #endregion
        }


        // Multiple items selected in row and column fields
        [Test]
        public void ParseTableMdx1()
        {
            string mdxString = @"SELECT NON EMPTY Hierarchize({DrilldownLevel({[UsageDate].[Usage_MonthAbbrev].[All]},,,INCLUDE_CALC_MEMBERS)}) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[Usage].[Inbound or Outbound].[All],[Usage].[Inbound or Outbound].[Inbound or Outbound].AllMembers}, {([Usage].[Call Type].[All])}), [Usage].[Inbound or Outbound].[Inbound or Outbound].AllMembers, [Usage].[Call Type])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON ROWS  FROM (SELECT ({[UsageDate].[Usage_Year].&[2010],[UsageDate].[Usage_Year].&[2011],[UsageDate].[Usage_Year].&[2012],[UsageDate].[Usage_Year].&[2013],[UsageDate].[Usage_Year].&[2014],[UsageDate].[Usage_Year].&[2015],[UsageDate].[Usage_Year].&[2016],[UsageDate].[Usage_Year].&[2017],[UsageDate].[Usage_Year].&[2018],[UsageDate].[Usage_Year].&[2019],[UsageDate].[Usage_Year].&[2020]},{[Usage].[Country].&[Algeria],[Usage].[Country].&[American samoa]}) ON COLUMNS  FROM [Model]) WHERE ([Measures].[Gross Billed Sum]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS";
            var daxFilters = DaxDrillParser.ConvertPivotTableMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            var yearValue = daxDic["[UsageDate].[Usage_Year]"];
            var yearValueStr = @"[UsageDate].[Usage_Year].&[2010]
[UsageDate].[Usage_Year].&[2011]
[UsageDate].[Usage_Year].&[2012]
[UsageDate].[Usage_Year].&[2013]
[UsageDate].[Usage_Year].&[2014]
[UsageDate].[Usage_Year].&[2015]
[UsageDate].[Usage_Year].&[2016]
[UsageDate].[Usage_Year].&[2017]
[UsageDate].[Usage_Year].&[2018]
[UsageDate].[Usage_Year].&[2019]
[UsageDate].[Usage_Year].&[2020]";
            Assert.AreEqual(yearValueStr, string.Join("\r\n", yearValue.Select(x => x.MDX)));

            var countryValue = daxDic["[Usage].[Country]"];
            var countryValueString = @"[Usage].[Country].&[Algeria]
[Usage].[Country].&[American samoa]";
            Assert.AreEqual(countryValueString, string.Join("\r\n", countryValue.Select(x => x.MDX)));

            #region User Friendly Result
            foreach (var daxFilter in daxFilters)
            {
                Console.WriteLine("Column={0} ; Table={1} ; Value={2}", daxFilter.ColumnName,
                    daxFilter.TableName, daxFilter.Value);
            }

            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value)
                {

                    Console.WriteLine(value.Value);
                }
            }
            #endregion
        }

        // multiple items selected in row field
        [Test]
        public void ParseTableMdx2()
        {
            string mdxString = @"SELECT NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[UsageDate].[Usage_Year].[All],[UsageDate].[Usage_Year].[Usage_Year].AllMembers}, {([UsageDate].[Usage_MonthAbbrev].[All])}), [UsageDate].[Usage_Year].[Usage_Year].AllMembers, [UsageDate].[Usage_MonthAbbrev])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[Usage].[Inbound or Outbound].[All],[Usage].[Inbound or Outbound].[Inbound or Outbound].AllMembers}, {([Usage].[Call Type].[All])}), [Usage].[Inbound or Outbound].[Inbound or Outbound].AllMembers, [Usage].[Call Type])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON ROWS  FROM (SELECT ({[Usage].[Call Type].&[MOC], [Usage].[Call Type].&[GPRS]}) ON COLUMNS  FROM [Model]) WHERE ([Usage].[Country].[All],[Measures].[Gross Billed Sum]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS";
            var daxFilters = DaxDrillParser.ConvertPivotTableMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            var callTypeValue = daxDic["[Usage].[Call Type]"].Select(x => x.MDX);
            string actual = string.Join("\r\n", callTypeValue);
            string expected = @"[Usage].[Call Type].&[MOC]
[Usage].[Call Type].&[GPRS]";
            Assert.AreEqual(expected, actual);

            #region User Friendly Result
            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value)
                {
                    Console.WriteLine(value);
                }
            }
            #endregion
        }

        // value containing "(" and ")" characters
        [Test]
        public void ParseTableMdx3()
        {
            string mdxString = @"SELECT NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[AsAtDate].[Year].[Year].AllMembers}, {([AsAtDate].[Month].[All])}), [AsAtDate].[Year].[Year].AllMembers, [AsAtDate].[Month])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[Activity].[Fin Activity].[All],[Activity].[Fin Activity].[Fin Activity].AllMembers}, {([ApTradeCreditor].[Supplier].[All])}), {[Activity].[Fin Activity].&[Other Capex]}, [ApTradeCreditor].[Supplier])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON ROWS  FROM (SELECT ({[ApTradeCreditor].[Supplier].&[HUAWEI TECHNOLOGIES (AUSTRALIA) PTY LTD]}) ON COLUMNS  FROM [Model]) WHERE ([ApTradeCreditor].[LiabilityAccount].[All],[ApTradeCreditor].[InvoiceNumber].&[AU1602160],[Measures].[TradeCreditorSum]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS";
            var daxFilters = DaxDrillParser.ConvertPivotTableMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            var supplierValue = daxDic["[ApTradeCreditor].[Supplier]"].Select(x => x.MDX);
            var expected = "[ApTradeCreditor].[Supplier].&[HUAWEI TECHNOLOGIES (AUSTRALIA) PTY LTD]";
            var actual = string.Join("\r\n", supplierValue);
            Assert.AreEqual(expected, actual);

            #region User Friendly Result
            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value)
                {
                    Console.WriteLine(value);
                }
            }
            #endregion
        }

        // value containing "(" and ")" characters, and with a filter applied to both rows and columns
        [Test]
        public void ParseTableMdx4()
        {
            string mdxString = @"SELECT NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[AsAtDate].[Year].[Year].AllMembers}, {([AsAtDate].[Month].[All])}), [AsAtDate].[Year].[Year].AllMembers, [AsAtDate].[Month])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[Activity].[Fin Activity].[All],[Activity].[Fin Activity].[Fin Activity].AllMembers}, {([ApTradeCreditor].[Supplier].[All])}), {[Activity].[Fin Activity].&[Other Capex]}, [ApTradeCreditor].[Supplier])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON ROWS  FROM (SELECT ({[AsAtDate].[Month].&[Apr]}) ON COLUMNS , ({[ApTradeCreditor].[Supplier].&[HUAWEI TECHNOLOGIES (AUSTRALIA) PTY LTD]}) ON ROWS  FROM [Model]) WHERE ([ApTradeCreditor].[LiabilityAccount].[All],[ApTradeCreditor].[InvoiceNumber].&[AU1602160],[Measures].[TradeCreditorSum]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS";
            var daxFilters = DaxDrillParser.ConvertPivotTableMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            Assert.AreEqual("[AsAtDate].[Month].&[Apr]", daxDic["[AsAtDate].[Month]"][0].MDX);
            Assert.AreEqual("[ApTradeCreditor].[Supplier].&[HUAWEI TECHNOLOGIES (AUSTRALIA) PTY LTD]", daxDic["[ApTradeCreditor].[Supplier]"][0].MDX);

            #region User Friendly Result
            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value)
                {
                    Console.WriteLine(value);
                }
            }
            #endregion
        }


        [Test]
        public void ParseTableMdx5()
        {
            string mdxString = @"SELECT NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[PaymentDate].[Payment_Year].[All],[PaymentDate].[Payment_Year].[Payment_Year].AllMembers}, {([PaymentDate].[Payment_Month].[All])}), [PaymentDate].[Payment_Year].[Payment_Year].AllMembers, [PaymentDate].[Payment_Month])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY Hierarchize(DrilldownMember(DrilldownMember(CrossJoin({[Activity].[Dim_1_2].[All],[Activity].[Dim_1_2].[Dim_1_2].AllMembers}, {([Activity].[Dim_1_3].[All],[APPmtDistn].[VendorName].[All])}), [Activity].[Dim_1_2].[Dim_1_2].AllMembers, [Activity].[Dim_1_3]), {[Activity].[Dim_1_3].&[Regulatory Fees]}, [APPmtDistn].[VendorName])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON ROWS  FROM (SELECT ({[PaymentDate].[Payment_Year].&[2.016E3]}) ON COLUMNS , ({[Activity].[Dim_1_3].&[Regulatory Fees]},{[APPmtDistn].[VendorName].&[AUSTRALIAN COMMUNICATIONS AND MEDIA AUTHORITY]}) ON ROWS  FROM [Model]) WHERE ([APPmtDistn].[InvoiceNum].[All],[Account].[IsCash].&[1],[Measures].[Total Payment Amount Aud]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS";
            var daxFilters = DaxDrillParser.ConvertPivotTableMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            Assert.AreEqual("[PaymentDate].[Payment_Year].&[2.016E3]", daxDic["[PaymentDate].[Payment_Year]"][0].MDX);
            Assert.AreEqual("[Activity].[Dim_1_3].&[Regulatory Fees]", daxDic["[Activity].[Dim_1_3]"][0].MDX);
            Assert.AreEqual("[APPmtDistn].[VendorName].&[AUSTRALIAN COMMUNICATIONS AND MEDIA AUTHORITY]",
                daxDic["[APPmtDistn].[VendorName]"][0].MDX);

            #region User Friendly Result
            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value)
                {
                    Console.WriteLine(value);
                }
            }
            #endregion
        }

        // filter applied on hierarchy on fields which are also on the pivot table
        [Test]
        public void ParseHierTableMdx()
        {
            string mdxString = @"SELECT NON EMPTY Hierarchize(DrilldownMember(CrossJoin({[TranDate].[Tran_Year].[All],[TranDate].[Tran_Year].[Tran_Year].AllMembers}, {([TranDate].[Tran_MonthAbbrev].[All])}), [TranDate].[Tran_Year].[Tran_Year].AllMembers, [TranDate].[Tran_MonthAbbrev])) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY Hierarchize({DrilldownLevel({[CounterCcy].[Counter Ccy].[All]},,,INCLUDE_CALC_MEMBERS)}) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME ON ROWS  FROM (SELECT ({[TranDate].[Tran_Year].&[2.016E3]},{[TranDate].[Tran_MonthAbbrev].&[Jun], [TranDate].[Tran_MonthAbbrev].&[May], [TranDate].[Tran_MonthAbbrev].&[Apr], [TranDate].[Tran_MonthAbbrev].&[Mar], [TranDate].[Tran_MonthAbbrev].&[Feb], [TranDate].[Tran_MonthAbbrev].&[Jan]}) ON COLUMNS  FROM (SELECT ({[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Jul],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Mar],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[May],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Nov],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Oct],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Sep],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.016E3]}) ON COLUMNS  FROM [Model])) WHERE ([Measures].[Trades CCA Sum]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS";
            var daxFilters = DaxDrillParser.ConvertPivotTableMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            #region User Friendly Result
            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value.Select(x => x.MDX))
                {
                    Console.WriteLine(value);
                }
            }
            #endregion
        }

        [Test]
        public void ParseHierPivotFilter()
        {
            string value = "[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Mar]";
            var daxFilter = DaxDrillParser.CreateDaxFilterFromColumn(value);
        }

        [Test]
        public void ParseHierCellMdx()
        {
            string mdxString = @" ([CounterCcy].[Counter Ccy].&[EUR], [Measures].[Trades CCA Sum], [TranDate].[Tran_MonthAbbrev].&[Apr], [TranDate].[Tran_Year].&[2.016E3], [TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.016E3],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Jul],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Mar],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[May],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Nov],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Oct],[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Sep])";

            var daxFilters = DaxDrillParser.ConvertPivotCellMdxToDaxFilterList(mdxString);
            var daxDic = DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilters);

            #region User Friendly Result - Dictionary
            foreach (var daxFilter in daxFilters)
            {
                Console.WriteLine("Column={0} ; Table={1} ; Value={2}", daxFilter.ColumnName,
                    daxFilter.TableName, daxFilter.Value);
            }

            foreach (var pair in daxDic)
            {
                Console.WriteLine(pair.Key + " ---------");
                foreach (var value in pair.Value.Select(x => x.MDX))
                {
                    Console.WriteLine(value);
                }
            }
            #endregion

            #region Assert
            // columns
            Assert.AreEqual("[CounterCcy].[Counter Ccy].&[EUR]", daxDic["[CounterCcy].[Counter Ccy]"][0].MDX);
            Assert.AreEqual("[TranDate].[Tran_MonthAbbrev].&[Apr]", daxDic["[TranDate].[Tran_MonthAbbrev]"][0].MDX);
            Assert.AreEqual("[TranDate].[Tran_Year].&[2.016E3]", daxDic["[TranDate].[Tran_Year]"][0].MDX);

            // hierarchies
            Assert.AreEqual("[TranDate].[Tran_YearMonthDay].&[2.016E3]", daxDic["[TranDate].[Tran_YearMonthDay]"][0].MDX);
            //Assert.AreEqual("[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Jul]", daxDic["[TranDate].[Tran_YearMonthDay]"][1].MDX);
            //Assert.AreEqual("[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Mar]", daxDic["[TranDate].[Tran_YearMonthDay]"][2].MDX);
            #endregion

        }

        [Test]
        public void ParseMdxWithHierarchy()
        {
            string pivotFieldNamesString = @"[Measures].[Trades CCA Sum]
[TranDate].[Tran_Year].[Tran_Year]
[TranDate].[Tran_YearMonthDay].[Tran_Year]
[TranDate].[Tran_YearMonthDay].[Tran_MonthAbbrev]
[TranDate].[Tran_YearMonthDay].[Tran_DayOfMonth]
[CounterCcy].[Counter Ccy].[Counter Ccy]
[TranDate].[Tran_MonthAbbrev].[Tran_MonthAbbrev]";
            string[] pivotFieldNames = pivotFieldNamesString.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string pivotItemValue = "[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Jul]";

            var creator = new DaxFilterCreator(pivotItemValue, pivotFieldNames);
            var daxFilter = creator.CreateDaxFilter();

            Assert.AreEqual("Tran_Year", daxFilter.ColumnNameHierarchy[0].ColumnName);
            Assert.AreEqual("Tran_MonthAbbrev", daxFilter.ColumnNameHierarchy[1].ColumnName);
            Assert.AreEqual("2.014E3", daxFilter.ValueHierarchy[0]);
            Assert.AreEqual("Jul", daxFilter.ValueHierarchy[1]);
        }

        [Test]
        public void PivotFieldIsHierarchy()
        {
            string pivotFieldValue = "[TranDate].[Tran_MonthAbbrev].[Tran_MonthAbbrev]";
            bool isHierarchy = DaxFilterCreator.PivotFieldIsHierarchy(pivotFieldValue);
            Assert.IsFalse(isHierarchy);
        }

        [Test]
        public void PivotFieldIsNotHierarchy()
        {
            string pivotFieldValue = "[TranDate].[Tran_YearMonthDay].[Tran_MonthAbbrev]";
            bool isHierarchy = DaxFilterCreator.PivotFieldIsHierarchy(pivotFieldValue);
            Assert.IsTrue(isHierarchy);
        }

        [Test]
        public void ParsePivotItemValue()
        {
            var itemString = "[TranDate].[Tran_YearMonthDay].[Tran_Year].&[2.014E3].&[Jul]";
            var daxFilter = DaxDrillParser.CreateDaxFilterFromHierarchy(itemString, null);

            Assert.AreEqual(true, daxFilter.IsHierarchy);
            Assert.AreEqual("Tran_YearMonthDay", daxFilter.ColumnName);
            Assert.AreEqual("TranDate", daxFilter.TableName);
            Assert.AreEqual("2.014E3", daxFilter.ValueHierarchy[0]);
            Assert.AreEqual("Jul", daxFilter.ValueHierarchy[1]);
        }


    }
}
