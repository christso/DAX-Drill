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
    public class UtilTests
    {
        [Test]
        public void SetPivotFieldPage()
        {
            var pageName = DaxDrillParser.CreatePivotFieldPageName("[PrdDate].[Prd_MonthAbbrev].[Prd_MonthAbbrev]", "May");
            Assert.AreEqual("[PrdDate].[Prd_MonthAbbrev].&[May]", pageName);
        }
    }
}
