using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace DG2NTT.DaxDrill.Tests
{
    public class ExcelTests
    {
        public void TemplateTest()
        {
            var xlApp = new Excel.Application();
            xlApp.Visible = true;
            xlApp.Quit();
            CleanUp();
        }
        private static void CleanUp()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
