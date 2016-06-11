using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DG2NTT.DaxDrill.Tests
{
    public class ExcelTests
    {
        public void StartExcelApp()
        {
            var xlApp = new Excel.Application();
            xlApp.Visible = true;
        }
    }
}
