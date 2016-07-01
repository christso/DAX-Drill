using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.Tests
{
    class Program
    {
        private static ExcelTests tests = new ExcelTests();
        static void Main(string[] args)
        {
            tests.TemplateTest();
            Console.ReadKey();
        }
    }
}
