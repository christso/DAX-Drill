using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            var tests = new ParseMdxTests();
            tests.AddMultiplePageFieldFilterToDic();
            Console.WriteLine("Test passed");
            Console.ReadKey();
        }
    }
}
