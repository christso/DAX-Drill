using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaxDrill.Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            var tests = new ParseXmlTests();
            tests.MeasureRefXmlTest();
            Console.WriteLine("Test passed");
            Console.ReadKey();
        }
    }
}
