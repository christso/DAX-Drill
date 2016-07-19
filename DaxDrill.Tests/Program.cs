using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.Tests
{
    class Program
    {
        private static DaxDrillTests tests = new DaxDrillTests();
        static void Main(string[] args)
        {
            tests.SetPivotFieldPage();
            Console.ReadKey();
        }
    }
}
