using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.DaxHelpers
{
    public class TableMdxParser
    {
        public TableMdxParser(string mdxString)
        {
            this.mdxString = mdxString;
        }
        private string mdxString;

        public string[] ConvertColumnMdxToArray()
        {
            int currentIndex = 0;
            var itemList = new List<string>();
            while (currentIndex >= 0)
            {
                var itemArray = ConvertColumnMdxToArray(ref currentIndex);

                for (int i = 0; i < itemArray.Length; i++)
                    itemList.Add(itemArray[i]);
            }
            return itemList.ToArray();
        }

        public string[] ConvertColumnMdxToArray(ref int currentIndex)
        {
            string mdx = mdxString;
            const string startPattern = "FROM (SELECT (";

            // start reading from the end of the pattern
            int beginIndex = mdx.IndexOf(startPattern, currentIndex);
            if (beginIndex < 0)
            {
                currentIndex = -1;
                return new string[0];
            }
            beginIndex += startPattern.Length;

            mdx = mdx.Substring(beginIndex, mdx.Length - beginIndex);
            currentIndex = beginIndex;

            // stop reading after the next occurrence of ") ON"
            int columnEndIndex = mdx.IndexOf(") ON COLUMNS");
            int endIndex = columnEndIndex;

            mdx = mdx.Substring(0, endIndex);
            currentIndex = currentIndex + endIndex;

            // remove the outer character "{" and "}"
            mdx = mdx.Replace("{", "").Replace("}", "");

            string[] itemStringArray = mdx.Split(',');

            return itemStringArray;
        }

        public string[] ConvertRowMdxToArray()
        {
            int currentIndex = 0;
            var itemList = new List<string>();
            while (currentIndex >= 0)
            {
                var itemArray = ConvertRowMdxToArray(ref currentIndex);

                for (int i = 0; i < itemArray.Length; i++)
                    itemList.Add(itemArray[i]);
            }
            return itemList.ToArray();
        }
        public string[] ConvertRowMdxToArray(ref int currentIndex)
        {
            string mdx = mdxString;
            // start reading from the end of the pattern
            int beginIndex = mdx.IndexOf("FROM (SELECT (", currentIndex);
            if (beginIndex < 0)
            {
                currentIndex = -1;
                return new string[0];
            }
            beginIndex = mdx.IndexOf(") ON COLUMNS", beginIndex);
            int endIndex = mdx.IndexOf(") ON ROWS", beginIndex);
            currentIndex = endIndex;
            if (endIndex < 0)
            {
                currentIndex = -1;
                return new string[0];
            }

            mdx = mdx.Substring(beginIndex, endIndex - beginIndex);
            endIndex = mdx.Length - 1;

            // remove the outer character "{" and "}"
            beginIndex = mdx.IndexOf("({") + 1;
            mdx = mdx.Substring(beginIndex, endIndex - beginIndex);
            mdx = mdx.Replace("{", "").Replace("}", "");
            string[] itemStringArray = mdx.Split(',');

            return itemStringArray;
        }
    }
}
