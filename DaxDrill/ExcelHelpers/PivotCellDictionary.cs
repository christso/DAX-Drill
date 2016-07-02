using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class PivotCellDictionary
    {
        private Dictionary<string, string> singleSelectDictionary;
        private Dictionary<string, List<string>> multiSelectDictionary;

        public Dictionary<string, string> SingleSelectDictionary
        {
            get
            {
                if (this.singleSelectDictionary == null)
                    this.singleSelectDictionary = new Dictionary<string, string>();
                return this.singleSelectDictionary;
            }
            set
            {
                this.singleSelectDictionary = value;
            }
        }

        public Dictionary<string, List<string>> MultiSelectDictionary
        {
            get
            {
                if (this.multiSelectDictionary == null)
                    this.multiSelectDictionary = new Dictionary<string, List<string>>();
                return this.multiSelectDictionary;
            }
            set
            {
                this.multiSelectDictionary = value;
            }
        }

        public void AddMultiSelectItem(string key, string value)
        {
            List<string> selectList = null;
            if (!MultiSelectDictionary.TryGetValue(key, out selectList))
            {
                selectList = new List<string>();
                MultiSelectDictionary[key] = selectList;
            }
            selectList.Add(value);
            return;
        }
    }
}
