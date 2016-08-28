using DG2NTT.DaxDrill.DaxHelpers;
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
        private DaxFilterCollection multiSelectDictionary;

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

        public DaxFilterCollection MultiSelectDictionary
        {
            get
            {
                if (this.multiSelectDictionary == null)
                    this.multiSelectDictionary = new DaxFilterCollection();
                return this.multiSelectDictionary;
            }
            set
            {
                this.multiSelectDictionary = value;
            }
        }

        public void AddMultiSelectItem(string key, DaxFilter value)
        {
            List<DaxFilter> selectList = null;
            if (!MultiSelectDictionary.TryGetValue(key, out selectList))
            {
                selectList = new List<DaxFilter>();
                MultiSelectDictionary[key] = selectList;
            }
            selectList.Add(value);
            return;
        }
    }
}
