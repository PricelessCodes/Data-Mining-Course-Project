using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Association_Rules_Algorithm
{
    class ItemSet
    {
        public string ID;
        public List<string> Items;

        public ItemSet()
        {
            ID = null;
            Items = null;
        }

        public ItemSet(string ID, List<string> Items)
        {
            this.ID = ID;
            this.Items = new List<string>();
            this.Items = Items;
        }
    }
}
