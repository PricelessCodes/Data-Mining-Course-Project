using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Association_Rules_Algorithm
{
    class Item
    {
        public int Count;
        public string I;

        public Item()
        {
            Count = 0;
            I = null;
        }

        public Item(string Item, int Count)
        {
            this.Count = Count;
            this.I = Item;
        }
    }
}
