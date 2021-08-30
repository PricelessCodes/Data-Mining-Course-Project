using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Association_Rules_Algorithm
{
    public class SetOfItems
    {
        public int Count;
        public List<string> Items;

        public SetOfItems()
        {
            Count = 0;
            Items = new List<string>();
        }

        public SetOfItems(int Count, List<string> Items)
        {
            this.Count = Count;
            this.Items = new List<string>();
            this.Items = Items;
        }

        public void Add(string Item)
        {
            this.Items.Add(Item);
        }
    }
}
