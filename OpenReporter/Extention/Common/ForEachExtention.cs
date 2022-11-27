using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class ForEachExtention
    {
        public static IEnumerable<TItem> ForEach<TItem>(this IEnumerable<TItem> Items, Action<TItem> ForAction)
        {
            foreach (var Item in Items)
                ForAction.Invoke(Item);
            return Items;
        }
    }
}
