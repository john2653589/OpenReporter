using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class OpenRowFindExtention
    {
        public static IOpenCell FindCell(this IOpenRow Row, int ColumnIndex)
        {
            var Cell = Row.Cells
               .FirstOrDefault(Item => Item.ColumnIndex == ColumnIndex);

            return Cell;
        }
        public static IOpenCell FindCell(this IOpenRow Row, string ColumnRef)
        {
            var Cell = Row.Cells
                .FirstOrDefault(Item => Item.ColumnRef == ColumnRef);

            return Cell;
        }

    }
}
