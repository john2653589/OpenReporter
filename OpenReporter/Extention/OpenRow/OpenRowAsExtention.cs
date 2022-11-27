using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class OpenRowAsExtention
    {
        public static IOpenRow AsCell(this IOpenRow Row, int ColumnIndex, Action<IOpenCell> CellAction)
        {
            var Cell = Row.FindCell(ColumnIndex);
            CellAction.Invoke(Cell);
            return Row;
        }
        public static IOpenRow AsCell(this IOpenRow Row, string ColumnRef, Action<IOpenCell> CellAction)
        {
            var Cell = Row.FindCell(ColumnRef);
            CellAction.Invoke(Cell);
            return Row;
        }
    }
}
