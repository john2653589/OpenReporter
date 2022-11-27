using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class OpenSheetAsExtention
    {
        public static IOpenSheet AsRows(this IOpenSheet Sheet, int RowIndex, Action<IOpenRow> RowAction)
        {
            var Row = Sheet.Rows.FirstOrDefault(Item => Item.AbsRowIndex == RowIndex);
            RowAction.Invoke(Row);
            return Sheet;
        }
        public static IOpenSheet AsCell(this IOpenSheet Sheet, int RowIndex, int ColumnIndex, Action<IOpenCell> CellAction)
        {
            var Position = CellPosition.Create(RowIndex, ColumnIndex);
            var Cell = Sheet
               .FindRows(Position.RowIndex).Cells
               .FirstOrDefault(Item => Item.ColumnIndex == ColumnIndex);
            CellAction.Invoke(Cell);
            return Sheet;
        }
        public static IOpenSheet AsCell(this IOpenSheet Sheet, string CellRef, Action<IOpenCell> CellAction)
        {
            var Position = CellPosition.Create(CellRef);
            Sheet.AsCell(Position.RowIndex, Position.ColumnIndex, CellAction);
            return Sheet;
        }
        public static IOpenSheet AsCell(this IOpenSheet Sheet, int RowIndex, string ColumnRef, Action<IOpenCell> CellAction)
        {
            var Position = CellPosition.Create(RowIndex, ColumnRef);
            Sheet.AsCell(Position.RowIndex, Position.ColumnIndex, CellAction);
            return Sheet;
        }
    }
}
