using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class OpenSheetFindExtention
    {
        public static IOpenRow FindRows(this IOpenSheet Sheet, int RowIndex)
        {
            var Row = Sheet.Rows.FirstOrDefault(Item => Item.AbsRowIndex == RowIndex);
            return Row;
        }
        public static IOpenCell FindCell(this IOpenSheet Sheet, int RowIndex, int ColumnIndex)
        {
            var Position = CellPosition.Create(RowIndex, ColumnIndex);
            var Cell = Sheet
               .FindRows(Position.RowIndex).Cells
               .FirstOrDefault(Item => Item.ColumnIndex == ColumnIndex);
            return Cell;
        }
        public static IOpenCell FindCell(this IOpenSheet Sheet, string CellRef)
        {
            var Position = CellPosition.Create(CellRef);
            var Cell = Sheet.FindCell(Position.RowIndex, Position.ColumnIndex);
            return Cell;
        }
        public static IOpenCell FindCell(this IOpenSheet Sheet, int RowIndex, string ColumnRef)
        {
            var Position = CellPosition.Create(RowIndex, ColumnRef);
            var Cell = Sheet.FindCell(Position.RowIndex, Position.ColumnIndex);
            return Cell;
        }
    }
}
