using Rugal.Net.OpenReporter.Extention;
using Rugal.Net.OpenReporter.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Rugal.Net.OpenReporter.Interface
{
    /// <summary>
    /// OpenReporter
    /// Version 1.0.0
    /// From Rugal Tu
    /// </summary>
    public interface IOpenReport
    {
        public IOpenReport ReadFile(string FullFileName);
        public IOpenSheet AsSheet(string SheetName) => FindSheet(SheetName);
        public IOpenSheet this[string SheetName] => FindSheet(SheetName);
        internal IOpenSheet FindSheet(string SheetName);

        public IOpenReport Save();

    }

    public interface IOpenSheet
    {
        public CellPosition ForPosition { get; set; }
        public int CurrentRowIndex => ForPosition?.RowIndex ?? 1;
        public IOpenReport Reporter { get; }
        public IEnumerable<IOpenRow> Rows => SelectRows();
        public IOpenSheet InsertRowFrom(int ToRowIndex, int FromRowIndex = -1);
        public IOpenSheet InsertRowFromClear(int ToRowIndex, int FromRowIndex = -1);
        public IOpenCell this[string _CellRef] => this.FindCell(_CellRef);
        public IOpenCell this[int _RowIndex, int _ColumnIndex] => this.FindCell(_RowIndex, _ColumnIndex);
        public IOpenCell this[int _RowIndex, string _ColumnRef] => this.FindCell(_RowIndex, _ColumnRef);
        internal CellPosition InitPosition()
        {
            ForPosition ??= new CellPosition();
            return ForPosition;
        }
        internal IEnumerable<IOpenRow> SelectRows();
    }

    public interface IOpenRow
    {
        public IOpenSheet Sheet { get; }
        public int AbsRowIndex { get; set; }
        public IEnumerable<IOpenCell> Cells => SelectCells();
        public IOpenCell this[string ColumnRef] => this.FindCell(ColumnRef);
        public IOpenCell this[int ColumnIndex] => this.FindCell(ColumnIndex);
        internal IEnumerable<IOpenCell> SelectCells();
    }

    public interface IOpenCell
    {
        #region Property
        public IOpenRow Row { get; }
        public CellPosition Position { get; set; }
        public string CellRef => Position.CellRef;
        public string ColumnRef => Position.ColumnRef;
        public int RowIndex => Position.RowIndex;
        public int ColumnIndex => Position.ColumnIndex;
        public object Value => GetValue();
        public TValue GetValue<TValue>() => ConvertValue<TValue>(Value);
        #endregion

        #region Method
        public object GetValue();
        public TValue ConvertValue<TValue>(object Value);
        public void SetValue(object Value);

        #endregion
    }
}