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
        public string AssignFileName { get; set; }
        public IOpenReport ReadFile(string _AssignFileName);
        public IOpenSheet AsSheet(string SheetName) => FindSheet(SheetName);
        public IOpenSheet this[string SheetName] => FindSheet(SheetName);
        public IOpenReport Save();
        public IOpenReport SaveAs(string SaveFullFileName);
        public IOpenReport SaveAs(string ExportFileName, string ExportPath);
        public IOpenReport SaveAsClose(string SaveFullFileName);
        public IOpenReport SaveAsClose(string ExportFileName, string ExportPath);
        public byte[] SaveAsAndReadClose(string SaveFullFileName);
        public byte[] SaveAsAndReadClose(string ExportFileName, string ExportPath);
        public byte[] SaveAsAndReadDelete(string SaveFullFileName);
        public byte[] SaveAsAndReadDelete(string ExportFileName, string ExportPath);
        public IOpenReport Close();
        public byte[] ReadXmlByte();
        public byte[] ReadXmlByteClose();
        internal IOpenSheet FindSheet(string SheetName);
    }
    public interface IOpenSheet
    {
        public IOpenReport Report { get; }
        public IEnumerable<IOpenRow> Rows => SelectRows();
        public IOpenCell this[string _CellRef] => this.FindCell(_CellRef);
        public IOpenCell this[int _RowIndex, string _ColumnRef] => this.FindCell(_RowIndex, _ColumnRef);
        public IOpenCell this[int _RowIndex, int _ColumnIndex] => this.FindCell(_RowIndex, _ColumnIndex);
        public CellPosition ForPosition { get; set; }
        public int CurrentRowIndex => ForPosition?.RowIndex ?? 1;
        public IOpenSheet InsertRowAfterFrom(int ToRowIndex = -1, int FromRowIndex = -1);
        public IOpenSheet InsertRowAfterFromClear(int ToRowIndex = -1, int FromRowIndex = -1);
        internal CellPosition InitPosition()
        {
            ForPosition ??= new CellPosition();
            return ForPosition;
        }
        public IOpenSheet PositionRow(int RowIndex)
        {
            InitPosition().TrySet_RowIndex(RowIndex);
            return this;
        }
        public IOpenRow CurrentRow()
        {
            var CurrentRow = this.FindRows(CurrentRowIndex);
            return CurrentRow;
        }
        public IOpenRow NextRow()
        {
            ForPosition.AddRow();
            var NextRow = CurrentRow();
            return NextRow;
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
        #endregion

        #region Method
        public object GetValue();
        public TValue GetValue<TValue>() => ConvertValue<TValue>(Value);
        public TValue ConvertValue<TValue>(object Value);
        public void SetValue(object Value);
        #endregion
    }
}