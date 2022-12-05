using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Rugal.OpenExcel.Core
{
    public class OpenExcelReader
    {
        public OpenExcelModel Excel { get; set; }
        public Worksheet Sheet { get; set; }
        public bool IsReadAll { get; set; }
        public MergeCells MergeCells => Sheet.Descendants<MergeCells>().FirstOrDefault();
        public IEnumerable<MergeCell> MergeCellList => MergeCells.Elements<MergeCell>();
        public SheetData SheetData => Sheet.Elements<SheetData>().FirstOrDefault();
        public IEnumerable<Row> Rows => SheetData.Elements<Row>();


        public (int Row, int Col) StartIdx { get; set; }
        public (int Row, int Col) EndIdx { get; set; }
        public List<List<Cell>> DataRange { get; set; }
        public int RowCount => DataRange.Count;
        public int ColCount => DataRange.Max(Item => Item.Count);
        public List<Cell> GetCells()
        {
            return DataRange.SelectMany(Item => Item).ToList();
        }

        #region Lambda Function
        /// <summary>
        /// 從讀取範圍內取出相對位置Row
        /// </summary>
        /// <param name="Row"></param>
        /// <returns></returns>
        public List<Cell> this[int Row] => DataRange[Row];
        /// <summary>
        /// 取出絕對位置Cell
        /// </summary>
        /// <param name="CellRef"></param>
        /// <returns></returns>
        public Cell this[string CellRef] => GetCellFromRef(CellRef);
        /// <summary>
        /// 從讀取範圍裡取出相對位置Cell
        /// </summary>
        /// <param name="Row">相對Row</param>
        /// <param name="Ref">絕對Ref</param>
        /// <returns></returns>
        public Cell this[int Row, string Ref] => GetCellFromArea(Row, Ref);
        /// <summary>
        /// 絕對位置，使用數字定位
        /// </summary>
        /// <param name="Row"></param>
        /// <param name="Col"></param>
        /// <returns></returns>
        public Cell this[int Row, int Col] => GetCellFromRef($"{Col.ToColumnRef()}{Row}", false);

        /// <summary>
        /// 讀取絕對位置Value
        /// </summary>
        /// <param name="CellRef"></param>
        /// <returns></returns>
        public string Value(string CellRef) => this[CellRef].Value(Excel);
        public string Value(int Row, string Ref) => this[Row, Ref].Value(Excel);
        public string Value_Abs(int Row, int Col) => GetAbsCell(Row, Col)?.Value(Excel);

        public DateTime ValueTime(string CellRef) => this[CellRef].ValueTime();
        public DateTime ValueTime(int Row, string Ref) => this[Row, Ref].ValueTime();
        #endregion
        public OpenExcelReader(OpenExcelModel _Excel, Worksheet _Sheet)
        {
            Excel = _Excel;
            Sheet = _Sheet;
            IsReadAll = true;
            RangeInit();

        }
        public OpenExcelReader(OpenExcelModel _Excel, Worksheet _Sheet, (int Row, int Col) _EndIdx, (int Row, int Col) _StartIdx = default) : this(_Excel, _Sheet)
        {
            _StartIdx = _StartIdx == default ? (1, 1) : _StartIdx;
            IsReadAll = false;
            SetRange(_EndIdx, _StartIdx);
        }
        public void SetRange((int Row, int Col) _EndIdx, (int Row, int Col) _StartIdx)
        {
            IsReadAll = false;
            EndIdx = _EndIdx;
            StartIdx = _StartIdx;
            RangeInit();
        }
        public void UpdateRange() => RangeInit();
        /// <summary>
        /// 絕對位置
        /// </summary>
        /// <param name="ToRow"></param>
        /// <param name="FromRow"></param>
        /// <param name="AddRowNum"></param>
        public void AddRowFrom(int ToRow, int FromRow, int AddRowNum)
        {
            for (var i = 0; i < AddRowNum; i++)
            {
                CopyRowFrom(ToRow, FromRow);
                ToRow++;
            }
            UpdateRange();
        }
        public void CopyRowFrom(int ToRow, int FromRow, bool IsCopyMerge = true)
        {
            var FromRowData = Rows.FirstOrDefault(Item => Item.RowIndex.Value == FromRow);
            var NewRow = FromRowData.CloneNode(true) as Row;
            NewRow.RemoveAllChildren();
            NewRow.RowIndex = new UInt32Value((uint)ToRow);

            var FromCells = FromRowData.Elements<Cell>().ToList();
            foreach (var Item in FromCells)
            {
                var OrgCellRef = Item.CellReference;
                var OrgRef = OrgCellRef.GetCellRef();
                var NewCell = Item.CloneNode(true) as Cell;
                NewCell.CellReference = new StringValue($"{OrgRef}{ToRow}");
                NewRow.Append(NewCell);
            }
            var LastRowIdx = Rows.Where(Item => Item.RowIndex <= ToRow).Max(Item => Item.RowIndex);
            var LastRow = Rows.FirstOrDefault(Item => Item.RowIndex == LastRowIdx);
            LastRow.InsertAfterSelf(NewRow);
            if (LastRow.RowIndex == NewRow.RowIndex)
                LastRow.Remove();

            if (IsCopyMerge)
                CopyRowMergeFrom(ToRow, FromRow);
        }
        public void CopyRowMergeFrom(int ToRow, int FromRow)
        {
            var FromMerges = MergeCellList.Where((Item) =>
            {
                var (From, To) = Item.Reference.GetMergeCellRow();
                return From == FromRow && To == FromRow;
            }).ToList();

            var NewMergeCells = FromMerges.Select((Item) =>
            {
                var OrgRef = Item.Reference;
                var (From, To) = OrgRef.GetMergeCellRef();
                var NewMerge = new MergeCell()
                {
                    Reference = new StringValue($"{From}{ToRow}:{To}{ToRow}"),
                };
                return NewMerge;
            }).ToList();

            var ExistData = MergeCellList.Where((Item) =>
            {
                var (From, To) = Item.Reference.GetMergeCellRow();
                return (ToRow == From || ToRow == To);
            }).ToList();

            foreach (var Item in ExistData)
                Item.Remove();

            foreach (var Item in NewMergeCells)
                MergeCells.Append(Item);
        }
        public List<Row> CopyRows(int StartRow, int EndRow) => Sheet.CopyRows(StartRow, EndRow);
        public List<MergeCell> CopyMeger(int StartRow, int EndRow) => Sheet.CopyMeger(StartRow, EndRow);
        private void RangeInit()
        {
            if (IsReadAll)
            {
                DataRange = Rows.Select(Item => Item.Elements<Cell>().ToList()).ToList();
                return;
            }
            var StartRow = StartIdx.Row;
            var EndRow = EndIdx.Row;

            DataRange = Rows
                .Where(Item => Item.RowIndex.Value >= StartRow && Item.RowIndex.Value <= EndRow)
                .Select(Item => Item.Elements<Cell>().ToList()).ToList();
        }
        public Cell GetCellFromRef(string CellRef, bool IsAddNew = true)
        {
            var GetCell = GetCells().FirstOrDefault(Item => Item.CellReference.Value == CellRef);
            if (GetCell is null && IsAddNew)
            {
                GetCell = new Cell()
                {
                    CellReference = new StringValue(CellRef),
                };
                var RowIdx = CellRef.GetCellRow();
                var Row = Rows.FirstOrDefault(Item => Item.RowIndex == RowIdx);
                Row?.Append(GetCell);
            }
            return GetCell;
        }
        public Cell GetAbsCell(int RowIdx, int ColIdx)
        {
            var FindRef = $"{ColIdx.ToColumnRef()}{RowIdx}";
            var GetCell = GetCells().FirstOrDefault(Item => Item.CellReference.Value == FindRef);
            return GetCell;
        }
        public Cell GetCellFromArea(int AreaRowIdx, string CellRef)
        {
            var CellList = DataRange[AreaRowIdx];
            var GetCell = CellList
                .FirstOrDefault(Item => Item.CellReference.GetCellRef() == CellRef);

            var RowIdx = CellList.First().CellReference.GetCellRow();
            if (GetCell is null)
            {
                GetCell = new Cell()
                {
                    CellReference = new StringValue($"{CellRef}{RowIdx}"),
                };
                var AddRow = Rows.FirstOrDefault(Item => Item.RowIndex == RowIdx);

                var CellColIdx = GetCell.CellReference.GetColumnIdx();
                var FindRow = AddRow.Elements<Cell>();
                var FindIdx = 0;
                foreach (var Item in FindRow)
                {
                    if (Item.CellReference.GetColumnIdx() > CellColIdx)
                        break;
                    FindIdx++;
                }
                AddRow.InsertAt(GetCell, FindIdx);
            }
            return GetCell;
        }




    }
}
