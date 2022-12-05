using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Rugal.OpenExcel.Core;
using System.Text.RegularExpressions;

namespace Rugal.OpenExcel.Core
{
    public static class OpenExcelExtention
    {
        #region Reader Convert
        public static OpenExcelReader Reader(this Worksheet Sheet, OpenExcelModel Excel)
        {
            var Ret = new OpenExcelReader(Excel, Sheet);
            return Ret;
        }
        public static OpenExcelReader Reader(this Worksheet Sheet, OpenExcelModel Excel, (int Row, int Col) EndIdx, (int Row, int Col) StartIdx = default)
        {
            var Ret = new OpenExcelReader(Excel, Sheet, EndIdx, StartIdx);
            return Ret;
        }
        #endregion

        #region Value Process
        public static string Value(this Cell Cell, OpenExcelModel Excel)
        {
            if (Cell.Count() == 0)
                return null;

            var StringTable = Excel.WorkbookPart.SharedStringTablePart.SharedStringTable;
            var Value = Cell?.CellValue?.InnerText;
            if (Cell?.DataType != null && Cell.DataType == CellValues.SharedString)
                if (int.TryParse(Value, out int IntValue))
                    Value = StringTable.ChildElements[IntValue].InnerText;

            return Value;
        }
        public static DateTime ValueTime(this Cell Cell)
        {
            if (Cell.Count() == 0)
                return default;

            DateTime Ret = default;
            var Value = Cell?.CellValue?.InnerText;
            if (int.TryParse(Value, out int IntValue))
                Ret = DateTime.FromOADate(IntValue);

            return Ret;
        }
        public static void SetValue(this Cell Cell, string Value)
        {
            Cell.DataType = new EnumValue<CellValues>(CellValues.String);
            Cell.CellValue = new CellValue(Value);
        }
        public static void SetValue(this Cell Cell, DateTime Value)
        {
            Cell.DataType = new EnumValue<CellValues>(CellValues.Date);
            Cell.CellValue = new CellValue(Value);
        }
        public static void SetValue(this Cell Cell, double Value)
        {
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            Cell.CellValue = new CellValue(Value);
        }
        public static void SetValue(this Cell Cell, int Value)
        {
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            Cell.CellValue = new CellValue(Value);
        }
        public static void SetCalc(this Cell Cell, string CalcString)
        {
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            Cell.CellFormula = new CellFormula(CalcString);
        }
        #endregion

        #region Row、Cell Convert
        public static (int From, int To) GetMergeCellRow(this StringValue CellLocate)
        {
            if (CellLocate is null)
                return (-1, -1);

            var Arr = CellLocate.Value.Split(':');
            var Ret = (Arr[0].GetCellRow(), Arr[1].GetCellRow());
            return Ret;
        }
        public static (string From, string To) GetMergeCellRef(this StringValue CellRef)
        {
            var RefArr = CellRef.Value.Split(':');
            var From = RefArr[0].GetCellRef();
            var To = RefArr[1].GetCellRef();
            return (From, To);
        }
        public static int GetCellRow(this StringValue CellRef) => GetCellRow(CellRef.Value);
        public static int GetCellRow(this string CellRef)
        {
            var GetInt = Regex.Replace(CellRef, "[a-z]", "", RegexOptions.IgnoreCase);
            if (int.TryParse(GetInt, out int RowIdx))
                return RowIdx;
            return -1;
        }
        public static string GetCellRef(this StringValue CellRef) => GetCellRef(CellRef.Value);
        public static string GetCellRef(this string CellRef)
        {
            var Ref = Regex.Replace(CellRef, "[0-9]", "");
            return Ref;
        }
        public static int GetColumnIdx(this StringValue CellRef)
        {
            var Ref = Regex.Replace(CellRef.Value, "[0-9]", "");
            var EngIdx = Ref.PadLeft(3).Select(ItemChar => "ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf(ItemChar));
            return EngIdx.Aggregate(0, (Current, Index) => (Current * 26) + (Index + 1));
        }
        public static string ToColumnRef(this int CellColIdx)
        {
            var RetRef = "";
            while (CellColIdx > 0)
            {
                var CharIdx = (CellColIdx - 1) % 26;
                var RefChar = Convert.ToChar('A' + CharIdx);
                RetRef = RefChar + RetRef;
                CellColIdx = (CellColIdx - CharIdx) / 26;
            }
            return RetRef;
        }

        #endregion

        #region Copy Process
        public static List<Row> CopyRows(this Worksheet Sheet, int StartRow, int EndRow)
        {
            var SheetData = Sheet.Elements<SheetData>().FirstOrDefault();
            var Ret = SheetData.Elements<Row>()
                .Where(Item => Item.RowIndex >= StartRow && Item.RowIndex <= EndRow)
                .Select(Item => Item.CloneNode(true) as Row)
                .ToList();
            return Ret;
        }
        public static List<MergeCell> CopyMeger(this Worksheet Sheet, int StartRow, int EndRow)
        {
            var MergeCells = Sheet.Descendants<MergeCells>().FirstOrDefault();
            var MergeCellList = MergeCells.Elements<MergeCell>();
            var Ret = MergeCellList
                .Where((Item) =>
                {
                    var Ref = Item.Reference;
                    var (FromRow, ToRow) = Ref.GetMergeCellRow();
                    var RetWhere = FromRow >= StartRow && FromRow <= EndRow && ToRow >= StartRow && ToRow <= EndRow;
                    return RetWhere;
                })
                .Select(Item => Item.CloneNode(true) as MergeCell)
                .ToList();
            return Ret;
        }
        public static void CopyRowTo(this List<Row> RowList, SheetData Rows, int AddRowNum)
        {
            foreach (var ItemRow in RowList)
            {
                var OrgRowIdx = ItemRow.RowIndex.Value;
                var NewRow = ItemRow.CloneNode(true) as Row;
                NewRow.RowIndex = new UInt32Value(OrgRowIdx + (uint)AddRowNum);

                foreach (var ItemCell in NewRow.Elements<Cell>().ToList())
                {
                    var OrgCellRef = ItemCell.CellReference.Value;
                    var OrgRef = OrgCellRef.GetCellRef();
                    var OrgRow = OrgCellRef.GetCellRow();
                    var NewCell = ItemCell.CloneNode(true) as Cell;
                    NewCell.CellReference = new StringValue($"{OrgRef}{OrgRow + AddRowNum}");
                    ItemCell.InsertAfterSelf(NewCell);
                    ItemCell.Remove();
                }
                Rows.Append(NewRow);
            }
        }
        public static void CopyMegerTo(this List<MergeCell> MergeCellList, MergeCells MergeCells, int AddRowNum)
        {
            foreach (var Item in MergeCellList)
            {
                var OrgRef = Item.Reference;
                var (FromRef, ToRef) = OrgRef.GetMergeCellRef();
                var (FromRow, ToRow) = OrgRef.GetMergeCellRow();

                var NewMerge = new MergeCell()
                {
                    Reference = $"{FromRef}{FromRow + AddRowNum}:{ToRef}{ToRow + AddRowNum}",
                };
                MergeCells.Append(NewMerge);
            }
        }
        #endregion
    }
}
