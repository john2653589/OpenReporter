using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Model
{
    public class CellPosition
    {
        public LockModeType LockMode => CheckLockMode();
        public string CellRef { get; private set; }
        public string ColumnRef { get; private set; }
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }

        public CellPosition() { }
        public CellPosition(string _CellRef)
        {
            Set(_CellRef);
        }
        public CellPosition(int _RowIndex, string _ColRef)
        {
            Set(_RowIndex, _ColRef);
        }
        public CellPosition(int _RowIndex, int _ColumnIndex)
        {
            Set(_RowIndex, _ColumnIndex);
        }

        #region Static Method
        public static string ColumnIndexToRef(int ColIndex)
        {
            var RetRef = "";
            while (ColIndex > 0)
            {
                var CharIdx = (ColIndex - 1) % 26;
                var RefChar = Convert.ToChar('A' + CharIdx);
                RetRef = RefChar + RetRef;
                ColIndex = (ColIndex - CharIdx) / 26;
            }
            return RetRef;
        }
        public static int RefToColumnIndex(string ColRef)
        {
            var Ref = Regex.Replace(ColRef, "[0-9]", "");
            var EnglishIdx = Ref.PadLeft(3).Select(ItemChar => "ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf(ItemChar));
            var ColumnIndex = EnglishIdx.Aggregate(0, (Current, Index) =>
            {
                return (Current * 26) + Index + 1;
            });
            return ColumnIndex;
        }
        public static string GetColumnRef(string CellRef)
        {
            var ColRef = Regex.Replace(CellRef, "[^A-Z]", "");
            return ColRef;
        }
        public static int GetRowIndex(string CellRef)
        {
            var RowIndex = int.Parse(Regex.Replace(CellRef, "[A-Z]", ""));
            return RowIndex;
        }
        public static int GetColumnIndex(string CellRef)
        {
            var ColRef = GetColumnRef(CellRef);
            var ColumnIndex = RefToColumnIndex(ColRef);
            return ColumnIndex;
        }
        #endregion

        #region Create Method
        public static CellPosition Create(string _CellRef) => new(_CellRef);
        public static CellPosition Create(int _RowIndex, string _ColRef) => new(_RowIndex, _ColRef);
        public static CellPosition Create(int _RowIndex, int _ColumnIndex) => new(_RowIndex, _ColumnIndex);
        #endregion

        #region Set Method
        public CellPosition Set(string _CellRef)
        {
            CellRef = _CellRef.ToUpper();
            TryInit(InitFromType.CellRef);
            return this;
        }
        public CellPosition Set(int _RowIndex, string _ColumnRef)
        {
            RowIndex = _RowIndex;
            ColumnRef = _ColumnRef.ToUpper();

            TryInit(InitFromType.RowIndexAndColumnRef);
            return this;
        }
        public CellPosition Set(int _RowIndex, int _ColumnIndex)
        {
            RowIndex = _RowIndex;
            ColumnIndex = _ColumnIndex;

            TryInit(InitFromType.RowIndexAndColumnIndex);
            return this;
        }

        public CellPosition SetWith_ColumnIndex(int _RowIndex)
        {
            Set(_RowIndex, ColumnIndex);
            return this;
        }
        public CellPosition SetWith_ColumnRef(int _RowIndex)
        {
            Set(_RowIndex, ColumnRef);
            return this;
        }
        public CellPosition SetWith_RowIndex(string _ColumnRef)
        {
            Set(RowIndex, _ColumnRef);
            return this;
        }
        public CellPosition SetWith_RowIndex(int _ColumnIndex)
        {
            Set(RowIndex, _ColumnIndex);
            return this;
        }

        public CellPosition SetClear_RowIndex(int _RowIndex)
        {
            Clear();
            RowIndex = _RowIndex;
            return this;
        }
        public CellPosition SetClear_ColumnRef(string _ColumnRef)
        {
            Clear();
            ColumnRef = _ColumnRef;
            return this;
        }
        public CellPosition SetClear_ColumnIndex(int _ColumnIndex)
        {
            Clear();
            ColumnIndex = _ColumnIndex;
            return this;
        }

        public CellPosition TrySet_RowIndex(int _RowIndex)
        {
            if (LockMode == LockModeType.Row)
                SetClear_RowIndex(_RowIndex);
            else
                Set(RowIndex, ColumnIndex);

            return this;
        }
        #endregion

        #region Count Method
        public CellPosition AddRow(int AddCount = 1)
        {
            TrySet_RowIndex(RowIndex + AddCount);
            return this;
        }
        #endregion

        #region Process Method
        private void Clear()
        {
            CellRef = null;
            ColumnRef = null;
            RowIndex = 0;
            ColumnIndex = 0;
        }
        private void TryInit(InitFromType InitFrom)
        {
            switch (InitFrom)
            {
                case InitFromType.CellRef:
                    ColumnRef = GetColumnRef(CellRef);
                    RowIndex = GetRowIndex(CellRef);
                    ColumnIndex = GetColumnIndex(CellRef);
                    break;
                case InitFromType.RowIndexAndColumnRef:
                    CellRef = $"{ColumnRef}{RowIndex}";
                    ColumnIndex = RefToColumnIndex(ColumnRef);
                    break;
                case InitFromType.RowIndexAndColumnIndex:
                    ColumnRef = ColumnIndexToRef(ColumnIndex);
                    CellRef = $"{ColumnRef}{RowIndex}";
                    break;
            }
        }
        private LockModeType CheckLockMode()
        {
            var CheckList = new[]
            {
                CellRef is not null,
                ColumnRef is not null,
                RowIndex > 0,
                ColumnIndex > 0,
            };
            var CheckCount = CheckList.Count(Item => Item);
            var ModeType = CheckCount == 4 ? LockModeType.Cell : LockModeType.Row;
            return ModeType;
        }
        #endregion

        enum InitFromType
        {
            CellRef,
            RowIndexAndColumnRef,
            RowIndexAndColumnIndex,
        }
    }
    public enum LockModeType
    {
        Row,
        Cell
    }
}
