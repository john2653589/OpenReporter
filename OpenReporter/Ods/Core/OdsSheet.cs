using Rugal.Net.OpenReporter.Extention;
using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using Rugal.Net.OpenReporter.Ods.Extention;
using System.Xml;

namespace Rugal.Net.OpenReporter.Ods.Core
{
    public class OdsSheet : IOpenSheet
    {
        #region Interface Property
        public CellPosition ForPosition { get; set; }
        public IOpenReport Report => OdsReport;
        #endregion

        #region Ods Property
        public OdsReport OdsReport { get; }
        public XmlNode SheetNode { get; }
        public XmlNamespaceManager NamespaceManager => OdsReport.NamespaceManager;
        public IEnumerable<XmlNode> RowNodes => SheetNode.NodeList_Row(NamespaceManager).Select();
        #endregion

        public OdsSheet(OdsReport _OdsReport, XmlNode _SheetNode)
        {
            OdsReport = _OdsReport;
            SheetNode = _SheetNode;
        }

        #region Interface Method
        public IEnumerable<IOpenRow> SelectRows()
        {
            var GetRows = new List<IOpenRow>();
            var Index = 1;
            foreach (var Item in RowNodes)
            {
                var Row = new OdsRow(Index, Item, this);
                if (Item.IsRepeatedRow(out var Count))
                    Index += Count;
                else
                {
                    GetRows.Add(Row);
                    Index++;
                }
            }
            return GetRows;
        }

        public IOpenSheet InsertRowAfterFrom(int ToRowIndex = -1, int FromRowIndex = -1)
        {
            if (ToRowIndex == -1)
                ToRowIndex = ForPosition.RowIndex;

            if (FromRowIndex == -1)
                FromRowIndex = ToRowIndex;

            var Row = SelectRows()
                .First(Item => Item.AbsRowIndex == FromRowIndex) as OdsRow;

            var CloneRow = Row.RowNode.CloneNode(true);
            var NewRow = new OdsRow(ToRowIndex, CloneRow, this);

            SheetNode.InsertAfter(NewRow.RowNode, Row.RowNode);
            return this;
        }

        public IOpenSheet InsertRowAfterFromClear(int ToRowIndex = -1, int FromRowIndex = -1)
        {
            if(ToRowIndex == -1)
                ToRowIndex = ForPosition.RowIndex;

            if (FromRowIndex == -1)
                FromRowIndex = ToRowIndex;

            var Row = SelectRows()
                .First(Item => Item.AbsRowIndex == FromRowIndex) as OdsRow;

            var CloneRow = Row.RowNode.CloneNode(true);
            var NewRow = new OdsRow(ToRowIndex, CloneRow, this);

            foreach (var Cell in NewRow.CellNodes)
                Cell.ClearValue();

            SheetNode.InsertAfter(NewRow.RowNode, Row.RowNode);
            return this;
        }


        #endregion

    }
}