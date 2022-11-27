using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Ods.Extention;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Rugal.Net.OpenReporter.Ods.Core
{
    public class OdsRow : IOpenRow
    {
        #region Interface Property
        public IOpenSheet Sheet => OdsSheet;
        public int AbsRowIndex { get; set; }
        #endregion

        #region Ods Property
        public OdsSheet OdsSheet { get; }
        public XmlNode RowNode { get; }
        public XmlNamespaceManager NamespaceManager => OdsSheet.NamespaceManager;
        public IEnumerable<XmlNode> CellNodes => RowNode.NodeList_Cell(NamespaceManager).Select();
        #endregion
        public OdsRow(int _AbsRowIndex, XmlNode _RowNode, OdsSheet _OdsSheet)
        {
            AbsRowIndex = _AbsRowIndex;
            RowNode = _RowNode;
            OdsSheet = _OdsSheet;
        }

        #region Interface Method
        public IEnumerable<IOpenCell> SelectCells()
        {
            var GetCells = new List<IOpenCell>();

            var Index = 1;
            foreach (var Item in CellNodes)
            {
                if (Item.IsRepeatedCell(out var Count))
                {
                    Index += Count;
                    continue;
                }

                var Row = new OdsCell(Index, Item, this);
                GetCells.Add(Row);
                if (Item.IsSpannedCell(out Count))
                    Index += Count;
                else
                    Index++;
            }

            return GetCells;
        }
        #endregion
    }
}
